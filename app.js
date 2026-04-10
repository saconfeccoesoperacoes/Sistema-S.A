/* Base de armazenamento em memória + migração opcional do navegador.
   O sistema passa a usar o Supabase como fonte principal.
   O browserStorage abaixo existe só para importar dados antigos uma vez, se houver. */
const browserStorage = (() => {
  try { return window.localStorage; } catch (e) { return null; }
})();

const APP_STORAGE_KEYS = [
  "configERP","usuarios","usuarioAtual","pedidos","packlist","estoque","expedicao",
  "modoExpedicao","preferenciaShopee","memoriaExpedicao","produtos","cores","logs",
  "producaoDetalhada","pedidosCompra","usuariosSchemaVersion","tema"
];

const __saMemoryStore = Object.create(null);

(function preloadFromBrowserStorage(){
  if(!browserStorage) return;
  for(const key of APP_STORAGE_KEYS){
    const value = browserStorage.getItem(key);
    if(value !== null && value !== undefined){
      __saMemoryStore[key] = String(value);
    }
  }
})();

const appStorage = {
  getItem(key){
    return Object.prototype.hasOwnProperty.call(__saMemoryStore, key) ? __saMemoryStore[key] : null;
  },
  setItem(key, value){
    __saMemoryStore[key] = String(value);
  },
  removeItem(key){
    delete __saMemoryStore[key];
  },
  clear(){
    Object.keys(__saMemoryStore).forEach(key => delete __saMemoryStore[key]);
  },
  key(index){
    return Object.keys(__saMemoryStore)[index] || null;
  },
  get length(){
    return Object.keys(__saMemoryStore).length;
  }
};

window.SA_APP_STORAGE = appStorage;

let dadosFileHandle = null;

const defaults = {
  empresaNome: "Nome da empresa",
  limiteEstoqueBaixo: 3,
  produtosPadrao: ["Calça Pantalona","Blusa Manga Curta Gola V","Blusa Manga Longa","Blusa Regata","Vestido","Macacão"],
  cores: ["Preto","Branco","Verde","Vermelho","Pink"],
  tamanhos: ["P","M","G","GG"],
  usuarios: [
    {usuario:"Rodolfo", senha:"SA12345", nivel:"MASTER", origem:"padrao"}
  ]
};

let config = JSON.parse(appStorage.getItem("configERP")) || {
  empresaNome: defaults.empresaNome,
  limiteEstoqueBaixo: defaults.limiteEstoqueBaixo
};

function normalizarUsuario(user){
  if(!user) return null;
  return {
    usuario: String(user.usuario || "").trim(),
    senha: String(user.senha || "").trim(),
    nivel: String(user.nivel || user.perfil || "OPER").toUpperCase(),
    origem: user.origem || "custom"
  };
}

function usuariosPadrao(){
  return defaults.usuarios.map(normalizarUsuario);
}

function mergeUsuarios(listaSalva, listaPadrao){
  const mapa = new Map();
  (Array.isArray(listaSalva) ? listaSalva : []).map(normalizarUsuario).filter(Boolean).forEach(u=>{
    mapa.set(u.usuario.toLowerCase(), u);
  });
  (Array.isArray(listaPadrao) ? listaPadrao : []).map(normalizarUsuario).filter(Boolean).forEach(u=>{
    if(!mapa.has(u.usuario.toLowerCase())) mapa.set(u.usuario.toLowerCase(), u);
  });
  return Array.from(mapa.values());
}

function sincronizarUsuariosPadrao(){
  usuarios = mergeUsuarios(usuarios, usuariosPadrao());
  saveLocal();
}

const USERS_SCHEMA_VERSION = "v7_rodolfo_only";
let usuarios = [];
const usuariosSalvos = JSON.parse(appStorage.getItem("usuarios") || "[]");
const schemaSalvo = appStorage.getItem("usuariosSchemaVersion");
if(schemaSalvo !== USERS_SCHEMA_VERSION){
  usuarios = usuariosPadrao();
  appStorage.setItem("usuariosSchemaVersion", USERS_SCHEMA_VERSION);
  appStorage.setItem("usuarios", JSON.stringify(usuarios));
} else {
  usuarios = mergeUsuarios(Array.isArray(usuariosSalvos) ? usuariosSalvos : [], usuariosPadrao());
}
let usuarioAtual = JSON.parse(appStorage.getItem("usuarioAtual")) || null;
if(usuarioAtual){
  const encontrado = usuarios.find(u => u.usuario.toLowerCase() === String(usuarioAtual.usuario || '').toLowerCase() && u.senha === usuarioAtual.senha);
  if(!encontrado) usuarioAtual = null;
}
let pedidos = JSON.parse(appStorage.getItem("pedidos")) || [];
let packlist = JSON.parse(appStorage.getItem("packlist")) || [];
let estoque = JSON.parse(appStorage.getItem("estoque")) || {};
let expedicao = JSON.parse(appStorage.getItem("expedicao")) || [];
let modoExpedicao = appStorage.getItem("modoExpedicao") || "auto";
let preferenciaShopee = appStorage.getItem("preferenciaShopee") || "shopee_agencia";
let memoriaExpedicao = JSON.parse(appStorage.getItem("memoriaExpedicao") || "{}");
let usuariosJsonConectado = false;
let pedidoAccessGranted = false;
let usuariosFileHandle = null;
let produtos = JSON.parse(appStorage.getItem("produtos")) || defaults.produtosPadrao.slice();
let cores = JSON.parse(appStorage.getItem("cores")) || defaults.cores.slice();
let logs = JSON.parse(appStorage.getItem("logs")) || [];
let producaoDetalhada = JSON.parse(appStorage.getItem("producaoDetalhada")) || [];
let pedidosCompra = (JSON.parse(appStorage.getItem("pedidosCompra")) || []).map(normalizarPedidoCompraItem).filter(Boolean);

sincronizarUsuariosPadrao();

let tamanhos = defaults.tamanhos.slice();

const PERMISSOES = {
  MASTER: ["ALL"],
  ADMIN: ["DELETE_GENERAL"],
  OPER: []
};

function getPermissoesUsuario(user = usuarioAtual){
  if(!user) return [];
  return PERMISSOES[user.nivel] || [];
}

function isMaster(user = usuarioAtual){
  return !!user && user.nivel === "MASTER";
}

function hasPermission(perm, user = usuarioAtual){
  if(!user) return false;
  const perms = getPermissoesUsuario(user);
  return perms.includes("ALL") || perms.includes(perm);
}

// ==========================
// UTIL
// ==========================
function key(p,c,t){ return `${p}|${c}|${t}`; }

function numero(v){
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function money(v){
  return Number(v || 0).toLocaleString("pt-BR", {style:"currency", currency:"BRL"});
}


const CLOUD_TABLE = (window.SA_CLOUD_CONFIG && window.SA_CLOUD_CONFIG.TABLE_NAME) || "app_state";
let cloudClient = null;
let cloudSaveTimer = null;
let cloudIsHydrating = false;
let cloudHydrated = false;

function getCloudClient(){
  try{
    if(cloudClient) return cloudClient;
    if(!window.SA_CLOUD_CONFIG) return null;
    const url = window.SA_CLOUD_CONFIG.SUPABASE_URL;
    const key = window.SA_CLOUD_CONFIG.SUPABASE_ANON_KEY;
    if(!url || !key || !window.supabase || !window.supabase.createClient) return null;
    cloudClient = window.supabase.createClient(url, key);
    return cloudClient;
  }catch(e){
    console.error("Erro ao criar cliente Supabase:", e);
    return null;
  }
}

function getCloudWorkspaceId(){
  return (window.SA_CLOUD_CONFIG && window.SA_CLOUD_CONFIG.WORKSPACE_ID) || "sa-principal";
}

function getCloudPayload(){
  return {
    config,
    usuarios,
    usuarioAtual,
    pedidos,
    packlist,
    estoque,
    expedicao,
    modoExpedicao,
    preferenciaShopee,
    memoriaExpedicao,
    produtos,
    cores,
    logs,
    producaoDetalhada,
    pedidosCompra,
    tema: appStorage.getItem("tema") || "light",
    usuariosSchemaVersion: appStorage.getItem("usuariosSchemaVersion") || USERS_SCHEMA_VERSION
  };
}

function applyCloudPayload(payload){
  if(!payload || typeof payload !== "object") return;

  if(payload.config && typeof payload.config === "object") config = payload.config;
  if(Array.isArray(payload.usuarios)) usuarios = mergeUsuarios(payload.usuarios, usuariosPadrao());
  if(Object.prototype.hasOwnProperty.call(payload, "usuarioAtual")) usuarioAtual = payload.usuarioAtual || null;
  if(Array.isArray(payload.pedidos)) pedidos = payload.pedidos;
  if(Array.isArray(payload.packlist)) packlist = payload.packlist;
  if(payload.estoque && typeof payload.estoque === "object") estoque = payload.estoque;
  if(Array.isArray(payload.expedicao)) expedicao = payload.expedicao;
  if(typeof payload.modoExpedicao === "string") modoExpedicao = payload.modoExpedicao;
  if(typeof payload.preferenciaShopee === "string") preferenciaShopee = payload.preferenciaShopee;
  if(payload.memoriaExpedicao && typeof payload.memoriaExpedicao === "object") memoriaExpedicao = payload.memoriaExpedicao;
  if(Array.isArray(payload.produtos)) produtos = payload.produtos;
  if(Array.isArray(payload.cores)) cores = payload.cores;
  if(Array.isArray(payload.logs)) logs = payload.logs;
  if(Array.isArray(payload.producaoDetalhada)) producaoDetalhada = payload.producaoDetalhada;
  if(Array.isArray(payload.pedidosCompra)) pedidosCompra = payload.pedidosCompra;

  appStorage.setItem("configERP", JSON.stringify(config));
  appStorage.setItem("usuarios", JSON.stringify(usuarios));
  appStorage.setItem("usuarioAtual", JSON.stringify(usuarioAtual));
  appStorage.setItem("pedidos", JSON.stringify(pedidos));
  appStorage.setItem("packlist", JSON.stringify(packlist));
  appStorage.setItem("estoque", JSON.stringify(estoque));
  appStorage.setItem("expedicao", JSON.stringify(expedicao));
  appStorage.setItem("modoExpedicao", modoExpedicao || "auto");
  appStorage.setItem("preferenciaShopee", preferenciaShopee || "shopee_agencia");
  appStorage.setItem("memoriaExpedicao", JSON.stringify(memoriaExpedicao || {}));
  appStorage.setItem("produtos", JSON.stringify(produtos));
  appStorage.setItem("cores", JSON.stringify(cores));
  appStorage.setItem("logs", JSON.stringify(logs));
  appStorage.setItem("producaoDetalhada", JSON.stringify(producaoDetalhada));
  appStorage.setItem("pedidosCompra", JSON.stringify(pedidosCompra));
  appStorage.setItem("usuariosSchemaVersion", payload.usuariosSchemaVersion || USERS_SCHEMA_VERSION);

  if(payload.tema === "dark"){
    document.body.classList.add("dark");
    appStorage.setItem("tema", "dark");
  }else{
    document.body.classList.remove("dark");
    appStorage.setItem("tema", "light");
  }
}

async function loadCloudState(){
  const client = getCloudClient();
  if(!client) return false;

  cloudIsHydrating = true;
  try{
    const { data, error } = await client
      .from(CLOUD_TABLE)
      .select("payload")
      .eq("workspace_id", getCloudWorkspaceId())
      .maybeSingle();

    if(error) throw error;
    if(data && data.payload){
      applyCloudPayload(data.payload);
    }
    cloudHydrated = true;
    return true;
  }catch(err){
    console.error("Erro ao carregar da nuvem:", err);
    notificar("Falha ao carregar da nuvem. Usando dados locais.", "error", 5000);
    return false;
  }finally{
    cloudIsHydrating = false;
  }
}

async function saveCloudState(){
  if(cloudIsHydrating) return;
  const client = getCloudClient();
  if(!client) return;

  try{
    const payload = getCloudPayload();
    const { error } = await client
      .from(CLOUD_TABLE)
      .upsert({
        workspace_id: getCloudWorkspaceId(),
        payload,
        updated_at: new Date().toISOString()
      });
    if(error) throw error;
  }catch(err){
    console.error("Erro ao salvar na nuvem:", err);
  }
}

function scheduleCloudSave(){
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = setTimeout(saveCloudState, 700);
}

function saveLocal(){
  appStorage.setItem("configERP", JSON.stringify(config));
  appStorage.setItem("usuarios", JSON.stringify(usuarios));
  appStorage.setItem("usuarioAtual", JSON.stringify(usuarioAtual));
  appStorage.setItem("pedidos", JSON.stringify(pedidos));
  appStorage.setItem("packlist", JSON.stringify(packlist));
  appStorage.setItem("estoque", JSON.stringify(estoque));
  appStorage.setItem("expedicao", JSON.stringify(expedicao));
  appStorage.setItem("modoExpedicao", modoExpedicao || "auto");
  appStorage.setItem("preferenciaShopee", preferenciaShopee || "shopee_agencia");
  appStorage.setItem("memoriaExpedicao", JSON.stringify(memoriaExpedicao || {}));
  appStorage.setItem("produtos", JSON.stringify(produtos));
  appStorage.setItem("cores", JSON.stringify(cores));
  appStorage.setItem("logs", JSON.stringify(logs));
  appStorage.setItem("producaoDetalhada", JSON.stringify(producaoDetalhada));
  appStorage.setItem("pedidosCompra", JSON.stringify(pedidosCompra));
  appStorage.setItem("usuariosSchemaVersion", USERS_SCHEMA_VERSION);
  scheduleCloudSave();
}

function notificar(msg, tipo="success", duracao=3000){
  const n = document.getElementById("notificacao");
  n.innerText = msg;
  n.className = tipo === "error" ? "show error" : "show";
  setTimeout(()=>{ n.className = ""; }, duracao);
}

function adicionarLog(acao, detalhes=""){
  if(!usuarioAtual) return;
  logs.unshift({
    dataHora: new Date().toLocaleString("pt-BR"),
    usuario: usuarioAtual.usuario,
    acao,
    detalhes
  });
  saveLocal();
  renderLogs();
}

function preencherSelect(id, itens){
  const select = document.getElementById(id);
  if(!select) return;
  select.innerHTML = "";
  itens.forEach(item=>{
    const o = document.createElement("option");
    o.value = item;
    o.textContent = item;
    select.appendChild(o);
  });
}

function atualizarBoasVindas(){
  const el = document.getElementById("bemVindo");
  const topbarUsuario = document.getElementById("topbarUsuario");
  el.innerText = usuarioAtual
    ? `${config.empresaNome} | BEM VINDO, ${usuarioAtual.usuario.toUpperCase()}!`
    : `${config.empresaNome} | BEM VINDO!`;

  if(topbarUsuario){
    topbarUsuario.innerText = usuarioAtual ? usuarioAtual.usuario : "Rodolfo";
  }
}

function getMasterPassword(){
  const master = usuarios.find(u => u.usuario === "Rodolfo" && u.nivel === "MASTER") || usuarios.find(u => u.nivel === "MASTER");
  return master ? master.senha : "";
}

function abrirPedidosProtegido(){
  if(!usuarioAtual){
    notificar("Faça login para acessar o sistema.", "error");
    return;
  }
  if(isMaster() || pedidoAccessGranted){
    pedidoAccessGranted = true;
    show("pedidos");
    return;
  }
  const senha = prompt("Área protegida. Digite a senha master para acessar Pedidos:");
  if(senha && senha === getMasterPassword()){
    pedidoAccessGranted = true;
    adicionarLog("PEDIDOS_LIBERADOS", `Aba Pedidos liberada por ${usuarioAtual.usuario}`);
    show("pedidos");
  } else {
    notificar("Senha master inválida.", "error");
  }
}

function aplicarPermissoes(){
  document.querySelectorAll(".danger").forEach(btn => {
    const scope = btn.dataset.dangerScope || "delete";
    const permitido = scope === "master" ? isMaster() : (isMaster() || hasPermission("DELETE_GENERAL"));
    btn.style.display = permitido ? "inline-block" : "none";
  });

  document.querySelectorAll('[data-master-only="true"]').forEach(el => {
    el.style.display = isMaster() ? "inline-flex" : "none";
  });
}


// ==========================
// USUÁRIOS JSON
// ==========================
async function carregarUsuariosJsonAuto(){
  try {
    if(location.protocol === "http:" || location.protocol === "https:"){
      const resp = await fetch("usuarios.json?_=" + Date.now(), {cache:"no-store"});
      if(resp.ok){
        const data = await resp.json();
        const lista = Array.isArray(data) ? data : data.usuarios;
        if(Array.isArray(lista) && lista.length){
          usuarios = lista.map(normalizarUsuario).filter(Boolean);
          usuariosJsonConectado = true;
          appStorage.setItem("usuarios", JSON.stringify(usuarios));
          atualizarStatusUsuariosJson("usuarios.json carregado automaticamente.");
        }
      }
    }
  } catch(e){}
}

function atualizarStatusUsuariosJson(msg){
  const el = document.getElementById("usuariosJsonStatus");
  if(el) el.textContent = msg;
}

function conteudoUsuariosJson(){
  return JSON.stringify({
    updatedAt: new Date().toISOString(),
    usuarios: usuarios.map(u => ({usuario:u.usuario, senha:u.senha, nivel:u.nivel, origem:u.origem || "custom"}))
  }, null, 2);
}

async function conectarUsuariosJson(){
  if(!isMaster()) return notificar("Apenas Rodolfo pode conectar o usuarios.json.", "error");
  if(!window.showOpenFilePicker){
    notificar("Seu navegador não suporta conexão direta com arquivo local. Use Exportar usuarios.json.", "error");
    return;
  }
  try{
    const [handle] = await window.showOpenFilePicker({multiple:false, types:[{description:"JSON", accept:{"application/json":[".json"]}}]});
    const file = await handle.getFile();
    const text = await file.text();
    const data = JSON.parse(text || "{}");
    const lista = Array.isArray(data) ? data : data.usuarios;
    if(!Array.isArray(lista)) throw new Error("Estrutura inválida.");
    usuarios = lista.map(normalizarUsuario).filter(Boolean);
    usuariosFileHandle = handle;
    usuariosJsonConectado = true;
    saveLocal();
    renderUsuarios();
    atualizarStatusUsuariosJson(`Conectado: ${file.name}`);
    notificar("usuarios.json conectado com sucesso!");
  }catch(e){
    notificar("Conexão com usuarios.json cancelada ou inválida.", "error");
  }
}

async function persistirUsuariosJson(){
  appStorage.setItem("usuarios", JSON.stringify(usuarios));
  if(usuariosFileHandle){
    const writable = await usuariosFileHandle.createWritable();
    await writable.write(conteudoUsuariosJson());
    await writable.close();
    usuariosJsonConectado = true;
    atualizarStatusUsuariosJson("usuarios.json sincronizado com sucesso.");
    return true;
  }
  atualizarStatusUsuariosJson(location.protocol.startsWith("http") ? "Use Exportar usuarios.json para atualizar o arquivo físico." : "HTML local: conecte o usuarios.json para sincronizar automaticamente.");
  return false;
}

function exportarUsuariosJson(){
  const blob = new Blob([conteudoUsuariosJson()], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "usuarios.json";
  a.click();
  URL.revokeObjectURL(url);
  notificar("usuarios.json exportado!");
}

// ==========================
// LOGIN / TEMA
// ==========================
function login() {
  const u = document.getElementById("loginUsuario").value.trim();
  const s = document.getElementById("loginSenha").value.trim();
  const user = usuarios.find(x => x.usuario.toLowerCase() === u.toLowerCase() && x.senha === s);

  if(!user){
    notificar("Usuário ou senha inválidos!", "error");
    return;
  }

  usuarioAtual = user;
  saveLocal();
  document.getElementById("loginTela").style.display = "none";
  atualizarBoasVindas();
  aplicarPermissoes();
  init();
  resetTimer();
  adicionarLog("LOGIN", "Usuário logou no sistema");
  notificar(`Bem-vindo, ${user.usuario}!`);
}

function logout(){
  usuarioAtual = null;
  pedidoAccessGranted = false;
  saveLocal();
  document.getElementById("loginTela").style.display = "flex";
  atualizarBoasVindas();
}

document.getElementById("btnLogout").onclick = ()=>{
  logout();
  notificar("Sessão encerrada!", "error");
};

document.getElementById("btnTema").onclick = ()=>{
  document.body.classList.toggle("dark");
  appStorage.setItem("tema", document.body.classList.contains("dark") ? "dark" : "light");
  scheduleCloudSave();
};


function showConfigTab(tab){
  document.querySelectorAll('.config-tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.configTab === tab);
  });
  document.querySelectorAll('.config-panel').forEach(panel => {
    panel.classList.toggle('active', panel.id === `configTab-${tab}`);
  });
}


function showProdTab(tab){
  document.querySelectorAll('.prod-tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.prodTab === tab);
  });
  document.querySelectorAll('.prod-panel').forEach(panel => {
    panel.classList.toggle('active', panel.id === `prodTab-${tab}`);
  });
}

function renderRecebimentoRows(){
  const tbody = document.getElementById('rec_tbody');
  if(!tbody) return;
  if(tbody.children.length) return;
  const linhas = 7;
  const htmlRows = [];
  for(let i=0;i<linhas;i++){
    htmlRows.push(`
      <tr>
        <td><input id="rec_cor_${i}" placeholder="Cor"></td>
        <td><input id="rec_p_${i}" type="number" min="0" value="0"></td>
        <td><input id="rec_m_${i}" type="number" min="0" value="0"></td>
        <td><input id="rec_g_${i}" type="number" min="0" value="0"></td>
        <td><input id="rec_gg_${i}" type="number" min="0" value="0"></td>
        <td><input id="rec_outros_${i}" type="number" min="0" value="0"></td>
        <td><input id="rec_defeito_${i}" placeholder="Apontar defeitos / observações"></td>
      </tr>
    `);
  }
  tbody.innerHTML = htmlRows.join('');
}

function sincronizarRecebimentoDaCostura(){
  renderRecebimentoRows();
  const recId = document.getElementById('rec_prod_id');
  const prodId = document.getElementById('prod_id');
  const recResp = document.getElementById('rec_responsavel');
  const prodResp = document.getElementById('prod_responsavel');
  const recData = document.getElementById('rec_data');
  if(recId && prodId && !recId.value.trim()) recId.value = prodId.value.trim();
  if(recResp && prodResp && !recResp.value.trim()) recResp.value = prodResp.value.trim();
  if(recData && !recData.value) recData.value = new Date().toISOString().slice(0,10);
  for(let i=0;i<7;i++){
    const cor = document.getElementById(`pc_cor_${i}`)?.value || '';
    const p = document.getElementById(`pc_p_${i}`)?.value || 0;
    const m = document.getElementById(`pc_m_${i}`)?.value || 0;
    const g = document.getElementById(`pc_g_${i}`)?.value || 0;
    const gg = document.getElementById(`pc_gg_${i}`)?.value || 0;
    const outros = document.getElementById(`pc_outros_${i}`)?.value || 0;
    const targetCor = document.getElementById(`rec_cor_${i}`);
    if(targetCor){
      targetCor.value = cor;
      document.getElementById(`rec_p_${i}`).value = p;
      document.getElementById(`rec_m_${i}`).value = m;
      document.getElementById(`rec_g_${i}`).value = g;
      document.getElementById(`rec_gg_${i}`).value = gg;
      document.getElementById(`rec_outros_${i}`).value = outros;
    }
  }
  notificar('Dados do romaneio carregados no recebimento!');
}

function exportarRomaneioRecebimento(){
  renderRecebimentoRows();
  const wb = XLSX.utils.book_new();
  const data = [];
  const prodId = document.getElementById('rec_prod_id')?.value.trim() || document.getElementById('prod_id')?.value.trim() || '';
  const responsavel = document.getElementById('rec_responsavel')?.value.trim() || document.getElementById('prod_responsavel')?.value.trim() || '';
  const dataRec = document.getElementById('rec_data')?.value || '';
  data.push([`ROMANEIO DE RECEBIMENTO - ID: ${prodId}`]);
  data.push([`DATA: ${dataRec}`, '', '', '', '', '', `RESPONSÁVEL: ${responsavel}`]);
  data.push([]);
  data.push(['COR','P','M','G','GG','OUTROS','DEFEITO']);
  for(let i=0;i<7;i++){
    data.push([
      document.getElementById(`rec_cor_${i}`)?.value || '',
      Number(document.getElementById(`rec_p_${i}`)?.value || 0),
      Number(document.getElementById(`rec_m_${i}`)?.value || 0),
      Number(document.getElementById(`rec_g_${i}`)?.value || 0),
      Number(document.getElementById(`rec_gg_${i}`)?.value || 0),
      Number(document.getElementById(`rec_outros_${i}`)?.value || 0),
      document.getElementById(`rec_defeito_${i}`)?.value || ''
    ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Recebimento');
  XLSX.writeFile(wb, `Recebimento_${prodId || Date.now()}.xlsx`);
  notificar('Romaneio de recebimento exportado!');
}

// ==========================
// TIMER
// ==========================
let timerInatividade;
function resetTimer(){
  clearTimeout(timerInatividade);
  timerInatividade = setTimeout(()=>{
    logout();
    notificar("Sessão expirada por inatividade!", "error");
  }, 30 * 60 * 1000);
}
["click","mousemove","keypress","scroll","touchstart"].forEach(evt=>{
  document.addEventListener(evt, resetTimer);
});

// ==========================
// NAVEGAÇÃO
// ==========================
function updateNavigationUI(id){
  const titulos = {
    dashboard: "Dashboard",
    pedidos: "Pedidos",
    estoque: "Estoque",
    packlist: "Packlist",
    pedidoCompra: "Pedido de Compra",
    producao: "Produção",
    expedicao: "Expedição",
    config: "Configurações",
    auditoria: "Auditoria",
    backup: "Backup",
    ajuda: "Ajuda"
  };

  const subtitulos = {
    dashboard: "Visão geral da operação, indicadores e alertas do sistema.",
    pedidos: "Cadastro, consulta e atualização dos pedidos em andamento.",
    estoque: "Controle de entradas, saídas e níveis por produto.",
    packlist: "Separação e impressão da lista de conferência dos itens.",
    pedidoCompra: "Monte a lista de compra sem alterar o estoque.",
    producao: "Acompanhamento de risco, corte, costura e custos de produção.",
    expedicao: "Leitura de códigos e organização da planilha de expedição.",
    config: "Ajustes gerais do sistema e gestão de usuários.",
    auditoria: "Histórico de ações realizadas dentro do sistema.",
    backup: "Exportação, importação e armazenamento portátil dos dados.",
    ajuda: "Contato rápido, orientações de operação e suporte do sistema."
  };

  document.querySelectorAll('.nav-btn[data-section]').forEach(btn=>{
    btn.classList.toggle('active', btn.dataset.section === id);
  });

  const topMap = {
    dashboard: 'dashboard',
    estoque: 'estoque',
    pedidos: 'pedidos',
    producao: 'producao',
    packlist: 'pedidos',
    pedidoCompra: 'pedidos',
    expedicao: 'pedidos',
    config: 'dashboard',
    auditoria: 'dashboard',
    backup: 'dashboard',
    ajuda: 'ajuda'
  };
  document.querySelectorAll('.topbar-link[data-topnav]').forEach(link=>{
    link.classList.toggle('active', link.dataset.topnav === topMap[id]);
  });

  const pageTitle = document.getElementById('pageTitle');
  const pageSubtitle = document.getElementById('pageSubtitle');
  if(pageTitle) pageTitle.innerText = titulos[id] || 'Sistema';
  if(pageSubtitle) pageSubtitle.innerText = subtitulos[id] || 'Painel do sistema.';
}

function show(id){ 
  if(id === "expedicao") renderExpedicao();
  ["dashboard","pedidos","estoque","packlist","pedidoCompra","producao","expedicao","config","auditoria","backup","ajuda"].forEach(t=>{
    const el = document.getElementById(t);
    if(el) el.classList.add("hidden");
  });
  const destino = document.getElementById(id);
  if(destino) destino.classList.remove("hidden");
  updateNavigationUI(id);

  if(id === "config") renderUsuarios();
  if(id === "auditoria") renderLogs();
  if(id === "estoque") renderEstoque();
  if(id === "pedidos") renderPedidosModulo();
  if(id === "pedidoCompra") renderPedidoCompra();
  if(id === "producao") renderProducaoDetalhada();
  if(id === "expedicao") setTimeout(focarCampoExpedicao, 80);
}


// ==========================
// PEDIDO DE COMPRA
// ==========================
function normalizarPedidoCompraItem(item){
  if(!item) return null;
  return {
    id: item.id || Date.now(),
    produto: String(item.produto || '').trim(),
    cor: String(item.cor || '').trim(),
    estoqueM: Number(item.estoqueM ?? 0) || 0,
    estoqueGG: Number(item.estoqueGG ?? 0) || 0,
    precisoM: Number(item.precisoM ?? item.m ?? 0) || 0,
    precisoGG: Number(item.precisoGG ?? item.gg ?? 0) || 0,
    obs: String(item.obs || '').trim(),
    criadoEm: item.criadoEm || new Date().toLocaleString('pt-BR')
  };
}

function calcularComprarM(item){
  return Math.max(0, Number(item?.precisoM || 0) - Number(item?.estoqueM || 0));
}

function calcularComprarGG(item){
  return Math.max(0, Number(item?.precisoGG || 0) - Number(item?.estoqueGG || 0));
}

function calcularTotalPedidoCompraItem(item){
  return calcularComprarM(item) + calcularComprarGG(item);
}

function limparFormularioPedidoCompra(){
  ['compraEstoqueM','compraEstoqueGG','compraM','compraGG'].forEach(id => {
    const el = document.getElementById(id);
    if(el) el.value = 0;
  });
  const obs = document.getElementById('compraObs');
  if(obs) obs.value = '';
}

function adicionarPedidoCompra(){
  const produto = (document.getElementById('compraProduto')?.value || '').trim();
  const cor = (document.getElementById('compraCor')?.value || '').trim();
  const estoqueM = parseInt(document.getElementById('compraEstoqueM')?.value || 0, 10) || 0;
  const estoqueGG = parseInt(document.getElementById('compraEstoqueGG')?.value || 0, 10) || 0;
  const precisoM = parseInt(document.getElementById('compraM')?.value || 0, 10) || 0;
  const precisoGG = parseInt(document.getElementById('compraGG')?.value || 0, 10) || 0;
  const obs = (document.getElementById('compraObs')?.value || '').trim();

  if(!produto || !cor){
    notificar('Selecione a peça e a cor.', 'error');
    return;
  }
  if((precisoM + precisoGG) <= 0){
    notificar('Informe ao menos uma quantidade em "Vou precisar".', 'error');
    return;
  }

  const chaveProduto = produto.toLowerCase();
  const chaveCor = cor.toLowerCase();
  const existente = pedidosCompra.find(item =>
    String(item.produto).toLowerCase() === chaveProduto &&
    String(item.cor).toLowerCase() === chaveCor
  );

  if(existente){
    existente.estoqueM += estoqueM;
    existente.estoqueGG += estoqueGG;
    existente.precisoM = Number(existente.precisoM || 0) + precisoM;
    existente.precisoGG = Number(existente.precisoGG || 0) + precisoGG;
    existente.obs = [existente.obs, obs].filter(Boolean).join(' | ');
  } else {
    pedidosCompra.push(normalizarPedidoCompraItem({
      id: Date.now(),
      produto,
      cor,
      estoqueM,
      estoqueGG,
      precisoM,
      precisoGG,
      obs
    }));
  }

  saveLocal();
  renderPedidoCompra();
  limparFormularioPedidoCompra();
  adicionarLog('PEDIDO_COMPRA_ADICIONADO', `${produto} / ${cor} / estoque M:${estoqueM} GG:${estoqueGG} / preciso M:${precisoM} GG:${precisoGG}`);
  notificar('Item adicionado ao pedido de compra!');
}

function removerPedidoCompra(id){
  pedidosCompra = pedidosCompra.filter(item => item.id !== id);
  saveLocal();
  renderPedidoCompra();
  adicionarLog('PEDIDO_COMPRA_REMOVIDO', `ID: ${id}`);
  notificar('Item removido do pedido de compra!');
}

function limparPedidoCompra(){
  if(!pedidosCompra.length){
    notificar('A lista já está vazia.', 'error');
    return;
  }
  if(!confirm('Confirma limpar toda a lista do pedido de compra?')) return;
  pedidosCompra = [];
  saveLocal();
  renderPedidoCompra();
  adicionarLog('PEDIDO_COMPRA_LIMPO', 'Lista de pedido de compra apagada');
  notificar('Pedido de compra limpo!');
}

function renderPedidoCompra(){
  const tbody = document.getElementById('listaPedidoCompra');
  const tfoot = document.getElementById('totaisPedidoCompra');
  const boxItens = document.getElementById('pedidoCompraItens');
  const boxTotal = document.getElementById('pedidoCompraTotalPecas');
  const boxComprarM = document.getElementById('pedidoCompraTotalComprarM');
  const boxComprarGG = document.getElementById('pedidoCompraTotalComprarGG');
  const boxComprarGeral = document.getElementById('pedidoCompraTotalComprarGeral');
  if(!tbody || !tfoot) return;

  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  let totalEstoque = 0;
  let totalComprarM = 0;
  let totalComprarGG = 0;
  let totalComprar = 0;

  if(!pedidosCompra.length){
    tbody.innerHTML = '<tr><td colspan="8"><div class="alerta-vazio" style="margin:0;">Nenhum item no pedido de compra.</div></td></tr>';
  }

  pedidosCompra.map(normalizarPedidoCompraItem).forEach((item, index) => {
    const estoqueTotal = Number(item.estoqueM || 0) + Number(item.estoqueGG || 0);
    const comprarM = calcularComprarM(item);
    const comprarGG = calcularComprarGG(item);
    const totalLinha = comprarM + comprarGG;

    totalEstoque += estoqueTotal;
    totalComprarM += comprarM;
    totalComprarGG += comprarGG;
    totalComprar += totalLinha;

    const detalhe = `Tenho: M ${item.estoqueM || 0} | GG ${item.estoqueGG || 0} • Preciso: M ${item.precisoM || 0} | GG ${item.precisoGG || 0}`;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td><strong>${item.produto}</strong></td>
      <td>${item.cor}<div style="font-size:11px;color:var(--muted);margin-top:4px;">${detalhe}</div></td>
      <td><strong>${comprarM}</strong></td>
      <td><strong>${comprarGG}</strong></td>
      <td><strong>${totalLinha}</strong></td>
      <td>${item.obs || '-'}</td>
      <td><button class="btn-mini delete" onclick="removerPedidoCompra(${item.id})">Remover</button></td>
    `;
    tbody.appendChild(tr);
  });

  if(pedidosCompra.length){
    tfoot.innerHTML = `
      <tr>
        <td colspan="3"><strong>Totais</strong></td>
        <td><strong>${totalComprarM}</strong></td>
        <td><strong>${totalComprarGG}</strong></td>
        <td><strong>${totalComprar}</strong></td>
        <td colspan="2"></td>
      </tr>
    `;
  }

  if(boxItens) boxItens.innerText = String(pedidosCompra.length);
  if(boxTotal) boxTotal.innerText = String(totalEstoque);
  if(boxComprarM) boxComprarM.innerText = String(totalComprarM);
  if(boxComprarGG) boxComprarGG.innerText = String(totalComprarGG);
  if(boxComprarGeral) boxComprarGeral.innerText = String(totalComprar);
}

function imprimirPedidoCompra(){
  if(!pedidosCompra.length){
    notificar('Nenhum item para imprimir.', 'error');
    return;
  }

  let linhas = '';
  let totalComprarM = 0;
  let totalComprarGG = 0;
  let totalComprar = 0;

  pedidosCompra.map(normalizarPedidoCompraItem).forEach((item, index) => {
    const comprarM = calcularComprarM(item);
    const comprarGG = calcularComprarGG(item);
    const totalLinha = comprarM + comprarGG;
    totalComprarM += comprarM;
    totalComprarGG += comprarGG;
    totalComprar += totalLinha;
    linhas += `<tr><td>${index+1}</td><td>${item.produto}</td><td>${item.cor}</td><td>${comprarM}</td><td>${comprarGG}</td><td>${totalLinha}</td><td>${item.obs || '-'}</td></tr>`;
  });

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Pedido de Compra</title><style>body{font-family:Arial;padding:20px}h1{margin:0 0 6px}p{margin:0 0 16px;color:#555}table{width:100%;border-collapse:collapse;margin-top:14px}th,td{border:1px solid #cfcfcf;padding:8px;text-align:left;font-size:13px}th{background:#f3f4f6}tfoot td{font-weight:700;background:#fafafa}</style></head><body><h1>Pedido de Compra</h1><p>Data: ${new Date().toLocaleString('pt-BR')}</p><table><thead><tr><th>#</th><th>Peça</th><th>Cor</th><th>M</th><th>GG</th><th>Total a comprar</th><th>Observação</th></tr></thead><tbody>${linhas}</tbody><tfoot><tr><td colspan="3">Totais</td><td>${totalComprarM}</td><td>${totalComprarGG}</td><td>${totalComprar}</td><td></td></tr></tfoot></table></body></html>`;

  const win = window.open('', '_blank');
  win.document.write(html);
  win.document.close();
  win.focus();
  win.print();
}

function exportarPedidoCompraExcel(){
  if(!pedidosCompra.length){
    notificar('Nenhum item para exportar.', 'error');
    return;
  }

  const agora = new Date();
  const dataBr = agora.toLocaleDateString('pt-BR');
  const horaBr = agora.toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'});

  const linhas = pedidosCompra.map(normalizarPedidoCompraItem).map((item, index) => {
    const comprarM = calcularComprarM(item);
    const comprarGG = calcularComprarGG(item);
    return [
      index + 1,
      item.produto || '',
      item.cor || '',
      comprarM,
      comprarGG,
      {f:`D${index+5}+E${index+5}`},
      item.obs || ''
    ];
  });

  const totalComprarM = pedidosCompra.reduce((a, item) => a + calcularComprarM(normalizarPedidoCompraItem(item)), 0);
  const totalComprarGG = pedidosCompra.reduce((a, item) => a + calcularComprarGG(normalizarPedidoCompraItem(item)), 0);

  const data = [
    ['PEDIDO DE COMPRA'],
    ['Data', dataBr, 'Hora', horaBr],
    [],
    ['#', 'Peça', 'Cor', 'M', 'GG', 'Total a comprar', 'Observações'],
    ...linhas,
    [],
    ['TOTAIS', '', '', totalComprarM, totalComprarGG, {f:`D${linhas.length+7}+E${linhas.length+7}`}, '']
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [
    {wch: 6},
    {wch: 28},
    {wch: 18},
    {wch: 10},
    {wch: 10},
    {wch: 16},
    {wch: 28}
  ];
  ws['!merges'] = [{s:{r:0,c:0}, e:{r:0,c:6}}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'PedidoCompra');
  XLSX.writeFile(wb, `pedido_compra_${new Date().toISOString().slice(0,10)}.xlsx`);
  notificar('Planilha exportada com o pedido final de compra!');
}

// ==========================
// INIT
// ==========================
async 
function aplicarAprimoramentosVisuais(){
  const btnTema = document.getElementById('btnTema');
  if(btnTema && !btnTema.dataset.enhanced){
    btnTema.dataset.enhanced = '1';
    btnTema.innerText = document.body.classList.contains('dark') ? 'Tema claro' : 'Tema escuro';
    btnTema.addEventListener('click', ()=>{
      setTimeout(()=>{
        btnTema.innerText = document.body.classList.contains('dark') ? 'Tema claro' : 'Tema escuro';
      }, 50);
    });
  }
}

async function init(){
  await carregarUsuariosJsonAuto();
  atualizarStatusUsuariosJson(usuariosJsonConectado ? "usuarios.json carregado." : "usuarios.json não conectado.");
  const modoSelect = document.getElementById("modoExpedicao");
  if(modoSelect) modoSelect.value = modoExpedicao;
  atualizarModoExpedicao();
  renderRecebimentoRows();
  preencherSelect("produto", produtos);
  preencherSelect("cor", cores);
  preencherSelect("tamanho", tamanhos);

  preencherSelect("produtoMov", produtos);
  preencherSelect("corMov", cores);
  preencherSelect("tamMov", tamanhos);

  preencherSelect("pedidoProduto", produtos);
  preencherSelect("pedidoCor", cores);
  preencherSelect("pedidoTamanho", tamanhos);

  preencherSelect("compraProduto", produtos);
  preencherSelect("compraCor", cores);

  document.getElementById("empresaNome").value = config.empresaNome || defaults.empresaNome;
  document.getElementById("limiteEstoqueBaixo").value = config.limiteEstoqueBaixo || 3;
  document.getElementById("pedidoData").value = new Date().toISOString().slice(0,10);

  resetDiario();
  buildProducaoTables();
  bindRecalculoProducao();
  render();
  aplicarAprimoramentosVisuais();
  updateNavigationUI("dashboard");

  document.getElementById("dataPrint").innerText = new Date().toLocaleDateString("pt-BR");
}

function render(){
  renderDashboard();
  renderPedidosModulo();
  renderEstoque();
  renderPacklist();
  renderUsuarios();
  showConfigTab("usuarios");
  renderLogs();
  renderProducaoDetalhada();
  renderPedidoCompra();
  renderExpedicao();
}

// ==========================
// DASHBOARD
// ==========================
function renderDashboard(){
  document.getElementById("pedidosHoje").innerText = pedidos.length;

  let totalVendidas = 0;
  let vendasPorProduto = {};
  let alertas = 0;

  pedidos.forEach(p=>{
    if(p.status !== "CANCELADO"){
      totalVendidas += Number(p.qtd || 0);
      vendasPorProduto[p.produto] = (vendasPorProduto[p.produto] || 0) + Number(p.qtd || 0);
    }
  });

  const totalEstoque = Object.values(estoque).reduce((a,b)=>a+b,0);

  Object.keys(estoque).forEach(k=>{
    const qtd = estoque[k] || 0;
    if(qtd > 0 && qtd <= Number(config.limiteEstoqueBaixo || 3)) alertas++;
  });

  let maisVendido = "-";
  let maior = 0;
  Object.keys(vendasPorProduto).forEach(prod=>{
    if(vendasPorProduto[prod] > maior){
      maior = vendasPorProduto[prod];
      maisVendido = prod;
    }
  });

  document.getElementById("pecasVendidas").innerText = totalVendidas;
  document.getElementById("pecasEstoque").innerText = totalEstoque;
  document.getElementById("maisVendidos").innerText = maisVendido;
  document.getElementById("alertasEstoque").innerText = alertas;

  renderAlertasDashboard();
}

function renderAlertasDashboard(){
  const lista = document.getElementById("listaAlertasDashboard");
  lista.innerHTML = "";

  const alertas = [];
  produtos.forEach(prod=>{
    cores.forEach(c=>{
      tamanhos.forEach(t=>{
        const qtd = estoque[key(prod,c,t)] || 0;
        if(qtd > 0 && qtd <= Number(config.limiteEstoqueBaixo || 3)){
          alertas.push(`${prod} — ${c} / ${t}: ${qtd} peça(s)`);
        }
      });
    });
  });

  if(alertas.length === 0){
    lista.innerHTML = `<div class="alerta-vazio">Nenhum alerta de estoque baixo no momento.</div>`;
    return;
  }

  alertas.slice(0,10).forEach(item=>{
    lista.innerHTML += `<div class="alerta-box">⚠ ${item}</div>`;
  });
}

// ==========================
// PEDIDOS
// ==========================
function criarPedido(){
  const codigo = document.getElementById("pedidoCodigo").value.trim() || ("PED-" + Date.now());
  const cliente = document.getElementById("pedidoCliente").value.trim();
  const produto = document.getElementById("pedidoProduto").value;
  const cor = document.getElementById("pedidoCor").value;
  const tamanho = document.getElementById("pedidoTamanho").value;
  const qtd = parseInt(document.getElementById("pedidoQtd").value);
  const status = document.getElementById("pedidoStatus").value;
  const data = document.getElementById("pedidoData").value || new Date().toISOString().slice(0,10);
  const obs = document.getElementById("pedidoObs").value.trim();

  if(!cliente || !produto || !cor || !tamanho || !qtd || qtd <= 0){
    notificar("Preencha os dados do pedido corretamente.", "error");
    return;
  }

  if(pedidos.some(p=>p.codigo === codigo)){
    notificar("Já existe um pedido com esse código.", "error");
    return;
  }

  const saldo = estoque[key(produto, cor, tamanho)] || 0;
  if(status !== "CANCELADO" && qtd > saldo){
    notificar("Estoque insuficiente para criar este pedido.", "error");
    return;
  }

  if(status !== "CANCELADO"){
    estoque[key(produto, cor, tamanho)] = saldo - qtd;
  }

  pedidos.unshift({codigo, cliente, produto, cor, tamanho, qtd, status, data, obs});
  saveLocal();
  render();
  adicionarLog("PEDIDO_CRIADO", `Código: ${codigo}, Cliente: ${cliente}, Produto: ${produto}, Qtd: ${qtd}`);
  notificar("Pedido salvo com sucesso!");

  document.getElementById("pedidoCodigo").value = "";
  document.getElementById("pedidoCliente").value = "";
  document.getElementById("pedidoQtd").value = "";
  document.getElementById("pedidoObs").value = "";
  document.getElementById("pedidoStatus").value = "PENDENTE";
  document.getElementById("pedidoData").value = new Date().toISOString().slice(0,10);
}

function statusPill(status){
  const s = String(status || "").toUpperCase();
  if(s === "CONCLUIDO") return `<span class="pill pill-success">${s}</span>`;
  if(s === "CANCELADO") return `<span class="pill pill-danger">${s}</span>`;
  if(s === "SEPARANDO") return `<span class="pill pill-warning">${s}</span>`;
  return `<span class="pill pill-neutral">${s}</span>`;
}

function renderPedidosModulo(){
  const lista = document.getElementById("listaPedidosModulo");
  if(!lista) return;

  const busca = (document.getElementById("filtroPedidoBusca").value || "").toLowerCase().trim();
  const statusFiltro = document.getElementById("filtroPedidoStatus").value;

  lista.innerHTML = "";

  const filtrados = pedidos.filter(p=>{
    const txt = `${p.codigo} ${p.cliente} ${p.produto}`.toLowerCase();
    const okBusca = !busca || txt.includes(busca);
    const okStatus = !statusFiltro || p.status === statusFiltro;
    return okBusca && okStatus;
  });

  if(filtrados.length === 0){
    lista.innerHTML = `<tr><td colspan="9">Nenhum pedido encontrado.</td></tr>`;
    return;
  }

  filtrados.forEach(p=>{
    const idx = pedidos.findIndex(x=>x.codigo === p.codigo);
    lista.innerHTML += `
      <tr>
        <td>${p.codigo}</td>
        <td>${p.cliente}</td>
        <td>${p.produto}</td>
        <td>${p.cor}</td>
        <td>${p.tamanho}</td>
        <td>${p.qtd}</td>
        <td>${statusPill(p.status)}</td>
        <td>${p.data}</td>
        <td>
          <button class="secondary" onclick="alterarStatusPedido(${idx}, 'CONCLUIDO')">Concluir</button>
          <button class="danger" onclick="cancelarPedido(${idx})">Cancelar</button>
        </td>
      </tr>
    `;
  });
}

function alterarStatusPedido(index, novoStatus){
  const pedido = pedidos[index];
  if(!pedido) return;
  pedido.status = novoStatus;
  saveLocal();
  renderPedidosModulo();
  adicionarLog("PEDIDO_STATUS", `Pedido ${pedido.codigo} alterado para ${novoStatus}`);
  notificar("Status atualizado!");
}

function cancelarPedido(index){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  const pedido = pedidos[index];
  if(!pedido) return;

  if(pedido.status !== "CANCELADO"){
    estoque[key(pedido.produto, pedido.cor, pedido.tamanho)] =
      (estoque[key(pedido.produto, pedido.cor, pedido.tamanho)] || 0) + Number(pedido.qtd || 0);
  }

  pedido.status = "CANCELADO";
  saveLocal();
  render();
  adicionarLog("PEDIDO_CANCELADO", `Pedido ${pedido.codigo} cancelado`);
  notificar("Pedido cancelado!");
}

// ==========================
// PACKLIST
// ==========================
function addPedidoPacklist(){
  const produto = document.getElementById("produto").value;
  const cor = document.getElementById("cor").value;
  const tamanho = document.getElementById("tamanho").value;
  const qtd = parseInt(document.getElementById("quantidade").value);

  if(!qtd || qtd <= 0){
    notificar("Quantidade inválida!", "error");
    return;
  }

  const saldo = estoque[key(produto, cor, tamanho)] || 0;
  if(qtd > saldo){
    notificar("Estoque insuficiente para este packlist!", "error");
    return;
  }

  packlist.push({produto, cor, tamanho, qtd});
  estoque[key(produto, cor, tamanho)] = saldo - qtd;

  saveLocal();
  render();
  adicionarLog("PACKLIST_ADICIONADO", `Produto: ${produto}, Cor: ${cor}, Tam: ${tamanho}, Qtd: ${qtd}`);
  notificar("Item adicionado ao packlist!");
}

function renderPacklist(){
  const lista = document.getElementById("listaPacklist");
  const totais = document.getElementById("totaisPacklist");

  lista.innerHTML = "";
  totais.innerHTML = "";

  let totalGeral = 0;
  const produtosTotais = {};

  packlist.forEach((o, i)=>{
    lista.innerHTML += `
      <tr>
        <td>${o.produto}</td>
        <td>${o.cor}</td>
        <td>${o.tamanho}</td>
        <td>${o.qtd}</td>
        <td><button class="danger" onclick="removerItemPacklist(${i})">Excluir</button></td>
      </tr>
    `;
    totalGeral += o.qtd;
    produtosTotais[o.produto] = (produtosTotais[o.produto] || 0) + o.qtd;
  });

  let footerHTML = `<tr><td colspan="3">Total Geral</td><td>${totalGeral}</td><td></td></tr>`;
  Object.keys(produtosTotais).forEach(p=>{
    footerHTML += `<tr><td colspan="3">${p}</td><td>${produtosTotais[p]}</td><td></td></tr>`;
  });

  totais.innerHTML = footerHTML;
  aplicarPermissoes();
}

function removerItemPacklist(index){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  const item = packlist[index];
  if(!item) return;

  estoque[key(item.produto, item.cor, item.tamanho)] =
    (estoque[key(item.produto, item.cor, item.tamanho)] || 0) + Number(item.qtd || 0);

  packlist.splice(index, 1);
  saveLocal();
  render();
  adicionarLog("PACKLIST_REMOVIDO", `Produto: ${item.produto}, Qtd: ${item.qtd}`);
  notificar("Item removido do packlist!");
}

function imprimirPacklist(){
  if(packlist.length === 0){
    notificar("Não há itens no packlist!", "error");
    return;
  }

  const acao = prompt("Escolha:\n1 - Imprimir / Salvar PDF\n2 - Excluir packlist");

  if(acao === "1"){
    let tabela = `<table>
      <thead>
        <tr><th>Produto</th><th>Cor</th><th>Tamanho</th><th>Qtd</th></tr>
      </thead><tbody>`;

    packlist.forEach(p=>{
      tabela += `<tr><td>${p.produto}</td><td>${p.cor}</td><td>${p.tamanho}</td><td>${p.qtd}</td></tr>`;
    });

    tabela += `</tbody><tfoot>${document.getElementById("totaisPacklist").innerHTML}</tfoot></table>`;
    const rodape = document.getElementById("packlistRodape").outerHTML;

    const printWindow = window.open("", "", "width=800,height=600");
    printWindow.document.write(`
      <html>
      <head>
        <title>Packlist</title>
        <style>
          body{font-family:sans-serif;padding:20px;}
          table{width:100%;border-collapse:collapse;margin-bottom:20px;}
          th,td{border:1px solid #000;padding:8px;text-align:left;}
          tfoot td{font-weight:bold;}
        

.prod-shell{display:flex;flex-direction:column;gap:16px;}
.prod-tabs{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:2px;}
.prod-tab-btn{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:12px;padding:12px 16px;font-weight:800;cursor:pointer;transition:.18s ease;box-shadow:var(--shadow);}
.prod-tab-btn:hover{transform:translateY(-1px);}
.prod-tab-btn.active{background:var(--primary-soft);border-color:var(--primary-border);color:var(--primary);}
.prod-panel{display:none;}
.prod-panel.active{display:block;}
.prod-note{color:var(--muted);margin:0 0 14px;line-height:1.6;}
.prod-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px;}

.config-shell{display:grid;grid-template-columns:220px 1fr;gap:18px;align-items:start;margin-top:10px;}
.config-side{display:flex;flex-direction:column;gap:10px;}
.config-tab-btn{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:12px;padding:12px 14px;text-align:left;font-weight:700;cursor:pointer;transition:.18s ease;}
.config-tab-btn:hover{transform:translateY(-1px);box-shadow:var(--shadow);}
.config-tab-btn.active{background:var(--primary-soft);border-color:rgba(12,74,110,.22);color:var(--primary);}
.config-panel{display:none;}
.config-panel.active{display:block;}
.config-info-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;margin-top:14px;}
.config-box{border:1px solid var(--line);background:var(--card);border-radius:14px;padding:16px;box-shadow:var(--shadow);}
.config-box h3,.config-box h4{margin:0 0 8px;}
.config-box p,.config-box li,.config-box small{color:var(--muted);}
.config-box ul{margin:10px 0 0 18px;padding:0;}
.users-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px;}
.users-toolbar{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;}
.users-toolbar .field{border:1px solid var(--line);background:var(--card);border-radius:14px;padding:14px;}
.user-file-status{margin-top:8px;font-size:13px;color:var(--muted);}
.master-badge{display:inline-flex;align-items:center;gap:8px;padding:7px 10px;border-radius:999px;background:var(--primary-soft);color:var(--primary);font-weight:700;font-size:12px;}
.inline-actions.wrap{flex-wrap:wrap;}
@media (max-width: 980px){.config-shell{grid-template-columns:1fr;}.config-side{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));}}

</style>
      </head>
      <body>${tabela}${rodape}</body>
      </html>
    `);

    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    printWindow.close();
    notificar("Use 'Salvar como PDF' do navegador.");
  } else if(acao === "2"){
    if(confirm("Tem certeza? Excluirá o packlist do dia.")){
      packlist.forEach(item=>{
        estoque[key(item.produto, item.cor, item.tamanho)] =
          (estoque[key(item.produto, item.cor, item.tamanho)] || 0) + Number(item.qtd || 0);
      });
      packlist = [];
      saveLocal();
      render();
      adicionarLog("PACKLIST_EXCLUIDO", "Todos os itens do packlist foram excluídos");
      notificar("Packlist excluído!");
    }
  } else {
    notificar("Opção inválida!", "error");
  }
}

// ==========================
// ESTOQUE
// ==========================
function entrada(){
  const p = document.getElementById("produtoMov").value;
  const c = document.getElementById("corMov").value;
  const t = document.getElementById("tamMov").value;
  const q = parseInt(document.getElementById("qtdMov").value);

  if(!q || q <= 0){
    notificar("Quantidade inválida!", "error");
    return;
  }

  estoque[key(p,c,t)] = (estoque[key(p,c,t)] || 0) + q;
  saveLocal();
  render();
  adicionarLog("ENTRADA_ESTOQUE", `Produto: ${p}, Cor: ${c}, Tamanho: ${t}, Qtd: ${q}`);
  notificar("Entrada no estoque registrada!");
}

function saida(){
  const p = document.getElementById("produtoMov").value;
  const c = document.getElementById("corMov").value;
  const t = document.getElementById("tamMov").value;
  const q = parseInt(document.getElementById("qtdMov").value);

  if(!q || q <= 0){
    notificar("Quantidade inválida!", "error");
    return;
  }

  const k = key(p,c,t);
  const atual = estoque[k] || 0;

  if(q > atual){
    notificar("Estoque insuficiente!", "error");
    return;
  }

  estoque[k] = atual - q;
  saveLocal();
  render();
  adicionarLog("SAIDA_ESTOQUE", `Produto: ${p}, Cor: ${c}, Tamanho: ${t}, Qtd: ${q}`);
  notificar("Saída do estoque registrada!");
}

function novoProduto(){
  const nome = prompt("Nome do novo produto:");
  const nomeNormalizado = String(nome || "").trim();
  if(!nomeNormalizado || produtos.some(p => p.toLowerCase() === nomeNormalizado.toLowerCase())){
    notificar("Produto inválido ou já existente!", "error");
    return;
  }

  produtos.push(nomeNormalizado);
  saveLocal();
  init();
  adicionarLog("PRODUTO_ADICIONADO", `Produto: ${nomeNormalizado}`);
  notificar(`Produto "${nomeNormalizado}" adicionado!`);
}

function novaCor(){
  const nome = prompt("Nome da nova cor:");
  const nomeNormalizado = String(nome || "").trim();
  if(!nomeNormalizado){
    notificar("Cor inválida!", "error");
    return;
  }
  if(cores.some(c => c.toLowerCase() === nomeNormalizado.toLowerCase())){
    notificar("Essa cor já existe no sistema!", "error");
    return;
  }

  cores.push(nomeNormalizado);
  saveLocal();
  init();
  adicionarLog("COR_ADICIONADA", `Cor: ${nomeNormalizado}`);
  notificar(`Cor "${nomeNormalizado}" adicionada!`);
}

function removerCor(){
  const corAtual = document.getElementById("corMov")?.value || "";
  const nome = prompt("Digite o nome da cor que deseja remover:", corAtual);
  const nomeNormalizado = String(nome || "").trim();
  if(!nomeNormalizado){
    notificar("Nenhuma cor informada.", "error");
    return;
  }
  const corEncontrada = cores.find(c => c.toLowerCase() === nomeNormalizado.toLowerCase());
  if(!corEncontrada){
    notificar("Cor não encontrada no sistema!", "error");
    return;
  }
  if(!confirm(`Confirma remover a cor "${corEncontrada}"?`)) return;

  cores = cores.filter(c => c.toLowerCase() !== corEncontrada.toLowerCase());
  Object.keys(estoque).forEach(k=>{
    const partes = k.split("|");
    if(partes[1] && partes[1].toLowerCase() === corEncontrada.toLowerCase()) delete estoque[k];
  });
  pedidos = pedidos.filter(p => String(p.cor || "").toLowerCase() !== corEncontrada.toLowerCase());
  packlist = packlist.filter(p => String(p.cor || "").toLowerCase() !== corEncontrada.toLowerCase());
  saveLocal();
  init();
  adicionarLog("COR_REMOVIDA", `Cor: ${corEncontrada}`);
  notificar(`Cor "${corEncontrada}" removida!`);
}

function retirarProduto(){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  const p = document.getElementById("produtoMov").value;
  if(confirm(`Confirma remover "${p}" do catálogo?`)){
    produtos = produtos.filter(x => x !== p);
    Object.keys(estoque).forEach(k=>{
      if(k.startsWith(p + "|")) delete estoque[k];
    });
    saveLocal();
    init();
    adicionarLog("PRODUTO_REMOVIDO", `Produto: ${p}`);
    notificar(`Produto "${p}" removido!`);
  }
}

function calcularTotalProduto(produtoNome){
  let total = 0;
  cores.forEach(c=>{
    tamanhos.forEach(t=>{
      total += estoque[key(produtoNome, c, t)] || 0;
    });
  });
  return total;
}

function classificarEstoque(total){
  if(total <= 2) return {classe:"critico", texto:"Crítico"};
  if(total <= Number(config.limiteEstoqueBaixo || 3) + 2) return {classe:"baixo", texto:"Baixo"};
  return {classe:"bom", texto:"Bom"};
}

function montarDetalhesProduto(produtoNome){
  let html = "";
  let encontrou = false;

  cores.forEach(corItem=>{
    let linhas = "";
    let totalCor = 0;

    tamanhos.forEach(tamItem=>{
      const qtd = estoque[key(produtoNome, corItem, tamItem)] || 0;
      totalCor += qtd;

      if(qtd > 0){
        linhas += `
          <div class="estoque-tamanho-linha">
            <span>Tam. ${tamItem}</span>
            <strong>${qtd}</strong>
          </div>
        `;
      }
    });

    if(totalCor > 0){
      encontrou = true;
      html += `
        <div class="estoque-cor-bloco">
          <div class="estoque-cor-titulo">${corItem} — ${totalCor} peça(s)</div>
          ${linhas}
        </div>
      `;
    }
  });

  if(!encontrou){
    html = `<div class="estoque-vazio">Sem estoque disponível para este produto.</div>`;
  }

  return html;
}

function toggleEstoqueCard(card){
  card.classList.toggle("open");
}

function renderEstoque(){
  const lista = document.getElementById("estoqueLista");
  const busca = (document.getElementById("buscaEstoque")?.value || "").toLowerCase().trim();

  lista.innerHTML = "";

  let produtosFiltrados = produtos.filter(prod => prod.toLowerCase().includes(busca));

  if(produtosFiltrados.length === 0){
    lista.innerHTML = `<div class="alerta-vazio">Nenhum produto encontrado na busca.</div>`;
    document.getElementById("pecasEstoque").innerText = Object.values(estoque).reduce((a,b)=>a+b,0);
    return;
  }

  produtosFiltrados.forEach(prod=>{
    const totalProduto = calcularTotalProduto(prod);
    const detalheHtml = montarDetalhesProduto(prod);
    const status = classificarEstoque(totalProduto);

    const card = document.createElement("div");
    card.className = "estoque-card";
    card.onclick = function(){
      toggleEstoqueCard(card);
    };

    card.innerHTML = `
      <div class="estoque-faixa faixa-${status.classe}"></div>
      <div class="estoque-card-header">
        <h3>${prod}</h3>
        <div class="estoque-toggle">⌄</div>
      </div>

      <div class="estoque-total">
        ${totalProduto}
        <span>${totalProduto === 1 ? "Peça" : "Peças"}</span>
      </div>

      <div class="estoque-status status-${status.classe}">
        Estoque ${status.texto}
      </div>

      <div class="estoque-detalhe">
        ${detalheHtml}
      </div>
    `;

    lista.appendChild(card);
  });

  const totalEstoque = Object.values(estoque).reduce((a,b)=>a+b,0);
  document.getElementById("pecasEstoque").innerText = totalEstoque;
}

function zerarEstoque(){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  if(confirm("Deseja zerar todo estoque?")){
    estoque = {};
    saveLocal();
    render();
    adicionarLog("ESTOQUE_ZERADO", "Todo o estoque foi zerado");
    notificar("Estoque zerado!");
  }
}

// ==========================
// PRODUÇÃO
// ==========================
function buildProducaoTables(){
  const rcBody = document.getElementById("rc_tbody");
  const pcBody = document.getElementById("pc_tbody");
  if(!rcBody || !pcBody) return;

  if(rcBody.innerHTML.trim() === ""){
    for(let i=0;i<12;i++){
      rcBody.innerHTML += `
        <tr>
          <td><input id="rc_cor_${i}"></td>
          <td><input id="rc_p_${i}" type="number"></td>
          <td><input id="rc_m_${i}" type="number"></td>
          <td><input id="rc_g_${i}" type="number"></td>
          <td><input id="rc_gg_${i}" type="number"></td>
          <td><input id="rc_outros_${i}" type="number"></td>
          <td><input id="rc_obs_${i}"></td>
          <td><input id="rc_subtotal_${i}" readonly></td>
        </tr>
      `;
    }
  }

  if(pcBody.innerHTML.trim() === ""){
    for(let i=0;i<7;i++){
      pcBody.innerHTML += `
        <tr>
          <td><input id="pc_cor_${i}"></td>
          <td><input id="pc_p_${i}" type="number"></td>
          <td><input id="pc_m_${i}" type="number"></td>
          <td><input id="pc_g_${i}" type="number"></td>
          <td><input id="pc_gg_${i}" type="number"></td>
          <td><input id="pc_outros_${i}" type="number"></td>
          <td><input id="pc_obs_${i}"></td>
          <td><input id="pc_subtotal_${i}" readonly></td>
        </tr>
      `;
    }
  }
}

function getInputValue(id){
  const el = document.getElementById(id);
  return el ? el.value : "";
}

function setInputValue(id, value){
  const el = document.getElementById(id);
  if(el) el.value = value;
}

function setTextValue(id, value){
  const el = document.getElementById(id);
  if(el) el.innerText = value;
}

function calcularSubtotalLinha(prefixo, index){
  const p = numero(getInputValue(`${prefixo}_p_${index}`));
  const m = numero(getInputValue(`${prefixo}_m_${index}`));
  const g = numero(getInputValue(`${prefixo}_g_${index}`));
  const gg = numero(getInputValue(`${prefixo}_gg_${index}`));
  const outros = numero(getInputValue(`${prefixo}_outros_${index}`));
  return p + m + g + gg + outros;
}

function calcularTotaisRiscoCorte(){
  let total = 0;

  for(let i=0;i<12;i++){
    const subtotal = calcularSubtotalLinha("rc", i);
    total += subtotal;
    setInputValue(`rc_subtotal_${i}`, subtotal);
  }

  const corte = numero(getInputValue("rc_corte"));
  const risco = numero(getInputValue("rc_risco_valor"));
  const ampliacao = numero(getInputValue("rc_ampliacao_valor"));
  const valorFinal = corte + risco + ampliacao;
  const valorPorPeca = total > 0 ? valorFinal / total : 0;

  setInputValue("rc_total_pecas", total);
  setInputValue("rc_valor_final", valorFinal.toFixed(2));
  setInputValue("rc_valor_peca", valorPorPeca.toFixed(4));

  setTextValue("rc_total_pecas_text", total);
  setTextValue("rc_valor_final_text", money(valorFinal));
  setTextValue("rc_valor_peca_text", money(valorPorPeca));

  return { totalPecas: total, corte, risco, ampliacao, valorFinal, valorPorPeca };
}

function calcularTotaisCostura(){
  let total = 0;

  for(let i=0;i<7;i++){
    const subtotal = calcularSubtotalLinha("pc", i);
    total += subtotal;
    setInputValue(`pc_subtotal_${i}`, subtotal);
  }

  const qtdRolo = numero(getInputValue("pc_qtd_rolo"));
  const transporte = numero(getInputValue("pc_transporte"));
  const aviamento = numero(getInputValue("pc_aviamento"));
  const riscoCorte = calcularTotaisRiscoCorte().valorFinal;
  const valorTecido = qtdRolo * 539.8;
  const valorCostura = total * 3;
  const custoTotal = valorTecido + valorCostura + riscoCorte + transporte + aviamento;
  const custoUnitario = total > 0 ? custoTotal / total : 0;

  setInputValue("pc_total_pecas", total);
  setInputValue("pc_valor_tecido", valorTecido.toFixed(2));
  setInputValue("pc_valor_costura", valorCostura.toFixed(2));
  setInputValue("pc_valor_risco_corte", riscoCorte.toFixed(2));
  setInputValue("pc_custo_total", custoTotal.toFixed(2));
  setInputValue("pc_custo_unitario", custoUnitario.toFixed(4));

  setTextValue("pc_total_pecas_text", total);
  setTextValue("pc_valor_tecido_text", money(valorTecido));
  setTextValue("pc_valor_costura_text", money(valorCostura));
  setTextValue("pc_valor_risco_corte_text", money(riscoCorte));
  setTextValue("pc_custo_total_text", money(custoTotal));
  setTextValue("pc_custo_unitario_text", money(custoUnitario));

  return {
    totalPecas: total,
    qtdRolo,
    valorTecido,
    valorCostura,
    riscoCorte,
    transporte,
    aviamento,
    custoTotal,
    custoUnitario
  };
}

function recalcularProducaoCompleta(){
  calcularTotaisRiscoCorte();
  calcularTotaisCostura();
}

function coletarLinhas(prefixo, qtdLinhas){
  const linhas = [];
  for(let i=0;i<qtdLinhas;i++){
    const cor = getInputValue(`${prefixo}_cor_${i}`).trim();
    const p = numero(getInputValue(`${prefixo}_p_${i}`));
    const m = numero(getInputValue(`${prefixo}_m_${i}`));
    const g = numero(getInputValue(`${prefixo}_g_${i}`));
    const gg = numero(getInputValue(`${prefixo}_gg_${i}`));
    const outros = numero(getInputValue(`${prefixo}_outros_${i}`));
    const obs = getInputValue(`${prefixo}_obs_${i}`).trim();
    const subtotal = p + m + g + gg + outros;

    if(cor || subtotal > 0 || obs){
      linhas.push({cor,p,m,g,gg,outros,obs,subtotal});
    }
  }
  return linhas;
}

function salvarProducaoPlanilha(){
  const resumoRisco = calcularTotaisRiscoCorte();
  const resumoCostura = calcularTotaisCostura();

  const registro = {
    id: getInputValue("prod_id") || ("PROD-" + Date.now()),
    dataInicio: getInputValue("prod_data_inicio"),
    dataFim: getInputValue("prod_data_fim"),
    modeloTecido: getInputValue("prod_modelo_tecido"),
    localProducao: getInputValue("prod_local_producao"),
    localCorte: getInputValue("rc_local_corte"),
    ampliacaoNome: getInputValue("rc_ampliacao_nome"),
    riscoNome: getInputValue("rc_risco_nome"),
    responsavel: getInputValue("prod_responsavel"),
    riscoCorte: {
      linhas: coletarLinhas("rc", 12),
      corte: resumoRisco.corte,
      risco: resumoRisco.risco,
      ampliacao: resumoRisco.ampliacao,
      totalPecas: resumoRisco.totalPecas,
      valorFinal: resumoRisco.valorFinal,
      valorPorPeca: resumoRisco.valorPorPeca
    },
    costura: {
      linhas: coletarLinhas("pc", 7),
      qtdRolo: resumoCostura.qtdRolo,
      valorTecido: resumoCostura.valorTecido,
      valorCostura: resumoCostura.valorCostura,
      riscoCorte: resumoCostura.riscoCorte,
      transporte: resumoCostura.transporte,
      aviamento: resumoCostura.aviamento,
      totalPecas: resumoCostura.totalPecas,
      custoTotal: resumoCostura.custoTotal,
      custoUnitario: resumoCostura.custoUnitario
    }
  };

  const indexExistente = producaoDetalhada.findIndex(x=>x.id === registro.id);
  if(indexExistente >= 0) producaoDetalhada[indexExistente] = registro;
  else producaoDetalhada.unshift(registro);

  saveLocal();
  renderProducaoDetalhada();
  adicionarLog("PRODUCAO_SALVA", `ID: ${registro.id}, Modelo: ${registro.modeloTecido}, Total peças: ${registro.costura.totalPecas}`);
  notificar("Produção salva com sucesso!");
}

function renderProducaoDetalhada(){
  const lista = document.getElementById("listaProducao");
  lista.innerHTML = "";

  if(producaoDetalhada.length === 0){
    lista.innerHTML = `<tr><td colspan="8">Nenhuma produção salva.</td></tr>`;
    return;
  }

  producaoDetalhada.forEach((item, i)=>{
    lista.innerHTML += `
      <tr>
        <td>${item.id || ""}</td>
        <td>${item.dataInicio || ""}</td>
        <td>${item.dataFim || ""}</td>
        <td>${item.modeloTecido || ""}</td>
        <td>${item.costura.totalPecas || 0}</td>
        <td>${money(item.costura.custoTotal || 0)}</td>
        <td>${money(item.costura.custoUnitario || 0)}</td>
        <td>
          <button class="primary" onclick="carregarProducaoDetalhada(${i})">Abrir</button>
          <button class="danger" onclick="removerProducaoDetalhada(${i})">Excluir</button>
        </td>
      </tr>
    `;
  });
}

function carregarProducaoDetalhada(index){
  const item = producaoDetalhada[index];
  if(!item) return;

  setInputValue("prod_id", item.id || "");
  setInputValue("prod_data_inicio", item.dataInicio || "");
  setInputValue("prod_data_fim", item.dataFim || "");
  setInputValue("prod_modelo_tecido", item.modeloTecido || "");
  setInputValue("prod_local_producao", item.localProducao || "");
  setInputValue("rc_local_corte", item.localCorte || "");
  setInputValue("rc_ampliacao_nome", item.ampliacaoNome || "");
  setInputValue("rc_risco_nome", item.riscoNome || "");
  setInputValue("prod_responsavel", item.responsavel || "");

  for(let i=0;i<12;i++){
    const linha = item.riscoCorte.linhas[i] || {};
    setInputValue(`rc_cor_${i}`, linha.cor || "");
    setInputValue(`rc_p_${i}`, linha.p || "");
    setInputValue(`rc_m_${i}`, linha.m || "");
    setInputValue(`rc_g_${i}`, linha.g || "");
    setInputValue(`rc_gg_${i}`, linha.gg || "");
    setInputValue(`rc_outros_${i}`, linha.outros || "");
    setInputValue(`rc_obs_${i}`, linha.obs || "");
  }

  for(let i=0;i<7;i++){
    const linha = item.costura.linhas[i] || {};
    setInputValue(`pc_cor_${i}`, linha.cor || "");
    setInputValue(`pc_p_${i}`, linha.p || "");
    setInputValue(`pc_m_${i}`, linha.m || "");
    setInputValue(`pc_g_${i}`, linha.g || "");
    setInputValue(`pc_gg_${i}`, linha.gg || "");
    setInputValue(`pc_outros_${i}`, linha.outros || "");
    setInputValue(`pc_obs_${i}`, linha.obs || "");
  }

  setInputValue("rc_corte", item.riscoCorte.corte || "");
  setInputValue("rc_risco_valor", item.riscoCorte.risco || "");
  setInputValue("rc_ampliacao_valor", item.riscoCorte.ampliacao || "");
  setInputValue("pc_qtd_rolo", item.costura.qtdRolo || "");
  setInputValue("pc_transporte", item.costura.transporte || "");
  setInputValue("pc_aviamento", item.costura.aviamento || "");

  recalcularProducaoCompleta();
  notificar("Produção carregada!");
}

function removerProducaoDetalhada(index){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  if(!confirm("Deseja excluir esta produção?")) return;
  producaoDetalhada.splice(index, 1);
  saveLocal();
  renderProducaoDetalhada();
  notificar("Produção removida!");
}

function limparFormularioProducao(){
  const ids = [
    "prod_id","prod_data_inicio","prod_data_fim","prod_modelo_tecido","prod_local_producao",
    "rc_local_corte","rc_ampliacao_nome","rc_risco_nome","prod_responsavel",
    "rc_corte","rc_risco_valor","rc_ampliacao_valor",
    "pc_qtd_rolo","pc_transporte","pc_aviamento"
  ];
  ids.forEach(id=>setInputValue(id, ""));

  for(let i=0;i<12;i++){
    ["cor","p","m","g","gg","outros","obs","subtotal"].forEach(campo=> setInputValue(`rc_${campo}_${i}`, ""));
  }
  for(let i=0;i<7;i++){
    ["cor","p","m","g","gg","outros","obs","subtotal"].forEach(campo=> setInputValue(`pc_${campo}_${i}`, ""));
  }

  recalcularProducaoCompleta();
  notificar("Formulário limpo!");
}

function bindRecalculoProducao(){
  const ids = [];
  for(let i=0;i<12;i++) ids.push(`rc_p_${i}`, `rc_m_${i}`, `rc_g_${i}`, `rc_gg_${i}`, `rc_outros_${i}`);
  for(let i=0;i<7;i++) ids.push(`pc_p_${i}`, `pc_m_${i}`, `pc_g_${i}`, `pc_gg_${i}`, `pc_outros_${i}`);
  ids.push("rc_corte","rc_risco_valor","rc_ampliacao_valor","pc_qtd_rolo","pc_transporte","pc_aviamento");

  ids.forEach(id=>{
    const el = document.getElementById(id);
    if(el){
      el.removeEventListener("input", recalcularProducaoCompleta);
      el.addEventListener("input", recalcularProducaoCompleta);
    }
  });
}

function exportarProducaoParaPlanilha(){
  if(typeof XLSX === "undefined"){
    notificar("Biblioteca XLSX não carregada no sistema.", "error");
    return;
  }

  const resumoRisco = calcularTotaisRiscoCorte();
  const resumoCostura = calcularTotaisCostura();

  const wb = XLSX.utils.book_new();

  const rcData = [];
  rcData.push(["ROMANEIO DE RISCO E CORTE","","","","","","","DATA:"]);
  rcData.push(["MODELO:", getInputValue("prod_modelo_tecido"), "", "", "", "", "", "ID: " + getInputValue("prod_id")]);
  rcData.push(["AMPLIAÇÃO:", getInputValue("rc_ampliacao_nome")]);
  rcData.push(["RISCO:", getInputValue("rc_risco_nome")]);
  rcData.push(["LOCAL CORTE:", getInputValue("rc_local_corte")]);
  rcData.push([]);
  rcData.push(["COR","P","M","G","GG","OUTROS","OBSERVAÇÃO","SUBTOTAL"]);

  for(let i=0;i<12;i++){
    rcData.push([
      getInputValue(`rc_cor_${i}`),
      numero(getInputValue(`rc_p_${i}`)),
      numero(getInputValue(`rc_m_${i}`)),
      numero(getInputValue(`rc_g_${i}`)),
      numero(getInputValue(`rc_gg_${i}`)),
      numero(getInputValue(`rc_outros_${i}`)),
      getInputValue(`rc_obs_${i}`),
      calcularSubtotalLinha("rc", i)
    ]);
  }

  rcData.push(["TOTAL","","","","","","",resumoRisco.totalPecas]);
  rcData.push(["CORTE:",resumoRisco.corte]);
  rcData.push(["RISCO:",resumoRisco.risco]);
  rcData.push(["AMPLIAÇÃO:",resumoRisco.ampliacao]);
  rcData.push([]);
  rcData.push(["VALOR FINAL:",resumoRisco.valorFinal]);
  rcData.push(["VALOR POR PEÇA:",resumoRisco.valorPorPeca]);
  rcData.push([]);
  rcData.push(["RESPONSAVEL:", getInputValue("prod_responsavel")]);

  const wsRC = XLSX.utils.aoa_to_sheet(rcData);
  XLSX.utils.book_append_sheet(wb, wsRC, "Risco e Corte");

  const pcData = [];
  pcData.push(["ROMANEIO DE PRODUÇÃO COSTURA           ID:", getInputValue("prod_id")]);
  pcData.push([]);
  pcData.push(["DATA INICIO:", getInputValue("prod_data_inicio"), "", "", "", "DATA FIM:", getInputValue("prod_data_fim")]);
  pcData.push([]);
  pcData.push(["MODELO/TECIDO:", getInputValue("prod_modelo_tecido")]);
  pcData.push([]);
  pcData.push(["LOCAL DE PRODUÇÃO:", getInputValue("prod_local_producao")]);
  pcData.push([]);
  pcData.push(["COR","P","M","G","GG","OUTROS","OBSERVAÇÃO","QTD TOTAL"]);

  for(let i=0;i<7;i++){
    pcData.push([
      getInputValue(`pc_cor_${i}`),
      numero(getInputValue(`pc_p_${i}`)),
      numero(getInputValue(`pc_m_${i}`)),
      numero(getInputValue(`pc_g_${i}`)),
      numero(getInputValue(`pc_gg_${i}`)),
      numero(getInputValue(`pc_outros_${i}`)),
      getInputValue(`pc_obs_${i}`),
      calcularSubtotalLinha("pc", i)
    ]);
  }

  while(pcData.length < 24) pcData.push([]);
  pcData.push(["","","","","","","",resumoCostura.totalPecas]);
  pcData.push([]);
  pcData.push(["CUSTO UNITARIO POR PEÇA E TOTAL"]);
  pcData.push(["QUANTIDADE DE ROLO:","",resumoCostura.qtdRolo]);
  pcData.push(["VALOR DO TECIDO:","",resumoCostura.valorTecido]);
  pcData.push(["VALOR COSTURA UN:","",resumoCostura.valorCostura]);
  pcData.push(["CORTE/ AMPLIA/ RISCO:","",resumoCostura.riscoCorte]);
  pcData.push(["TRANSPORTE:","",resumoCostura.transporte]);
  pcData.push(["AVIAMENTO:","",resumoCostura.aviamento]);
  pcData.push(["CUSTO TOTAL:","",resumoCostura.custoTotal]);
  pcData.push(["CUSTO UNITARIO PÇ:","",resumoCostura.custoUnitario]);
  pcData.push([]);
  pcData.push(["RESPONSAVEL:", getInputValue("prod_responsavel")]);

  const wsPC = XLSX.utils.aoa_to_sheet(pcData);
  XLSX.utils.book_append_sheet(wb, wsPC, "Romaneio Costura");

  const recData = [];
  recData.push(["RECEBIMENTO __/___/___ - RESPONSAVEL:"]);
  recData.push(["COR","P","M","G","GG","OUTROS","DEFEITO"]);
  recData.push([]);
  recData.push([]);
  recData.push([]);
  recData.push([]);
  recData.push([]);

  const wsREC = XLSX.utils.aoa_to_sheet(recData);
  XLSX.utils.book_append_sheet(wb, wsREC, "Recebimento");

  XLSX.writeFile(wb, `Romaneio_${getInputValue("prod_id") || Date.now()}.xlsx`);
  notificar("Planilha exportada com sucesso!");
}

// ==========================
// CONFIG
// ==========================
function salvarConfiguracoesGerais(){
  config.empresaNome = document.getElementById("empresaNome").value.trim() || defaults.empresaNome;
  config.limiteEstoqueBaixo = Math.max(1, parseInt(document.getElementById("limiteEstoqueBaixo").value) || 3);
  saveLocal();
  atualizarBoasVindas();
  renderDashboard();
  adicionarLog("CONFIG_SALVA", "Configurações gerais atualizadas");
  notificar("Configurações salvas!");
}

function atualizarPainelUsuarios(){
  const liberado = isMaster();
  const panel = document.getElementById("masterUserPanel");
  const blocked = document.getElementById("masterUserPanelBlocked");
  if(panel) panel.classList.toggle("hidden", !liberado);
  if(blocked) blocked.classList.toggle("hidden", liberado);
}

function renderUsuarios(){
  atualizarPainelUsuarios();
  const lista = document.getElementById("listaUsuarios");
  if(!lista) return;
  lista.innerHTML = "";

  const listaOrdenada = usuarios.slice().sort((a,b)=> a.usuario.localeCompare(b.usuario, 'pt-BR'));
  if(listaOrdenada.length === 0){
    lista.innerHTML = `<tr><td colspan="4">Nenhum usuário cadastrado.</td></tr>`;
    return;
  }

  listaOrdenada.forEach(u=>{
    const podeExcluir = isMaster() && u.usuario !== "Rodolfo";
    const safeName = u.usuario.replace(/'/g, "\'");
    lista.innerHTML += `<tr>
      <td>${u.usuario}</td>
      <td>${u.nivel}</td>
      <td>${u.origem || "custom"}</td>
      <td>${podeExcluir ? `<button class="btn-mini delete" onclick="removerUsuario('${safeName}')">Remover</button>` : `<span style="color:var(--muted)">${u.usuario === "Rodolfo" ? "Protegido" : "Sem permissão"}</span>`}</td>
    </tr>`;
  });
}

async function adicionarUsuario(){
  show("config");
  showConfigTab("usuarios");
  document.getElementById("novoUsuarioNome")?.focus();
}

async function adicionarUsuarioFormulario(){
  const nome = document.getElementById("novoUsuarioNome")?.value || "";
  const senha = document.getElementById("novoUsuarioSenha")?.value || "";
  const nivel = document.getElementById("novoUsuarioNivel")?.value || "OPER";
  const ok = await criarUsuario(nome, senha, nivel);
  if(ok){
    document.getElementById("novoUsuarioNome").value = "";
    document.getElementById("novoUsuarioSenha").value = "";
    document.getElementById("novoUsuarioNivel").value = "OPER";
  }
}

async function criarUsuario(nome, senha, nivel){
  if(!isMaster()) return notificar("Apenas Rodolfo pode adicionar usuários.", "error");
  nome = String(nome || "").trim();
  senha = String(senha || "").trim();
  nivel = String(nivel || "OPER").trim().toUpperCase();
  if(!nome) return notificar("Informe o nome do usuário.", "error");
  if(!senha) return notificar("Informe a senha do usuário.", "error");
  if(usuarios.some(u => u.usuario.toLowerCase() === nome.toLowerCase())){
    return notificar("Já existe um usuário com esse nome.", "error");
  }
  if(!["ADMIN","OPER"].includes(nivel)) nivel = "OPER";
  usuarios.push(normalizarUsuario({usuario:nome, senha:senha, nivel:nivel, origem:"custom"}));
  saveLocal();
  await persistirUsuariosJson();
  renderUsuarios();
  adicionarLog("USUARIO_ADICIONADO", `Usuário: ${nome}, Nível: ${nivel}`);
  notificar(`Usuário "${nome}" adicionado com sucesso!`);
  return true;
}

async function removerUsuario(nome){
  if(!isMaster()) return notificar("Apenas Rodolfo pode remover usuários.", "error");
  if(!nome) return notificar("Usuário inválido.", "error");
  if(nome === "Rodolfo") return notificar("O usuário master não pode ser removido.", "error");
  const idx = usuarios.findIndex(u => u.usuario === nome);
  if(idx === -1) return notificar("Usuário não encontrado.", "error");
  if(!confirm(`Remover o usuário ${nome}?`)) return;
  const removido = usuarios.splice(idx, 1)[0];
  saveLocal();
  await persistirUsuariosJson();
  renderUsuarios();
  adicionarLog("USUARIO_REMOVIDO", `Usuário: ${removido.usuario}`);
  notificar(`Usuário "${removido.usuario}" removido!`);
}

async function resetUsuarios(){
  if(!isMaster()) return notificar("Apenas Rodolfo pode resetar usuários padrão.", "error");
  if(confirm("Deseja restaurar apenas o usuário master Rodolfo? Isso apagará os demais usuários.")){
    usuarios = usuariosPadrao();
    saveLocal();
    await persistirUsuariosJson();
    renderUsuarios();
    adicionarLog("USUARIOS_RESETADOS", "Usuários restaurados para o padrão mínimo");
    notificar("Usuários restaurados!");
  }
}

// ==========================
// LOGS
// ==========================
function renderLogs(){
  const tbody = document.getElementById("listaLogs");
  if(!tbody) return;
  tbody.innerHTML = "";
  if(logs.length === 0){
    tbody.innerHTML = `<tr><td colspan="4">Nenhum log registrado.</td></tr>`;
    return;
  }
  logs.forEach(l=>{
    tbody.innerHTML += `
      <tr>
        <td>${l.dataHora}</td>
        <td>${l.usuario}</td>
        <td>${l.acao}</td>
        <td>${l.detalhes}</td>
      </tr>
    `;
  });
}

function limparLogs(){
  if(!isMaster()) return notificar("Apenas Rodolfo pode limpar logs.", "error");
  if(confirm("Tem certeza que deseja limpar todos os logs?")){
    logs = [];
    saveLocal();
    renderLogs();
    notificar("Logs limpos!");
  }
}

// ==========================
// BACKUP / IMPORTAÇÃO
// ==========================
function exportarDados(){
  const data = {
    config,
    usuarios,
    usuarioAtual,
    pedidos,
    packlist,
    estoque,
    produtos,
    logs,
    producaoDetalhada
  };

  const blob = new Blob([JSON.stringify(data, null, 2)], {type: "application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "backup_erp_sa_confeccoes_v5_1.json";
  a.click();
  URL.revokeObjectURL(url);
  notificar("Backup exportado!");
}

function importarDados(){
  const fileInput = document.getElementById("importFile");
  if(fileInput.files.length === 0){
    notificar("Selecione um arquivo JSON!", "error");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function(e){
    try {
      const data = JSON.parse(e.target.result);
      config = data.config || config;
      usuarios = data.usuarios || usuarios;
      usuarioAtual = data.usuarioAtual || usuarioAtual;
      pedidos = data.pedidos || [];
      packlist = data.packlist || [];
      estoque = data.estoque || {};
      produtos = data.produtos || defaults.produtosPadrao.slice();
      logs = data.logs || [];
      producaoDetalhada = data.producaoDetalhada || [];
      saveLocal();
      init();
      atualizarBoasVindas();

      if(usuarioAtual){
        document.getElementById("loginTela").style.display = "none";
        aplicarPermissoes();
      }

      notificar("Dados importados com sucesso!");
    } catch(err){
      notificar("Erro ao importar JSON!", "error");
    }
  };

  reader.readAsText(file);
}

async function saveDados() {
  if(!window.showSaveFilePicker){
    notificar("Seu navegador não suporta salvamento portátil.", "error");
    return;
  }

  if(!dadosFileHandle){
    try{
      dadosFileHandle = await window.showSaveFilePicker({
        suggestedName: "dados_sa_confeccoes_v5_1.json",
        types: [{description: "JSON S.A Confecções", accept: {"application/json": [".json"]}}]
      });
    } catch(e){
      return;
    }
  }

  try {
    const writable = await dadosFileHandle.createWritable();
    await writable.write(JSON.stringify({
      config, usuarios, usuarioAtual, pedidos, packlist, estoque, produtos, logs, producaoDetalhada
    }, null, 2));
    await writable.close();
    notificar("Dados salvos em arquivo portátil!");
  } catch(e){
    notificar("Erro ao salvar dados: " + e.message, "error");
  }
}

async function loadDados() {
  if(!window.showOpenFilePicker){
    notificar("Seu navegador não suporta carregamento portátil.", "error");
    return;
  }

  try {
    const [fileHandle] = await window.showOpenFilePicker({
      multiple: false,
      types: [{description: "JSON S.A Confecções", accept: {"application/json": [".json"]}}]
    });

    dadosFileHandle = fileHandle;
    const file = await fileHandle.getFile();
    const text = await file.text();
    const data = JSON.parse(text);

    config = data.config || config;
    usuarios = data.usuarios || usuarios;
    usuarioAtual = data.usuarioAtual || usuarioAtual;
    pedidos = data.pedidos || [];
    packlist = data.packlist || [];
    estoque = data.estoque || {};
    produtos = data.produtos || defaults.produtosPadrao.slice();
    logs = data.logs || [];
    producaoDetalhada = data.producaoDetalhada || [];

    saveLocal();
    init();
    renderUsuarios();
    renderLogs();

    if(usuarioAtual){
      document.getElementById("loginTela").style.display = "none";
      atualizarBoasVindas();
      aplicarPermissoes();
    }

    notificar("Dados carregados do arquivo portátil!");
  } catch(e){
    notificar("Erro ao carregar dados: " + e.message, "error");
  }
}

// ==========================
// RESET DIÁRIO
// ==========================
function resetDiario(){
  const hoje = new Date().toLocaleDateString("pt-BR");
  const ultimo = appStorage.getItem("ultimoDiaERP");

  if(ultimo !== hoje){
    packlist = [];
    saveLocal();
    appStorage.setItem("ultimoDiaERP", hoje);
  }
}

// ==========================
// ONLOAD
// ==========================
window.onload = async ()=>{
  await loadCloudState();

  if(appStorage.getItem("tema") === "dark"){
    document.body.classList.add("dark");
  } else {
    document.body.classList.remove("dark");
  }

  atualizarBoasVindas();

  document.getElementById("loginUsuario").addEventListener("keypress", function(e){
    if(e.key === "Enter") login();
  });
  document.getElementById("loginSenha").addEventListener("keypress", function(e){
    if(e.key === "Enter") login();
  });

  if(usuarioAtual){
    document.getElementById("loginTela").style.display = "none";
    init();
    aplicarPermissoes();
    resetTimer();
  } else {
    document.getElementById("loginTela").style.display = "flex";
    init();
  }
};
function focarCampoExpedicao(){
  const input = document.getElementById("codigoExpedicao");
  if(input) input.focus();
}

function atualizarModoExpedicao(){
  const select = document.getElementById("modoExpedicao");
  const texto = document.getElementById("textoModoExpedicao");
  if(select){
    modoExpedicao = select.value || "auto";
    appStorage.setItem("modoExpedicao", modoExpedicao);
    scheduleCloudSave();
  }

  const mensagens = {
    auto: "No automático: códigos numéricos com 11 dígitos entram como Mercado Envios. BR entra como Shopee no destino BR ativo.",
    shopee_agencia: "Os próximos códigos BR serão lançados como Shopee Agência até você trocar o modo.",
    shopee_direta: "Os próximos códigos BR serão lançados como Shopee Entrega Direta até você trocar o modo.",
    mercado_envios: "Todos os próximos códigos serão lançados como Mercado Envios até você trocar o modo.",
    mercado_flex: "Todos os próximos códigos serão lançados como Mercado Flex até você trocar o modo."
  };

  if(texto) texto.innerText = mensagens[modoExpedicao] || mensagens.auto;
  atualizarPainelExpedicao();
}

function setPreferenciaShopee(tipo){
  if(!["shopee_agencia","shopee_direta"].includes(tipo)) return;
  preferenciaShopee = tipo;
  appStorage.setItem("preferenciaShopee", preferenciaShopee);
  scheduleCloudSave();
  atualizarPainelExpedicao();
  notificar(`Destino BR alterado para ${nomePreferenciaShopee()}.`);
  focarCampoExpedicao();
}

function nomePreferenciaShopee(){
  if(preferenciaShopee === "shopee_direta") return "Shopee Entrega Direta";
  return "Shopee Agência";
}

function identificarTipoExpedicao(codigo){
  const c = String(codigo || "").trim().toUpperCase();
  if(!c) return null;

  const manualMap = {
    shopee_agencia: {tipo: "Shopee Xpress", caixa: "CAIXA SHOPEE XPRESS (AGÊNCIA)", classe: "expedicao-aviso-xpress"},
    shopee_direta: {tipo: "Shopee Entrega Direta", caixa: "CAIXA SHOPEE ENTREGA DIRETA / TIGER FLEX", classe: "expedicao-aviso-direta"},
    mercado_envios: {tipo: "Mercado Envios", caixa: "CAIXA MERCADO ENVIOS (AGÊNCIA / CORREIOS)", classe: "expedicao-aviso-mercado"},
    mercado_flex: {tipo: "Mercado Flex", caixa: "CAIXA MERCADO FLEX / TIGER FLEX", classe: "expedicao-aviso-ok"}
  };

  const prefixo = c.slice(0,2);
  const detalhes = {prefixo, tamanho: c.length};

  if(modoExpedicao === "mercado_envios") return {...manualMap.mercado_envios, ...detalhes};
  if(modoExpedicao === "mercado_flex") return {...manualMap.mercado_flex, ...detalhes};
  if(modoExpedicao === "shopee_agencia" && prefixo === "BR") return {...manualMap.shopee_agencia, ...detalhes};
  if(modoExpedicao === "shopee_direta" && prefixo === "BR") return {...manualMap.shopee_direta, ...detalhes};

  if(/^\d{11}$/.test(c)){
    return {...manualMap.mercado_envios, ...detalhes};
  }

  if(prefixo === "BR" && c.length >= 10){
    return {...(preferenciaShopee === "shopee_direta" ? manualMap.shopee_direta : manualMap.shopee_agencia), ...detalhes};
  }

  if(prefixo === "46"){
    return {...manualMap.mercado_flex, ...detalhes};
  }

  return {tipo: "Não reconhecido", caixa: "VERIFICAR MANUALMENTE", classe: "expedicao-aviso-erro", ...detalhes};
}

function handleBipExpedicao(event){
  if(event.key === "Enter"){
    event.preventDefault();
    processarCodigoExpedicao();
  }
}

function atualizarPainelExpedicao(ultimoItem = null){
  const ultimoCodigo = document.getElementById("ultimoCodigoExpedicao");
  const ultimaLeitura = document.getElementById("ultimaLeituraExpedicao");
  const pref = document.getElementById("preferenciaShopee");

  if(pref) pref.innerText = nomePreferenciaShopee();
  if(ultimoCodigo) ultimoCodigo.innerText = ultimoItem ? ultimoItem.codigo : (expedicao[0]?.codigo || "--");
  if(ultimaLeitura) ultimaLeitura.innerText = ultimoItem ? ultimoItem.dataHora : (expedicao[0]?.dataHora || "--");
}

function processarCodigoExpedicao(){
  const input = document.getElementById("codigoExpedicao");
  const aviso = document.getElementById("expedicaoAviso");
  const codigo = input.value.trim().toUpperCase();

  if(!codigo){
    notificar("Digite ou bipa um código.", "error");
    return;
  }

  const jaExiste = expedicao.some(item => item.codigo === codigo);
  if(jaExiste){
    aviso.className = "expedicao-aviso-erro";
    aviso.innerHTML = `⚠ Código já bipado: <strong>${codigo}</strong>`;
    input.value = "";
    input.focus();
    notificar("Esse código já foi lançado.", "error");
    return;
  }

  const info = identificarTipoExpedicao(codigo);
  const agora = new Date().toLocaleString("pt-BR");
  const item = {
    codigo,
    prefixo: info.prefixo || codigo.slice(0,2),
    tamanho: info.tamanho || codigo.length,
    tipo: info.tipo,
    caixa: info.caixa,
    dataHora: agora
  };

  expedicao.unshift(item);

  if(codigo.startsWith("BR")){
    if(modoExpedicao === "shopee_agencia" || info.tipo === "Shopee Xpress"){
      preferenciaShopee = "shopee_agencia";
    } else if(modoExpedicao === "shopee_direta" || info.tipo === "Shopee Entrega Direta"){
      preferenciaShopee = "shopee_direta";
    }
    appStorage.setItem("preferenciaShopee", preferenciaShopee);
    appStorage.setItem("memoriaExpedicao", JSON.stringify(memoriaExpedicao));
  }

  saveLocal();
  renderExpedicao();
  atualizarPainelExpedicao(item);

  if(info.tipo === "Não reconhecido"){
    aviso.className = "expedicao-aviso-erro";
    aviso.innerHTML = `⚠ <strong>${codigo}</strong><br>Tipo não reconhecido.<br><strong>VERIFICAR MANUALMENTE</strong>`;
    notificar("Código não reconhecido.", "error");
  } else {
    aviso.className = info.classe;
    aviso.innerHTML = `✅ <strong>${codigo}</strong><br>Tipo: <strong>${info.tipo}</strong><br>SEPARE EM: <strong>${info.caixa}</strong>`;
    adicionarLog("EXPEDICAO_BIPADA", `Código: ${codigo} | Tipo: ${info.tipo}`);
    notificar("Código lançado na expedição!");
  }

  input.value = "";
  input.focus();
}

function renderExpedicao(){
  const lista = document.getElementById("listaExpedicao");
  if(!lista) return;

  lista.innerHTML = "";

  let totalXpress = 0;
  let totalMercado = 0;
  let totalDireta = 0;
  let totalFlex = 0;

  if(expedicao.length === 0){
    lista.innerHTML = `<tr><td colspan="7">Nenhum código bipado ainda.</td></tr>`;
  } else {
    expedicao.forEach((item, i)=>{
      if(item.tipo === "Shopee Xpress") totalXpress++;
      if(item.tipo === "Mercado Envios") totalMercado++;
      if(item.tipo === "Shopee Entrega Direta") totalDireta++;
      if(item.tipo === "Mercado Flex") totalFlex++;

      lista.innerHTML += `
        <tr>
          <td>${i + 1}</td>
          <td>${item.codigo}<br><small style="color:var(--muted)">${item.tamanho || item.codigo.length} caracteres</small></td>
          <td>${item.prefixo || item.codigo.slice(0,2)}</td>
          <td>${item.tipo}</td>
          <td><strong>${item.caixa}</strong></td>
          <td>${item.dataHora}</td>
          <td><button class="danger" onclick="removerExpedicao(${i})">Excluir</button></td>
        </tr>
      `;
    });
  }

  document.getElementById("totalShopeeXpress").innerText = totalXpress;
  document.getElementById("totalMercadoEnvios").innerText = totalMercado;
  document.getElementById("totalShopeeDireta").innerText = totalDireta;
  document.getElementById("totalMercadoFlex").innerText = totalFlex;
  atualizarPainelExpedicao();
}


function removerExpedicao(index){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  const item = expedicao[index];
  if(!item) return;

  expedicao.splice(index, 1);
  saveLocal();
  renderExpedicao();
  adicionarLog("EXPEDICAO_REMOVIDA", `Código removido: ${item.codigo}`);
  notificar("Código removido da expedição.");
}

function limparExpedicao(){
  if(!(isMaster() || hasPermission("DELETE_GENERAL"))) return notificar("Acesso negado!", "error");
  if(!confirm("Deseja limpar toda a planilha de expedição?")) return;

  expedicao = [];
  saveLocal();
  renderExpedicao();
  atualizarPainelExpedicao();

  const aviso = document.getElementById("expedicaoAviso");
  if(aviso){
    aviso.className = "alerta-vazio";
    aviso.innerHTML = "Aguardando leitura...";
  }

  adicionarLog("EXPEDICAO_LIMPA", "Planilha de expedição zerada");
  focarCampoExpedicao();
  notificar("Expedição limpa!");
}

function imprimirExpedicao(){
  if(expedicao.length === 0){
    notificar("Não há itens na expedição para imprimir.", "error");
    return;
  }

  const totalXpress = expedicao.filter(x=>x.tipo === "Shopee Xpress").length;
  const totalMercado = expedicao.filter(x=>x.tipo === "Mercado Envios").length;
  const totalDireta = expedicao.filter(x=>x.tipo === "Shopee Entrega Direta").length;
  const totalFlex = expedicao.filter(x=>x.tipo === "Mercado Flex").length;

  let linhas = "";
  expedicao.forEach((item, i)=>{
    linhas += `
      <tr>
        <td>${i + 1}</td>
        <td>${item.codigo}</td>
        <td>${item.prefixo || item.codigo.slice(0,2)}</td>
        <td>${item.tipo}</td>
        <td>${item.caixa}</td>
        <td>${item.dataHora}</td>
      </tr>
    `;
  });

  const janela = window.open("", "", "width=1000,height=700");
  janela.document.write(`
    <html>
      <head>
        <title>Expedição</title>
        <style>
          body{font-family:Arial,sans-serif;padding:20px;color:#111}
          h1,h2{margin:0 0 12px}
          .resumo{margin-bottom:20px}
          .resumo div{margin-bottom:6px;font-weight:700}
          table{width:100%;border-collapse:collapse}
          th,td{border:1px solid #ccc;padding:8px;text-align:left;font-size:13px}
          th{background:#f3f4f6}
        </style>
      </head>
      <body>
        <h1>Planilha de Expedição</h1>
        <h2>S.A Confecções</h2>
        <div class="resumo">
          <div>Shopee Xpress: ${totalXpress}</div>
          <div>Shopee Entrega Direta: ${totalDireta}</div>
          <div>Mercado Envios: ${totalMercado}</div>
          <div>Mercado Flex: ${totalFlex}</div>
          <div>Data: ${new Date().toLocaleString("pt-BR")}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Código</th>
              <th>Prefixo</th>
              <th>Tipo</th>
              <th>Caixa</th>
              <th>Data/Hora</th>
            </tr>
          </thead>
          <tbody>${linhas}</tbody>
        </table>
      </body>
    </html>
  `);
  janela.document.close();
  janela.focus();
  janela.print();
}


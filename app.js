const CONFIG = window.SA_CONFIG || {};
if (!CONFIG.SUPABASE_URL || !CONFIG.SUPABASE_ANON_KEY) {
  console.warn("Preencha o config.js com a URL e a anon key do Supabase.");
}
const supabaseClient = window.supabase.createClient(CONFIG.SUPABASE_URL || "https://example.supabase.co", CONFIG.SUPABASE_ANON_KEY || "public-anon-key");

const state = {
  session: null,
  profile: null,
  settings: null,
  products: [],
  inventory: [],
  purchaseOrders: [],
  productions: [],
  shipments: [],
  audit: [],
};

const defaults = {
  workspace_id: "sa-principal",
  company_name: "S.A Confecções",
  low_stock_limit: 3,
  accent_color: "#165dff",
  notes: "",
};

const sectionMap = {
  dashboard: { el: "dashboardSection", title: "Dashboard", subtitle: "Visão geral da operação em nuvem." },
  products: { el: "productsSection", title: "Produtos", subtitle: "Cadastre, edite e controle a base de produtos." },
  inventory: { el: "inventorySection", title: "Estoque", subtitle: "Entradas, saídas, ajustes e saldo por variante." },
  purchaseOrders: { el: "purchaseOrdersSection", title: "Pedido de compra", subtitle: "Controle compras abertas, recebidas e canceladas." },
  production: { el: "productionSection", title: "Produção", subtitle: "Acompanhe lotes, responsáveis e quantidades prontas." },
  shipments: { el: "shipmentsSection", title: "Expedição", subtitle: "Leitura rápida de etiquetas e histórico de despachos." },
  settings: { el: "settingsSection", title: "Configurações", subtitle: "Empresa, workspace, perfil e preferências." },
  audit: { el: "auditSection", title: "Auditoria", subtitle: "Histórico de ações gravadas no banco." },
};

document.addEventListener("DOMContentLoaded", initApp);

async function initApp() {
  bindStaticEvents();

  const { data: { session } } = await supabaseClient.auth.getSession();
  if (session) {
    state.session = session;
    await bootstrapApp();
  } else {
    showAuth();
  }

  supabaseClient.auth.onAuthStateChange(async (event, session) => {
    state.session = session;
    if (event === "SIGNED_IN" || event === "TOKEN_REFRESHED") {
      await bootstrapApp();
    }
    if (event === "SIGNED_OUT") {
      resetState();
      showAuth();
    }
  });
}

function bindStaticEvents() {
  document.querySelectorAll(".auth-tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".auth-tab").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const tab = btn.dataset.authTab;
      ["login", "signup", "reset"].forEach(name => {
        document.getElementById(`${name}Form`).classList.toggle("hidden", name !== tab);
      });
    });
  });

  document.getElementById("loginForm").addEventListener("submit", handleLogin);
  document.getElementById("signupForm").addEventListener("submit", handleSignup);
  document.getElementById("resetForm").addEventListener("submit", handleResetPassword);
  document.getElementById("logoutBtn").addEventListener("click", async () => {
    await supabaseClient.auth.signOut();
    toast("Sessão encerrada.");
  });
  document.getElementById("themeToggle").addEventListener("click", toggleTheme);
  document.getElementById("refreshAllBtn").addEventListener("click", async () => {
    await loadAllData();
    toast("Dados atualizados.");
  });
  document.getElementById("productForm").addEventListener("submit", saveProduct);
  document.getElementById("productResetBtn").addEventListener("click", clearProductForm);
  document.getElementById("inventoryForm").addEventListener("submit", createInventoryMovement);
  document.getElementById("purchaseForm").addEventListener("submit", createPurchaseOrder);
  document.getElementById("productionForm").addEventListener("submit", createProductionOrder);
  document.getElementById("shipmentForm").addEventListener("submit", createShipment);
  document.getElementById("settingsForm").addEventListener("submit", saveSettings);
  document.getElementById("profileForm").addEventListener("submit", saveProfile);

  document.getElementById("shipmentCode").addEventListener("input", (e) => {
    const code = e.target.value.trim();
    const info = classifyShipment(code, document.getElementById("shipmentMode").value);
    document.getElementById("shipmentChannel").value = info.channel;
    setShipmentFeedback(info);
  });

  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.addEventListener("click", () => setSection(btn.dataset.section));
  });

  document.querySelectorAll("[data-jump]").forEach(btn => {
    btn.addEventListener("click", () => setSection(btn.dataset.jump));
  });

  ["productSearch","inventorySearch","poSearch","productionSearch","shipmentSearch","auditSearch"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("input", renderAllTables);
  });
}

async function handleLogin(e) {
  e.preventDefault();
  const email = document.getElementById("loginEmail").value.trim();
  const password = document.getElementById("loginPassword").value;
  const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
  if (error) return toast(error.message, "error");
  toast("Login realizado com sucesso.");
}

async function handleSignup(e) {
  e.preventDefault();
  const email = document.getElementById("signupEmail").value.trim();
  const password = document.getElementById("signupPassword").value;
  const full_name = document.getElementById("signupName").value.trim();
  const redirectTo = window.location.origin + window.location.pathname;
  const { error } = await supabaseClient.auth.signUp({
    email,
    password,
    options: { data: { full_name }, emailRedirectTo: redirectTo },
  });
  if (error) return toast(error.message, "error");
  toast("Conta criada. Verifique seu email se a confirmação estiver habilitada.");
}

async function handleResetPassword(e) {
  e.preventDefault();
  const email = document.getElementById("resetEmail").value.trim();
  const redirectTo = window.location.origin + window.location.pathname;
  const { error } = await supabaseClient.auth.resetPasswordForEmail(email, { redirectTo });
  if (error) return toast(error.message, "error");
  toast("Link de recuperação enviado.");
}

async function bootstrapApp() {
  try {
    await ensureProfile();
    await ensureSettings();
    await loadAllData();
    applyUserContext();
    showApp();
    setSection("dashboard");
  } catch (error) {
    console.error(error);
    toast(error.message || "Erro ao carregar o sistema.", "error");
  }
}

async function ensureProfile() {
  const user = state.session?.user;
  if (!user) throw new Error("Usuário não autenticado.");
  const { data, error } = await supabaseClient
    .from("profiles")
    .select("*")
    .eq("id", user.id)
    .maybeSingle();

  if (error) throw error;

  if (!data) {
    const payload = {
      id: user.id,
      email: user.email,
      full_name: user.user_metadata?.full_name || user.email?.split("@")[0] || "Usuário",
      role: "MASTER",
    };
    const { error: insertError, data: inserted } = await supabaseClient
      .from("profiles")
      .upsert(payload)
      .select()
      .single();
    if (insertError) throw insertError;
    state.profile = inserted;
    return;
  }
  state.profile = data;
}

async function ensureSettings() {
  const { data, error } = await supabaseClient
    .from("app_settings")
    .select("*")
    .limit(1)
    .maybeSingle();
  if (error) throw error;

  if (!data) {
    const payload = { ...defaults, created_by: state.profile.id, updated_by: state.profile.id };
    const { error: insertError, data: inserted } = await supabaseClient
      .from("app_settings")
      .insert(payload)
      .select()
      .single();
    if (insertError) throw insertError;
    state.settings = inserted;
    await addAudit("INSERT", "app_settings", "Configurações iniciais criadas");
    applyAccentColor();
    return;
  }
  state.settings = data;
  applyAccentColor();
}

async function loadAllData() {
  await Promise.all([
    loadProducts(),
    loadInventory(),
    loadPurchaseOrders(),
    loadProductions(),
    loadShipments(),
    loadAudit(),
  ]);
  populateProductSelects();
  renderDashboard();
  renderAllTables();
  populateSettingsForms();
}

async function loadProducts() {
  const { data, error } = await supabaseClient
    .from("products")
    .select("*")
    .order("created_at", { ascending: false });
  if (error) throw error;
  state.products = data || [];
}

async function loadInventory() {
  const { data, error } = await supabaseClient
    .from("inventory_balance")
    .select("*")
    .order("product_name", { ascending: true });
  if (error) throw error;
  state.inventory = data || [];
}

async function loadPurchaseOrders() {
  const { data, error } = await supabaseClient
    .from("purchase_orders")
    .select("*, products(name)")
    .order("created_at", { ascending: false });
  if (error) throw error;
  state.purchaseOrders = data || [];
}

async function loadProductions() {
  const { data, error } = await supabaseClient
    .from("production_orders")
    .select("*, products(name)")
    .order("created_at", { ascending: false });
  if (error) throw error;
  state.productions = data || [];
}

async function loadShipments() {
  const { data, error } = await supabaseClient
    .from("shipments")
    .select("*")
    .order("created_at", { ascending: false })
    .limit(200);
  if (error) throw error;
  state.shipments = data || [];
}

async function loadAudit() {
  const { data, error } = await supabaseClient
    .from("audit_logs")
    .select("*")
    .order("created_at", { ascending: false })
    .limit(300);
  if (error) throw error;
  state.audit = data || [];
}

function populateProductSelects() {
  const selects = ["inventoryProduct", "poProduct", "productionProduct"];
  selects.forEach(id => {
    const select = document.getElementById(id);
    select.innerHTML = `<option value="">Selecione</option>`;
    state.products.forEach(product => {
      const option = document.createElement("option");
      option.value = product.id;
      option.textContent = `${product.name}${product.sku ? ` • ${product.sku}` : ""}`;
      select.appendChild(option);
    });
  });
}

function renderDashboard() {
  document.getElementById("metricProducts").textContent = state.products.filter(p => p.active).length;
  document.getElementById("metricStock").textContent = state.inventory.reduce((acc, item) => acc + Number(item.balance || 0), 0);
  document.getElementById("metricPO").textContent = state.purchaseOrders.filter(po => ["aberto","enviado"].includes(po.status)).length;
  document.getElementById("metricProd").textContent = state.productions.filter(po => ["aberto","em_andamento"].includes(po.status)).length;
  const today = new Date().toISOString().slice(0,10);
  document.getElementById("metricShip").textContent = state.shipments.filter(item => (item.created_at || "").slice(0,10) === today).length;

  const lowLimit = Number(state.settings?.low_stock_limit || defaults.low_stock_limit);
  const low = state.inventory.filter(item => Number(item.balance) <= lowLimit);
  const lowEl = document.getElementById("lowStockList");
  if (!low.length) {
    lowEl.className = "stack-list empty-state";
    lowEl.textContent = "Nenhum item crítico no momento.";
  } else {
    lowEl.className = "stack-list";
    lowEl.innerHTML = low.slice(0, 10).map(item => `
      <div class="stack-item">
        <strong>${escapeHtml(item.product_name)}</strong>
        <span>${escapeHtml(item.color)} • ${escapeHtml(item.size)} • saldo ${Number(item.balance)}</span>
      </div>
    `).join("");
  }

  const activityEl = document.getElementById("recentActivity");
  const recent = state.audit.slice(0, 8);
  if (!recent.length) {
    activityEl.className = "stack-list empty-state";
    activityEl.textContent = "Sem movimentações recentes.";
  } else {
    activityEl.className = "stack-list";
    activityEl.innerHTML = recent.map(log => `
      <div class="stack-item">
        <strong>${escapeHtml(log.action)}</strong>
        <span>${formatDateTime(log.created_at)} • ${escapeHtml(log.summary || "")}</span>
      </div>
    `).join("");
  }
}

function renderAllTables() {
  renderProductsTable();
  renderInventoryTable();
  renderPurchaseTable();
  renderProductionTable();
  renderShipmentTable();
  renderAuditTable();
}

function renderProductsTable() {
  const q = valueOf("productSearch").toLowerCase();
  const rows = state.products.filter(p =>
    [p.name, p.sku, p.category, ...(p.colors || []), ...(p.sizes || [])].join(" ").toLowerCase().includes(q)
  );
  document.getElementById("productsTable").innerHTML = rows.map(p => `
    <tr>
      <td>${escapeHtml(p.name)}</td>
      <td>${escapeHtml(p.sku || "-")}</td>
      <td>${escapeHtml(p.category || "-")}</td>
      <td>${escapeHtml((p.colors || []).join(", ") || "-")}</td>
      <td>${escapeHtml((p.sizes || []).join(", ") || "-")}</td>
      <td>${statusPill(p.active ? "Ativo" : "Inativo", p.active ? "ok" : "neutral")}</td>
      <td>
        <div class="table-actions">
          <button class="btn secondary xs" onclick="window.editProduct('${p.id}')">Editar</button>
          <button class="btn danger xs" onclick="window.toggleProductActive('${p.id}', ${!p.active})">${p.active ? "Inativar" : "Ativar"}</button>
        </div>
      </td>
    </tr>
  `).join("") || `<tr><td colspan="7" class="empty-state">Nenhum produto encontrado.</td></tr>`;
}

function renderInventoryTable() {
  const q = valueOf("inventorySearch").toLowerCase();
  const lowLimit = Number(state.settings?.low_stock_limit || defaults.low_stock_limit);
  const rows = state.inventory.filter(item =>
    [item.product_name, item.color, item.size].join(" ").toLowerCase().includes(q)
  );
  document.getElementById("inventoryTable").innerHTML = rows.map(item => {
    const balance = Number(item.balance || 0);
    const risk = balance <= 0 ? statusPill("Crítico", "danger") : balance <= lowLimit ? statusPill("Baixo", "warn") : statusPill("OK", "ok");
    return `
      <tr>
        <td>${escapeHtml(item.product_name)}</td>
        <td>${escapeHtml(item.color)}</td>
        <td>${escapeHtml(item.size)}</td>
        <td>${balance}</td>
        <td>${money(item.avg_unit_cost || 0)}</td>
        <td>${risk}</td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="6" class="empty-state">Sem saldo calculado.</td></tr>`;
}

function renderPurchaseTable() {
  const q = valueOf("poSearch").toLowerCase();
  const rows = state.purchaseOrders.filter(item =>
    [item.code, item.supplier_name, item.products?.name, item.color, item.size, item.status].join(" ").toLowerCase().includes(q)
  );
  document.getElementById("purchaseTable").innerHTML = rows.map(item => `
    <tr>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.supplier_name)}</td>
      <td>${escapeHtml(item.products?.name || "-")}</td>
      <td>${escapeHtml(item.color)} / ${escapeHtml(item.size)}</td>
      <td>${Number(item.quantity)}</td>
      <td>${statusPill(labelStatus(item.status), statusKind(item.status))}</td>
      <td>${money((item.quantity || 0) * (item.unit_cost || 0))}</td>
      <td>
        <div class="table-actions">
          <button class="btn secondary xs" onclick="window.advancePurchaseOrder('${item.id}')">Avançar status</button>
        </div>
      </td>
    </tr>
  `).join("") || `<tr><td colspan="8" class="empty-state">Nenhum pedido de compra.</td></tr>`;
}

function renderProductionTable() {
  const q = valueOf("productionSearch").toLowerCase();
  const rows = state.productions.filter(item =>
    [item.batch_code, item.products?.name, item.responsible_name, item.color, item.size, item.status].join(" ").toLowerCase().includes(q)
  );
  document.getElementById("productionTable").innerHTML = rows.map(item => `
    <tr>
      <td>${escapeHtml(item.batch_code)}</td>
      <td>${escapeHtml(item.products?.name || "-")}</td>
      <td>${escapeHtml(item.responsible_name || "-")}</td>
      <td>${escapeHtml(item.color)} / ${escapeHtml(item.size)}</td>
      <td>${Number(item.planned_qty || 0)}</td>
      <td>${Number(item.done_qty || 0)}</td>
      <td>${statusPill(labelStatus(item.status), statusKind(item.status))}</td>
      <td>
        <div class="table-actions">
          <button class="btn secondary xs" onclick="window.advanceProduction('${item.id}')">Avançar status</button>
          <button class="btn primary xs" onclick="window.completeProductionToStock('${item.id}')">Lançar no estoque</button>
        </div>
      </td>
    </tr>
  `).join("") || `<tr><td colspan="8" class="empty-state">Nenhuma ordem de produção.</td></tr>`;
}

function renderShipmentTable() {
  const q = valueOf("shipmentSearch").toLowerCase();
  const rows = state.shipments.filter(item =>
    [item.shipping_code, item.channel, item.mode, item.status, item.order_reference].join(" ").toLowerCase().includes(q)
  );
  document.getElementById("shipmentTable").innerHTML = rows.map(item => `
    <tr>
      <td>${formatDateTime(item.created_at)}</td>
      <td>${escapeHtml(item.shipping_code)}</td>
      <td>${escapeHtml(item.channel)}</td>
      <td>${escapeHtml(item.mode)}</td>
      <td>${statusPill(labelStatus(item.status), statusKind(item.status))}</td>
      <td>${escapeHtml(item.order_reference || "-")}</td>
    </tr>
  `).join("") || `<tr><td colspan="6" class="empty-state">Nenhuma leitura registrada.</td></tr>`;
}

function renderAuditTable() {
  const q = valueOf("auditSearch").toLowerCase();
  const rows = state.audit.filter(item =>
    [item.action, item.table_name, item.summary, item.user_email].join(" ").toLowerCase().includes(q)
  );
  document.getElementById("auditTable").innerHTML = rows.map(item => `
    <tr>
      <td>${formatDateTime(item.created_at)}</td>
      <td>${escapeHtml(item.action)}</td>
      <td>${escapeHtml(item.table_name)}</td>
      <td>${escapeHtml(item.summary || "-")}</td>
      <td>${escapeHtml(item.user_email || "-")}</td>
    </tr>
  `).join("") || `<tr><td colspan="5" class="empty-state">Nenhum log disponível.</td></tr>`;
}

async function saveProduct(e) {
  e.preventDefault();
  const payload = {
    name: valueOf("productName"),
    sku: valueOf("productSku"),
    category: valueOf("productCategory"),
    colors: listOf("productColors"),
    sizes: listOf("productSizes"),
    notes: valueOf("productNotes"),
    active: document.getElementById("productActive").value === "true",
    updated_by: state.profile.id,
  };
  const id = valueOf("productId");
  let error;
  if (id) {
    ({ error } = await supabaseClient.from("products").update(payload).eq("id", id));
    if (!error) await addAudit("UPDATE", "products", `Produto atualizado: ${payload.name}`);
  } else {
    payload.created_by = state.profile.id;
    ({ error } = await supabaseClient.from("products").insert(payload));
    if (!error) await addAudit("INSERT", "products", `Produto criado: ${payload.name}`);
  }
  if (error) return toast(error.message, "error");
  clearProductForm();
  await loadProducts();
  populateProductSelects();
  renderDashboard();
  renderProductsTable();
  toast("Produto salvo com sucesso.");
}

window.editProduct = function(id) {
  const item = state.products.find(p => p.id === id);
  if (!item) return;
  setValue("productId", item.id);
  setValue("productName", item.name);
  setValue("productSku", item.sku || "");
  setValue("productCategory", item.category || "");
  setValue("productColors", (item.colors || []).join(", "));
  setValue("productSizes", (item.sizes || []).join(", "));
  setValue("productNotes", item.notes || "");
  document.getElementById("productActive").value = item.active ? "true" : "false";
  setSection("products");
};

window.toggleProductActive = async function(id, nextValue) {
  const item = state.products.find(p => p.id === id);
  if (!item) return;
  const { error } = await supabaseClient.from("products").update({ active: nextValue, updated_by: state.profile.id }).eq("id", id);
  if (error) return toast(error.message, "error");
  await addAudit("UPDATE", "products", `Produto ${nextValue ? "ativado" : "inativado"}: ${item.name}`);
  await loadProducts();
  renderProductsTable();
  renderDashboard();
  toast("Status do produto atualizado.");
};

function clearProductForm() {
  ["productId","productName","productSku","productCategory","productColors","productSizes","productNotes"].forEach(id => setValue(id, ""));
  document.getElementById("productActive").value = "true";
}

async function createInventoryMovement(e) {
  e.preventDefault();
  const movement_type = valueOf("inventoryType");
  const qty = Number(valueOf("inventoryQty"));
  const payload = {
    product_id: valueOf("inventoryProduct"),
    color: valueOf("inventoryColor"),
    size: valueOf("inventorySize"),
    movement_type,
    quantity: qty,
    unit_cost: Number(valueOf("inventoryCost") || 0),
    notes: valueOf("inventoryNotes"),
    created_by: state.profile.id,
  };
  const { error } = await supabaseClient.from("inventory_movements").insert(payload);
  if (error) return toast(error.message, "error");
  await addAudit("INSERT", "inventory_movements", `${movement_type} de ${qty} unidades`);
  document.getElementById("inventoryForm").reset();
  await loadInventory();
  await loadAudit();
  renderDashboard();
  renderInventoryTable();
  toast("Movimentação registrada.");
}

async function createPurchaseOrder(e) {
  e.preventDefault();
  const payload = {
    code: `PC-${Date.now().toString().slice(-6)}`,
    supplier_name: valueOf("poSupplier"),
    product_id: valueOf("poProduct"),
    color: valueOf("poColor"),
    size: valueOf("poSize"),
    quantity: Number(valueOf("poQty")),
    unit_cost: Number(valueOf("poUnitCost") || 0),
    status: valueOf("poStatus"),
    notes: valueOf("poNotes"),
    created_by: state.profile.id,
    updated_by: state.profile.id,
  };
  const { error } = await supabaseClient.from("purchase_orders").insert(payload);
  if (error) return toast(error.message, "error");
  await addAudit("INSERT", "purchase_orders", `Pedido de compra criado: ${payload.code}`);
  document.getElementById("purchaseForm").reset();
  await loadPurchaseOrders();
  await loadAudit();
  renderDashboard();
  renderPurchaseTable();
  toast("Pedido de compra criado.");
}

window.advancePurchaseOrder = async function(id) {
  const item = state.purchaseOrders.find(p => p.id === id);
  if (!item) return;
  const order = ["aberto","enviado","recebido"];
  const next = order[Math.min(order.indexOf(item.status) + 1, order.length - 1)];
  const { error } = await supabaseClient.from("purchase_orders").update({ status: next, updated_by: state.profile.id }).eq("id", id);
  if (error) return toast(error.message, "error");

  if (next === "recebido") {
    const { error: stockError } = await supabaseClient.from("inventory_movements").insert({
      product_id: item.product_id,
      color: item.color,
      size: item.size,
      movement_type: "entrada",
      quantity: item.quantity,
      unit_cost: item.unit_cost || 0,
      notes: `Entrada automática do ${item.code}`,
      created_by: state.profile.id,
    });
    if (stockError) return toast(stockError.message, "error");
  }

  await addAudit("UPDATE", "purchase_orders", `Status de ${item.code} para ${next}`);
  await loadPurchaseOrders();
  await loadInventory();
  await loadAudit();
  renderDashboard();
  renderPurchaseTable();
  renderInventoryTable();
  toast("Status do pedido de compra atualizado.");
};

async function createProductionOrder(e) {
  e.preventDefault();
  const payload = {
    product_id: valueOf("productionProduct"),
    batch_code: valueOf("productionBatch"),
    responsible_name: valueOf("productionResponsible"),
    status: valueOf("productionStatus"),
    color: valueOf("productionColor"),
    size: valueOf("productionSize"),
    planned_qty: Number(valueOf("productionPlannedQty")),
    done_qty: Number(valueOf("productionDoneQty")),
    notes: valueOf("productionNotes"),
    created_by: state.profile.id,
    updated_by: state.profile.id,
  };
  const { error } = await supabaseClient.from("production_orders").insert(payload);
  if (error) return toast(error.message, "error");
  await addAudit("INSERT", "production_orders", `Ordem de produção criada: ${payload.batch_code}`);
  document.getElementById("productionForm").reset();
  await loadProductions();
  await loadAudit();
  renderDashboard();
  renderProductionTable();
  toast("Ordem de produção salva.");
}

window.advanceProduction = async function(id) {
  const item = state.productions.find(p => p.id === id);
  if (!item) return;
  const order = ["aberto","em_andamento","finalizado"];
  const next = order[Math.min(order.indexOf(item.status) + 1, order.length - 1)];
  const { error } = await supabaseClient.from("production_orders").update({ status: next, updated_by: state.profile.id }).eq("id", id);
  if (error) return toast(error.message, "error");
  await addAudit("UPDATE", "production_orders", `Status de ${item.batch_code} para ${next}`);
  await loadProductions();
  await loadAudit();
  renderDashboard();
  renderProductionTable();
  toast("Status da produção atualizado.");
};

window.completeProductionToStock = async function(id) {
  const item = state.productions.find(p => p.id === id);
  if (!item) return toast("Ordem não encontrada.", "error");
  const qty = Number(item.done_qty || 0);
  if (qty <= 0) return toast("Informe quantidade pronta maior que zero.", "error");

  const { error } = await supabaseClient.from("inventory_movements").insert({
    product_id: item.product_id,
    color: item.color,
    size: item.size,
    movement_type: "entrada",
    quantity: qty,
    unit_cost: 0,
    notes: `Entrada automática da produção ${item.batch_code}`,
    created_by: state.profile.id,
  });
  if (error) return toast(error.message, "error");

  await supabaseClient.from("production_orders").update({ status: "finalizado", updated_by: state.profile.id }).eq("id", id);
  await addAudit("INSERT", "inventory_movements", `Produção ${item.batch_code} lançada no estoque`);
  await loadProductions();
  await loadInventory();
  await loadAudit();
  renderDashboard();
  renderProductionTable();
  renderInventoryTable();
  toast("Produção lançada no estoque.");
};

async function createShipment(e) {
  e.preventDefault();
  const code = valueOf("shipmentCode");
  const info = classifyShipment(code, valueOf("shipmentMode"));
  const payload = {
    shipping_code: code,
    order_reference: valueOf("shipmentOrderRef"),
    channel: info.channel,
    mode: info.mode,
    status: valueOf("shipmentStatus"),
    notes: valueOf("shipmentNotes"),
    created_by: state.profile.id,
  };
  const { error } = await supabaseClient.from("shipments").insert(payload);
  if (error) return toast(error.message, "error");
  await addAudit("INSERT", "shipments", `Leitura de expedição registrada: ${code}`);
  document.getElementById("shipmentForm").reset();
  document.getElementById("shipmentChannel").value = "";
  setShipmentFeedback({ tone: "neutral", message: "Digite ou bipa um código para classificar automaticamente a expedição." });
  await loadShipments();
  await loadAudit();
  renderDashboard();
  renderShipmentTable();
  toast("Expedição registrada.");
}

function classifyShipment(code, mode) {
  const cleaned = (code || "").trim().toUpperCase();
  const finalMode = mode === "auto" ? detectMode(cleaned) : mode;
  let channel = "Manual";
  if (cleaned.startsWith("AN")) channel = "Mercado Envios";
  else if (cleaned.startsWith("BR")) channel = "Shopee";
  else if (!cleaned) channel = "";
  else channel = "Manual";

  const message =
    !cleaned ? "Digite ou bipa um código para classificar automaticamente a expedição." :
    channel === "Mercado Envios" ? "Código identificado como Mercado Envios. Pode jogar na caixa correta." :
    channel === "Shopee" && finalMode === "shopee_direta" ? "Código Shopee Direta identificado." :
    channel === "Shopee" ? "Código Shopee identificado. Verifique coleta/agência." :
    "Código não reconhecido automaticamente. Revise o modo.";

  const tone =
    channel === "Mercado Envios" ? "warn" :
    channel === "Shopee" ? "info" :
    cleaned ? "warn" : "neutral";

  return { channel, mode: finalMode, tone, message };
}

function detectMode(code) {
  if (code.startsWith("AN")) return "mercado_envios";
  if (code.startsWith("BR")) return "shopee_agencia";
  return "manual";
}

function setShipmentFeedback(info) {
  const box = document.getElementById("shipmentFeedback");
  box.className = `callout ${info.tone}`;
  box.textContent = info.message;
}

async function saveSettings(e) {
  e.preventDefault();
  const payload = {
    company_name: valueOf("settingsCompanyName"),
    workspace_id: valueOf("settingsWorkspaceId"),
    low_stock_limit: Number(valueOf("settingsLowStock")),
    accent_color: valueOf("settingsAccent"),
    notes: valueOf("settingsNotes"),
    updated_by: state.profile.id,
  };
  const { error, data } = await supabaseClient
    .from("app_settings")
    .update(payload)
    .eq("id", state.settings.id)
    .select()
    .single();
  if (error) return toast(error.message, "error");
  state.settings = data;
  applyAccentColor();
  applyUserContext();
  await addAudit("UPDATE", "app_settings", "Configurações gerais atualizadas");
  toast("Configurações salvas.");
}

async function saveProfile(e) {
  e.preventDefault();
  const full_name = valueOf("profileName");
  const { error, data } = await supabaseClient
    .from("profiles")
    .update({ full_name })
    .eq("id", state.profile.id)
    .select()
    .single();
  if (error) return toast(error.message, "error");
  state.profile = data;
  applyUserContext();
  await addAudit("UPDATE", "profiles", "Perfil atualizado");
  toast("Perfil atualizado.");
}

async function addAudit(action, table_name, summary) {
  if (!state.profile || !state.session) return;
  await supabaseClient.from("audit_logs").insert({
    action,
    table_name,
    summary,
    user_id: state.profile.id,
    user_email: state.session.user.email,
  });
}

function populateSettingsForms() {
  const settings = state.settings || defaults;
  setValue("settingsCompanyName", settings.company_name || defaults.company_name);
  setValue("settingsWorkspaceId", settings.workspace_id || defaults.workspace_id);
  setValue("settingsLowStock", settings.low_stock_limit || defaults.low_stock_limit);
  setValue("settingsAccent", settings.accent_color || defaults.accent_color);
  setValue("settingsNotes", settings.notes || "");

  setValue("profileName", state.profile?.full_name || "");
  setValue("profileEmail", state.profile?.email || state.session?.user?.email || "");
  setValue("profileRole", state.profile?.role || "OPER");
}

function applyUserContext() {
  const companyName = state.settings?.company_name || defaults.company_name;
  const workspaceId = state.settings?.workspace_id || defaults.workspace_id;
  document.getElementById("companyNameSidebar").textContent = companyName;
  document.getElementById("workspaceChip").textContent = `Workspace: ${workspaceId}`;
  document.getElementById("userName").textContent = state.profile?.full_name || "Usuário";
  document.getElementById("userMeta").textContent = `${state.profile?.role || "OPER"} • ${state.session?.user?.email || ""}`;
  document.getElementById("userAvatar").textContent = initials(state.profile?.full_name || state.session?.user?.email || "SA");
}

function setSection(key) {
  Object.entries(sectionMap).forEach(([name, cfg]) => {
    document.getElementById(cfg.el).classList.toggle("hidden", name !== key);
    document.querySelector(`.nav-btn[data-section="${name}"]`)?.classList.toggle("active", name === key);
  });
  document.getElementById("pageTitle").textContent = sectionMap[key].title;
  document.getElementById("pageSubtitle").textContent = sectionMap[key].subtitle;
}

function showAuth() {
  document.getElementById("authScreen").classList.remove("hidden");
  document.getElementById("appShell").classList.add("hidden");
}

function showApp() {
  document.getElementById("authScreen").classList.add("hidden");
  document.getElementById("appShell").classList.remove("hidden");
}

function resetState() {
  state.profile = null;
  state.products = [];
  state.inventory = [];
  state.purchaseOrders = [];
  state.productions = [];
  state.shipments = [];
  state.audit = [];
}

function toggleTheme() {
  document.body.classList.toggle("dark");
  localStorage.setItem("sa_theme", document.body.classList.contains("dark") ? "dark" : "light");
}

(function restoreTheme() {
  if (localStorage.getItem("sa_theme") === "dark") document.body.classList.add("dark");
})();

function applyAccentColor() {
  const color = state.settings?.accent_color || defaults.accent_color;
  document.documentElement.style.setProperty("--primary", color);
}

function listOf(id) {
  return valueOf(id).split(",").map(v => v.trim()).filter(Boolean);
}
function valueOf(id) {
  return document.getElementById(id)?.value?.trim() || "";
}
function setValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value ?? "";
}
function initials(text) {
  return text.split(" ").filter(Boolean).slice(0,2).map(t => t[0].toUpperCase()).join("") || "SA";
}
function money(value) {
  return Number(value || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}
function formatDateTime(value) {
  if (!value) return "-";
  return new Date(value).toLocaleString("pt-BR");
}
function toast(message, type = "success") {
  const el = document.getElementById("toast");
  el.textContent = message;
  el.classList.remove("hidden");
  el.style.background = type === "error"
    ? "linear-gradient(135deg,#ef4444,#dc2626)"
    : "linear-gradient(135deg,var(--primary),var(--primary-2))";
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => el.classList.add("hidden"), 3200);
}
function escapeHtml(text) {
  return String(text ?? "").replace(/[&<>"']/g, (m) => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[m]));
}
function statusPill(text, kind = "neutral") {
  return `<span class="status-pill status-${kind}">${escapeHtml(text)}</span>`;
}
function labelStatus(status) {
  const map = {
    aberto: "Aberto",
    enviado: "Enviado",
    recebido: "Recebido",
    cancelado: "Cancelado",
    em_andamento: "Em andamento",
    finalizado: "Finalizado",
    separado: "Separado",
    embalado: "Embalado",
    despachado: "Despachado",
  };
  return map[status] || status || "-";
}
function statusKind(status) {
  if (["recebido","finalizado","despachado"].includes(status)) return "ok";
  if (["cancelado"].includes(status)) return "danger";
  if (["aberto","enviado","em_andamento","separado","embalado"].includes(status)) return "warn";
  return "neutral";
}
(function () {
  "use strict";

  const STORAGE_KEY = "bakeryApp";
  const CURRENCY_KEYS = new Set([
    "totalRevenue",
    "totalProductionCost",
    "grossProfitTotal",
    "totalStockValueCost"
  ]);
  const METRIC_ANIMATION_MS = 760;
  const DRAWER_BREAKPOINT = 900;

  let bakeryApp = loadBakeryApp();
  let activeTab = "dashboard";
  let statusTimer = null;
  const metricPrevious = {};
  const metricFrameHandles = {};

  const el = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheElements();
    bindEvents();
    setLogoFallback();
    ensureActiveSheet();
    setDefaultSaleDateIfEmpty();
    renderAll(true);
    registerServiceWorker();
  }

  function cacheElements() {
    el.appShell = document.getElementById("appShell");
    el.drawerBackdrop = document.getElementById("drawerBackdrop");
    el.sidebarToggleBtn = document.getElementById("sidebarToggleBtn");
    el.sheetList = document.getElementById("sheetList");
    el.newSheetBtn = document.getElementById("newSheetBtn");
    el.renameSheetBtn = document.getElementById("renameSheetBtn");
    el.deleteSheetBtn = document.getElementById("deleteSheetBtn");
    el.activeSheetName = document.getElementById("activeSheetName");
    el.activeSheetMeta = document.getElementById("activeSheetMeta");
    el.statusMessage = document.getElementById("statusMessage");
    el.tabButtons = Array.from(document.querySelectorAll(".tab-btn"));
    el.dashboardTab = document.getElementById("dashboardTab");
    el.inventoryTab = document.getElementById("inventoryTab");
    el.salesTab = document.getElementById("salesTab");
    el.metricEls = Array.from(document.querySelectorAll("[data-metric]"));
    el.runSystemCheckBtn = document.getElementById("runSystemCheckBtn");
    el.systemCheckResult = document.getElementById("systemCheckResult");

    el.addProductForm = document.getElementById("addProductForm");
    el.productName = document.getElementById("productName");
    el.unitCost = document.getElementById("unitCost");
    el.unitPrice = document.getElementById("unitPrice");
    el.openingStock = document.getElementById("openingStock");
    el.inventoryTableBody = document.getElementById("inventoryTableBody");

    el.addSaleForm = document.getElementById("addSaleForm");
    el.saleDate = document.getElementById("saleDate");
    el.saleItem = document.getElementById("saleItem");
    el.saleQty = document.getElementById("saleQty");
    el.salesTableBody = document.getElementById("salesTableBody");
  }

  function bindEvents() {
    el.sidebarToggleBtn.addEventListener("click", toggleDrawer);
    el.drawerBackdrop.addEventListener("click", closeDrawer);
    window.addEventListener("resize", () => {
      if (window.innerWidth > DRAWER_BREAKPOINT) {
        closeDrawer();
      }
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closeDrawer();
      }
    });

    el.newSheetBtn.addEventListener("click", createNewSheet);
    el.renameSheetBtn.addEventListener("click", renameActiveSheet);
    el.deleteSheetBtn.addEventListener("click", deleteActiveSheet);
    el.sheetList.addEventListener("click", handleSheetClick);

    el.tabButtons.forEach((button) => {
      button.addEventListener("click", () => setActiveTab(button.dataset.tab));
    });

    el.addProductForm.addEventListener("submit", onAddProduct);
    el.inventoryTableBody.addEventListener("change", onInventoryChange);
    el.inventoryTableBody.addEventListener("click", onInventoryActionClick);

    el.addSaleForm.addEventListener("submit", onAddSale);
    el.salesTableBody.addEventListener("click", onSalesActionClick);

    el.runSystemCheckBtn.addEventListener("click", onRunSystemCheck);
  }

  function setLogoFallback() {
    const logo = document.getElementById("brandLogo");
    const fallback = document.getElementById("brandLogoFallback");

    if (!logo || !fallback) {
      return;
    }

    const sourceCandidates = [];
    const initialSrc = logo.getAttribute("src");
    if (initialSrc) {
      sourceCandidates.push(initialSrc);
    }

    ["assets/logo.png", "assets/logo.webp", "logo.png", "logo.webp"].forEach((path) => {
      if (!sourceCandidates.includes(path)) {
        sourceCandidates.push(path);
      }
    });

    let sourceIndex = 0;

    function loadNextSource() {
      if (sourceIndex >= sourceCandidates.length) {
        logo.style.display = "none";
        fallback.style.display = "grid";
        return;
      }

      logo.src = sourceCandidates[sourceIndex];
      sourceIndex += 1;
    }

    logo.addEventListener("error", loadNextSource);

    logo.addEventListener("load", () => {
      logo.style.display = "block";
      fallback.style.display = "none";
    });

    loadNextSource();
  }

  function loadBakeryApp() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return createDefaultAppState();
      }
      const parsed = JSON.parse(raw);
      return normalizeAppState(parsed);
    } catch (error) {
      return createDefaultAppState();
    }
  }

  function saveBakeryApp() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(bakeryApp));
  }

  function createDefaultAppState() {
    const initialSheet = createSheet("Main Stock Sheet");
    return {
      activeSheetId: initialSheet.id,
      sheets: [initialSheet]
    };
  }

  function normalizeAppState(value) {
    if (!value || typeof value !== "object") {
      return createDefaultAppState();
    }

    const rawSheets = Array.isArray(value.sheets) ? value.sheets : [];
    const normalizedSheets = rawSheets
      .map((sheet, index) => normalizeSheet(sheet, index))
      .filter(Boolean);

    if (!normalizedSheets.length) {
      return createDefaultAppState();
    }

    const normalizedActiveId = typeof value.activeSheetId === "string" ? value.activeSheetId : "";
    const activeExists = normalizedSheets.some((sheet) => sheet.id === normalizedActiveId);

    return {
      activeSheetId: activeExists ? normalizedActiveId : normalizedSheets[0].id,
      sheets: normalizedSheets
    };
  }

  function normalizeSheet(sheet, index) {
    if (!sheet || typeof sheet !== "object") {
      return null;
    }

    const nowISO = new Date().toISOString();
    const products = Array.isArray(sheet.products) ? sheet.products : [];
    const sales = Array.isArray(sheet.sales) ? sheet.sales : [];

    return {
      id: typeof sheet.id === "string" && sheet.id.trim() ? sheet.id : createId(),
      name: typeof sheet.name === "string" && sheet.name.trim() ? sheet.name.trim() : `Stock Sheet ${index + 1}`,
      createdAtISO: isISODateTime(sheet.createdAtISO) ? sheet.createdAtISO : nowISO,
      updatedAtISO: isISODateTime(sheet.updatedAtISO) ? sheet.updatedAtISO : nowISO,
      products: products.map(normalizeProduct).filter(Boolean),
      sales: sales.map(normalizeSale).filter(Boolean)
    };
  }

  function normalizeProduct(product) {
    if (!product || typeof product !== "object") {
      return null;
    }

    return {
      name: typeof product.name === "string" ? product.name.trim() : "",
      unitCost: toNumberOrZero(product.unitCost),
      unitPrice: toNumberOrZero(product.unitPrice),
      openingStock: toIntegerOrZero(product.openingStock),
      stockOut: toIntegerOrZero(product.stockOut)
    };
  }

  function normalizeSale(sale) {
    if (!sale || typeof sale !== "object") {
      return null;
    }

    return {
      id: typeof sale.id === "string" && sale.id.trim() ? sale.id : createId(),
      dateISO: isISODate(sale.dateISO) ? sale.dateISO : todayISODate(),
      itemName: typeof sale.itemName === "string" ? sale.itemName.trim() : "",
      qtySold: toIntegerOrZero(sale.qtySold),
      revenue: toNumberOrZero(sale.revenue),
      productionCost: toNumberOrZero(sale.productionCost),
      grossProfit: toNumberOrZero(sale.grossProfit)
    };
  }

  function ensureActiveSheet() {
    if (!Array.isArray(bakeryApp.sheets)) {
      bakeryApp = createDefaultAppState();
      saveBakeryApp();
      return;
    }

    if (!bakeryApp.sheets.length) {
      const replacement = createSheet("Main Stock Sheet");
      bakeryApp.sheets.push(replacement);
      bakeryApp.activeSheetId = replacement.id;
      saveBakeryApp();
      return;
    }

    const hasActive = bakeryApp.sheets.some((sheet) => sheet.id === bakeryApp.activeSheetId);
    if (!hasActive) {
      bakeryApp.activeSheetId = bakeryApp.sheets[0].id;
      saveBakeryApp();
    }
  }

  function getActiveSheet() {
    ensureActiveSheet();
    return bakeryApp.sheets.find((sheet) => sheet.id === bakeryApp.activeSheetId) || null;
  }

  function createSheet(name) {
    const nowISO = new Date().toISOString();
    return {
      id: createId(),
      name: name || `Stock Sheet ${bakeryApp.sheets.length + 1}`,
      createdAtISO: nowISO,
      updatedAtISO: nowISO,
      products: [],
      sales: []
    };
  }

  function touchSheet(sheet) {
    sheet.updatedAtISO = new Date().toISOString();
  }

  function createId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }
    return `id_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  }

  function recalcActiveSheet() {
    const emptyTotals = {
      totalItemsAvailable: 0,
      totalItemsSold: 0,
      totalRevenue: 0,
      totalProductionCost: 0,
      grossProfitTotal: 0,
      totalStockValueCost: 0
    };

    const sheet = getActiveSheet();
    if (!sheet) {
      return {
        sheet: null,
        products: [],
        sales: [],
        totals: emptyTotals,
        validation: {
          isValid: false,
          errors: ["No active stock sheet."]
        }
      };
    }

    const errors = [];
    const products = [];
    const sales = [];
    const totals = { ...emptyTotals };
    const nameIndex = new Map();
    const soldQtyByName = new Map();

    sheet.sales.forEach((sale, saleIndex) => {
      const itemName = typeof sale.itemName === "string" ? sale.itemName.trim() : "";
      const qtySold = toIntegerOrZero(sale.qtySold);
      const revenue = toNumberOrZero(sale.revenue);
      const productionCost = toNumberOrZero(sale.productionCost);
      const grossProfit = toNumberOrZero(sale.grossProfit);

      if (!itemName) {
        errors.push(`Sale #${saleIndex + 1}: item name is required.`);
      }

      if (!Number.isInteger(qtySold) || qtySold < 1) {
        errors.push(`Sale #${saleIndex + 1}: qty sold must be an integer >= 1.`);
      }

      if (
        !Number.isFinite(revenue) ||
        !Number.isFinite(productionCost) ||
        !Number.isFinite(grossProfit) ||
        revenue < 0 ||
        productionCost < 0
      ) {
        errors.push(`Sale #${saleIndex + 1}: monetary values must be valid non-negative numbers.`);
      }

      const normalizedName = normalizeName(itemName);
      if (normalizedName) {
        soldQtyByName.set(normalizedName, (soldQtyByName.get(normalizedName) || 0) + qtySold);
      }

      totals.totalItemsSold += qtySold;
      totals.totalRevenue += revenue;
      totals.totalProductionCost += productionCost;
      totals.grossProfitTotal += grossProfit;

      sales.push({
        id: sale.id,
        dateISO: isISODate(sale.dateISO) ? sale.dateISO : todayISODate(),
        itemName,
        qtySold,
        revenue,
        productionCost,
        grossProfit
      });
    });

    sheet.products.forEach((product, productIndex) => {
      const name = typeof product.name === "string" ? product.name.trim() : "";
      const unitCost = toNumberOrZero(product.unitCost);
      const unitPrice = toNumberOrZero(product.unitPrice);
      const openingStock = toIntegerOrZero(product.openingStock);
      const stockOut = toIntegerOrZero(product.stockOut);
      const stockAvailable = openingStock - stockOut;
      const stockValueCost = stockAvailable * unitCost;

      const normalizedName = normalizeName(name);
      if (!normalizedName) {
        errors.push(`Product #${productIndex + 1}: name is required.`);
      } else if (nameIndex.has(normalizedName)) {
        errors.push(`Duplicate product name found: "${name}".`);
      } else {
        nameIndex.set(normalizedName, productIndex);
      }

      if (!Number.isFinite(unitCost) || unitCost < 0) {
        errors.push(`Product "${name || `#${productIndex + 1}`}": unit cost must be >= 0.`);
      }

      if (!Number.isFinite(unitPrice) || unitPrice < 0) {
        errors.push(`Product "${name || `#${productIndex + 1}`}": unit price must be >= 0.`);
      }

      if (!Number.isInteger(openingStock) || openingStock < 0) {
        errors.push(`Product "${name || `#${productIndex + 1}`}": opening stock must be an integer >= 0.`);
      }

      if (!Number.isInteger(stockOut) || stockOut < 0) {
        errors.push(`Product "${name || `#${productIndex + 1}`}": stock out must be an integer >= 0.`);
      }

      if (stockAvailable < 0) {
        errors.push(`Product "${name || `#${productIndex + 1}`}": stock available cannot be negative.`);
      }

      const totalSoldForProduct = soldQtyByName.get(normalizedName) || 0;
      if (stockOut < totalSoldForProduct) {
        errors.push(`Product "${name || `#${productIndex + 1}`}": stock out is lower than logged sales.`);
      }

      totals.totalItemsAvailable += stockAvailable;
      totals.totalStockValueCost += stockValueCost;

      products.push({
        name,
        unitCost,
        unitPrice,
        openingStock,
        stockOut,
        stockAvailable,
        stockValueCost
      });
    });

    sales.forEach((sale, saleIndex) => {
      const exists = nameIndex.has(normalizeName(sale.itemName));
      if (!exists) {
        errors.push(`Sale #${saleIndex + 1}: "${sale.itemName}" does not match any product.`);
      }
    });

    totals.totalRevenue = roundCurrency(totals.totalRevenue);
    totals.totalProductionCost = roundCurrency(totals.totalProductionCost);
    totals.grossProfitTotal = roundCurrency(totals.grossProfitTotal);
    totals.totalStockValueCost = roundCurrency(totals.totalStockValueCost);
    totals.totalItemsSold = Math.round(totals.totalItemsSold);
    totals.totalItemsAvailable = Math.round(totals.totalItemsAvailable);

    Object.keys(totals).forEach((key) => {
      if (!Number.isFinite(totals[key])) {
        errors.push(`Total "${key}" is invalid (NaN or infinite).`);
      }
    });

    return {
      sheet,
      products,
      sales,
      totals,
      validation: {
        isValid: errors.length === 0,
        errors
      }
    };
  }
  function renderAll(isInitialRender) {
    ensureActiveSheet();
    const computed = recalcActiveSheet();
    renderSidebar();
    renderTopBar(computed);
    renderTabs();
    renderDashboard(computed, isInitialRender);
    renderInventory(computed);
    renderSales(computed);
  }

  function renderSidebar() {
    const fragment = document.createDocumentFragment();
    const orderedSheets = [...bakeryApp.sheets].sort((a, b) => {
      return new Date(b.updatedAtISO).getTime() - new Date(a.updatedAtISO).getTime();
    });

    orderedSheets.forEach((sheet) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `sheet-item${sheet.id === bakeryApp.activeSheetId ? " active" : ""}`;
      button.dataset.sheetId = sheet.id;

      const title = document.createElement("span");
      title.className = "sheet-title";
      title.textContent = sheet.name;

      const date = document.createElement("span");
      date.className = "sheet-date";
      date.textContent = `Created ${formatHumanDate(sheet.createdAtISO)}`;

      button.appendChild(title);
      button.appendChild(date);
      fragment.appendChild(button);
    });

    el.sheetList.innerHTML = "";
    el.sheetList.appendChild(fragment);
  }

  function renderTopBar(computed) {
    if (!computed.sheet) {
      el.activeSheetName.textContent = "No Sheet";
      el.activeSheetMeta.textContent = "";
      return;
    }

    el.activeSheetName.textContent = computed.sheet.name;
    el.activeSheetMeta.textContent = `Created ${formatHumanDate(
      computed.sheet.createdAtISO
    )} • Updated ${formatHumanDateTime(computed.sheet.updatedAtISO)}`;
  }

  function renderTabs() {
    el.tabButtons.forEach((button) => {
      const tab = button.dataset.tab;
      button.classList.toggle("active", tab === activeTab);
    });

    el.dashboardTab.classList.toggle("active", activeTab === "dashboard");
    el.inventoryTab.classList.toggle("active", activeTab === "inventory");
    el.salesTab.classList.toggle("active", activeTab === "sales");
  }

  function renderDashboard(computed, isInitialRender) {
    const totals = computed.totals;
    const metrics = {
      totalRevenue: totals.totalRevenue,
      totalProductionCost: totals.totalProductionCost,
      grossProfitTotal: totals.grossProfitTotal,
      totalItemsSold: totals.totalItemsSold,
      totalItemsAvailable: totals.totalItemsAvailable,
      totalStockValueCost: totals.totalStockValueCost
    };

    Object.entries(metrics).forEach(([key, value]) => {
      const metricEl = el.metricEls.find((node) => node.dataset.metric === key);
      if (!metricEl) {
        return;
      }

      animateMetric(
        metricEl,
        key,
        value,
        CURRENCY_KEYS.has(key) ? formatCurrency : formatInteger,
        isInitialRender
      );
    });

    const grossValueEl = el.metricEls.find((node) => node.dataset.metric === "grossProfitTotal");
    if (grossValueEl) {
      grossValueEl.classList.remove("profit-positive", "profit-negative", "profit-neutral");
      if (totals.grossProfitTotal > 0) {
        grossValueEl.classList.add("profit-positive");
      } else if (totals.grossProfitTotal < 0) {
        grossValueEl.classList.add("profit-negative");
      } else {
        grossValueEl.classList.add("profit-neutral");
      }
    }
  }

  function renderInventory(computed) {
    const products = computed.products;

    if (!products.length) {
      el.inventoryTableBody.innerHTML = `<tr><td class="empty-row" colspan="8">No products yet. Add your first product above.</td></tr>`;
      return;
    }

    const fragment = document.createDocumentFragment();

    products.forEach((product, index) => {
      const row = document.createElement("tr");
      if (product.stockAvailable <= 3) {
        row.classList.add("low-stock");
      }

      row.innerHTML = `
        <td>
          <input class="table-input" type="text" data-action="edit-product" data-field="name" data-index="${index}" value="${escapeForAttr(product.name)}" maxlength="80">
        </td>
        <td>
          <input class="table-input" type="number" min="0" step="0.01" data-action="edit-product" data-field="unitCost" data-index="${index}" value="${formatNumberForInput(product.unitCost)}">
        </td>
        <td>
          <input class="table-input" type="number" min="0" step="0.01" data-action="edit-product" data-field="unitPrice" data-index="${index}" value="${formatNumberForInput(product.unitPrice)}">
        </td>
        <td>
          <input class="table-input" type="number" min="0" step="1" data-action="edit-product" data-field="openingStock" data-index="${index}" value="${product.openingStock}">
        </td>
        <td><span class="readonly-pill">${product.stockOut}</span></td>
        <td><span class="readonly-pill">${product.stockAvailable}</span></td>
        <td>${formatCurrency(product.stockValueCost)}</td>
        <td><button type="button" class="btn small btn-danger" data-action="delete-product" data-index="${index}">Delete</button></td>
      `;

      fragment.appendChild(row);
    });

    el.inventoryTableBody.innerHTML = "";
    el.inventoryTableBody.appendChild(fragment);
  }

  function renderSales(computed) {
    renderSaleItemOptions(computed.products);

    if (!computed.sales.length) {
      el.salesTableBody.innerHTML = `<tr><td class="empty-row" colspan="7">No sales logged yet.</td></tr>`;
      return;
    }

    const salesSorted = [...computed.sales].sort((a, b) => {
      return new Date(b.dateISO).getTime() - new Date(a.dateISO).getTime();
    });

    const fragment = document.createDocumentFragment();
    salesSorted.forEach((sale) => {
      const row = document.createElement("tr");
      const grossClass = sale.grossProfit > 0 ? "gross-positive" : sale.grossProfit < 0 ? "gross-negative" : "";
      row.innerHTML = `
        <td>${formatHumanDate(sale.dateISO)}</td>
        <td>${escapeForHtml(sale.itemName)}</td>
        <td>${sale.qtySold}</td>
        <td>${formatCurrency(sale.revenue)}</td>
        <td>${formatCurrency(sale.productionCost)}</td>
        <td class="${grossClass}">${formatCurrency(sale.grossProfit)}</td>
        <td><button type="button" class="btn small btn-danger" data-action="delete-sale" data-sale-id="${sale.id}">Delete</button></td>
      `;
      fragment.appendChild(row);
    });

    el.salesTableBody.innerHTML = "";
    el.salesTableBody.appendChild(fragment);
  }

  function renderSaleItemOptions(products) {
    const previous = el.saleItem.value;
    el.saleItem.innerHTML = "";

    const saleSubmitButton = el.addSaleForm.querySelector("button[type='submit']");

    if (!products.length) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "Add products first";
      option.selected = true;
      el.saleItem.appendChild(option);
      el.saleItem.disabled = true;
      el.saleQty.disabled = true;
      saleSubmitButton.disabled = true;
      return;
    }

    products.forEach((product) => {
      const option = document.createElement("option");
      option.value = product.name;
      option.textContent = `${product.name} (${product.stockAvailable} available)`;
      el.saleItem.appendChild(option);
    });

    const stillExists = products.some((product) => product.name === previous);
    if (stillExists) {
      el.saleItem.value = previous;
    }

    el.saleItem.disabled = false;
    el.saleQty.disabled = false;
    saleSubmitButton.disabled = false;
  }

  function animateMetric(targetEl, key, targetValue, formatter, immediate) {
    const startValue = key in metricPrevious ? metricPrevious[key] : 0;
    const endValue = Number.isFinite(targetValue) ? targetValue : 0;

    if (metricFrameHandles[key]) {
      cancelAnimationFrame(metricFrameHandles[key]);
      metricFrameHandles[key] = null;
    }

    if (immediate || Math.abs(startValue - endValue) < 0.0001) {
      targetEl.textContent = formatter(endValue);
      metricPrevious[key] = endValue;
      return;
    }

    let startTime = null;

    function step(timestamp) {
      if (!startTime) {
        startTime = timestamp;
      }

      const elapsed = timestamp - startTime;
      const progress = Math.min(elapsed / METRIC_ANIMATION_MS, 1);
      const eased = easeOutCubic(progress);
      const nextValue = startValue + (endValue - startValue) * eased;

      targetEl.textContent = formatter(nextValue);

      if (progress < 1) {
        metricFrameHandles[key] = requestAnimationFrame(step);
      } else {
        metricPrevious[key] = endValue;
      }
    }

    metricFrameHandles[key] = requestAnimationFrame(step);
  }

  function onAddProduct(event) {
    event.preventDefault();
    const sheet = getActiveSheet();
    if (!sheet) {
      return;
    }

    const name = el.productName.value.trim();
    const unitCost = Number(el.unitCost.value);
    const unitPrice = Number(el.unitPrice.value);
    const openingStock = Number(el.openingStock.value);

    if (!name) {
      showStatus("Product name is required.", "error");
      return;
    }

    if (sheet.products.some((product) => normalizeName(product.name) === normalizeName(name))) {
      showStatus("Product names must be unique within a sheet.", "error");
      return;
    }

    if (!isNonNegativeNumber(unitCost) || !isNonNegativeNumber(unitPrice)) {
      showStatus("Unit cost and unit price must be numbers >= 0.", "error");
      return;
    }

    if (!Number.isInteger(openingStock) || openingStock < 0) {
      showStatus("Opening stock must be an integer >= 0.", "error");
      return;
    }

    sheet.products.push({
      name,
      unitCost: roundCurrency(unitCost),
      unitPrice: roundCurrency(unitPrice),
      openingStock,
      stockOut: 0
    });

    touchSheet(sheet);
    saveBakeryApp();
    renderAll(false);
    el.addProductForm.reset();
    showStatus(`Product "${name}" added.`, "success");
  }

  function onInventoryChange(event) {
    const input = event.target.closest("[data-action='edit-product']");
    if (!input) {
      return;
    }

    const index = Number(input.dataset.index);
    const field = input.dataset.field;
    if (!Number.isInteger(index) || !field) {
      return;
    }

    const sheet = getActiveSheet();
    if (!sheet || !sheet.products[index]) {
      return;
    }

    const product = sheet.products[index];

    if (field === "name") {
      const oldName = product.name;
      const nextName = input.value.trim();

      if (!nextName) {
        showStatus("Product name cannot be empty.", "error");
        renderAll(false);
        return;
      }

      const duplicate = sheet.products.some((item, i) => {
        return i !== index && normalizeName(item.name) === normalizeName(nextName);
      });

      if (duplicate) {
        showStatus("Duplicate product name is not allowed.", "error");
        renderAll(false);
        return;
      }

      product.name = nextName;
      if (normalizeName(oldName) !== normalizeName(nextName)) {
        sheet.sales.forEach((sale) => {
          if (normalizeName(sale.itemName) === normalizeName(oldName)) {
            sale.itemName = nextName;
          }
        });
      }

      touchSheet(sheet);
      saveBakeryApp();
      renderAll(false);
      showStatus(`Renamed "${oldName}" to "${nextName}".`, "success");
      return;
    }

    const numericValue = Number(input.value);
    if (!Number.isFinite(numericValue) || numericValue < 0) {
      showStatus("Values must be valid numbers and cannot be negative.", "error");
      renderAll(false);
      return;
    }

    if (field === "openingStock") {
      if (!Number.isInteger(numericValue)) {
        showStatus("Opening stock must be an integer value.", "error");
        renderAll(false);
        return;
      }

      if (numericValue < product.stockOut) {
        showStatus("Opening stock cannot be lower than stock out.", "error");
        renderAll(false);
        return;
      }

      product.openingStock = numericValue;
    } else if (field === "unitCost" || field === "unitPrice") {
      product[field] = roundCurrency(numericValue);
    } else {
      return;
    }

    touchSheet(sheet);
    saveBakeryApp();
    renderAll(false);
  }

  function onInventoryActionClick(event) {
    const deleteBtn = event.target.closest("[data-action='delete-product']");
    if (!deleteBtn) {
      return;
    }

    const index = Number(deleteBtn.dataset.index);
    if (!Number.isInteger(index)) {
      return;
    }

    const sheet = getActiveSheet();
    if (!sheet || !sheet.products[index]) {
      return;
    }

    const product = sheet.products[index];
    const hasSales = sheet.sales.some((sale) => normalizeName(sale.itemName) === normalizeName(product.name));

    if (hasSales) {
      showStatus("Delete related sales first before deleting this product.", "error");
      return;
    }

    sheet.products.splice(index, 1);
    touchSheet(sheet);
    saveBakeryApp();
    renderAll(false);
    showStatus(`Product "${product.name}" deleted.`, "success");
  }

  function onAddSale(event) {
    event.preventDefault();
    const sheet = getActiveSheet();
    if (!sheet) {
      return;
    }

    const dateISO = el.saleDate.value || todayISODate();
    const itemName = el.saleItem.value.trim();
    const qtySold = Number(el.saleQty.value);

    if (!isISODate(dateISO)) {
      showStatus("Please choose a valid sale date.", "error");
      return;
    }

    if (!itemName) {
      showStatus("Please select an item.", "error");
      return;
    }

    if (!Number.isInteger(qtySold) || qtySold < 1) {
      showStatus("Quantity sold must be an integer >= 1.", "error");
      return;
    }

    const computed = recalcActiveSheet();
    const computedProduct = computed.products.find(
      (product) => normalizeName(product.name) === normalizeName(itemName)
    );

    if (!computedProduct) {
      showStatus("Selected product no longer exists.", "error");
      return;
    }

    if (qtySold > computedProduct.stockAvailable) {
      showStatus(
        `Cannot add sale. ${computedProduct.name} has only ${computedProduct.stockAvailable} in stock.`,
        "error"
      );
      return;
    }

    const product = sheet.products.find((p) => normalizeName(p.name) === normalizeName(itemName));

    if (!product) {
      showStatus("Product lookup failed while adding sale.", "error");
      return;
    }

    const revenue = roundCurrency(qtySold * product.unitPrice);
    const productionCost = roundCurrency(qtySold * product.unitCost);
    const grossProfit = roundCurrency(revenue - productionCost);

    product.stockOut += qtySold;
    sheet.sales.push({
      id: createId(),
      dateISO,
      itemName: product.name,
      qtySold,
      revenue,
      productionCost,
      grossProfit
    });

    touchSheet(sheet);
    saveBakeryApp();
    renderAll(false);
    el.saleQty.value = "";
    setDefaultSaleDateIfEmpty();
    showStatus(`Sale logged for ${qtySold} x ${product.name}.`, "success");
  }

  function onSalesActionClick(event) {
    const deleteBtn = event.target.closest("[data-action='delete-sale']");
    if (!deleteBtn) {
      return;
    }

    const saleId = deleteBtn.dataset.saleId;
    if (!saleId) {
      return;
    }

    const sheet = getActiveSheet();
    if (!sheet) {
      return;
    }

    const saleIndex = sheet.sales.findIndex((sale) => sale.id === saleId);
    if (saleIndex < 0) {
      return;
    }

    const sale = sheet.sales[saleIndex];
    const product = sheet.products.find(
      (item) => normalizeName(item.name) === normalizeName(sale.itemName)
    );

    if (!product) {
      showStatus("Sale cannot be deleted because the related product is missing.", "error");
      return;
    }

    if (product.stockOut - sale.qtySold < 0) {
      showStatus("Sale cannot be deleted due to invalid stock rollback.", "error");
      return;
    }

    product.stockOut -= sale.qtySold;
    sheet.sales.splice(saleIndex, 1);
    touchSheet(sheet);
    saveBakeryApp();
    renderAll(false);
    showStatus("Sale deleted and stock rolled back.", "success");
  }

  function onRunSystemCheck() {
    const result = runSystemCheck();
    el.systemCheckResult.classList.remove("pass", "fail");

    if (result.passed) {
      el.systemCheckResult.classList.add("pass");
      el.systemCheckResult.textContent = `PASS\n${result.messages.join("\n")}`;
    } else {
      el.systemCheckResult.classList.add("fail");
      el.systemCheckResult.textContent = `FAIL\n${result.messages.join("\n")}`;
    }
  }

  function runSystemCheck() {
    const computed = recalcActiveSheet();
    const messages = [];
    const tolerance = 0.01;

    const rowTotals = {
      totalRevenue: roundCurrency(sumBy(computed.sales, "revenue")),
      totalProductionCost: roundCurrency(sumBy(computed.sales, "productionCost")),
      grossProfitTotal: roundCurrency(sumBy(computed.sales, "grossProfit")),
      totalItemsSold: Math.round(sumBy(computed.sales, "qtySold")),
      totalItemsAvailable: Math.round(sumBy(computed.products, "stockAvailable")),
      totalStockValueCost: roundCurrency(sumBy(computed.products, "stockValueCost"))
    };

    Object.keys(rowTotals).forEach((key) => {
      const expected = rowTotals[key];
      const actual = computed.totals[key];
      const mismatch = CURRENCY_KEYS.has(key)
        ? Math.abs(expected - actual) > tolerance
        : expected !== actual;

      if (mismatch) {
        messages.push(`- ${key} mismatch. Expected ${expected}, got ${actual}.`);
      }
    });

    const hasNegativeStock = computed.products.some((product) => product.stockAvailable < 0);
    if (hasNegativeStock) {
      messages.push("- Negative stock available detected.");
    }

    if (!computed.validation.isValid) {
      computed.validation.errors.forEach((error) => {
        messages.push(`- ${error}`);
      });
    }

    if (!messages.length) {
      messages.push("- All totals and invariants are valid.");
    }

    return {
      passed: messages.length === 1 && messages[0] === "- All totals and invariants are valid.",
      messages
    };
  }

  function handleSheetClick(event) {
    const sheetButton = event.target.closest(".sheet-item");
    if (!sheetButton) {
      return;
    }

    const sheetId = sheetButton.dataset.sheetId;
    if (!sheetId || sheetId === bakeryApp.activeSheetId) {
      closeDrawer();
      return;
    }

    const exists = bakeryApp.sheets.some((sheet) => sheet.id === sheetId);
    if (!exists) {
      return;
    }

    bakeryApp.activeSheetId = sheetId;
    saveBakeryApp();
    renderAll(false);
    closeDrawer();
  }

  function createNewSheet() {
    const fallbackName = `Stock Sheet ${bakeryApp.sheets.length + 1}`;
    const input = window.prompt("New stock sheet name:", fallbackName);
    if (input === null) {
      return;
    }

    const name = input.trim() || fallbackName;
    const sheet = createSheet(name);
    bakeryApp.sheets.unshift(sheet);
    bakeryApp.activeSheetId = sheet.id;
    activeTab = "dashboard";
    saveBakeryApp();
    renderAll(false);
    showStatus(`Created "${name}".`, "success");
    closeDrawer();
  }

  function renameActiveSheet() {
    const sheet = getActiveSheet();
    if (!sheet) {
      return;
    }

    const input = window.prompt("Rename stock sheet:", sheet.name);
    if (input === null) {
      return;
    }

    const nextName = input.trim();
    if (!nextName) {
      showStatus("Stock sheet name cannot be empty.", "error");
      return;
    }

    sheet.name = nextName;
    touchSheet(sheet);
    saveBakeryApp();
    renderAll(false);
    showStatus("Stock sheet renamed.", "success");
  }

  function deleteActiveSheet() {
    const sheet = getActiveSheet();
    if (!sheet) {
      return;
    }

    bakeryApp.sheets = bakeryApp.sheets.filter((item) => item.id !== sheet.id);
    ensureActiveSheet();
    bakeryApp.activeSheetId = bakeryApp.sheets[0].id;
    saveBakeryApp();
    renderAll(false);
    showStatus(`Deleted "${sheet.name}".`, "success");
  }

  function setActiveTab(tab) {
    const validTabs = new Set(["dashboard", "inventory", "sales"]);
    if (!validTabs.has(tab)) {
      return;
    }
    activeTab = tab;
    renderTabs();
  }

  function toggleDrawer() {
    el.appShell.classList.toggle("drawer-open");
  }

  function closeDrawer() {
    el.appShell.classList.remove("drawer-open");
  }

  function showStatus(message, type) {
    clearTimeout(statusTimer);
    el.statusMessage.textContent = message;
    el.statusMessage.classList.remove("error", "success");
    if (type === "error" || type === "success") {
      el.statusMessage.classList.add(type);
    }

    statusTimer = setTimeout(() => {
      el.statusMessage.textContent = "";
      el.statusMessage.classList.remove("error", "success");
    }, 3800);
  }

  function setDefaultSaleDateIfEmpty() {
    if (!el.saleDate.value) {
      el.saleDate.value = todayISODate();
    }
  }

  function registerServiceWorker() {
    if (!("serviceWorker" in navigator)) {
      return;
    }

    if (window.location.protocol === "file:") {
      return;
    }

    window.addEventListener("load", () => {
      navigator.serviceWorker.register("./service-worker.js").catch(() => {
        // Keep UI silent if service worker registration fails.
      });
    });
  }

  function sumBy(items, key) {
    return items.reduce((sum, item) => {
      const value = Number(item[key]);
      if (!Number.isFinite(value)) {
        return sum;
      }
      return sum + value;
    }, 0);
  }

  function toNumberOrZero(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
  }

  function toIntegerOrZero(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
      return 0;
    }
    return Math.trunc(number);
  }

  function roundCurrency(value) {
    return Math.round((value + Number.EPSILON) * 100) / 100;
  }

  function formatCurrency(value) {
    const rounded = roundCurrency(value);
    const sign = rounded < 0 ? "-" : "";
    const abs = Math.abs(rounded);
    return `${sign}₦${abs.toLocaleString("en-NG", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    })}`;
  }

  function formatInteger(value) {
    const rounded = Math.round(value);
    return rounded.toLocaleString("en-NG");
  }

  function formatHumanDate(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "Invalid Date";
    }
    return date.toLocaleDateString("en-NG", {
      day: "2-digit",
      month: "short",
      year: "numeric"
    });
  }

  function formatHumanDateTime(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "Invalid Date";
    }
    return date.toLocaleString("en-NG", {
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    });
  }

  function todayISODate() {
    const now = new Date();
    const offsetMs = now.getTimezoneOffset() * 60000;
    return new Date(now.getTime() - offsetMs).toISOString().slice(0, 10);
  }

  function normalizeName(name) {
    return String(name || "").trim().toLowerCase();
  }

  function isNonNegativeNumber(value) {
    return Number.isFinite(value) && value >= 0;
  }

  function isISODate(value) {
    if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return false;
    }
    const date = new Date(`${value}T00:00:00`);
    return !Number.isNaN(date.getTime());
  }

  function isISODateTime(value) {
    if (typeof value !== "string") {
      return false;
    }
    const date = new Date(value);
    return !Number.isNaN(date.getTime());
  }

  function formatNumberForInput(number) {
    const rounded = roundCurrency(number);
    return String(rounded);
  }

  function escapeForHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function escapeForAttr(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function easeOutCubic(x) {
    return 1 - Math.pow(1 - x, 3);
  }
})();

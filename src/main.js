import ExcelJS from 'exceljs/dist/exceljs.min.js';
import logoAdd from './assets/logo-add.png';
import logoSolutions from './assets/logo-solutions.png';
import logoBio from './assets/logo-bio.png';

// Division configurations
const DIVISIONS = {
  add: {
    id: 'add',
    name: 'Food Additives',
    address: '243 Horreya Ave, Sportin Alexandria , Egypt',
    website: 'www.awa-foodadd.net',
    phone: '(+203) 4292030 - 4211200',
    fax: '+203 4292060',
    logo: logoAdd,
    themeColor: '#3B4E87',
    sheetName: 'ADD',
    defaultItems: [
      { desc: 'Nisin', origin: 'Netherland', qty: '', price: '' },
      { desc: 'Natamysin', origin: 'Netherland', qty: '', price: '' },
      { desc: 'Rennet', origin: 'Turky', qty: '', price: '' },
      { desc: 'Natural red color ', origin: 'egypt', qty: '', price: '' }
    ],
    defaultTerms: [
      '1.The above price including vat 14%',
      "2.Quotation validity /  2days from the quotation's date."
    ]
  },
  solutions: {
    id: 'solutions',
    name: 'Food Solutions',
    address: 'Area 11,12,13 - Block 15, 4th Industrial Zone, New Borg El-Arab',
    website: 'www.awa-foodsolutions.com',
    phone: '+203 5890155 / 165',
    fax: '+203 5890145',
    logo: logoSolutions,
    themeColor: '#3B4E87',
    sheetName: 'solutions',
    defaultItems: [
      { desc: '', origin: 'Netherland', qty: '', price: '' },
      { desc: '', origin: 'Netherland', qty: '', price: '' },
      { desc: '', origin: 'Turky', qty: '', price: '' },
      { desc: '', origin: 'egypt', qty: '', price: '' }
    ],
    defaultTerms: [
      '1.The above price including vat 14%',
      "2.Quotation validity /  2days from the quotation's date."
    ]
  },
  bio: {
    id: 'bio',
    name: 'Bio Ingredients',
    address: '243 Horreya Ave, Sportin Alexandria , Egypt',
    website: 'awa-bioingredients.com',
    phone: '(+203) 4292030 - 4211200',
    fax: '+203 4292060',
    logo: logoBio,
    themeColor: '#3B4E87',
    sheetName: 'Bio',
    defaultItems: [],
    defaultTerms: [
      '1.The above price including vat 14%',
      "2.Quotation validity /  2days from the quotation's date."
    ]
  }
};

// Global App State
const quoteData = {
  preparedBy: '',
  quoteNo: '',
  customerId: '123',
  date: '',
  validUntil: '',
  validityDays: 3,
  currency: 'EGP',
  vatRate: 14,
  vatIncluded: false,
  itemCount: 12,
  unit: 'kg',
  customerName: '[Company Name]',
  customerContact: '[contact name]',
  customerPhone: '[Phone]',
  lines: [], // Array of variable size
  terms: [], // Array of terms lines (3 items)
  footerContact: 'If you have any questions about this price quote, please contact',
  footerThanks: 'Thank You For Your Business!'
};

let currentDiv = 'add';
let materialsList = [];
let activeRowEditIdx = 0; // The active row accordion index expanded in the sidebar

// Dom elements cache
const DOM = {
  divSelectors: document.querySelectorAll('.btn-div'),
  inputPreparedBy: document.getElementById('input-prepared-by'),
  inputQuoteNo: document.getElementById('input-quote-no'),
  inputDate: document.getElementById('input-date'),
  inputCustomerId: document.getElementById('input-customer-id'),
  inputCustomerName: document.getElementById('input-customer-name'),
  inputCustomerContact: document.getElementById('input-customer-contact'),
  inputCustomerPhone: document.getElementById('input-customer-phone'),
  inputCurrency: document.getElementById('input-currency'),
  inputValidity: document.getElementById('input-validity'),
  inputVat: document.getElementById('input-vat'),
  inputItemCount: document.getElementById('input-item-count'),
  inputUnit: document.getElementById('input-unit'),
  customUnitGroup: document.getElementById('custom-unit-group'),
  inputCustomUnit: document.getElementById('input-custom-unit'),
  checkboxVatIncluded: document.getElementById('checkbox-vat-included'),
  inputFooterContact: document.getElementById('input-footer-contact'),
  inputFooterThanks: document.getElementById('input-footer-thanks'),
  inputTerm1: document.getElementById('input-term-1'),
  inputTerm2: document.getElementById('input-term-2'),
  inputTerm3: document.getElementById('input-term-3'),
  
  sidebarItemsAccordion: document.getElementById('sidebar-items-accordion'),

  btnNew: document.getElementById('btn-new'),
  btnPrint: document.getElementById('btn-print'),
  btnShare: document.getElementById('btn-share'),
  btnExport: document.getElementById('btn-export'),
  
  // Sheet preview elements
  sheetLogo: document.getElementById('sheet-logo'),
  companyAddress: document.getElementById('company-address'),
  companyWebsite: document.getElementById('company-website'),
  companyPhone: document.getElementById('company-phone'),
  companyFax: document.getElementById('company-fax'),
  sheetPreparedBy: document.getElementById('sheet-prepared-by'),
  sheetDate: document.getElementById('sheet-date'),
  sheetQuoteNo: document.getElementById('sheet-quote-no'),
  sheetCustomerId: document.getElementById('sheet-customer-id'),
  sheetValidUntil: document.getElementById('sheet-valid-until'),
  sheetCustomerName: document.getElementById('sheet-customer-name'),
  sheetCustomerContact: document.getElementById('sheet-customer-contact'),
  sheetCustomerPhone: document.getElementById('sheet-customer-phone'),
  sheetCurrency: document.getElementById('sheet-currency'),
  itemsBody: document.getElementById('items-body'),
  sheetSubtotal: document.getElementById('sheet-subtotal'),
  sheetVatAmount: document.getElementById('sheet-vat-amount'),
  sheetVatPercent: document.getElementById('sheet-vat-percent'),
  sheetTotal: document.getElementById('sheet-total'),
  vatCalcRow: document.getElementById('vat-calc-row'),
  termsList: document.getElementById('terms-list'),
  sheetFooterContact: document.getElementById('sheet-footer-contact'),
  sheetFooterThanks: document.getElementById('sheet-footer-thanks'),
  toast: document.getElementById('toast'),
  headerQty: document.getElementById('header-qty'),
  headerPrice: document.getElementById('header-price')
};

// Initialize Application
function init() {
  // Set default dates
  const today = new Date();
  quoteData.date = formatDate(today);
  quoteData.validUntil = formatDate(addDays(today, quoteData.validityDays));
  
  DOM.sheetDate.textContent = quoteData.date;
  DOM.sheetValidUntil.textContent = quoteData.validUntil;

  // Set default values in controls
  DOM.inputPreparedBy.value = '';
  quoteData.preparedBy = '[salesperson name]';
  DOM.sheetPreparedBy.textContent = quoteData.preparedBy;
  
  DOM.inputQuoteNo.value = 'Q' + Math.floor(100000 + Math.random() * 900000);
  quoteData.quoteNo = DOM.inputQuoteNo.value;
  DOM.sheetQuoteNo.textContent = quoteData.quoteNo;

  DOM.inputCustomerId.value = quoteData.customerId;
  DOM.inputCustomerName.value = quoteData.customerName;
  DOM.inputCustomerContact.value = quoteData.customerContact;
  DOM.inputCustomerPhone.value = quoteData.customerPhone;
  DOM.inputCurrency.value = quoteData.currency;
  DOM.inputValidity.value = quoteData.validityDays;
  DOM.inputVat.value = quoteData.vatRate;
  DOM.inputItemCount.value = quoteData.itemCount;
  DOM.inputUnit.value = quoteData.unit;
  DOM.checkboxVatIncluded.checked = quoteData.vatIncluded;
  
  DOM.inputDate.value = quoteData.date;
  DOM.inputFooterContact.value = quoteData.footerContact;
  DOM.inputFooterThanks.value = quoteData.footerThanks;

  DOM.sheetCustomerName.textContent = quoteData.customerName;
  DOM.sheetCustomerContact.textContent = quoteData.customerContact;
  DOM.sheetCustomerPhone.textContent = quoteData.customerPhone;
  DOM.sheetFooterContact.textContent = quoteData.footerContact;
  DOM.sheetFooterThanks.textContent = quoteData.footerThanks;

  // Event Listeners for controls
  setupControlListeners();

  // Setup button actions
  DOM.btnNew.addEventListener('click', () => resetQuote(currentDiv));
  DOM.btnPrint.addEventListener('click', () => window.print());
  DOM.btnShare.addEventListener('click', sharePDF);
  DOM.btnExport.addEventListener('click', exportToExcel);

  // Switch to default division Additives
  switchDivision('add');

  // Load materials for autocomplete
  fetch('/materials.json')
    .then(res => res.json())
    .then(data => {
      materialsList = data;
      console.log(`Loaded ${materialsList.length} materials for autocomplete.`);
    })
    .catch(err => console.error("Error loading materials list:", err));
}

// Helpers
function formatDate(date) {
  const d = new Date(date);
  let month = '' + (d.getMonth() + 1);
  let day = '' + d.getDate();
  const year = d.getFullYear();

  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;

  return [year, month, day].join('-');
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + parseInt(days));
  return result;
}

function showToast(msg) {
  DOM.toast.textContent = msg;
  DOM.toast.classList.add('show');
  setTimeout(() => {
    DOM.toast.classList.remove('show');
  }, 3000);
}

// Setup Event Listeners for UI Side Panel
function setupControlListeners() {
  // Division select
  DOM.divSelectors.forEach(btn => {
    btn.addEventListener('click', (e) => {
      const divId = btn.getAttribute('data-div');
      switchDivision(divId);
    });
  });

  // Sidebar Inputs -> bind to quoteData and Sheet Preview
  DOM.inputPreparedBy.addEventListener('input', (e) => {
    quoteData.preparedBy = e.target.value || '[salesperson name]';
    DOM.sheetPreparedBy.textContent = quoteData.preparedBy;
  });

  DOM.inputQuoteNo.addEventListener('input', (e) => {
    quoteData.quoteNo = e.target.value;
    DOM.sheetQuoteNo.textContent = quoteData.quoteNo;
  });

  DOM.inputDate.addEventListener('input', (e) => {
    quoteData.date = e.target.value;
    DOM.sheetDate.textContent = quoteData.date;
    
    // Recalculate Validity
    const dateVal = new Date(quoteData.date);
    const newValidUntil = formatDate(addDays(dateVal, quoteData.validityDays));
    quoteData.validUntil = newValidUntil;
    DOM.sheetValidUntil.textContent = newValidUntil;
  });

  DOM.inputCustomerId.addEventListener('input', (e) => {
    quoteData.customerId = e.target.value;
    DOM.sheetCustomerId.textContent = quoteData.customerId;
  });

  DOM.inputCustomerName.addEventListener('input', (e) => {
    quoteData.customerName = e.target.value || '[Company Name]';
    DOM.sheetCustomerName.textContent = quoteData.customerName;
  });

  DOM.inputCustomerContact.addEventListener('input', (e) => {
    quoteData.customerContact = e.target.value || '[contact name]';
    DOM.sheetCustomerContact.textContent = quoteData.customerContact;
  });

  DOM.inputCustomerPhone.addEventListener('input', (e) => {
    quoteData.customerPhone = e.target.value || '[Phone]';
    DOM.sheetCustomerPhone.textContent = quoteData.customerPhone;
  });

  DOM.inputCurrency.addEventListener('input', (e) => {
    quoteData.currency = e.target.value;
    DOM.sheetCurrency.textContent = quoteData.currency;
  });

  DOM.inputValidity.addEventListener('input', (e) => {
    const days = parseInt(e.target.value) || 0;
    quoteData.validityDays = days;
    const dateVal = new Date(quoteData.date);
    const newValidUntil = formatDate(addDays(dateVal, days));
    quoteData.validUntil = newValidUntil;
    DOM.sheetValidUntil.textContent = newValidUntil;
  });

  DOM.inputVat.addEventListener('input', (e) => {
    quoteData.vatRate = parseFloat(e.target.value) || 0;
    DOM.sheetVatPercent.textContent = quoteData.vatRate;
    calculateTotals();
  });

  DOM.checkboxVatIncluded.addEventListener('change', (e) => {
    quoteData.vatIncluded = e.target.checked;
    calculateTotals();
  });

  // Handle Item Count changes dynamically
  DOM.inputItemCount.addEventListener('input', (e) => {
    let count = parseInt(e.target.value) || 12;
    if (count < 1) count = 1;
    if (count > 30) count = 30;
    quoteData.itemCount = count;
    resizeQuoteLines(count);
  });

  // Handle Measurement Unit changes
  DOM.inputUnit.addEventListener('change', (e) => {
    updateMeasurementUnit(e.target.value);
  });

  DOM.inputCustomUnit.addEventListener('input', (e) => {
    if (DOM.inputUnit.value === 'custom') {
      updateMeasurementUnit('custom');
    }
  });

  DOM.inputFooterContact.addEventListener('input', (e) => {
    quoteData.footerContact = e.target.value || '';
    DOM.sheetFooterContact.textContent = quoteData.footerContact;
  });

  DOM.inputFooterThanks.addEventListener('input', (e) => {
    quoteData.footerThanks = e.target.value || '';
    DOM.sheetFooterThanks.textContent = quoteData.footerThanks;
  });

  // Handle Terms edits from the sidebar
  [DOM.inputTerm1, DOM.inputTerm2, DOM.inputTerm3].forEach((input, index) => {
    input.addEventListener('input', (e) => {
      quoteData.terms[index] = e.target.value;
      renderSheetTerms();
    });
  });
}

// Load default terms
function resetTermsToDefault(defaultTerms) {
  quoteData.terms = [...defaultTerms];
  
  DOM.inputTerm1.value = quoteData.terms[0] || '';
  DOM.inputTerm2.value = quoteData.terms[1] || '';
  DOM.inputTerm3.value = quoteData.terms[2] || '';
  
  renderSheetTerms();
}

function renderSheetTerms() {
  DOM.termsList.innerHTML = '';
  for (let i = 0; i < 3; i++) {
    const val = quoteData.terms[i] || '';
    const item = document.createElement('div');
    item.className = 'terms-item';
    item.textContent = val;
    DOM.termsList.appendChild(item);
  }
}

// Render preview items table (completely read-only)
function renderItemTable() {
  DOM.itemsBody.innerHTML = '';
  
  quoteData.lines.forEach((line, idx) => {
    const row = document.createElement('tr');
    if (idx === activeRowEditIdx) {
      row.classList.add('active-row-edit');
    }
    
    row.innerHTML = `
      <td colspan="2" class="desc-cell-container text-trebuchet">${line.desc || ''}</td>
      <td class="center-align text-trebuchet text-center">${line.origin || ''}</td>
      <td class="center-align text-trebuchet text-center">${line.qty || ''}</td>
      <td class="center-align text-trebuchet text-center">${line.price || ''}</td>
      <td class="amount-cell text-trebuchet amount-val" id="amount-${idx}">0.00</td>
    `;
    
    // Allow clicking a row on A4 sheet to open its accordion editor in the sidebar
    row.addEventListener('click', () => {
      setActiveRowEdit(idx);
    });
    
    DOM.itemsBody.appendChild(row);
  });

  // Update all amounts on sheet
  quoteData.lines.forEach((_, i) => {
    calculateRowAmount(i);
  });
}

// Change the active editing row index
function setActiveRowEdit(idx) {
  activeRowEditIdx = idx;
  
  // Re-render accordion items to update visual expansion state
  renderSidebarItemEditors();
  
  // Refresh highlights on sheet preview
  const rows = DOM.itemsBody.querySelectorAll('tr');
  rows.forEach((row, rIdx) => {
    if (rIdx === idx) {
      row.classList.add('active-row-edit');
      
      // Smoothly scroll the sidebar accordion item into view
      const activeAccordion = DOM.sidebarItemsAccordion.children[idx];
      if (activeAccordion) {
        activeAccordion.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    } else {
      row.classList.remove('active-row-edit');
    }
  });
}

// Render dynamic collapsible item editor list in sidebar
function renderSidebarItemEditors() {
  DOM.sidebarItemsAccordion.innerHTML = '';
  
  quoteData.lines.forEach((line, idx) => {
    const item = document.createElement('div');
    item.className = 'accordion-item';
    if (idx === activeRowEditIdx) {
      item.classList.add('expanded', 'active-sidebar-edit');
    }
    
    const headerText = line.desc ? `${idx + 1}. ${line.desc}` : `Item Row ${idx + 1} (Empty)`;
    
    item.innerHTML = `
      <div class="accordion-header">
        <span class="header-title">${headerText}</span>
        <span class="accordion-arrow">▼</span>
      </div>
      <div class="accordion-content">
        <div class="form-group">
          <label>Description</label>
          <div class="desc-input-wrapper">
            <input type="text" class="sidebar-desc-input" value="${line.desc}" placeholder="Search or type description..." autocomplete="off">
          </div>
        </div>
        <div class="form-group">
          <label>Origin</label>
          <input type="text" class="sidebar-origin-input" value="${line.origin}" placeholder="Origin (e.g. Egypt)">
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>QTY / ${quoteData.unit}</label>
            <input type="text" class="sidebar-qty-input" value="${line.qty}" placeholder="-">
          </div>
          <div class="form-group">
            <label>Price / ${quoteData.unit}</label>
            <input type="text" class="sidebar-price-input" value="${line.price}" placeholder="-">
          </div>
        </div>
      </div>
    `;
    
    // Bind toggle action on header click
    const header = item.querySelector('.accordion-header');
    header.addEventListener('click', () => {
      if (activeRowEditIdx === idx) {
        setActiveRowEdit(-1); // Collapse all
      } else {
        setActiveRowEdit(idx); // Expand clicked
      }
    });
    
    DOM.sidebarItemsAccordion.appendChild(item);
  });
  
  setupSidebarItemListeners();
}

// Bind input change handlers inside active sidebar accordion item
function setupSidebarItemListeners() {
  if (activeRowEditIdx === -1) return;
  
  const activeItem = DOM.sidebarItemsAccordion.children[activeRowEditIdx];
  if (!activeItem) return;
  
  const descInput = activeItem.querySelector('.sidebar-desc-input');
  const originInput = activeItem.querySelector('.sidebar-origin-input');
  const qtyInput = activeItem.querySelector('.sidebar-qty-input');
  const priceInput = activeItem.querySelector('.sidebar-price-input');
  const descWrapper = descInput.parentElement;
  
  const updateState = () => {
    const idx = activeRowEditIdx;
    quoteData.lines[idx].desc = descInput.value;
    quoteData.lines[idx].origin = originInput.value;
    
    // Clean inputs: allow only numbers and decimals
    let qtyVal = qtyInput.value.replace(/[^0-9.]/g, '');
    let priceVal = priceInput.value.replace(/[^0-9.]/g, '');
    
    if (qtyInput.value !== qtyVal) qtyInput.value = qtyVal;
    if (priceInput.value !== priceVal) priceInput.value = priceVal;
    
    quoteData.lines[idx].qty = qtyVal;
    quoteData.lines[idx].price = priceVal;
    
    // Update matching row values on sheet preview dynamically
    const rows = DOM.itemsBody.querySelectorAll('tr');
    if (rows[idx]) {
      rows[idx].querySelector('.desc-cell-container').textContent = descInput.value;
      rows[idx].querySelector('td:nth-child(2)').textContent = originInput.value;
      rows[idx].querySelector('td:nth-child(3)').textContent = qtyVal;
      rows[idx].querySelector('td:nth-child(4)').textContent = priceVal;
    }
    
    calculateRowAmount(idx);
    calculateTotals();
    
    // Update header label inside sidebar accordion
    const headerTitle = activeItem.querySelector('.header-title');
    headerTitle.textContent = descInput.value ? `${idx + 1}. ${descInput.value}` : `Item Row ${idx + 1} (Empty)`;
  };
  
  descInput.addEventListener('input', updateState);
  originInput.addEventListener('input', updateState);
  qtyInput.addEventListener('input', updateState);
  priceInput.addEventListener('input', updateState);

  // Material Auto-complete Dropdown logic inside sidebar
  let activeMatchIdx = -1;
  
  const closeDropdown = () => {
    const existing = descWrapper.querySelector('.autocomplete-dropdown');
    if (existing) existing.remove();
    activeMatchIdx = -1;
  };
  
  const selectItem = (item) => {
    const idx = activeRowEditIdx;
    descInput.value = item.desc;
    quoteData.lines[idx].desc = item.desc;
    
    if (item.origin) {
      originInput.value = item.origin;
      quoteData.lines[idx].origin = item.origin;
    }
    if (item.price) {
      priceInput.value = item.price.toString();
      quoteData.lines[idx].price = item.price.toString();
    }
    
    updateState();
    closeDropdown();
    qtyInput.focus();
  };
  
  descInput.addEventListener('input', (e) => {
    closeDropdown();
    const val = descInput.value.trim().toLowerCase();
    if (!val) return;
    
    // Filter matching materials starting with typed text
    const matches = materialsList.filter(m => 
      m.desc.toLowerCase().startsWith(val)
    ).slice(0, 40);
    
    if (matches.length === 0) return;
    
    const dropdown = document.createElement('div');
    dropdown.className = 'autocomplete-dropdown';
    
    matches.forEach((match, mIdx) => {
      const el = document.createElement('div');
      el.className = 'autocomplete-item';
      el.textContent = match.desc;
      
      el.addEventListener('click', () => selectItem(match));
      dropdown.appendChild(el);
    });
    
    descWrapper.appendChild(dropdown);
  });
  
  // Keyboard bindings for autocomplete navigation
  descInput.addEventListener('keydown', (e) => {
    const dropdown = descWrapper.querySelector('.autocomplete-dropdown');
    if (!dropdown) return;
    
    const items = dropdown.querySelectorAll('.autocomplete-item');
    if (items.length === 0) return;
    
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeMatchIdx = (activeMatchIdx + 1) % items.length;
      highlightItem(items);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeMatchIdx = (activeMatchIdx - 1 + items.length) % items.length;
      highlightItem(items);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeMatchIdx >= 0 && activeMatchIdx < items.length) {
        const selectedText = items[activeMatchIdx].textContent;
        const match = materialsList.find(m => m.desc === selectedText);
        if (match) selectItem(match);
      }
    } else if (e.key === 'Escape') {
      closeDropdown();
    }
  });
  
  function highlightItem(items) {
    items.forEach((item, index) => {
      if (index === activeMatchIdx) {
        item.classList.add('active');
        item.scrollIntoView({ block: 'nearest' });
      } else {
        item.classList.remove('active');
      }
    });
  }
  
  // Dismiss dropdown on document click
  document.addEventListener('click', (e) => {
    if (!descWrapper.contains(e.target)) {
      closeDropdown();
    }
  });
}

// Calculate individual row amount
function calculateRowAmount(idx) {
  const line = quoteData.lines[idx];
  const qty = parseFloat(line.qty) || 0;
  const price = parseFloat(line.price) || 0;
  const amount = qty * price;
  
  const amtCell = document.getElementById(`amount-${idx}`);
  if (amtCell) {
    amtCell.textContent = amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return amount;
}

// Calculate global subtotal, vat, and grand total
function calculateTotals() {
  let subtotal = 0;
  quoteData.lines.forEach((line, idx) => {
    const qty = parseFloat(line.qty) || 0;
    const price = parseFloat(line.price) || 0;
    subtotal += qty * price;
  });

  let vatAmount = 0;
  let total = subtotal;

  if (quoteData.vatRate > 0) {
    if (quoteData.vatIncluded) {
      // VAT is already included in subtotal: subtotal = Base + VAT
      // Base = subtotal / (1 + vatRate/100)
      const base = subtotal / (1 + (quoteData.vatRate / 100));
      vatAmount = subtotal - base;
      DOM.vatCalcRow.style.display = 'table-row';
      DOM.vatCalcRow.querySelector('td:nth-child(2)').textContent = `VAT Included (${quoteData.vatRate}%)`;
    } else {
      // VAT is added on top of subtotal
      vatAmount = subtotal * (quoteData.vatRate / 100);
      total = subtotal + vatAmount;
      DOM.vatCalcRow.style.display = 'table-row';
      DOM.vatCalcRow.querySelector('td:nth-child(2)').textContent = `VAT (${quoteData.vatRate}%)`;
    }
  } else {
    DOM.vatCalcRow.style.display = 'none';
  }

  DOM.sheetSubtotal.textContent = subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  DOM.sheetVatAmount.textContent = vatAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  DOM.sheetTotal.textContent = total.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Reset current quote to defaults
function resetQuote(divId) {
  const confirmReset = confirm("Are you sure you want to initialize a new quote? Current entries will be cleared.");
  if (!confirmReset) return;

  const nextNo = 'Q' + Math.floor(100000 + Math.random() * 900000);
  quoteData.quoteNo = nextNo;
  DOM.inputQuoteNo.value = nextNo;
  DOM.sheetQuoteNo.textContent = nextNo;

  quoteData.customerId = '123';
  DOM.inputCustomerId.value = '123';
  DOM.sheetCustomerId.textContent = '123';

  DOM.inputCustomerName.value = '[Company Name]';
  DOM.inputCustomerContact.value = '[contact name]';
  DOM.inputCustomerPhone.value = '[Phone]';
  DOM.sheetCustomerName.textContent = '[Company Name]';
  DOM.sheetCustomerContact.textContent = '[contact name]';
  DOM.sheetCustomerPhone.textContent = '[Phone]';
  
  quoteData.customerName = '[Company Name]';
  quoteData.customerContact = '[contact name]';
  quoteData.customerPhone = '[Phone]';

  const today = new Date();
  quoteData.date = formatDate(today);
  quoteData.validityDays = 3;
  quoteData.validUntil = formatDate(addDays(today, 3));
  
  DOM.inputDate.value = quoteData.date;
  DOM.sheetDate.textContent = quoteData.date;
  DOM.sheetValidUntil.textContent = quoteData.validUntil;
  DOM.inputValidity.value = 3;
  
  quoteData.itemCount = 12;
  DOM.inputItemCount.value = 12;

  activeRowEditIdx = 0; // Reset active row edit back to first row
  switchDivision(divId);
  showToast("Quotation sheet reset.");
}

// Resize the lines list (preserving current typed inputs)
function resizeQuoteLines(newCount) {
  const diff = newCount - quoteData.lines.length;
  if (diff > 0) {
    for (let i = 0; i < diff; i++) {
      quoteData.lines.push({ desc: '', origin: '', qty: '', price: '' });
    }
  } else if (diff < 0) {
    quoteData.lines.splice(newCount);
  }

  // Adjust activeRowEditIdx if it is now out of bounds
  if (activeRowEditIdx >= newCount) {
    activeRowEditIdx = newCount - 1;
  }
  if (activeRowEditIdx < 0 && newCount > 0) {
    activeRowEditIdx = 0;
  }

  renderItemTable();
  renderSidebarItemEditors();
  calculateTotals();
}

// Switch between division styles & defaults
function switchDivision(divId) {
  currentDiv = divId;
  
  DOM.divSelectors.forEach(btn => {
    if (btn.getAttribute('data-div') === divId) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  const config = DIVISIONS[divId];
  
  // Set division colors and assets
  document.documentElement.style.setProperty('--brand-navy', config.themeColor);
  DOM.sheetLogo.src = config.logo;
  DOM.companyAddress.textContent = config.address;
  DOM.companyWebsite.textContent = config.website;
  DOM.companyPhone.textContent = config.phone;
  DOM.companyFax.textContent = config.fax;

  // Re-generate lines based on the division defaults and current item count
  const newLines = [];
  for (let i = 0; i < quoteData.itemCount; i++) {
    const preset = config.defaultItems[i] || { desc: '', origin: '', qty: '', price: '' };
    newLines.push({ ...preset });
  }
  quoteData.lines = newLines;

  renderItemTable();
  renderSidebarItemEditors();
  resetTermsToDefault(config.defaultTerms);
  updateMeasurementUnit(DOM.inputUnit.value);
  calculateTotals();
}

// Helper to update the measurement unit labels
function updateMeasurementUnit(unitVal) {
  if (unitVal === 'custom') {
    DOM.customUnitGroup.style.display = 'block';
    const customVal = DOM.inputCustomUnit.value.trim() || 'unit';
    quoteData.unit = customVal;
  } else {
    DOM.customUnitGroup.style.display = 'none';
    quoteData.unit = unitVal;
  }

  // Update sheet headers text content
  if (DOM.headerQty && DOM.headerPrice) {
    DOM.headerQty.textContent = `QTY/${quoteData.unit}`;
    DOM.headerPrice.textContent = `Price/${quoteData.unit}`;
  }
  
  // Re-render accordion elements to show correct units
  renderSidebarItemEditors();
}

// Web Share PDF API inside Mobile Devices
function sharePDF() {
  const element = document.getElementById('printable-area');
  
  // Temporarily disable active row edit highlight for clean PDF printout
  const activeRow = DOM.itemsBody.querySelector('.active-row-edit');
  if (activeRow) activeRow.classList.remove('active-row-edit');
  
  const opt = {
    margin:       0.1,
    filename:     `AWA-Quotation-${quoteData.quoteNo || 'Draft'}.pdf`,
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2.5, useCORS: true, logging: false },
    jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' }
  };
  
  showToast("Generating PDF...");
  
  // Generate PDF and try Web Share API
  html2pdf().set(opt).from(element).toPdf().get('pdf').then(function (pdf) {
    // Restore active row highlight
    if (activeRow) activeRow.classList.add('active-row-edit');
    
    const blob = pdf.output('blob');
    const file = new File([blob], opt.filename, { type: 'application/pdf' });
    
    // Check if navigator.share with files is supported
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      navigator.share({
        files: [file],
        title: `AWA Quotation #${quoteData.quoteNo}`,
        text: `Attached is AWA quotation #${quoteData.quoteNo} for your review.`
      })
      .then(() => showToast("Shared successfully!"))
      .catch(err => {
        if (err.name !== 'AbortError') {
          console.error("Error sharing:", err);
          fallbackDownload(opt, element);
        }
      });
    } else {
      // Fallback: download directly
      fallbackDownload(opt, element);
    }
  });
}

function fallbackDownload(opt, element) {
  showToast("PDF downloaded successfully!");
  html2pdf().set(opt).from(element).save();
}

// Convert image to base64
async function getBase64FromUrl(url) {
  const response = await fetch(url);
  const blob = await response.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      resolve(reader.result.split(',')[1]);
    };
    reader.readAsDataURL(blob);
  });
}

// EXCEL EXPORTING LOGIC USING EXCELJS (with dynamic offsets)
async function exportToExcel() {
  const config = DIVISIONS[currentDiv];
  showToast("Exporting Excel...");
  
  const N = quoteData.lines.length;
  
  // Calculate dynamic rows offsets
  const itemStartRow = 17;
  const itemEndRow = 17 + N - 1;
  const subtotalRow = 17 + N;
  const termsHeaderRow = 17 + N + 1;
  const termsLine1 = 17 + N + 2;
  const termsLine2 = 17 + N + 3;
  const termsLine3 = 17 + N + 4;
  const totalRow = 17 + N + 5;
  const footerContactRow = 17 + N + 11;
  const footerThanksRow = 17 + N + 13;

  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(config.sheetName);
    
    // Page setup
    sheet.views = [{ showGridLines: true }];
    
    // Column widths
    sheet.columns = [
      { key: 'A', width: 40.6 },
      { key: 'B', width: 6.0 },
      { key: 'C', width: 13.7 },
      { key: 'D', width: 8.9 },
      { key: 'E', width: 9.9 },
      { key: 'F', width: 19.6 },
      { key: 'G', width: 9.2 },
      { key: 'H', width: 33.0 }
    ];

    // Styles & Helpers
    const fontArialBold = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    const fontTrebuchet = { name: 'Trebuchet MS', size: 10 };
    const fontTrebuchetBold = { name: 'Trebuchet MS', size: 10, bold: true };
    const fillNavy = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B4E87' } };
    
    const alignLeft = { horizontal: 'left', vertical: 'middle' };
    const alignCenter = { horizontal: 'center', vertical: 'middle' };
    const alignRight = { horizontal: 'right', vertical: 'middle' };
    
    const borderThinAll = {
      left: { style: 'thin', color: { indexed: 64 } },
      right: { style: 'thin', color: { indexed: 64 } },
      top: { style: 'thin', color: { indexed: 64 } },
      bottom: { style: 'thin', color: { indexed: 64 } }
    };
    const borderThinLeft = { left: { style: 'thin', color: { indexed: 64 } } };

    // Row heights
    sheet.getRow(1).height = 72;
    sheet.mergeCells('E1:F1');
    const qCell = sheet.getCell('E1');
    qCell.value = 'QUOTATION';
    qCell.font = { name: 'Arial', size: 24, bold: true, color: { argb: 'FF8699C2' } };
    qCell.alignment = alignRight;

    // Insert Image
    try {
      const base64 = await getBase64FromUrl(config.logo);
      const logoId = workbook.addImage({
        base64: base64,
        extension: 'png'
      });
      sheet.addImage(logoId, {
        tl: { col: 0.1, row: 1 },
        ext: { width: 220, height: 50 }
      });
    } catch (e) {
      console.error("Could not load image into Excel:", e);
    }

    // Row 3: Address & Date
    sheet.getCell('A3').value = config.address;
    sheet.getCell('A3').font = fontTrebuchet;
    sheet.getCell('A3').alignment = alignLeft;
    
    sheet.getCell('E3').value = 'DATE';
    sheet.getCell('E3').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('E3').alignment = alignRight;
    
    const dCell = sheet.getCell('F3');
    dCell.value = new Date(quoteData.date);
    dCell.numFmt = 'yyyy-mm-dd';
    dCell.font = fontTrebuchet;
    dCell.alignment = alignCenter;
    dCell.border = borderThinAll;

    // Row 4: Website & Quote #
    sheet.getCell('A4').value = `Website: ${config.website}`;
    sheet.getCell('A4').font = fontTrebuchet;
    sheet.getCell('A4').alignment = alignLeft;
    
    sheet.getCell('E4').value = 'QUOTE ';
    sheet.getCell('E4').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('E4').alignment = alignRight;
    
    const qnCell = sheet.getCell('F4');
    qnCell.value = quoteData.quoteNo;
    qnCell.font = fontTrebuchet;
    qnCell.alignment = alignCenter;
    qnCell.border = { left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    // Row 5: Phone & Customer ID
    sheet.getCell('A5').value = `Phone: ${config.phone}`;
    sheet.getCell('A5').font = fontTrebuchet;
    sheet.getCell('A5').alignment = alignLeft;
    
    sheet.getCell('E5').value = 'CUSTOMER ID';
    sheet.getCell('E5').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('E5').alignment = alignRight;
    
    const cidCell = sheet.getCell('F5');
    cidCell.value = quoteData.customerId;
    cidCell.font = fontTrebuchet;
    cidCell.alignment = alignCenter;
    cidCell.border = { left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    // Row 6: Fax & Valid Until
    sheet.getCell('A6').value = `Fax: ${config.fax}`;
    sheet.getCell('A6').font = fontTrebuchet;
    sheet.getCell('A6').alignment = alignLeft;
    
    sheet.getCell('E6').value = 'VALID UNTIL';
    sheet.getCell('E6').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('E6').alignment = alignRight;
    
    const vuCell = sheet.getCell('F6');
    vuCell.value = { formula: `F3+${quoteData.validityDays}` };
    vuCell.numFmt = 'yyyy-mm-dd';
    vuCell.font = fontTrebuchet;
    vuCell.alignment = alignCenter;
    vuCell.border = { left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    // Row 7: Prepared By
    sheet.getCell('A7').value = `Prepared by: ${quoteData.preparedBy}`;
    sheet.getCell('A7').font = fontTrebuchet;
    sheet.getCell('A7').alignment = alignLeft;

    // Row 9: Customer Header
    sheet.mergeCells('A9:C9');
    const custHeader = sheet.getCell('A9');
    custHeader.value = 'CUSTOMER';
    custHeader.font = fontArialBold;
    custHeader.fill = fillNavy;
    custHeader.alignment = alignLeft;

    // Customer details Rows 10-12
    sheet.getCell('A10').value = quoteData.customerName;
    sheet.getCell('A10').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('A10').alignment = alignLeft;

    sheet.getCell('A11').value = quoteData.customerContact;
    sheet.getCell('A11').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('A11').alignment = alignLeft;

    sheet.getCell('A12').value = quoteData.customerPhone;
    sheet.getCell('A12').font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    sheet.getCell('A12').alignment = alignLeft;

    // Row 15: Currency Header & Value
    sheet.getCell('E15').value = 'Currency';
    sheet.getCell('E15').font = fontArialBold;
    sheet.getCell('E15').fill = fillNavy;
    sheet.getCell('E15').alignment = alignLeft;

    sheet.getCell('F15').value = quoteData.currency;
    sheet.getCell('F15').font = fontTrebuchetBold;
    sheet.getCell('F15').alignment = alignLeft;

    // Row 16: Table Headers
    sheet.getRow(16).height = 18;
    sheet.mergeCells('A16:B16');
    sheet.getCell('A16').value = 'DESCRIPTION';
    sheet.getCell('A16').font = fontArialBold;
    sheet.getCell('A16').fill = fillNavy;
    sheet.getCell('A16').alignment = alignLeft;
    sheet.getCell('A16').border = { left: { style: 'thin' }, top: { style: 'thin' } };

    const colHeaders = {
      'C16': 'Origin',
      'D16': `QTY/${quoteData.unit}`,
      'E16': `Price/${quoteData.unit}`,
      'F16': 'AMOUNT'
    };
    
    Object.entries(colHeaders).forEach(([cellAddr, label]) => {
      const cell = sheet.getCell(cellAddr);
      cell.value = label;
      cell.font = fontArialBold;
      cell.fill = fillNavy;
      cell.alignment = alignCenter;
      cell.border = { top: { style: 'thin' } };
      if (cellAddr === 'F16') {
        cell.border = { top: { style: 'thin' }, right: { style: 'thin' } };
      }
    });

    // Rows 17 to (17+N-1): Table Rows
    for (let i = 0; i < N; i++) {
      const rNum = itemStartRow + i;
      const line = quoteData.lines[i];
      
      sheet.mergeCells(`A${rNum}:B${rNum}`);
      const descCell = sheet.getCell(`A${rNum}`);
      descCell.value = line.desc;
      descCell.font = fontTrebuchet;
      descCell.alignment = alignLeft;
      descCell.border = borderThinLeft;

      const originCell = sheet.getCell(`C${rNum}`);
      originCell.value = line.origin;
      originCell.font = fontTrebuchet;
      originCell.alignment = alignCenter;
      originCell.border = borderThinAll;

      const qtyCell = sheet.getCell(`D${rNum}`);
      qtyCell.value = line.qty !== '' ? parseFloat(line.qty) : null;
      qtyCell.font = fontTrebuchet;
      qtyCell.alignment = alignCenter;
      qtyCell.border = borderThinAll;
      if (qtyCell.value !== null) qtyCell.numFmt = '#,##0.00';

      const priceCell = sheet.getCell(`E${rNum}`);
      priceCell.value = line.price !== '' ? parseFloat(line.price) : null;
      priceCell.font = fontTrebuchet;
      priceCell.alignment = alignCenter;
      priceCell.border = borderThinAll;
      if (priceCell.value !== null) priceCell.numFmt = '#,##0.00';

      const amtCell = sheet.getCell(`F${rNum}`);
      amtCell.value = { formula: `D${rNum}*E${rNum}` };
      amtCell.font = fontTrebuchet;
      amtCell.alignment = alignRight;
      amtCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      amtCell.border = borderThinAll;
      amtCell.numFmt = '#,##0.00';
    }

    // Row: Subtotal
    sheet.getRow(subtotalRow).height = 14.25;
    sheet.getCell(`E${subtotalRow}`).value = 'Subtotal';
    sheet.getCell(`E${subtotalRow}`).font = fontTrebuchet;
    sheet.getCell(`E${subtotalRow}`).alignment = alignLeft;
    sheet.getCell(`E${subtotalRow}`).border = { top: { style: 'thin' } };

    const subValCell = sheet.getCell(`F${subtotalRow}`);
    subValCell.value = { formula: `SUM(F${itemStartRow}:F${itemEndRow})` };
    subValCell.font = fontTrebuchet;
    subValCell.alignment = alignRight;
    subValCell.border = { top: { style: 'thin' } };
    subValCell.numFmt = '#,##0.00';

    // Row: Terms & Conditions Title
    sheet.getRow(termsHeaderRow).height = 14.25;
    sheet.mergeCells(`A${termsHeaderRow}:C${termsHeaderRow}`);
    const termsTitle = sheet.getCell(`A${termsHeaderRow}`);
    termsTitle.value = 'TERMS AND CONDITIONS';
    termsTitle.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    termsTitle.fill = fillNavy;
    termsTitle.alignment = alignLeft;
    termsTitle.border = {
      left: { style: 'thin', color: { argb: 'FF3B4E87' } },
      top: { style: 'thin', color: { argb: 'FF3B4E87' } },
      bottom: { style: 'thin', color: { argb: 'FF3B4E87' } }
    };

    // Terms Lines (3 rows)
    const termsRows = [termsLine1, termsLine2, termsLine3];
    termsRows.forEach((rNum, idx) => {
      sheet.mergeCells(`A${rNum}:C${rNum}`);
      const tCell = sheet.getCell(`A${rNum}`);
      tCell.value = quoteData.terms[idx] || '';
      tCell.font = fontTrebuchet;
      tCell.alignment = alignLeft;
    });

    // Handle VAT calculation in Excel sheet E/F columns (on termsHeaderRow)
    let totalFormula = `F${subtotalRow}`;
    if (quoteData.vatRate > 0 && !quoteData.vatIncluded) {
      sheet.getCell(`E${termsHeaderRow}`).value = `VAT (${quoteData.vatRate}%)`;
      sheet.getCell(`E${termsHeaderRow}`).font = fontTrebuchet;
      sheet.getCell(`E${termsHeaderRow}`).alignment = alignLeft;

      const vatValCell = sheet.getCell(`F${termsHeaderRow}`);
      vatValCell.value = { formula: `F${subtotalRow}*(${quoteData.vatRate}/100)` };
      vatValCell.font = fontTrebuchet;
      vatValCell.alignment = alignRight;
      vatValCell.numFmt = '#,##0.00';
      
      totalFormula = `F${subtotalRow}+F${termsHeaderRow}`;
    }

    // Row: TOTAL
    sheet.getRow(totalRow).height = 15.0;
    sheet.getCell(`E${totalRow}`).value = 'TOTAL';
    sheet.getCell(`E${totalRow}`).font = { name: 'Trebuchet MS', size: 11, bold: true, color: { argb: 'FF101010' } };
    sheet.getCell(`E${totalRow}`).alignment = alignLeft;

    const totValCell = sheet.getCell(`F${totalRow}`);
    totValCell.value = { formula: totalFormula };
    totValCell.font = fontTrebuchetBold;
    totValCell.alignment = alignRight;
    totValCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8DCE7' } };
    totValCell.border = {
      top: { style: 'thin' },
      bottom: { style: 'double' }
    };
    totValCell.numFmt = '#,##0.00';

    // Row: Closing contact note
    sheet.mergeCells(`A${footerContactRow}:F${footerContactRow}`);
    const fContact = sheet.getCell(`A${footerContactRow}`);
    fContact.value = quoteData.footerContact;
    fContact.font = { name: 'Trebuchet MS', size: 10, color: { argb: 'FF101010' } };
    fContact.alignment = alignCenter;

    // Row: Thank you business
    sheet.getRow(footerThanksRow).height = 15.75;
    sheet.mergeCells(`A${footerThanksRow}:F${footerThanksRow}`);
    const fThanks = sheet.getCell(`A${footerThanksRow}`);
    fThanks.value = quoteData.footerThanks;
    fThanks.font = { name: 'Trebuchet MS', size: 12, bold: true, italic: true, color: { argb: 'FF101010' } };
    fThanks.alignment = alignCenter;

    // Save Buffer & download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `AWA_Quote_${quoteData.quoteNo || 'Draft'}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
    
    showToast("Excel exported successfully!");
  } catch (err) {
    console.error("Export Error:", err);
    showToast("Failed to export Excel.");
  }
}

// Start app on DOM load
window.addEventListener('DOMContentLoaded', init);

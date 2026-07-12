import ExcelJS from 'exceljs/dist/exceljs.min.js';

// Division configurations
const DIVISIONS = {
  add: {
    id: 'add',
    name: 'Food Additives',
    address: '243 Horreya Ave, Sportin Alexandria , Egypt',
    website: 'www.awa-foodadd.net',
    phone: '(+203) 4292030 - 4211200',
    fax: '+203 4292060',
    logo: '/assets/logo-add.png',
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
    logo: '/assets/logo-solutions.png',
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
    logo: '/assets/logo-bio.png',
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
let currentDiv = 'add';
let quoteData = {
  preparedBy: '',
  quoteNo: '123456',
  customerId: '123',
  date: '',
  validUntil: '',
  validityDays: 3,
  currency: 'EGP',
  vatRate: 14,
  vatIncluded: false,
  itemCount: 12,
  customerName: '[Company Name]',
  customerContact: '[contact name]',
  customerPhone: '[Phone]',
  lines: [], // Array of variable size
  terms: [], // Array of terms lines (3 items)
  footerContact: 'If you have any questions about this price quote, please contact',
  footerThanks: 'Thank You For Your Business!'
};

// Global materials data for autocomplete
let materialsList = [];

// Dom elements cache
const DOM = {
  divSelectors: document.querySelectorAll('.btn-div'),
  inputPreparedBy: document.getElementById('input-prepared-by'),
  inputQuoteNo: document.getElementById('input-quote-no'),
  inputCustomerId: document.getElementById('input-customer-id'),
  inputCurrency: document.getElementById('input-currency'),
  inputValidity: document.getElementById('input-validity'),
  inputVat: document.getElementById('input-vat'),
  inputItemCount: document.getElementById('input-item-count'),
  checkboxVatIncluded: document.getElementById('checkbox-vat-included'),
  btnNew: document.getElementById('btn-new'),
  btnPrint: document.getElementById('btn-print'),
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
  toast: document.getElementById('toast')
};

// Initialize Application
function init() {
  // Set default dates
  const today = new Date();
  quoteData.date = formatDate(today);
  quoteData.validUntil = formatDate(addDays(today, quoteData.validityDays));
  
  DOM.sheetDate.value = quoteData.date;
  DOM.sheetValidUntil.value = quoteData.validUntil;

  // Set default values in controls
  DOM.inputPreparedBy.value = '';
  quoteData.preparedBy = '[salesperson name]';
  DOM.sheetPreparedBy.textContent = quoteData.preparedBy;
  
  DOM.inputQuoteNo.value = quoteData.quoteNo;
  DOM.inputCustomerId.value = quoteData.customerId;
  DOM.inputCurrency.value = quoteData.currency;
  DOM.inputValidity.value = quoteData.validityDays;
  DOM.inputVat.value = quoteData.vatRate;
  DOM.inputItemCount.value = quoteData.itemCount;
  DOM.checkboxVatIncluded.checked = quoteData.vatIncluded;

  // Event Listeners for controls
  setupControlListeners();
  
  // Event Listeners for editable content on sheet
  setupSheetEditableListeners();

  // Setup button actions
  DOM.btnNew.addEventListener('click', () => resetQuote(currentDiv));
  DOM.btnPrint.addEventListener('click', () => window.print());
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

  DOM.inputCustomerId.addEventListener('input', (e) => {
    quoteData.customerId = e.target.value;
    DOM.sheetCustomerId.textContent = quoteData.customerId;
  });

  DOM.inputCurrency.addEventListener('input', (e) => {
    quoteData.currency = e.target.value;
    DOM.sheetCurrency.textContent = quoteData.currency;
  });

  DOM.inputValidity.addEventListener('input', (e) => {
    const days = parseInt(e.target.value) || 0;
    quoteData.validityDays = days;
    const dateVal = new Date(DOM.sheetDate.value);
    const newValidUntil = formatDate(addDays(dateVal, days));
    quoteData.validUntil = newValidUntil;
    DOM.sheetValidUntil.value = newValidUntil;
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
}

// Setup Event Listeners for Sheet Preview editable texts
function setupSheetEditableListeners() {
  DOM.sheetPreparedBy.addEventListener('blur', (e) => {
    quoteData.preparedBy = e.target.textContent;
    DOM.inputPreparedBy.value = quoteData.preparedBy;
  });

  DOM.sheetQuoteNo.addEventListener('blur', (e) => {
    quoteData.quoteNo = e.target.textContent;
    DOM.inputQuoteNo.value = quoteData.quoteNo;
  });

  DOM.sheetCustomerId.addEventListener('blur', (e) => {
    quoteData.customerId = e.target.textContent;
    DOM.inputCustomerId.value = quoteData.customerId;
  });

  DOM.sheetDate.addEventListener('change', (e) => {
    quoteData.date = e.target.value;
    const newValidUntil = formatDate(addDays(new Date(e.target.value), quoteData.validityDays));
    quoteData.validUntil = newValidUntil;
    DOM.sheetValidUntil.value = newValidUntil;
  });

  DOM.sheetValidUntil.addEventListener('change', (e) => {
    quoteData.validUntil = e.target.value;
    const d1 = new Date(DOM.sheetDate.value);
    const d2 = new Date(e.target.value);
    const diffTime = Math.abs(d2 - d1);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    quoteData.validityDays = diffDays;
    DOM.inputValidity.value = diffDays;
  });

  // Customer bindings
  DOM.sheetCustomerName.addEventListener('blur', (e) => quoteData.customerName = e.target.textContent);
  DOM.sheetCustomerContact.addEventListener('blur', (e) => quoteData.customerContact = e.target.textContent);
  DOM.sheetCustomerPhone.addEventListener('blur', (e) => quoteData.customerPhone = e.target.textContent);
  DOM.sheetCurrency.addEventListener('blur', (e) => {
    quoteData.currency = e.target.textContent;
    DOM.inputCurrency.value = quoteData.currency;
  });

  DOM.sheetFooterContact.addEventListener('blur', (e) => quoteData.footerContact = e.target.textContent);
  DOM.sheetFooterThanks.addEventListener('blur', (e) => quoteData.footerThanks = e.target.textContent);
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
  resetTermsToDefault(config.defaultTerms);
  calculateTotals();
}

// Resize the lines list (preserving current typed inputs)
function resizeQuoteLines(newCount) {
  const currentLines = [...quoteData.lines];
  const newLines = [];
  
  for (let i = 0; i < newCount; i++) {
    if (currentLines[i]) {
      newLines.push(currentLines[i]); // Keep existing
    } else {
      // Pad with blank row
      newLines.push({ desc: '', origin: '', qty: '', price: '' });
    }
  }
  
  quoteData.lines = newLines;
  renderItemTable();
  calculateTotals();
}

// Renders rows in the preview table using text input fields for easy typing
function renderItemTable() {
  DOM.itemsBody.innerHTML = '';
  
  quoteData.lines.forEach((line, idx) => {
    const row = document.createElement('tr');
    
    // Note: inputs are type="text" to allow normal keyboard inputs without arrows blocking
    row.innerHTML = `
      <td colspan="2" class="desc-cell-container"><input type="text" class="desc-input text-trebuchet" value="${line.desc}" placeholder="Item description..." autocomplete="off"></td>
      <td class="center-align"><input type="text" class="origin-input text-trebuchet text-center" value="${line.origin}" placeholder="Origin..."></td>
      <td class="center-align"><input type="text" class="qty-input num-input text-trebuchet text-center" value="${line.qty}" placeholder="-"></td>
      <td class="center-align"><input type="text" class="price-input num-input text-trebuchet text-center" value="${line.price}" placeholder="-"></td>
      <td class="amount-cell text-trebuchet amount-val" id="amount-${idx}">0.00</td>
    `;
    
    DOM.itemsBody.appendChild(row);
  });

  setupTableListeners();
  
  // Update all amounts on sheet
  quoteData.lines.forEach((_, i) => {
    calculateRowAmount(i);
  });
}

// Load default terms
function resetTermsToDefault(defaultTerms) {
  DOM.termsList.innerHTML = '';
  quoteData.terms = [...defaultTerms];
  
  for (let i = 0; i < 3; i++) {
    const val = quoteData.terms[i] || '';
    const item = document.createElement('div');
    item.className = 'terms-item';
    item.contentEditable = 'true';
    item.textContent = val;
    
    item.addEventListener('blur', (e) => {
      quoteData.terms[i] = e.target.textContent;
    });
    
    DOM.termsList.appendChild(item);
  }
}

// Set listeners on inputs inside preview table
function setupTableListeners() {
  const rows = DOM.itemsBody.querySelectorAll('tr');
  rows.forEach((row, idx) => {
    const descInput = row.querySelector('.desc-input');
    const originInput = row.querySelector('.origin-input');
    const qtyInput = row.querySelector('.qty-input');
    const priceInput = row.querySelector('.price-input');
    const descCell = descInput.parentElement;
    
    const updateState = () => {
      quoteData.lines[idx].desc = descInput.value;
      quoteData.lines[idx].origin = originInput.value;
      
      // Clean inputs: allow only numbers, dots, and common floats
      let qtyVal = qtyInput.value.replace(/[^0-9.]/g, '');
      let priceVal = priceInput.value.replace(/[^0-9.]/g, '');
      
      // Update fields if they typed invalid characters
      if (qtyInput.value !== qtyVal) qtyInput.value = qtyVal;
      if (priceInput.value !== priceVal) priceInput.value = priceVal;
      
      quoteData.lines[idx].qty = qtyVal;
      quoteData.lines[idx].price = priceVal;
      
      calculateRowAmount(idx);
    };

    descInput.addEventListener('input', updateState);
    originInput.addEventListener('input', updateState);
    qtyInput.addEventListener('input', updateState);
    priceInput.addEventListener('input', updateState);

    // Autocomplete Dropdown Functionality
    let activeItemIdx = -1;

    const closeDropdown = () => {
      const existing = descCell.querySelector('.autocomplete-dropdown');
      if (existing) existing.remove();
      activeItemIdx = -1;
    };

    const selectItem = (item) => {
      descInput.value = item.desc;
      originInput.value = item.origin;
      priceInput.value = item.price !== '' && item.price !== 0 ? item.price.toString() : '';
      
      // Update state data structure
      quoteData.lines[idx].desc = item.desc;
      quoteData.lines[idx].origin = item.origin;
      quoteData.lines[idx].price = priceInput.value;
      
      calculateRowAmount(idx);
      closeDropdown();
      qtyInput.focus(); // Shift focus to Quantity input automatically
    };

    descInput.addEventListener('input', (e) => {
      closeDropdown();
      const val = descInput.value.trim().toLowerCase();
      if (!val) return;

      // Filter: only show options starting with typed text
      const matches = materialsList.filter(item => 
        item.desc.toLowerCase().startsWith(val)
      ).slice(0, 40); // limit popup list to 40 items

      if (matches.length === 0) return;

      const dropdown = document.createElement('div');
      dropdown.className = 'autocomplete-dropdown';

      matches.forEach((match, mIdx) => {
        const itemElement = document.createElement('div');
        itemElement.className = 'autocomplete-item';
        itemElement.textContent = match.desc;
        if (mIdx === 0) {
          itemElement.classList.add('active');
          activeItemIdx = 0;
        }

        // Prevent input blur prior to executing selection
        itemElement.addEventListener('mousedown', (e) => {
          e.preventDefault();
          selectItem(match);
        });

        dropdown.appendChild(itemElement);
      });

      descCell.appendChild(dropdown);
    });

    descInput.addEventListener('keydown', (e) => {
      const dropdown = descCell.querySelector('.autocomplete-dropdown');
      if (!dropdown) return;

      const items = dropdown.querySelectorAll('.autocomplete-item');
      if (items.length === 0) return;

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        items[activeItemIdx]?.classList.remove('active');
        activeItemIdx = (activeItemIdx + 1) % items.length;
        items[activeItemIdx].classList.add('active');
        items[activeItemIdx].scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        items[activeItemIdx]?.classList.remove('active');
        activeItemIdx = (activeItemIdx - 1 + items.length) % items.length;
        items[activeItemIdx].classList.add('active');
        items[activeItemIdx].scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (activeItemIdx >= 0 && activeItemIdx < items.length) {
          const selectedText = items[activeItemIdx].textContent;
          const match = materialsList.find(m => m.desc === selectedText);
          if (match) {
            selectItem(match);
          }
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        closeDropdown();
      }
    });

    descInput.addEventListener('blur', () => {
      setTimeout(closeDropdown, 150);
    });
  });
}

// Calculates amount for a single row
function calculateRowAmount(idx) {
  const line = quoteData.lines[idx];
  const qty = parseFloat(line.qty) || 0;
  const price = parseFloat(line.price) || 0;
  const amount = qty * price;
  
  const amountCell = document.getElementById(`amount-${idx}`);
  if (qty > 0 && price > 0) {
    amountCell.textContent = amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  } else {
    amountCell.textContent = '0.00';
  }
  calculateTotals();
}

// Calculate subtotal, VAT, and total
function calculateTotals() {
  let subtotal = 0;
  quoteData.lines.forEach((line) => {
    const qty = parseFloat(line.qty) || 0;
    const price = parseFloat(line.price) || 0;
    subtotal += qty * price;
  });

  let vatAmount = 0;
  let total = subtotal;

  if (quoteData.vatRate > 0) {
    if (quoteData.vatIncluded) {
      // VAT is already included in prices: total = subtotal
      vatAmount = subtotal - (subtotal / (1 + quoteData.vatRate / 100));
      DOM.vatCalcRow.style.display = 'none';
    } else {
      // VAT needs to be added: total = subtotal + vat
      vatAmount = subtotal * (quoteData.vatRate / 100);
      total = subtotal + vatAmount;
      DOM.vatCalcRow.style.display = 'table-row';
    }
  } else {
    DOM.vatCalcRow.style.display = 'none';
  }

  DOM.sheetSubtotal.textContent = subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  DOM.sheetVatAmount.textContent = vatAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  DOM.sheetTotal.textContent = total.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Reset current quote to default blanks
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
  
  DOM.sheetDate.value = quoteData.date;
  DOM.sheetValidUntil.value = quoteData.validUntil;
  DOM.inputValidity.value = 3;
  
  quoteData.itemCount = 12;
  DOM.inputItemCount.value = 12;

  switchDivision(divId);
  showToast("Quotation sheet reset.");
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
      'D16': 'QTY/kg',
      'E16': 'Price/kg',
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

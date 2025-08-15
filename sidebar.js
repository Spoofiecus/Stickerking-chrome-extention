document.addEventListener('DOMContentLoaded', () => {
  const addStickerBtn = document.getElementById('add-sticker');
  const calculateBtn = document.getElementById('calculate');
  const copyQuoteBtn = document.getElementById('copy-quote');
  const stickersDiv = document.getElementById('stickers');
  const resultsDiv = document.getElementById('results');
  const vinylCostInput = document.getElementById('vinyl-cost');
  const vatRateInput = document.getElementById('vat-rate');
  const includeVatCheckbox = document.getElementById('include-vat');
  const darkModeToggle = document.getElementById('dark-mode-toggle');
  const materialSelect = document.getElementById('material-select');

  let stickerCount = 0;

  // Initialize material selection
  materialSelect.value = 'unspecified';

  // Dark Mode Toggle
  darkModeToggle.addEventListener('change', () => {
    document.body.classList.toggle('dark-mode', darkModeToggle.checked);
    localStorage.setItem('darkMode', darkModeToggle.checked);
  });

  // Load dark mode preference
  if (localStorage.getItem('darkMode') === 'true') {
    darkModeToggle.checked = true;
    document.body.classList.add('dark-mode');
  }

  // Collapsible Sections (only for Settings)
  document.querySelectorAll('.section-toggle').forEach(button => {
    button.addEventListener('click', () => {
      const targetId = button.getAttribute('data-target');
      const content = document.getElementById(targetId);
      const icon = button.querySelector('.toggle-icon');
      content.classList.toggle('active');
      icon.classList.toggle('active');
      button.setAttribute('aria-expanded', content.classList.contains('active'));
    });
  });

  // Add a new sticker input field
  function addStickerInput() {
    stickerCount++;
    const stickerInput = document.createElement('div');
    stickerInput.className = 'sticker-input';
    stickerInput.setAttribute('data-id', stickerCount);
    stickerInput.innerHTML = `
      <div class="input-group">
        <label for="width-${stickerCount}">Width (mm)</label>
        <input type="number" id="width-${stickerCount}" min="1" title="Enter the width of the sticker in millimeters" aria-label="Sticker width in millimeters">
      </div>
      <div class="input-group">
        <label for="height-${stickerCount}">Height (mm)</label>
        <input type="number" id="height-${stickerCount}" min="1" title="Enter the height of the sticker in millimeters" aria-label="Sticker height in millimeters">
      </div>
      <div class="input-group">
        <label for="quantity-${stickerCount}">Quantity</label>
        <input type="number" id="quantity-${stickerCount}" min="1" value="1" title="Enter the number of stickers needed" aria-label="Sticker quantity">
      </div>
      <button class="remove-button" data-id="${stickerCount}" aria-label="Remove sticker"></button>
    `;
    stickersDiv.appendChild(stickerInput);
    stickerInput.style.animation = 'fadeInUp 0.3s ease forwards';
  }

  // Remove a sticker input field
  stickersDiv.addEventListener('click', (event) => {
    if (event.target.closest('.remove-button')) {
      const button = event.target.closest('.remove-button');
      const id = button.getAttribute('data-id');
      const inputDiv = document.querySelector(`.sticker-input[data-id="${id}"]`);
      if (inputDiv) {
        inputDiv.style.animation = 'fadeOutDown 0.3s ease forwards';
        setTimeout(() => inputDiv.remove(), 300);
      }
    }
  });

  // Calculate sticker price
  function calculatePrice(width, height) {
    const VINYL_COST = parseFloat(vinylCostInput.value) || 420.00;
    const ROLL_WIDTH = 650;
    const BLEED = 1;
    const MIN_PRICE_PER_STICKER = 0.20;

    const W = parseFloat(width);
    const H = parseFloat(height);

    if (isNaN(W) || isNaN(H) || W <= 0 || H <= 0) {
      return { price: 'Invalid dimensions', stickersPerRow: 0 };
    }

    // Horizontal Orientation
    const W_bleed_horizontal = W + BLEED;
    const S_raw_horizontal = ROLL_WIDTH / W_bleed_horizontal;
    const decimal_part_horizontal = S_raw_horizontal - Math.floor(S_raw_horizontal);
    const S_rounded_horizontal = decimal_part_horizontal >= 0.95 ? Math.ceil(S_raw_horizontal) : Math.floor(S_raw_horizontal);
    const H_meters_horizontal = H / 1000;
    const Area_horizontal = 0.65 * H_meters_horizontal;
    const Row_Cost_horizontal = Area_horizontal * VINYL_COST;
    const P_horizontal = S_rounded_horizontal > 0 ? Row_Cost_horizontal / S_rounded_horizontal : Infinity;

    // Vertical Orientation
    const H_bleed_vertical = H + BLEED;
    const S_raw_vertical = ROLL_WIDTH / H_bleed_vertical;
    const decimal_part_vertical = S_raw_vertical - Math.floor(S_raw_vertical);
    const S_rounded_vertical = decimal_part_vertical >= 0.95 ? Math.ceil(S_raw_vertical) : Math.floor(S_raw_vertical);
    const W_meters_vertical = W / 1000;
    const Area_vertical = 0.65 * W_meters_vertical;
    const Row_Cost_vertical = Area_vertical * VINYL_COST;
    const P_vertical = S_rounded_vertical > 0 ? Row_Cost_vertical / S_rounded_vertical : Infinity;

    const price = Math.min(P_horizontal, P_vertical);
    const adjustedPrice = Math.max(price, MIN_PRICE_PER_STICKER);
    const stickersPerRow = price < P_vertical ? S_rounded_horizontal : S_rounded_vertical;

    return { price: adjustedPrice.toFixed(2), stickersPerRow };
  }

  // Event listener for calculating prices
  calculateBtn.addEventListener('click', () => {
    resultsDiv.innerHTML = ''; // Clear previous results
    // Re-add the copy quote button since innerHTML clears it
    resultsDiv.innerHTML = `
      <button id="copy-quote" class="copy-quote-button">
        <i class="fas fa-copy"></i>
      </button>
    `;

    if (materialSelect.value === 'unspecified') {
      alert('Quote not generated. Reason: no Material specified.');
      return;
    }

    let quote = "Dear Customer. Thank you for reaching out to us.\nBelow is your Quote based on your request:\n\n";
    // Add material to the quote
    const material = materialSelect.value === 'gloss' ? 'Gloss' : 'MATT';
    quote += `Material: ${material}\n\n`;

    let totalCostExclVat = 0;
    const vatRate = parseFloat(vatRateInput.value) / 100 || 0.15;
    const includeVat = includeVatCheckbox.checked;
    const MIN_ORDER_AMOUNT = 100.00;
    const roundedCorners = document.getElementById('rounded-corners').checked;

    const stickerInputs = stickersDiv.querySelectorAll('.sticker-input');
    stickerInputs.forEach((input, index) => {
      const widthInput = input.querySelector(`#width-${input.getAttribute('data-id')}`);
      const heightInput = input.querySelector(`#height-${input.getAttribute('data-id')}`);
      const quantityInput = input.querySelector(`#quantity-${input.getAttribute('data-id')}`);
      if (widthInput && heightInput && quantityInput) {
        const width = widthInput.value;
        const height = heightInput.value;
        const quantity = parseInt(quantityInput.value) || 1;
        if (width && height) {
          const { price, stickersPerRow } = calculatePrice(width, height);
          if (price === 'Invalid dimensions') {
            resultsDiv.innerHTML += `<p>Sticker ${index + 1}: Invalid dimensions</p>`;
            quote += `Sticker ${index + 1} (${width}x${height}mm): Invalid dimensions\n`;
          } else {
            const rows = Math.ceil(quantity / stickersPerRow);
            const totalStickers = rows * stickersPerRow;
            const totalPriceExclVatPerSticker = (price * totalStickers).toFixed(2);
            const totalPriceInclVat = (totalPriceExclVatPerSticker * (1 + vatRate)).toFixed(2);
            resultsDiv.innerHTML += `<p>${width}x${height}mm - R${price} excl VAT per sticker (${stickersPerRow} stickers per row)<br>${rows} rows - ${totalStickers} stickers<br>R${totalPriceExclVatPerSticker} Excl VAT</p>`;
            quote += `${width}x${height}mm - R${price} excl VAT per sticker (${stickersPerRow} stickers per row)\n${rows} rows - ${totalStickers} stickers\nR${totalPriceExclVatPerSticker} Excl VAT\n`;
            if (includeVat) {
              resultsDiv.innerHTML += `<p style="margin-left: 20px;">Incl VAT: R${totalPriceInclVat}</p>`;
              quote += `Incl VAT: R${totalPriceInclVat}\n`;
            }
            totalCostExclVat += parseFloat(totalPriceExclVatPerSticker);
          }
        }
      }
    });

    if (totalCostExclVat > 0) {
      const totalCostInclVat = (totalCostExclVat * (1 + vatRate)).toFixed(2);
      const totalTextExcl = `Total: R${totalCostExclVat.toFixed(2)} Exclusive of VAT`;
      resultsDiv.innerHTML += `<p><strong>${totalTextExcl}</strong></p>`;
      quote += `\n${totalTextExcl}\n`;
      if (includeVat) {
        const totalTextIncl = `Total Incl VAT: R${totalCostInclVat}. the complete order total`;
        resultsDiv.innerHTML += `<p><strong>${totalTextIncl}</strong></p>`;
        quote += `${totalTextIncl}\n`;
      }

      if (totalCostExclVat < MIN_ORDER_AMOUNT) {
        const minimumMessage = 'YOUR ORDER IS UNDER R100.00 EXCL VAT. WE HAVE A MINIMUM ORDER AMOUNT OF R100.00 EXCL VAT';
        resultsDiv.innerHTML += `<p style="color: #E74C3C; text-transform: uppercase;">${minimumMessage}</p>`;
        quote += `\n${minimumMessage}\n`;
      }

      if (roundedCorners) {
        quote += `\nCutline with rounded Corners\n`;
      }

      quote += `\nPlease let us know if this quote is accepted so we can proceed with printing.\n`;
      localStorage.setItem('quote', quote);
      resultsDiv.classList.add('show');
    }

    // Re-attach the event listener for the copy quote button
    const newCopyQuoteBtn = document.getElementById('copy-quote');
    newCopyQuoteBtn.addEventListener('click', () => {
      const quote = localStorage.getItem('quote');
      if (quote) {
        navigator.clipboard.writeText(quote).then(() => {
          alert('Quote copied to clipboard!');
        });
      } else {
        alert('No quote to copy. Please calculate prices first.');
      }
    });
  });

  // Event listener for adding stickers
  addStickerBtn.addEventListener('click', () => addStickerInput());

  // Add initial sticker input on load
  addStickerInput();

  // Ensure the container adjusts dynamically on resize
  window.addEventListener('resize', () => {
    const container = document.querySelector('.container');
    container.style.width = '100%';
  });
});
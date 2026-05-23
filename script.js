const DRAFT_KEY = "deliveryFormDraft";
const MAX_ENTRY_ROWS = 10;

let currentInvoiceDisplay = [];

function getEntryRowCount() {

  return document.querySelectorAll("#entryBody tr").length;

}

function createEntryRowHTML(rowNum) {

  return `

    <td class="col-shrink col-num row-num-cell">

      <span class="row-num">${rowNum}</span>

    </td>

    <td class="col-item-cell">

      <textarea
        class="item-input"
        placeholder="Item"
      ></textarea>

    </td>

    <td class="col-shrink col-qty-cell">

      <input
        type="text"
        inputmode="numeric"
        class="qty-input"
        placeholder="Qty"
      >

    </td>

    <td class="col-fill col-batch-cell">

      <textarea
        class="batch-input"
        placeholder="A010126--1, SH010126--2 (4 per row)"
        onblur="formatBatchInput(this)"
      ></textarea>

    </td>

    <td class="col-shrink col-action action-cell">

      <button
        type="button"
        class="delete-btn"
        onclick="deleteRow(this)"
      >
        Delete
      </button>

    </td>

  `;

}

function updateRowNumbers() {

  const rows =
    document.querySelectorAll("#entryBody tr");

  rows.forEach((row, index) => {

    const numCell =
      row.querySelector(".row-num");

    if (numCell) {

      numCell.textContent = index + 1;

    }

  });

  const addBtn =
    document.getElementById("addRowBtn");

  if (addBtn) {

    addBtn.disabled =
      rows.length >= MAX_ENTRY_ROWS;

  }

}

function formatBatchInput(textarea) {

  const codes =
    splitBatchCodes(textarea.value);

  const isSplitScreen =
    window.innerWidth <= 1200;

  const perLine =
    isSplitScreen ? 2 : 4;

  const lines = [];

  for (let i = 0; i < codes.length; i += perLine) {

    lines.push(
      codes
        .slice(i, i + perLine)
        .join(", ")
    );

  }

  textarea.value =
    lines.join("\n");

  // AUTO CALCULATE QTY

  const row = textarea.closest("tr");

  if (!row) return;

  const qtyInput =
    row.querySelector(".qty-input");

  if (!qtyInput) return;



  let totalQty = 0;

  codes.forEach(code => {

    // SUPPORT PLACEHOLDER
    // Example: 0-4

    if (/^0-\d+$/.test(code.trim())) {

      const placeholderQty =
        parseInt(code.split("-")[1]) || 0;

      totalQty += placeholderQty;

      return;

    }

    const match =
      code.match(/--(\d+)$/);

    if (match) {

      totalQty +=
        parseInt(match[1]) || 0;

    }

  });

  // ONLY AUTO FILL IF EMPTY

  if (
    qtyInput.value.trim() === ""
  ) {

    qtyInput.value = totalQty;

  }

  saveDraft();

}

function renderCurrentInvoiceTable() {

  renderRecordsTableBody(
    "currentRecordBody",
    currentInvoiceDisplay,
    {

      deleteHandler: "deleteCurrentInvoiceRow",

      getDeleteIndex: (_record, index) => index

    }
  );

}

function getDeliveryExportRows() {

  const rows = [];

  document.querySelectorAll("#entryBody tr").forEach(row => {

    const item =
      row.querySelector(".item-input")
        ?.value
        .trim() || "";

    const qty =
      row.querySelector(".qty-input")
        ?.value
        .trim() || "";

    const batch =
      row.querySelector(".batch-input")
        ?.value
        .trim() || "";

    if (!item && !batch) return;

    rows.push({

      "Item(s)": item,

      "Quantity": qty,

      "Manufactured Code":
        batch
          ? formatBatchForExport(batch)
          : ""

    });

  });

  return rows;

}

function exportDeliveryExcel() {

  const invoice =
    document.getElementById("invoiceNo")
      .value
      .trim();

  if (!invoice) {

    alert("Please enter Invoice No");

    return;

  }

  const rows = getDeliveryExportRows();

  if (rows.length === 0) {

    alert("No delivery rows to export. Enter item and manufactured code.");

    return;

  }

  const worksheet =
    XLSX.utils.json_to_sheet(rows);

  worksheet["!cols"] = [

    { wch: 26 },

    { wch: 8 },

    { wch: 50 }

  ];

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Delivery Record"
  );

  const safeInvoice =
    invoice.replace(/[^a-zA-Z0-9_-]/g, "_");

  XLSX.writeFile(
    workbook,
    `Delivery_Record_${safeInvoice}.xlsx`
  );

}

function rebuildPrintBody() {

  const printBody =
    document.getElementById("printBody");

  printBody.innerHTML = "";

  const rows =
    document.querySelectorAll("#entryBody tr");

  let validRowCount = 0;

  rows.forEach(row => {

    let item =
      row.querySelector(".item-input")
        ?.value
        .trim() || "";

    const qty =
      row.querySelector(".qty-input")
        ?.value
        .trim() || "";

    const batchInput =
      row.querySelector(".batch-input");

    if (!batchInput) return;

    let batch =
      batchInput.value.trim();

    if (!item) return;

    validRowCount++;

    // PRINT ONLY:
    // Auto line break before "(Lock"

    if (item.includes("(Lock")) {

      item =
        item.replace(/\s*\(Lock/g, "<br>(Lock");

    }

    // NO BATCH = NA internally

    if (!batch) {

      batch = "NA";

    }

    // BUILD SLOT ARRAY

    const batchCodes =
      batch === "NA"
        ? []
        : splitBatchCodes(batch);

    const qtyNum =
      parseInt(qty) || 0;

    const slots = [];


    let occupiedQty = 0;

    batchCodes.forEach(code => {

      const match = code.match(/--(\d+)$/);

      const codeQty = match
        ? parseInt(match[1])
        : 1;

      slots.push(code);

      occupiedQty += codeQty;

    });

    // ONLY add invisible spacing
    // if occupied qty < total qty

    const remainingQty =
      qtyNum - occupiedQty;

    for (let i = 0; i < remainingQty; i++) {

      slots.push("");

    }

    // BUILD PRINT LINES (FIXED 4 PER ROW)

    const lines = [];

    for (let i = 0; i < slots.length; i += 4) {

      const rowSlots = slots.slice(i, i + 4);

      lines.push(`
    <div class="print-batch-row">
      ${rowSlots.map((slot, index) => {

        if (!slot) {
          return `<span class="print-slot empty">&nbsp;</span>`;
        }

        const isLastInRow = index === rowSlots.length - 1;
        const isLastRow = (i + rowSlots.length) >= slots.length;

        const shouldAddComma = !(isLastInRow && isLastRow);

        return `<span class="print-slot">
    ${slot}${shouldAddComma ? "," : ""}
  </span>`;

      }).join("")}
    </div>
  `);

    }

    const wrappedBatch = lines.join("");

    printBody.innerHTML += `

      <tr>

        <td class="print-col-item">
          ${item}
        </td>

        <td class="print-col-qty">
          ${qty}
        </td>

        <td class="print-col-batch batch-cell">
          ${wrappedBatch}
        </td>

      </tr>

    `;

  });

  // Dynamic Delivered On spacing

  const deliverySection =
    document.getElementById("deliverySection");

  if (deliverySection) {

    if (validRowCount <= 4) {

      deliverySection.style.marginTop = "90pt";

    }

    else if (validRowCount <= 10) {

      deliverySection.style.marginTop = "54pt";

    }

    else {

      deliverySection.style.marginTop = "18pt";

    }

  }

}

function deleteCurrentInvoiceRow(index) {

  const removed =
    currentInvoiceDisplay[index];

  if (removed) {

    const stored =
      JSON.parse(localStorage.getItem("records"))
      || [];

    const storedIndex =
      stored.findIndex(record =>

        record.invoice === removed.invoice

        && record.rawCode === removed.rawCode

        && record.item === removed.item

      );

    if (storedIndex >= 0) {

      stored.splice(storedIndex, 1);

      localStorage.setItem(
        "records",
        JSON.stringify(stored)
      );

    }

  }

  currentInvoiceDisplay.splice(index, 1);

  renderCurrentInvoiceTable();

  rebuildPrintBody();

  saveDraft();

}

function saveDraft() {

  const rows =
    [...document.querySelectorAll("#entryBody tr")]
      .map(row => ({

        item:
          row.querySelector(".item-input")?.value ?? "",

        qty:
          row.querySelector(".qty-input")?.value ?? "",

        batch:
          row.querySelector(".batch-input")?.value ?? ""

      }));

  sessionStorage.setItem(DRAFT_KEY, JSON.stringify({

    invoice:
      document.getElementById("invoiceNo").value,

    deliveryDate:
      document.getElementById("deliveryDate").value,

    rows,

    currentInvoiceDisplay,

    printBodyHTML:
      document.getElementById("printBody").innerHTML,

    printDeliveryDate:
      document.getElementById("printDeliveryDate").innerText

  }));

}

function loadDraft() {

  const raw = sessionStorage.getItem(DRAFT_KEY);

  if (!raw) return;

  const draft = JSON.parse(raw);

  document.getElementById("invoiceNo").value =
    draft.invoice || "";

  document.getElementById("deliveryDate").value =
    draft.deliveryDate || "";

  const tbody = document.getElementById("entryBody");

  tbody.innerHTML = "";

  const rowData =
    draft.rows?.length ? draft.rows : [{}];

  rowData.forEach(r => {

    addRow();

    const tr = tbody.lastElementChild;

    tr.querySelector(".item-input").value =
      r.item || "";

    tr.querySelector(".qty-input").value =
      r.qty || "";

    tr.querySelector(".batch-input").value =
      r.batch || "";

  });

  if (draft.currentInvoiceDisplay) {

    currentInvoiceDisplay =
      draft.currentInvoiceDisplay.map(row => {

        if (row.rawCode) {

          return row;

        }

        if (row.batch && row.invoice && row.item) {

          return createRecordFromBatch(

            row.invoice,

            row.deliveryDate || "",

            row.item,

            row.batch

          );

        }

        return row;

      });

    renderCurrentInvoiceTable();

  } else if (draft.currentRecordHTML) {

    document.getElementById("currentRecordBody").innerHTML =
      draft.currentRecordHTML;

  } else {

    currentInvoiceDisplay = [];

    renderCurrentInvoiceTable();

  }

  document.getElementById("printDeliveryDate").innerText =
    draft.printDeliveryDate || "";

  if (draft.printBodyHTML) {

    document.getElementById("printBody").innerHTML =
      draft.printBodyHTML;

  } else {

    rebuildPrintBody();

  }

  updateRowNumbers();

}

function goToRecords() {

  saveDraft();

  window.open("records.html", "_blank");

}

document.addEventListener("DOMContentLoaded", () => {

  loadDraft();

  if (!sessionStorage.getItem(DRAFT_KEY)) {

    updateRowNumbers();

  }

  const container =
    document.querySelector(".container.screen-only");

  if (container) {

    container.addEventListener("input", saveDraft);

    container.addEventListener("change", saveDraft);

  }

});

function addRow() {

  if (getEntryRowCount() >= MAX_ENTRY_ROWS) {

    alert("Maximum 10 rows allowed.");

    return;

  }

  const tbody =
    document.getElementById("entryBody");

  const row = document.createElement("tr");

  row.innerHTML =
    createEntryRowHTML(getEntryRowCount() + 1);

  tbody.appendChild(row);

  updateRowNumbers();

  saveDraft();

}

function deleteRow(button) {

  const tbody =
    document.getElementById("entryBody");

  if (tbody.querySelectorAll("tr").length <= 1) {

    alert("At least one row is required.");

    return;

  }

  button.closest("tr").remove();

  updateRowNumbers();

  saveDraft();

}

function formatDate(dateString) {

  if (!dateString) return "";

  const months = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
  ];

  const date = new Date(dateString);

  const day =
    String(date.getDate()).padStart(2, "0");

  const month =
    months[date.getMonth()];

  const year =
    date.getFullYear();

  return `${day} ${month} ${year}`;

}

function saveRecord() {

  const invoice =
    document.getElementById("invoiceNo")
      .value
      .trim();

  const deliveryDate =
    document.getElementById("deliveryDate")
      .value;

  if (!invoice) {

    alert("Please enter Invoice No");

    return;

  }

  const rows =
    document.querySelectorAll("#entryBody tr");

  let records =
    JSON.parse(localStorage.getItem("records"))
    || [];

  currentInvoiceDisplay = [];

  rows.forEach(row => {

    const item =
      row.querySelector(".item-input")
        .value
        .trim();

    const qty =
      row.querySelector(".qty-input")
        .value
        .trim();

    const batchInput =
      row.querySelector(".batch-input");

    formatBatchInput(batchInput);

    const batch =
      batchInput.value.trim();

    if (!item || !batch) return;
    const finalBatch =
      batch || "NA";


    splitBatchCodes(finalBatch).forEach(code => {

      const record =
        createRecordFromBatch(

          invoice,

          formatDate(deliveryDate),

          item,

          code

        );

      currentInvoiceDisplay.push(record);

      records.push(record);

    });

  });

  renderCurrentInvoiceTable();

  rebuildPrintBody();

  document.getElementById(
    "printDeliveryDate"
  ).innerText =
    formatDate(deliveryDate);

  localStorage.setItem(
    "records",
    JSON.stringify(records)
  );

  saveDraft();

}

function newInvoice() {

  sessionStorage.removeItem(DRAFT_KEY);

  document.getElementById("invoiceNo")
    .value = "";


  document.getElementById("entryBody")
    .innerHTML = "";

  currentInvoiceDisplay = [];

  document.getElementById(
    "currentRecordBody"
  ).innerHTML = "";

  document.getElementById("printBody").innerHTML = "";

  document.getElementById("printDeliveryDate").innerText = "";

  addRow();

}

function clearEntryContent() {

  document.querySelectorAll("#entryBody tr")
    .forEach(row => {

      row.querySelector(".item-input").value = "";

      row.querySelector(".qty-input").value = "";

      row.querySelector(".batch-input").value = "";

    });

  currentInvoiceDisplay = [];

  renderCurrentInvoiceTable();

  rebuildPrintBody();

  saveDraft();

}

document.addEventListener("keydown", function (e) {

  const active =
    document.activeElement;

  const selectors = [

    ".item-input",

    ".qty-input",

    ".batch-input"

  ];

  let currentType = null;

  selectors.forEach(selector => {

    if (active.matches(selector)) {

      currentType = selector;

    }

  });

  if (!currentType) return;

  const currentInputs = [

    ...document.querySelectorAll(currentType)

  ];

  const currentIndex =
    currentInputs.indexOf(active);

  // ENTER / DOWN
  // SAME COLUMN NEXT ROW

  if (

    e.key === "Enter"

    ||

    e.key === "ArrowDown"

  ) {

    e.preventDefault();

    if (currentIndex < currentInputs.length - 1) {

      currentInputs[currentIndex + 1].focus();

    }

  }

  // UP

  if (e.key === "ArrowUp") {

    e.preventDefault();

    if (currentIndex > 0) {

      currentInputs[currentIndex - 1].focus();

    }

  }

  // RIGHT

  if (e.key === "ArrowRight") {

    e.preventDefault();

    moveHorizontal(active, 1);

  }

  // LEFT

  if (e.key === "ArrowLeft") {

    e.preventDefault();

    moveHorizontal(active, -1);

  }

});

function moveHorizontal(active, direction) {

  const row =
    active.closest("tr");

  if (!row) return;

  const rowInputs = [

    row.querySelector(".item-input"),

    row.querySelector(".qty-input"),

    row.querySelector(".batch-input")

  ];

  const currentIndex =
    rowInputs.indexOf(active);

  const nextIndex =
    currentIndex + direction;

  if (

    nextIndex >= 0

    &&

    nextIndex < rowInputs.length

  ) {

    rowInputs[nextIndex].focus();

  }

}
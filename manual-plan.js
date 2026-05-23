let originalData = [];

const excelFile = document.getElementById("excelFile");

const dataTableHead = document.querySelector("#dataTable thead");
const dataTableBody = document.querySelector("#dataTable tbody");

const planningTable = document.getElementById("planningTable");

const copyBtn = document.getElementById("copyBtn");

excelFile.addEventListener("change", handleFile);

function handleFile(e) {

    const file = e.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (event) {

        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, { type: "array" });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        const jsonData = XLSX.utils.sheet_to_json(firstSheet);

        originalData = jsonData;

        displayTable(originalData);

        generatePlanning(originalData);
    };

    reader.readAsArrayBuffer(file);
}

function displayTable(data) {

    dataTableHead.innerHTML = "";
    dataTableBody.innerHTML = "";

    if (data.length === 0) return;

    // TABLE HEADER
    const headers = Object.keys(data[0]);

    let headRow = "<tr>";

    headers.forEach(header => {
        headRow += `<th>${header}</th>`;
    });

    headRow += "</tr>";

    dataTableHead.innerHTML = headRow;

    // TABLE BODY
    data.forEach(row => {

        let bodyRow = "<tr>";

        headers.forEach(header => {
            bodyRow += `<td>${row[header] ?? ""}</td>`;
        });

        bodyRow += "</tr>";

        dataTableBody.innerHTML += bodyRow;
    });

}

function generatePlanning(data) {

    planningTable.innerHTML = "";

    const perRow = 5;

    let tr;

    data.forEach((row, index) => {

        if (index % perRow === 0) {

            tr = document.createElement("tr");

            planningTable.appendChild(tr);
        }

        const td = document.createElement("td");

        // AUTO ANALYZE EXCEL CONTENT
        // CHANGE COLUMN NAME HERE IF NEEDED

        const invoice =
            row.Invoice ||
            row.invoice ||
            row.INVOICE ||
            "";

        const qty =
            row.Qty ||
            row.qty ||
            row.QTY ||
            "";

        td.innerHTML = `
            <div>${qty} -- ${invoice}</div>
        `;

        tr.appendChild(td);

    });

}

copyBtn.addEventListener("click", copyPlanning);

function copyPlanning() {

    const range = document.createRange();

    range.selectNode(planningTable);

    window.getSelection().removeAllRanges();

    window.getSelection().addRange(range);

    try {
        document.execCommand("copy");
        alert("Planning copied! Paste into Excel.");
    }
    catch (err) {
        alert("Copy failed.");
    }

    window.getSelection().removeAllRanges();
}



document
  .getElementById("filterInvoice")
  .addEventListener("input", applyFilters);

document
  .getElementById("filterItem")
  .addEventListener("input", applyFilters);

document
  .getElementById("filterOperator")
  .addEventListener("input", applyFilters);

document
  .getElementById("filterDate")
  .addEventListener("input", applyFilters);


  function applyFilters(){

    const invoiceFilter =
        document
        .getElementById("filterInvoice")
        .value
        .toLowerCase();

    const itemFilter =
        document
        .getElementById("filterItem")
        .value
        .toLowerCase();

    const operatorFilter =
        document
        .getElementById("filterOperator")
        .value
        .toLowerCase();

    const dateFilter =
        document
        .getElementById("filterDate")
        .value
        .toLowerCase();

    const filteredData = originalData.filter(row => {

        const invoice =
            String(
                row.Invoice ||
                row.invoice ||
                ""
            ).toLowerCase();

        const item =
            String(
                row.Item ||
                row.item ||
                ""
            ).toLowerCase();

        const operator =
            String(
                row.Operator ||
                row.operator ||
                ""
            ).toLowerCase();

        const date =
            String(
                row.Date ||
                row.date ||
                ""
            ).toLowerCase();

        return (
            invoice.includes(invoiceFilter) &&
            item.includes(itemFilter) &&
            operator.includes(operatorFilter) &&
            date.includes(dateFilter)
        );

    });

    displayTable(filteredData);

    generatePlanning(filteredData);
}

function goBackPlanning(){

  window.location.href =
    "planning.html";

}
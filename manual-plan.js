let originalData = [];

const excelFile = document.getElementById("excelFile");

const dataTableHead = document.querySelector("#dataTable thead");
const dataTableBody = document.querySelector("#dataTable tbody");

const planningTable = document.getElementById("planningTable");

const copyBtn = document.getElementById("copyBtn");

const exportBtn =
    document.getElementById("exportBtn");

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

        populateFilters(originalData);
    };

    reader.readAsArrayBuffer(file);
}

function populateFilters(data) {

    populateSelect(
        "filterInvoice",
        data,
        ["Invoice", "invoice"]
    );

    populateSelect(
        "filterItem",
        data,
        ["Item", "item"]
    );

    populateSelect(
        "filterOperator",
        data,
        ["Operator", "operator"]
    );

    populateSelect(
        "filterDate",
        data,
        ["Date", "date"]
    );

}


function populateSelect(
    selectId,
    data,
    keys
) {

    const select =
        document.getElementById(selectId);

    const currentValue =
        select.value;

    select.innerHTML =
        `<option value="">All</option>`;

    const uniqueValues =
        [...new Set(

            data.map(row => {

                for (const key of keys) {

                    if (row[key]) {

                        return String(row[key]).trim();

                    }

                }

                return "";

            })

        )]

            .filter(Boolean)

            .sort((a, b) =>
                a.localeCompare(b)
            );

    uniqueValues.forEach(value => {

        const option =
            document.createElement("option");

        option.value = value;

        option.textContent = value;

        select.appendChild(option);

    });

    select.value = currentValue;

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

    const operatorOrder = [

        "R",
        "MR",
        "E",
        "P",
        "S",
        "MA",
        "A",
        "Y",
        "H",
        "SH",
        "MH"

    ];

    data.sort((a, b) => {

        const opA =
            String(
                a.Operator ||
                a.operator ||
                ""
            );

        const opB =
            String(
                b.Operator ||
                b.operator ||
                ""
            );

        return (
            operatorOrder.indexOf(opA)
            -
            operatorOrder.indexOf(opB)
        );

    });

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
    .addEventListener("change", applyFilters);

document
    .getElementById("filterItem")
    .addEventListener("change", applyFilters);

document
    .getElementById("filterOperator")
    .addEventListener("change", applyFilters);

document
    .getElementById("filterDate")
    .addEventListener("change", applyFilters);


function applyFilters() {

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

            (!invoiceFilter || invoice === invoiceFilter)

            &&

            (!itemFilter || item === itemFilter)

            &&

            (!operatorFilter || operator === operatorFilter)

            &&

            (!dateFilter || date === dateFilter)

        );

    });

    displayTable(filteredData);

    generatePlanning(filteredData);
}

function goBackPlanning() {

    window.location.href =
        "planning.html";

}




exportBtn.addEventListener(
    "click",
    exportDailyOuputExcel
);

function exportDailyOuputExcel() {

    const filteredData = getFilteredData();

    const grouped = {};

    filteredData.forEach(row => {

        const operator =
            row.Operator ||
            row.operator ||
            "";

        const item =
            row.Item ||
            row.item ||
            "";

        const qty =
            Number(
                row.Qty ||
                row.qty ||
                0
            );

        const invoice =
            row.Invoice ||
            row.invoice ||
            "";

        const key = `${operator}|||${item}`;

        if (!grouped[key]) {

            grouped[key] = {

                operator,
                item,
                totalQty: 0,
                matrix: []

            };

        }

        grouped[key].totalQty += qty;

        grouped[key].matrix.push(
            `${qty}--${invoice}`
        );

    });

    const exportData = [];

    Object.values(grouped).forEach(group => {

        const matrixRows = [];

        for (let i = 0; i < group.matrix.length; i += 5) {

            matrixRows.push(
                group.matrix.slice(i, i + 5)
            );

        }

        matrixRows.forEach((rowData, index) => {

            exportData.push({

                Operator:
                    index === 0
                        ? group.operator
                        : "",

                Item:
                    index === 0
                        ? group.item
                        : "",

                "Total Qty":
                    index === 0
                        ? group.totalQty
                        : "",

                Matrix1:
                    rowData[0] || "",

                Matrix2:
                    rowData[1] || "",

                Matrix3:
                    rowData[2] || "",

                Matrix4:
                    rowData[3] || "",

                Matrix5:
                    rowData[4] || ""

            });

        });

    });

    const worksheet =
        XLSX.utils.json_to_sheet(exportData);

    // COLUMN WIDTH
    worksheet["!cols"] = [

        { wch: 8 }, // Operator
        { wch: 28 }, // Item
        { wch: 8 }, // Total Qty
        { wch: 12 },
        { wch: 12 },
        { wch: 12 },
        { wch: 12 },
        { wch: 12 }

    ];

    const range =
        XLSX.utils.decode_range(
            worksheet["!ref"]
        );

    // STYLE ALL CELLS
    for (let R = range.s.r; R <= range.e.r; ++R) {

        for (let C = range.s.c; C <= range.e.c; ++C) {

            const cellAddress =
                XLSX.utils.encode_cell({
                    r: R,
                    c: C
                });

            if (!worksheet[cellAddress]) continue;

            // HEADER ROW
            if (R === 0) {

                worksheet[cellAddress].s = {

                    font: {
                        name: "Cambria",
                        sz: 12,
                        bold: true
                    },

                    alignment: {
                        vertical: "center",
                        horizontal: "center",
                        wrapText: true
                    },

                    border: {

                        top: {
                            style: "thick",
                            color: { rgb: "000000" }
                        },

                        bottom: {
                            style: "thick",
                            color: { rgb: "000000" }
                        },

                        left: {
                            style: "medium",
                            color: { rgb: "000000" }
                        },

                        right: {
                            style: "medium",
                            color: { rgb: "000000" }
                        }

                    }

                };

            }

            // NORMAL ROW
            else {

                worksheet[cellAddress].s = {

                    font: {
                        name: "Cambria",
                        sz: 12
                    },

                    alignment: {
                        vertical: "center",
                        horizontal: "center",
                        wrapText: true
                    },

                    border: {

                        top: {
                            style: "thin",
                            color: { rgb: "000000" }
                        },

                        bottom: {
                            style: "thin",
                            color: { rgb: "000000" }
                        },

                        left: {
                            style: "medium",
                            color: { rgb: "000000" }
                        },

                        right: {
                            style: "medium",
                            color: { rgb: "000000" }
                        }

                    }

                };

            }

        }

    }

    const workbook =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Planning"
    );

    XLSX.writeFile(
        workbook,
        "Welding Production Delivery Record.xlsx"
    );

}



function getFilteredData(){

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

    return originalData.filter(row => {

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

            (!invoiceFilter || invoice === invoiceFilter)

            &&

            (!itemFilter || item === itemFilter)

            &&

            (!operatorFilter || operator === operatorFilter)

            &&

            (!dateFilter || date === dateFilter)

        );

    });

}
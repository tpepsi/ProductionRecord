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

        populateFilters(originalData);

        populateYear();

        populateMonth();

        populateDate();



        populateMonth();

        populateDate();

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
            <div>${qty}--${invoice}</div>
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

    const yearFilter =
        document
            .getElementById("filterYear")
            .value;

    const monthFilter =
        document
            .getElementById("filterMonth")
            .value;

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

            &&

            (
                !yearFilter ||

                date.endsWith(yearFilter)
            )

            &&

            (
                !monthFilter ||

                monthOrder[
                Number(date.split("/")[1]) - 1
                ] === monthFilter
            )

        );

    });

    displayTable(filteredData);

    generatePlanning(filteredData);
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
        XLSX.utils.aoa_to_sheet([]);

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
    const sortedGroups =
        Object.values(grouped).sort((a, b) => {


            // 第一层：Operator 顺序

            const opA =
                operatorOrder.indexOf(a.operator);

            const opB =
                operatorOrder.indexOf(b.operator);


            if (opA !== opB) {

                return opA - opB;

            }


            // 第二层：同 Operator，Item A-Z

            return String(a.item)
                .localeCompare(
                    String(b.item)
                );


        });






        const workbook =
            XLSX.utils.book_new();

        const selectedDate =
            document
                .getElementById("filterDate")
                .value || "All Date";


        // 把整张表往下移两行
        XLSX.utils.sheet_add_aoa(
            worksheet,
            [
                [`Welding Production Output - ${selectedDate}`]
            ],
            {
                origin: "A1"
            }
        );


        // Merge A1:H1
        worksheet["!merges"] = [
            {
                s: { r: 0, c: 0 },   // A1
                e: { r: 0, c: 7 }    // H1
            }
        ];


        // Style A1
        worksheet["A1"].s = {

            font: {
                name: "Cambria",
                sz: 14,
                bold: true
            },

            alignment: {
                horizontal: "center",
                vertical: "center"
            }

        };

        // 原本资料从第4列开始
        XLSX.utils.sheet_add_json(

            worksheet,

            exportData,

            {

                origin: "A2",

                skipHeader: false

            }

        );

        // Column Width
        worksheet["!cols"] = [

            { wch: 8 },
            { wch: 28 },
            { wch: 8 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 }

        ];

        // Style
        const range = XLSX.utils.decode_range(worksheet["!ref"]);

        for (let R = range.s.r; R <= range.e.r; R++) {

            for (let C = range.s.c; C <= range.e.c; C++) {

                const addr = XLSX.utils.encode_cell({ r: R, c: C });

                if (!worksheet[addr]) continue;

                worksheet[addr].s = {

                    font: {

                        name: "Cambria",

                        sz: 12,

                        bold: R === 1

                    },

                    alignment: {

                        horizontal: "center",

                        vertical: "center",

                        wrapText: true

                    },

                    border: {

                        top: {

                            style: R === 1 ? "thick" : "thin",

                            color: { rgb: "000000" }

                        },

                        bottom: {

                            style: R === 1 ? "thick" : "thin",

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

        XLSX.utils.book_append_sheet(

            workbook,

            worksheet,

            "Welding Production Output"

        );

        const fileDate =
            selectedDate
                .replace(/\//g, "-");

        XLSX.writeFile(

            workbook,

            `Welding Production Output - ${fileDate}.xlsx`

        );

    }



function getFilteredData() {

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


function populateYear() {

            const select =
                document.getElementById("filterYear");

            select.innerHTML =
                `<option value="">All Years</option>`;

            const years =
                [...new Set(

                    originalData.map(row => {

                        const date =
                            row.Date ||
                            row.date ||
                            "";

                        if (!date) return "";

                        return date.split("/")[2];

                    })

                )]
                    .filter(Boolean)
                    .sort();

            years.forEach(year => {

                select.innerHTML +=
                    `<option value="${year}">
                ${year}
            </option>`;

            });

        }


const monthOrder = [

        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December"

    ];


    function populateMonth() {

        const year =
            document.getElementById("filterYear").value;

        const select =
            document.getElementById("filterMonth");

        select.innerHTML =
            `<option value="">All Months</option>`;

        let months =
            originalData.filter(row => {

                if (!year) return true;

                const date =
                    row.Date ||
                    row.date ||
                    "";

                return date.endsWith(year);

            });

        months =
            [...new Set(

                months.map(row => {

                    const date =
                        row.Date ||
                        row.date;

                    const month =
                        Number(date.split("/")[1]);

                    return monthOrder[month - 1];

                })

            )];

        months.sort(
            (a, b) =>
                monthOrder.indexOf(a) - monthOrder.indexOf(b)
        );

        months.forEach(month => {

            select.innerHTML +=

                `<option value="${month}">
            ${month}
        </option>`;

        });

    }


    function populateDate() {

        const year =
            document.getElementById("filterYear").value;

        const month =
            document.getElementById("filterMonth").value;

        const select =
            document.getElementById("filterDate");

        select.innerHTML =
            `<option value="">All Dates</option>`;

        let dates =
            originalData.filter(row => {

                const date =
                    row.Date ||
                    row.date ||
                    "";

                if (!date) return false;

                const parts =
                    date.split("/");

                const y = parts[2];

                const m =
                    monthOrder[
                    Number(parts[1]) - 1
                    ];

                return (

                    (!year || y === year)

                    &&

                    (!month || m === month)

                );

            });

        dates =
            [...new Set(

                dates.map(row =>

                    row.Date ||
                    row.date

                )

            )];

        dates.sort((a, b) => {

            const pa = a.split("/");

            const pb = b.split("/");

            return new Date(

                pa[2],
                pa[1] - 1,
                pa[0]

            ) -

                new Date(

                    pb[2],
                    pb[1] - 1,
                    pb[0]

                );

        });

        dates.forEach(date => {

            select.innerHTML +=

                `<option value="${date}">
            ${date}
        </option>`;

        });

    }





    document
        .getElementById("filterYear")
        .addEventListener("change", () => {

            populateMonth();

            populateDate();

            applyFilters();

        });

    document
        .getElementById("filterMonth")
        .addEventListener("change", () => {

            populateDate();

            applyFilters();

        });
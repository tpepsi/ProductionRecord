function getAllRecords(){

  return JSON.parse(
    localStorage.getItem("records")
  ) || [];

}

function renderPlanningPage(){

  const invoiceFilter =
    document.getElementById("filterInvoice")
    ?.value
    .trim()
    .toLowerCase() || "";

  const itemFilter =
    document.getElementById("filterItem")
    ?.value
    .trim()
    .toLowerCase() || "";

  const operatorFilter =
    document.getElementById("filterOperator")
    ?.value
    .trim()
    .toLowerCase() || "";

  const dateFilter =
    document.getElementById("filterDate")
    ?.value
    .trim()
    .toLowerCase() || "";

  let records = getAllRecords();

  records = records.filter(record => {

    const invoice =
      (record.invoice || "")
      .toLowerCase();

    const item =
      (record.item || "")
      .toLowerCase();

    const operator =
      (record.operator || "")
      .toLowerCase();

    const date =
      (record.date || "")
      .toLowerCase();

    return (

      invoice.includes(invoiceFilter)

      &&

      item.includes(itemFilter)

      &&

      operator.includes(operatorFilter)

      &&

      date.includes(dateFilter)

    );

  });

  renderLeftTable(records);

  renderPlanningMatrix(records);

}

function renderLeftTable(records){

  const body =
    document.getElementById(
      "planningRecordBody"
    );

  body.innerHTML = "";

  records.forEach(record => {

    body.innerHTML += `

      <tr>

        <td>${record.invoice || ""}</td>

        <td>${record.item || ""}</td>

        <td>${record.operator || ""}</td>

        <td>${record.date || ""}</td>

        <td>${record.qty || ""}</td>

        <td>${record.combine || ""}</td>

      </tr>

    `;

  });

}

function renderPlanningMatrix(records){

  const matrix =
    document.getElementById(
      "planningMatrix"
    );

  matrix.innerHTML = "";

  const combineList =
    records.map(record => {

      const qty =
        record.qty || "";

      const invoice =
        record.invoice || "";

      return `${qty}--${invoice}`;

    });

  const perRow = 5;

  for(let i = 0; i < combineList.length; i += perRow){

    const row =
      document.createElement("tr");

    combineList
      .slice(i, i + perRow)
      .forEach(value => {

        const td =
          document.createElement("td");

        td.textContent = value;

        row.appendChild(td);

      });

    matrix.appendChild(row);

  }

}

function copyPlanningTable(){

  const rows =
    document.querySelectorAll(
      "#planningMatrix tr"
    );

  let output = "";

  rows.forEach(row => {

    const cells =
      row.querySelectorAll("td");

    const rowText =
      [...cells]
      .map(td => td.innerText)
      .join("\t");

    output += rowText + "\n";

  });

  navigator.clipboard.writeText(output);

  alert(
    "Planning copied.\nPaste directly into Excel."
  );

}

document.addEventListener(
  "DOMContentLoaded",
  renderPlanningPage
);
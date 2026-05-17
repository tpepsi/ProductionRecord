let records = JSON.parse(localStorage.getItem("records")) || [];

let batchCodes = [];

renderTable(records);

renderSummary();

function handleBatch(event){

  if(event.key === "Enter"){

    event.preventDefault();

    const input =
      document.getElementById("batchInput");

    const value =
      input.value.trim().toUpperCase();

    if(value === ""){
      return;
    }

    if(!validateBatch(value)){

      document.getElementById("error").innerText =
        "Invalid Batch Format";

      return;

    }

    document.getElementById("error").innerText = "";

    batchCodes.push(value);

    document.getElementById("combinedCode").value =
      batchCodes.join(", ");

    calculateTotalQty();

    input.value = "";

  }

}

function validateBatch(batch){

  const normal = /^[A-Z]\d{6}--\d+$/;

  const noOperator = /^NA\d{6}--\d+$/;

  const noDate = /^[A-Z]NA--\d+$/;

  return (
    normal.test(batch) ||
    noOperator.test(batch) ||
    noDate.test(batch)
  );

}

function calculateTotalQty(){

  let total = 0;

  batchCodes.forEach(code => {

    const qty =
      Number(code.split("--")[1]);

    total += qty;

  });

  document.getElementById("totalQty").innerText =
    total;

}

function saveRecord(){

  const invoice =
    document.getElementById("invoice").value.trim();

  const item =
    document.getElementById("item").value.trim();

  const error =
    document.getElementById("error");

  error.innerText = "";

  if(invoice === "" || item === ""){

    error.innerText =
      "Please fill Invoice No and Item";

    return;

  }

  if(batchCodes.length === 0){

    error.innerText =
      "Please enter at least one Batch Code";

    return;

  }

  batchCodes.forEach(batch => {

    let operator = "NA";

    let date = "NA";

    const split =
      batch.split("--");

    const qty =
      split[1];

    const firstPart =
      split[0];

    // Y010125

    if(/^[A-Z]\d{6}$/.test(firstPart)){

      operator =
        firstPart.charAt(0);

      const rawDate =
        firstPart.substring(1);

      const day =
        rawDate.substring(0,2);

      const month =
        rawDate.substring(2,4);

      const year =
        rawDate.substring(4,6);

      date =
        `${day}/${month}/20${year}`;

    }

    // NA010125

    else if(/^NA\d{6}$/.test(firstPart)){

      operator = "NA";

      const rawDate =
        firstPart.substring(2);

      const day =
        rawDate.substring(0,2);

      const month =
        rawDate.substring(2,4);

      const year =
        rawDate.substring(4,6);

      date =
        `${day}/${month}/20${year}`;

    }

    // YNA

    else if(/^[A-Z]NA$/.test(firstPart)){

      operator =
        firstPart.charAt(0);

      date = "NA";

    }

    records.push({

      invoice,
      item,
      operator,
      date,
      qty,

      combine:
        `${qty}--${invoice}`

    });

  });

  localStorage.setItem(
    "records",
    JSON.stringify(records)
  );

  renderTable(records);

  renderSummary();

  newInv();

}

function renderTable(data){

  const table =
    document.getElementById("recordTable");

  table.innerHTML = "";

  data.forEach((record,index) => {

    table.innerHTML += `
      <tr>

        <td>${record.invoice}</td>

        <td>${record.item}</td>

        <td>${record.operator}</td>

        <td>${record.date}</td>

        <td>${record.qty}</td>

        <td>${record.combine}</td>

        <td>
          <button
            class="delete-btn"
            onclick="deleteRecord(${index})"
          >
            Delete
          </button>
        </td>

      </tr>
    `;

  });

}

function deleteRecord(index){

  records.splice(index,1);

  localStorage.setItem(
    "records",
    JSON.stringify(records)
  );

  renderTable(records);

  renderSummary();

}

function filterTable(){

  const keyword =
    document.getElementById("searchInput")
    .value
    .toLowerCase();

  const filtered =
    records.filter(record =>

      record.invoice.toLowerCase().includes(keyword)

      ||

      record.item.toLowerCase().includes(keyword)

      ||

      record.operator.toLowerCase().includes(keyword)

    );

  renderTable(filtered);

}

function renderSummary(){

  const summaryTable =
    document.getElementById("summaryTable");

  summaryTable.innerHTML = "";

  const summary = {};

  records.forEach(record => {

    if(!summary[record.operator]){

      summary[record.operator] = 0;

    }

    summary[record.operator] +=
      Number(record.qty);

  });

  for(let operator in summary){

    summaryTable.innerHTML += `
      <tr>
        <td>${operator}</td>
        <td>${summary[operator]}</td>
      </tr>
    `;

  }

}

function newInv(){

  document.getElementById("invoice").value = "";

  document.getElementById("item").value = "";

  document.getElementById("batchInput").value = "";

  document.getElementById("combinedCode").value = "";

  document.getElementById("totalQty").innerText = "0";

  document.getElementById("error").innerText = "";

  batchCodes = [];

}

function exportCSV(){

  let csv =
`Inv No,Item,Operator,Date,Quantity,Combine INV&QTY\n`;

  records.forEach(record => {

    csv +=
`${record.invoice},
${record.item},
${record.operator},
${record.date},
${record.qty},
${record.combine}\n`;

  });

  const blob =
    new Blob([csv], {type:'text/csv'});

  const link =
    document.createElement("a");

  link.href =
    URL.createObjectURL(blob);

  link.download =
    "production_record.csv";

  link.click();

}
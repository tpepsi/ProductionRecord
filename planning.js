function goBack(){
  window.location.href = "records.html";
}

const records = JSON.parse(localStorage.getItem("productionRecords")) || [];

const tbody = document.querySelector("#recordTable tbody");

records.forEach(record => {

  const tr = document.createElement("tr");

  tr.innerHTML = `
    <td>${record.invoice}</td>
    <td>${record.item}</td>
    <td>${record.operator}</td>
    <td>${record.date}</td>
    <td>${record.quantity}</td>
    <td>${record.combine}</td>
  `;

  tbody.appendChild(tr);
});

function generatePlanning(){

  const planningTable = document.querySelector("#planningTable tbody");

  planningTable.innerHTML = "";

  const combineData = records.map(r => r.combine);

  for(let i = 0; i < combineData.length; i += 5){

    const row = combineData.slice(i, i + 5);

    const tr = document.createElement("tr");

    row.forEach(data => {

      const td = document.createElement("td");

      td.textContent = data;

      tr.appendChild(td);
    });

    planningTable.appendChild(tr);
  }
}

function copyPlanning(){

  const rows = document.querySelectorAll("#planningTable tr");

  let output = "";

  rows.forEach(row => {

    const cols = row.querySelectorAll("td");

    let rowData = [];

    cols.forEach(col => {
      rowData.push(col.textContent);
    });

    output += rowData.join("\t") + "\n";
  });

  navigator.clipboard.writeText(output);

  alert("Planning data copied for Excel!");
}
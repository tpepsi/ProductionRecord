function splitBatchCodes(batchText){

  return batchText
    .split(/[,\n]/)
    .map(code => code.trim())
    .filter(code => code.length > 0);

}

function parseBatchCode(code){

  code = code.trim();

  if(!code){

    return {

      operator: "INVALID",

      mfgDate: "INVALID",

      quantity: "INVALID"

    };

  }

  let match =
    code.match(/^([A-Za-z]+)(\d{6})--(\d+)$/);

  if(match){

    return {

      operator: match[1],

      mfgDate: match[2],

      quantity: match[3]

    };

  }

  match =
    code.match(/^(\d{6})--(\d+)$/);

  if(match){

    return {

      operator: "NA",

      mfgDate: match[1],

      quantity: match[2]

    };

  }

  match =
    code.match(/^([A-Za-z]+)--(\d+)$/);

  if(match){

    return {

      operator: match[1],

      mfgDate: "NA",

      quantity: match[2]

    };

  }

  return {

    operator: "INVALID",

    mfgDate: "INVALID",

    quantity: "INVALID"

  };

}

function batchCodeDisplayLines(batchText){

  const batchList = splitBatchCodes(batchText);

  const lines = [];

  for(let i = 0; i < batchList.length; i += 4){

    const chunk = batchList.slice(i, i + 4);

    let line = chunk.join(", ");

    if(i + 4 < batchList.length){

      line += ",";

    }

    lines.push(line);

  }

  return lines;

}

function wrapBatchCode(batchText){

  return batchCodeDisplayLines(batchText)
    .map(line => `<span class="batch-line">${line}</span>`)
    .join("");

}

function formatBatchForExport(batchText){

  return batchCodeDisplayLines(
    batchText.replace(/\n/g, ",")
  ).join("\n");

}

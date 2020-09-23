const mysql = require("mysql");
const xl = require("excel4node");
const wb = new xl.Workbook();

const worksheetTitles = [
  "Domestic Standard Rates",
  "Domestic Expedited Rates",
  "Domestic Next Day Rates",
  "International Economy Rates",
  "International Expedited Rates",
];

const queries = [
  `\`locale\` = "domestic" AND \`shipping_speed\` = "standard"`,
  `\`locale\` = "domestic" AND \`shipping_speed\` = "expedited"`,
  `\`locale\` = "domestic" AND \`shipping_speed\` = "nextDay"`,
  `\`locale\` = "international" AND \`shipping_speed\` = "intlEconomy"`,
  `\`locale\` = "international" AND \`shipping_speed\` = "intlExpedited"`,
];

const style = wb.createStyle({
  font: {
    color: '#000000',
    size: 12,
  },
  numberFormat: '#0.00',
});

const worksheets = worksheetTitles.map((ws, i) => [wb.addWorksheet(ws), queries[i]]);

const connection = mysql.createConnection({
  user: "root",
  password: "my-secret-pw",
  database: "boingus"
});

connection.connect(function(err) {
  // 32 mins to connection. mysql had to be fed some insecure options given that the chosen package does not work w/ the default auth configuration in mysql 8. you can see the solve in ./config
  // the decision to not use an ORM was a heavy one to make: it was purely out of the time constraint. i figured 2 libraries were enough to learn in 2 hours, without adding the overhead of an ORM. that said, dealing with raw SQL is a fool's errand.
  // 57 mins: dealing with mysql. so particular with their quote syntax!
  // 1:30: creating the xls now...
  // 2:00 have the data ordered, and writing to the xls successfully. dealing w/ async probs..
  // 2:43 2/5 worksheets are populating correctly.
  // 2:58 all worksheets populating correctly. missing start/end weight
  // 3:10 all data is present. missing formatting details such as column width

  // the solution follows. please note the database name and password above.
  let promises = [];

  for (const [ws, qur] of worksheets) {
    promises.push(gatherData(ws, qur));
  }

  Promise.all(promises)
    .then(() => {
      console.log("writing to results.xlsx");
      wb.write('results.xlsx');
    });
})

function gatherData(ws, qur) {
  return new Promise((resolve, reject) => {
    const query = `\
    SELECT * FROM \`rates\`\
    WHERE \`client_id\` = 1240\
    AND ${qur}\
    ORDER BY \`zone\`, \`start_weight\`;`

    connection.query(query, function (error, results, fields) {
      let heightMod = 0;
      const z = results[0].zone;

      let weights = [],
        zones = new Set();

      while (results[heightMod] && results[heightMod].zone === z) {
        weights.push([results[heightMod].start_weight, results[heightMod].end_weight]);
        heightMod++;
      }

      const pair = i => [
        2+(i % heightMod), 
        3+(Math.floor(i/heightMod))
      ];

      for (let i = 0; i < results.length; i++) {
        let [x,y] = pair(i);
        ws.cell(x, y).number(results[i].rate).style(style);
        zones.add(results[i].zone);
      }

      populateHeader(ws, Array.from(zones));
      populateWeights(ws, weights);
      resolve();
    });

  })
}

function populateHeader(ws, zones) {
  ws.cell(1,1).string('Start Weight');
  ws.cell(1,2).string('End Weight');

  for (let i = 0; i < zones.length; i++) {
    ws.cell(1,i+3).string(`Zone ${zones[i]}`);
  }
}

function populateWeights(ws, weights) {
  for (let i = 0; i < weights.length; i++) {
    ws.cell(2+i,1).number(weights[i][0]).style(style);
    ws.cell(2+i,2).number(weights[i][1]).style(style);
  }
}
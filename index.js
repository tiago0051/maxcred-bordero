const mysql = require("mysql");
const Excel = require("exceljs");
const path = require("path");

require("dotenv").config();

const workbook = new Excel.Workbook();

const connection = mysql.createConnection(process.env);

connection.connect();

async function executeQuery(str, worksheet) {
  const results = await new Promise((resolve, reject) => {
    connection.query(str, (error, results, fields) => {
      if (error) reject(error);

      resolve({ results, fields });
    });
  });
  return { results, worksheet };
}

var date = new Date("2022-07-01");
const querys = [];

while (date < new Date("2023-08-01")) {
  date.setDate(1);
  date.setMonth(date.getMonth() + 1);

  const dateStr = `${date.getFullYear()}-${date.getMonth() + 1}`;

  console.log(date);

  const query = `SELECT caixa.idcaixa, caixa.dateTime as data, Contrato.idContrato, Contrato.ValorFace as face, Contrato.FatorCompra as compra, caixa.valor, Contrato.DataVencimento as vencimento, caixa.sobre FROM caixa
INNER JOIN Contrato ON Contrato.idContrato = caixa.idContrato
WHERE (dateTime > "${dateStr}-01" AND dateTime < "${dateStr}-31")
AND (sobre like "%COMPRA%") AND unidade = 3`;

  const worksheet = workbook.addWorksheet(dateStr);

  querys.push({ query, worksheet });
}

const promiseMap = querys.map(
  async (q) => await executeQuery(q.query, q.worksheet)
);

Promise.all(promiseMap)
  .then(async (q) => {
    q.map(async ({ results, worksheet }) => {
      const header = results.fields.map((f) => ({
        key: f.name,
        header: f.name,
      }));

      worksheet.columns = header;

      results.results.forEach((r) => {
        worksheet.addRow(r);
      });

      const exportPath = path.resolve(__dirname, "bordero.xlsx");

      await workbook.xlsx.writeFile(exportPath);
    });
  })
  .finally(() => connection.end());

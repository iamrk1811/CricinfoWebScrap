const request = require("request");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

const cricinfo = "https://www.espncricinfo.com/series";
const seriesId = "ipl-2020-21-1210595";

const handleBatsmanUtil = (cheerioHeader, content) => {
  singleBatsmanTable = [];
  cheerioHeader(content)
    .find("tr")
    .each((trIndex, tr) => {
      const row = [];
      cheerioHeader(tr)
        .find("td")
        .each((tdIndex, td) => {
          const style = cheerioHeader(td).attr("style") || "";
          if (!style.includes("display:none")) {
            row.push(cheerioHeader(td).text());
          }
        });
      // TODO
      if (row.length > 1) singleBatsmanTable.push(row);
    });
  return singleBatsmanTable;
};

const handleBatsman = (cheerioHeader, batsmanTables) => {
  const tables = [];
  const table1 = handleBatsmanUtil(cheerioHeader, batsmanTables[0]);
  const table2 = handleBatsmanUtil(cheerioHeader, batsmanTables[1]);
  tables.push(table1);
  tables.push(table2);
  return tables;
};

const handleBowlerUtil = (cheerioHeader, content) => {
  const singleBowlerTable = [];
  cheerioHeader(content)
    .find("tr")
    .each((trIndex, tr) => {
      const row = [];
      cheerioHeader(tr)
        .find("td")
        .each((tdIndex, td) => {
          const style = cheerioHeader(td).attr("style") || "";
          if (!style.includes("display:none")) {
            row.push(cheerioHeader(td).text());
          }
        });
      // TODO
      if (row.length > 1) singleBowlerTable.push(row);
    });
  return singleBowlerTable;
};

const handleBowler = (cheerioHeader, bowlerTables) => {
  const tables = [];
  const table1 = handleBowlerUtil(cheerioHeader, bowlerTables[0]);
  const table2 = handleBowlerUtil(cheerioHeader, bowlerTables[1]);
  tables.push(table1);
  tables.push(table2);
  return tables;
};

const workbook = new ExcelJS.Workbook();

const battingColumns = [
  { header: "Batting", key: "Batting", width: 10 },
  { header: "Out / Not Out", key: "Outnot", width: 10 },
  { header: "R", key: "R", width: 10 },
  { header: "B", key: "B", width: 10 },
  { header: "4S", key: "S4", width: 10 },
  { header: "6S", key: "S6", width: 10 },
  { header: "SR", key: "Sr", width: 10 },
];

const bowlerColumns = [
  { header: "Bowling", key: "Bowling", width: 10 },
  { header: "O", key: "O", width: 10 },
  { header: "M", key: "M", width: 10 },
  { header: "R", key: "R", width: 10 },
  { header: "W", key: "W", width: 10 },
  { header: "ECON", key: "Econ", width: 10 },
  { header: "0S", key: "S0", width: 10 },
  { header: "4S", key: "S4", width: 10 },
  { header: "6S", key: "S6", width: 10 },
  { header: "WD", key: "Wd", width: 10 },
  { header: "NB", key: "Nb", width: 10 },
];

const filldata = () => {
  request(
    cricinfo + "/" + seriesId + "/match-results",
    (error, response, body) => {
      if (!error && response.statusCode === 200) {
        const matchPage = cheerio.load(body);
        // storing all the scorecard links in the list
        let list = [];
        matchPage('[data-hover="Scorecard"]').each((index, element) => {
          list.push(
            "https://www.espncricinfo.com/" + matchPage(element).attr("href")
          );
        });

        // retrieve all innings name and create sheet names
        list.forEach((value, index) => {
          let id = value.replace("/full-scorecard", "").split("-");
          const sheetNames = [
            id[id.length - 1] + "-1-Batsman",
            id[id.length - 1] + "-1-Bowler",
            id[id.length - 1] + "-2-Batsman",
            id[id.length - 1] + "-2-Bowler",
          ];
          const data = [];
          request(value, (error, response, body) => {
            const inningsPage = cheerio.load(body);
            const batsmanTables = handleBatsman(
              inningsPage,
              inningsPage(".batsman")
            );

            const bowlerTables = handleBowler(
              inningsPage,
              inningsPage(".bowler")
            );
            data.push(batsmanTables[0]);
            data.push(bowlerTables[0]);
            data.push(batsmanTables[1]);
            data.push(bowlerTables[1]);
            for (let i = 0; i < 4; i++) {
              const sheet = workbook.addWorksheet(sheetNames[i]);
              const table = data[i];

              if (parseInt(i) % 2 == 0) {
                sheet.columns = battingColumns;
                // batsman code
                table.forEach((arr) => {
                  // arr represents a single row
                  const row = [];
                  arr.forEach((cellData) => {
                    row.push(cellData);
                  });
                  sheet.addRow(row);
                });
              } else {
                sheet.columns = bowlerColumns;
                table.forEach((arr) => {
                  // arr represents a single row
                  const row = [];
                  arr.forEach((cellData, index) => {
                    row.push(cellData);
                  });
                  sheet.addRow(row);
                });
              }
              sheet.getRow(1).eachCell((cell) => {
                // bold first row
                cell.font = { bold: true };
              });
            }
            if (index === list.length - 1) {
              workbook.xlsx.writeFile("rakib.xlsx");
              console.log("Saved");
            }
          });
        });
      }
    }
  );
};

filldata();

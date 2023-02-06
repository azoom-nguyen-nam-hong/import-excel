const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const app = express();
const XlsxPopulate = require("xlsx-populate");

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, file.fieldname + "-" + Date.now() + ".xlsx");
  },
});

const upload = multer({ storage: storage });

app.post("/upload", upload.single("file"), async (req, res) => {
  const type = req.body.type;
  let data;
  if (type === "member-sol") {
    data = await processMemberSol(req.file.path);
  } else if (type === "leader-sol") {
    data = await processLeaderSol(req.file.path);
  } else if (type === "kpi-sol") {
    data = await processKpiSol(req.file.path);
  } else if (type === "business") {
    data = await processBusinessDepartment(req.file.path);
  }
  res.send(data);
});

const getNumber = (str) => {
  if (typeof str === "number") {
    return str;
  }
  try {
    return Number(str.replace(/[^0-9-]+/g, ""));
  } catch (error) {
    return 0;
  }
};

const getMonth = (str) => {
  const year = str.split("年")[0];
  const month = str.split("年")[1].split("月")[0];
  if (month.length === 1) {
    return `${year}/0${month}`;
  }
  return `${year}/${month}`;
};

const division = (a, b) => {
  if (typeof a === "string") {
    a = a.replace(/[台件]/g, "");
  }
  if (typeof b === "string") {
    b = b.replace(/[台件]/g, "");
  }
  return a === 0 || b === 0 ? 0 : +(a / b).toFixed(2);
};

const processMemberSol = async (filePath) => {
  const result = XlsxPopulate.fromFileAsync(filePath).then((workbook) => {
    const data = workbook.sheet(0);
    let users = [];
    const colStart = 3;
    const maxColumn = data.usedRange().endCell().columnNumber();
    const maxRow = data.usedRange().endCell().rowNumber();
    const name = data.row(1).cell(2).value();
    for (let i = colStart; i <= maxColumn; i++) {
      let dataUser = {};
      for (let j = 2; j <= maxRow; j++) {
        if (data.row(j).cell(1).value() !== undefined) {
          if (data.row(j).cell(1).value() === "part") {
            dataUser[data.row(j).cell(1).value()] = getMonth(
              data.row(j).cell(i).value()
            );
            continue;
          } else {
            dataUser[data.row(j).cell(1).value()] = getNumber(
              data.row(j).cell(i).value()
            );
          }
          dataUser = { ...dataUser, name };
        }
      }
      users.push(dataUser);
    }

    let usersJson = [];
    for (const user of users) {
      usersJson.push({
        name: user["name"],
        part: user["part"],
        consignmentUnitBudget: user["budgetUnit"],
        numberOfUnitsEntrusted: user["actualUnits"],
        achievementRateOfUnits: division(
          user["actualUnits"],
          user["budgetUnit"]
        ),
        budgetForTheNumberOfSubdivisionsEntrusted: user["budgetMinutes"],
        numberOfUnitsEntrustedForSale: user["achievementsMinutes"],
        achievementRateOfUnits: division(
          user["achievementsMinutes"],
          user["budgetMinutes"]
        ),
        entrustedSimpleGrossProfitBudget: user["grossProfitTotalBudget"],
        consignmentSimpleGrossProfitResults: user["totalGrossProfit"],
        contractedSimpleGrossProfitAverage: division(
          user["totalGrossProfit"],
          user["actualUnits"]
        ),
        grossProfitRate: division(
          user["totalGrossProfit"],
          user["grossProfitTotalBudget"]
        ),
      });
    }
    return usersJson;
  });
  return result;
};

const processLeaderSol = async (filePath) => {
  const result = XlsxPopulate.fromFileAsync(filePath).then((workbook) => {
    const data = workbook.sheet(0);
    let users = [];
    const colStart = 3;
    const maxColumn = data.usedRange().endCell().columnNumber();
    const maxRow = data.usedRange().endCell().rowNumber();
    const name = data.row(1).cell(2).value();
    for (let i = colStart; i <= maxColumn; i++) {
      let dataUser = {};
      for (let j = 2; j <= maxRow; j++) {
        if (data.row(j).cell(1).value() !== undefined) {
          if (data.row(j).cell(1).value() === "part") {
            dataUser[data.row(j).cell(1).value()] = getMonth(
              data.row(j).cell(i).value()
            );
            continue;
          } else {
            dataUser[data.row(j).cell(1).value()] = getNumber(
              data.row(j).cell(i).value()
            );
          }
          dataUser = { ...dataUser, name };
        }
      }
      users.push(dataUser);
    }

    let result = [];
    for (const user of users) {
      result.push({
        name: user["name"],
        part: user["part"],
        budgetForNetIncreaseInVehicles: user["budgetUnit"],
        numberOfUnitsEntrusted: user["actualUnits"],
        achievementRateOfUnits: division(
          user["actualUnits"],
          user["budgetUnit"]
        ),
        reducedNumberOfVehicles: user["reducedNumberOfVehicles"],
        netIncrease: user["netIncrease"],
        achievementRateOfNetIncreaseInVehicles: division(
          user["netIncrease"],
          user["budgetUnit"]
        ),
        budgetForTheNumberOfSubdivisionsEntrusted: user["budgetMinutes"],
        numberOfUnitsEntrustedForSale: user["achievementsMinutes"],
        achievementRateOfNumberOfUnitsSold: division(
          user["achievementsMinutes"],
          user["budgetMinutes"]
        ),
        entrustedSimpleGrossProfitBudget: user["budgetUnit"] * 8000,
        consignmentSimpleGrossProfitResults: user["totalGrossProfit"],
        contractedSimpleGrossProfitAverage: division(
          user["totalGrossProfit"],
          user["actualUnits"]
        ),
        carReductionGrossProfit: user["budgetUnit"] * 8000,
        netIncreaseInGrossProfit:
          user["totalGrossProfit"] - user["budgetUnit"] * 8000,
        pullFullCarTotal: user["pullFullCarTotal"],
        pullAverageFullOccupancy: division(
          this.pullFullCarTotal,
          user["actualUnits"]
        ),
        grossProfitRate: division(
          user["totalGrossProfit"] - user["budgetUnit"] * 8000,
          user["budgetUnit"] * 8000
        ),
      });
    }
    return result;
  });
  return result;
};

const processKpiSol = async (filePath) => {
  const result = XlsxPopulate.fromFileAsync(filePath).then((workbook) => {
    const data = workbook.sheet(0);
    let users = [];
    const colStart = 2;
    const maxColumn = data.usedRange().endCell().columnNumber();
    const maxRow = data.usedRange().endCell().rowNumber();
    const name = data.row(1).cell(1).value();

    for (let j = 4; j <= maxRow; j++) {
      let dataUser = {};
      for (let i = colStart; i <= maxColumn; i++) {
        if (data.row(2).cell(i).value() !== undefined) {
          const month = getNumber(data.row(j).cell(1).value());
          dataUser[data.row(2).cell(i).value()] = getNumber(
            data.row(j).cell(i).value()
          );
          dataUser = { ...dataUser, name, month };
        }
      }
      if (dataUser["name"] !== undefined) {
        users.push(dataUser);
      }
    }

    let result = [];
    for (const user of users) {
      result.push({
        name: user["name"],
        month: user["month"],
        targetNumberOfProposals: user["goalKpi"],
        actualNumberOfProposals: user["total"],
        proposalNumberAchievementRate: division(user["total"], user["goalKpi"]),
        visitTarget: user["goalKdi"],
        actualNumberOfVisits: user["numberOfAppointments"],
        achievementRateOfVisits: division(
          user["numberOfAppointments"],
          user["goalKdi"]
        ),
      });
    }
    return result;
  });
  return result;
};

const processBusinessDepartment = async (filePath) => {
  const result = XlsxPopulate.fromFileAsync(filePath).then((workbook) => {
    const data = workbook.sheet(0);
    const dataArray = data
      .usedRange()
      .value()
      .map((item) => item.filter((item) => item !== undefined))
      .filter((item) => item.length !== 0);
    const month = dataArray[0][0].match(/(\d+)月/)[1];
    const salesSection = dataArray[0].filter(
      (item) => item === "営業１課" || item === "営業２課"
    ).length;

    let users = [];
    for (let i = 3; i < salesSection + 5; i++) {
      let dataUser = {};
      for (let j = 1; j < 33; j++) {
        if (data.row(j).cell(1).value() !== undefined) {
          dataUser[data.row(j).cell(1).value()] = data.row(j).cell(i).value();
          dataUser = { ...dataUser, month };
        }
      }
      users.push(dataUser);
    }

    let result = [];
    for (const user of users.slice(0, users.length - 2)) {
      const branch = getBranchNumber(user[dataArray[0][0]]);
      result.push({
        month,
        branch,
        data: {
          feeBudget: user["feeBudget"],
          feeRecord: user["feeRecord"],
          commissionAchievementRate: division(
            user["feeRecord"],
            user["feeBudget"]
          ),
          directManagementGrossProfitBudget: user["grossProfitBudget"],
          grossProfitPerformanceOfDirectManagement: user["grossProfitResults"],
          grossProfitAchievementRateOfDirectManagement: division(
            user["grossProfitResults"],
            user["grossProfitBudget"]
          ),
          numberOfUnitsUnderDirectManagementBudget: user["unitBudget"],
          actualNumberOfDirectlyManagedUnits: user["numberOfUnits"],
          achievementRateOfDirectlyManagedUnits: division(
            user["numberOfUnits"],
            user["unitBudget"]
          ),
          numberOfInquiries: user["numberOfInquiries"],
          numberOfDealsClosed: user["numberOfContractsClosed"],
          closingRate: division(
            user["numberOfContractsClosed"],
            user["numberOfInquiries"]
          ),
        },
      });
    }
    return result;
  });
  return result;
};

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Listening on port ${port}...`);
});

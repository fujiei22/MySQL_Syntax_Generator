const fs = require("fs");
const path = require("path");
const readline = require("readline");
const xlsx = require("xlsx");

// 建立 readline 介面
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

let workbook;  // 將 workbook 定義在 try 區塊外部

// 詢問使用者輸入 Excel 檔案名稱
function askForFileName() {
  rl.question("請輸入 Excel 檔案名稱：", (fileName) => {
    const filePath = path.join(__dirname, "xlsx", fileName);

    // 檢查檔案是否存在
    if (fs.existsSync(filePath)) {
      // 載入 Excel 檔案
      try {
        workbook = xlsx.readFile(filePath);  // 移到外部，以便在 selectSheet 中使用

        // 取得工作表名稱列表
        const sheetNames = workbook.SheetNames;

        // 列出工作表並讓使用者選擇
        selectSheet(sheetNames);
      } catch (error) {
        console.error("讀取檔案失敗，請重新輸入有效的 Excel 檔案名稱。");
        askForFileName();
      }
    } else {
      console.error(
        `檔案 ${fileName} 不存在，請重新輸入有效的 Excel 檔案名稱。`
      );
      askForFileName();
    }
  });
}

// 讓使用者選擇工作表
function selectSheet(sheetNames) {
  console.log("請選擇要操作的工作表：");
  sheetNames.forEach((sheet, index) => {
    console.log(`${index + 1}. ${sheet}`);
  });

  rl.question("請輸入選擇的工作表編號：", (selectedIndex) => {
    const sheetIndex = parseInt(selectedIndex) - 1;

    if (Number.isNaN(sheetIndex) || sheetIndex < 0 || sheetIndex >= sheetNames.length) {
      console.log("無效的選擇，請輸入有效的工作表編號。");
      selectSheet(sheetNames);
      return;
    }

    const selectedSheet = sheetNames[sheetIndex];

    // 指定工作表名稱
    const worksheet = workbook.Sheets[selectedSheet];

    // 將 Excel 資料轉換為 JSON 格式
    const excelData = xlsx.utils.sheet_to_json(worksheet);

    // 詢問使用者輸入資料表名稱
    askForTableName(excelData);
  });
}

// 詢問使用者輸入資料表名稱
function askForTableName(excelData) {
  // 取得 Excel 表頭
  const headers = Object.keys(excelData[0]);

  // 詢問使用者要生成的 SQL 語句類型
  rl.question("請輸入資料表名稱：", (tableName) => {
    // 產生唯一的檔案名稱，包含年月日時分秒
    const currentDate = new Date();
    currentDate.setHours(currentDate.getHours() + 8);  // 加 8 小時
    const formattedDate = currentDate
      .toISOString()
      .replace(/[-T:]/g, "")
      .split(".")[0];

    let outputFileName;
    let promptMessage;

    // 詢問使用者要生成的 SQL 語句類型
    rl.question(
      "要生成 INSERT (輸入 1) 還是 UPDATE (輸入 2) 語句？: ",
      (answer) => {
        if (answer === "1") {
          outputFileName = path.join(
            __dirname,
            "output",
            `insert_${formattedDate}.sql`
          );
          promptMessage = "INSERT SQL 語句已寫入到";
        } else if (answer === "2") {
          outputFileName = path.join(
            __dirname,
            "output",
            `update_${formattedDate}.sql`
          );
          promptMessage = "UPDATE SQL 語句已寫入到";
        } else {
          console.log(
            '無效的選擇。請輸入 "1" 生成 INSERT 語句，輸入 "2" 生成 UPDATE 語句。'
          );
          rl.close();
          return;
        }

        // 確保目錄存在，如果不存在就建立
        const outputDirectory = path.dirname(outputFileName);
        if (!fs.existsSync(outputDirectory)) {
          fs.mkdirSync(outputDirectory, { recursive: true });
        }

        // 開啟檔案以寫入 SQL 語法
        const outputFileStream = fs.createWriteStream(outputFileName);

        if (answer === "1") {
          // 生成 INSERT SQL 語句並寫入檔案
          for (const row of excelData) {
            const insertValues = headers
              .map((header) => sanitizeSqlValue(row[header]))
              .join(", ");
            const insertSql = `INSERT INTO ${tableName} (${headers.join(
              ", "
            )}) VALUES (${insertValues});`;

            // 寫入 SQL 語句到檔案
            outputFileStream.write(insertSql + "\n");
          }
        } else if (answer === "2") {
          // 詢問使用者要用作 WHERE 條件的表頭欄位
          rl.question("請輸入要用作 WHERE 條件的欄位：", (whereColumn) => {
            // 檢查使用者輸入的表頭欄位是否有效
            if (headers.includes(whereColumn)) {
              // 生成 UPDATE SQL 語句並寫入檔案
              for (const row of excelData) {
                const whereCondition = `${whereColumn} = '${sanitizeSqlValue(
                  row[whereColumn]
                )}'`;
                const updateValues = headers
                  .filter((header) => header !== whereColumn)
                  .map(
                    (header) => `${header} = '${sanitizeSqlValue(row[header])}'`
                  )
                  .join(", ");
                const updateSql = `UPDATE ${tableName} SET ${updateValues} WHERE ${whereCondition};`;

                // 寫入 SQL 語句到檔案
                outputFileStream.write(updateSql + "\n");
              }

              console.log(`${promptMessage} ${outputFileName}`);
            } else {
              console.log("無效的欄位名稱，請重新輸入。");
            }

            // 關閉檔案
            outputFileStream.end();
            rl.close();
          });
          return;
        }

        // 提示使用者輸入成功
        console.log(`${promptMessage} ${outputFileName}`);
        // 關閉檔案
        outputFileStream.end();
        rl.close();
      }
    );
  });
}

// 對 xlsx 資料中的特定符號進行處理
function sanitizeSqlValue(value) {
  if (typeof value === "string") {
    // 反引號轉換為單引號，再將單引號跳脫處理
    value = value.replace(/`/g, "'").replace(/'/g, "''");

    // 單引號進行跳脫處理
    value = value.replace(/'/g, "''");

    // 雙引號進行跳脫處理
    value = value.replace(/"/g, '""');

    return value;
  }
  return value;
}

// 啟動程式
askForFileName();
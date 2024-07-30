---
title: 讀書會專案分享：填表單自動生成文件 240730

---

# Project Sharing: Automated Document Generation via Form Filling 240730

What this project aims to do:
Every time we have a small gathering, we ask speakers to fill out a receipt. Previously, we collected information via email and had the speakers sign on the day of the event. However, with the advent of AI, we hope to gradually automate these processes. This time, we’ll be sharing the process of generating Google Docs automatically from form submissions!

# 讀書會專案分享：填表單自動生成文件 240730

### 這個專案想做什麼
每次小聚的時候，我們都會請講者填寫領據，以前是 email 收集資訊，當天請講者簽名，但現在有 AI 了，希望這些流程的東西都可以慢慢自動化，這次要來分享的就是從填表單到自動產生 google doc 文件的過程！


### 步驟

首先要有一個文件範本，我一開始是直接拿一個現有的領據文件，丟上 chatGPT 問他說我想要大家填寫 google 表單以後做出這樣的文件，請問我該做什麼，chatGPT 建議我把所有要代換掉的地方留下 placeholder，所以變成下圖這個樣子：

![Screenshot 2024-07-30 at 13.23.42](https://hackmd.io/_uploads/rJie0l8FR.png)

但接下來我按照 chatGPT 給的範例程式碼去執行並沒有成功，因此我在上次讀書會的時候，請教了 mosky，她就現場 demo 看看，原本以為會失敗，但直接成功，她精準而且有效的建議是：
1. 用英文問，不用擔心文法錯誤，AI 看得懂
2. 清楚的說出你想*使用的工具*跟你**想做的事情** 

而我發現她的問題跟我的問題最大的差別是，她有問出: “I want to **replace the values** *from google spreasheet to the google doc*” 

所以整個流程是：
1. 先做好你的 google form
2. 做好你的 google doc 文件範本，需要代換的欄位留下 placeholder
2. 將 form 連結至 spreadsheet
3. 在 google doc 設定 app script 
3. 執行完會成功得到一份將 spreadsheet 裡面的 value 代換進 placeholder 的文件

回家以後我也照著試試看，也差不多成功了。
但我跑了幾次以後發現，每次產生的檔案都叫做 xxx 檔案的副本，多了以後我看得很昏，所以我就再請 chatGPT 幫我加上自動改檔名的功能。
另外 spreasheet 可能有很多列，我希望可以指定某一列產生該列的文件。

![Screenshot 2024-07-30 at 13.42.11](https://hackmd.io/_uploads/r1_UCg8YC.png)


例如這個就是指定第二列產生文件，以下是結果：

![Screenshot 2024-07-30 at 13.36.05](https://hackmd.io/_uploads/SJjIRgLKR.png)

附上程式碼給大家參考：

```
function replaceValuesInDoc() {
  // Spreadsheet and Template IDs
  var spreadsheetId = '1Lc_39H_HU28EtuB1rm3r13ywuAIsrO4si_HCdAc1Uaw';
  var templateId = '1XomzcFjJPUBNvdjKQyMiZGftmM1gsDqFotwb3AIWiq4';

  // Specify the row number you want to process (1-based index)
  var rowToProcess = 1; // Process row 9

  // Open the spreadsheet and get the first sheet
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
  var data = sheet.getDataRange().getValues();

  // Log the number of rows in the sheet for debugging
  Logger.log('Total number of rows in the sheet: ' + data.length);

  // Check if the specified row exists
  if (rowToProcess < 1 || rowToProcess > data.length) {
    Logger.log('Error: Specified row ' + rowToProcess + ' does not exist.');
    return;
  }

  // Get the data for the specified row (convert 1-based to 0-based index)
  var rowIndex = rowToProcess - 1;
  var row = data[rowIndex];
  
  // Log the data for the specified row for debugging
  Logger.log('Data for row ' + rowToProcess + ': ' + JSON.stringify(row));

  var replacements = {
    '{{name}}': row[1],
    '{{id_number}}': row[2],
    '{{address}}': row[3],
    '{{phone_number}}': row[4],
    '{{email}}': row[5],
    '{{timestamp}}': row[0],
    '{{銀行名稱}}': row[6],
    '{{分行}}': row[8],
    '{{帳號}}': row[8],
    '{{戶名}}': row[9]
  };

 // Log replacements for debugging purposes
  Logger.log('Replacements: ' + JSON.stringify(replacements));

  // Create a name for the new document
  var docName = '李慕約公司領款收據_' + row[1]; // Example: "李慕約公司領款收據_name"
  
  // Log the new document name for debugging purposes
  Logger.log('Creating document: ' + docName);

  // Make a copy of the template with the new name
  var docCopy = DriveApp.getFileById(templateId).makeCopy(docName);
  var docCopyId = docCopy.getId();
  var doc = DocumentApp.openById(docCopyId);
  var body = doc.getBody();

  // Replace placeholders with actual values
  for (var key in replacements) {
    Logger.log('Replacing ' + key + ' with ' + replacements[key]);
    body.replaceText(key, replacements[key]);
  }

  // Save and close the document
  doc.saveAndClose();

  Logger.log('Document created successfully for row ' + rowToProcess);
}

```


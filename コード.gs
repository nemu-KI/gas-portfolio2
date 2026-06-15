function onFormSubmit(e) {
  if (!e) {
    showMessage("エラー: この関数はフォーム送信時に自動実行されます。エディタから直接実行することはできません。");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ====================================================
  // フォームの回答1シートから最新行のデータを取得
  // ====================================================
  const latestRowData = e.values;

  // 動的に列の位置を取得するために、ヘッダーを取得
  const formSheet = ss.getSheetByName("フォームの回答 1");
  if (!formSheet) {
    showMessage("エラー: 「フォームの回答 1」シートが見つかりません。");
    return;
  }
  const header = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

  // 各項目の列インデックスをすべて動的に取得
  const nameIndex = header.indexOf("氏名");
  const deptIndex = header.indexOf("部署");
  const destIndex = header.indexOf("出張先");
  const startIndex = header.indexOf("出張開始日");
  const endIndex = header.indexOf("出張終了日");
  const purposeIndex = header.indexOf("出張目的");
  const transitIndex = header.indexOf("交通費（円）");
  const lodgingIndex = header.indexOf("宿泊費（円）");
  const otherIndex = header.indexOf("その他経費（円）");
  const remarkIndex = header.indexOf("備考");

  if (transitIndex === -1 || lodgingIndex === -1 || otherIndex === -1 || nameIndex === -1 || remarkIndex === -1) {
    showMessage("エラー: フォーム回答シートに必要な列が見つかりません。ヘッダー名を確認してください。");
    return;
  }

  // ====================================================
  // 自動採番ロジック（年月-3桁連番）
  // ====================================================
  const historySheet = ss.getSheetByName("申請履歴");
  if (!historySheet) {
    showMessage("エラー: 「申請履歴」シートが見つかりません。");
    return;
  }

  // 現在の年月を取得 (例: "202606")
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0"); // 1月が0から始まるため+1し、2桁にゼロ埋め
  const currentYearMonth = `${year}${month}`;

  // 申請履歴シートのA列（申請番号）から最大連番を探す
  const lastRowHistory = historySheet.getLastRow();
  let maxSerialNumber = 0;

  if (lastRowHistory > 1) {
    // A列の2行目から最終行までの「申請番号」をまとめて取得
    const appNumbers = historySheet.getRange(2, 1, lastRowHistory - 1, 1).getValues();
    
    for (let i = 0; i < appNumbers.length; i++) {
      const fullNumber = appNumbers[i][0].toString(); // 例: "202606-002"
      
      // 現在の年月から始まっている場合のみ連番をチェック
      if (fullNumber.startsWith(currentYearMonth)) {
        // ハイフンの後ろの文字列（連番部分）を切り出して数値化
        const serialPart = Number(fullNumber.split("-")[1]);
        if (serialPart > maxSerialNumber) {
          maxSerialNumber = serialPart; // 最大値を更新
        }
      }
    }
  }

  // 最大連番+1で新しい番号を生成し、3桁にゼロ埋め
  const newSerialNumber = String(maxSerialNumber + 1).padStart(3, "0");
  const applicationNumber = `${currentYearMonth}-${newSerialNumber}`; // 完成: "202606-001"

  // ====================================================
  // 合計金額を計算
  // ====================================================
  const transitCost = Number(latestRowData[transitIndex]);
  const lodgingCost = Number(latestRowData[lodgingIndex]);
  const otherCost = Number(latestRowData[otherIndex]);

  const totalAmount = transitCost + lodgingCost + otherCost;

  // ====================================================
  // 「申請履歴」シートに1行追記（申請番号を先頭、合計経費を備考の前に挿入）
  // ====================================================
  // 備考より前のデータ、備考のデータを切り分けて組み替え
  const beforeRemarkData = latestRowData.slice(0, remarkIndex);
  const remarkData = latestRowData[remarkIndex];
  
  // 先頭に [applicationNumber] を結合する
  const outputRowData = [applicationNumber, ...beforeRemarkData, totalAmount, remarkData];

  // 申請履歴シートに追記
  historySheet.appendRow(outputRowData);

  // ====================================================
  // メールフォルダを取得（なければ自動生成）
  // ====================================================
  const ssId = ss.getId();
  const ssFile = DriveApp.getFileById(ssId);
  const parentFolder = ssFile.getParents().next();

  const folderName = "メール";
  const subFolders = parentFolder.getFoldersByName(folderName);
  let mailFolder;

  if (subFolders.hasNext()) {
    mailFolder = subFolders.next();
  } else {
    mailFolder = parentFolder.createFolder(folderName);
  }

  // メール文面への埋め込み用変数の整理
  const applicantName = latestRowData[nameIndex].toString().trim();
  const department = deptIndex !== -1 ? latestRowData[deptIndex] : "未設定";
  const destination = destIndex !== -1 ? latestRowData[destIndex] : "未設定";
  const purpose = purposeIndex !== -1 ? latestRowData[purposeIndex] : "未設定";
  const remark = remarkData !== "" ? remarkData : "特になし";

  const startDate = startIndex !== -1 ? latestRowData[startIndex] : "-";
  const endDate = endIndex !== -1 ? latestRowData[endIndex] : "-";

  const fTransit = transitCost.toLocaleString();
  const fLodging = lodgingCost.toLocaleString();
  const fOther = otherCost.toLocaleString();
  const fTotal = totalAmount.toLocaleString();

  // ====================================================
  // 申請者への受付確認テキストを生成・出力
  // ====================================================
  const applicantMailBody = `件名：出張経費申請を受け付けました

${applicantName} 様

以下の内容で出張経費申請を受け付けました。

【申請内容】
申請番号：${applicationNumber}
出張先：${destination}
出張期間：${startDate} ～ ${endDate}
出張目的：${purpose}
交通費：${fTransit}円
宿泊費：${fLodging}円
その他経費：${fOther}円
合計金額：${fTotal}円
備考：${remark}

審査完了後に改めてご連絡いたします。

総務部`;

  // ファイル名：「【申請者】申請番号_氏名.txt」
  mailFolder.createFile(`【申請者】${applicationNumber}_${applicantName}.txt`, applicantMailBody, MimeType.PLAIN_TEXT);

  // ====================================================
  // 経理担当者への申請内容通知テキストを生成・出力
  // ====================================================
  const keiriMailBody = `件名：出張経費申請が届きました

経理担当者 様

以下の出張経費申請が届きましたのでご確認ください。

【申請者情報】
申請番号：${applicationNumber}
氏名：${applicantName}
部署：${department}

【申請内容】
出張先：${destination}
出張期間：${startDate} ～ ${endDate}
出張目的：${purpose}
交通費：${fTransit}円
宿泊費：${fLodging}円
その他経費：${fOther}円
合計金額：${fTotal}円
備考：${remark}

総務部`;

  // ファイル名：「【経理】申請番号_氏名.txt」
  mailFolder.createFile(`【経理】${applicationNumber}_${applicantName}.txt`, keiriMailBody, MimeType.PLAIN_TEXT);
}

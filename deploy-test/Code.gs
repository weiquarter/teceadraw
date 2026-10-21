/** 部署測試：確認共用雲端硬碟裡的 Sheet 綁定 GAS 能以「任何人」存取部署，並回傳 JSON。 */

function doGet(e) {
  const ss = SpreadsheetApp.getActive();
  const out = {
    ok: true,
    message: '部署成功，匿名存取正常',
    sheetName: ss ? ss.getName() : null,
    timeZone: Session.getScriptTimeZone(),
    serverTime: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
    echo: (e && e.parameter && e.parameter.q) || null,
  };
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body = null;
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = String(err); }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, method: 'POST', body: body }))
    .setMimeType(ContentService.MimeType.JSON);
}

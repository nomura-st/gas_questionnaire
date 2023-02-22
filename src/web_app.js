/**
 * SHEETNAME_ROADSHOW のシートの情報を、GETの応答としてical形式で出力する
 * @param {*} e 引数など(対象の映画館)
 * @returns
 */
function doGet(e) {
  // *** データを取得する対象のSpreadSheet ***
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETNAME_ROADSHOW);
  let result = customTable(sheet, 2, 2, 9, false);
  let theater;
  if (e && e.parameter && e.parameter.theater) {
    switch (e.parameter.theater) {
      case "united_maebashi":
        theater = "ユナイテッド・シネマ 前橋";
        break;
      case "109_takasaki":
        theater = "109シネマズ高崎";
        break;
      case "aeon_takasaki":
        theater = "イオンシネマ高崎";
        break;
    }
  }
  let calendarName = `上映` + (theater ? `(${theater})` : "(全体)");

  if (theater) {
    result = result.filter((r) => r[7] == theater);
  }
  //  return ContentService.createTextOutput(JSON.stringify(result));
  let ievents = result
    .map((r) => {
      let type = r[1] ? `(${r[1]})` : "";
      let date = r[2].replaceAll("/", "");
      let st = r[4].replaceAll(":", "");
      let et = r[5].replaceAll(":", "");
      et = ("00" + et).slice(-6);
      return (
        `BEGIN:VEVENT\r\n` +
        `UID:${r[7]}_${r[0]}_${date}T${st}_${et}\r\n` +
        `DTSTAMP:${date}T${st}\r\n` +
        `DTSTART:${date}T${st}\r\n` +
        `DTEND:${date}T${et}\r\n` +
        `SUMMARY:${r[0]}${type}\r\n` +
        `LOCATION:${r[7]}\r\n` +
        `DESCRIPTION:${r[7]} (${r[8]})\r\n` +
        `END:VEVENT\r\n`
      );
    })
    .join("");

  let ical =
    `BEGIN:VCALENDAR\r\n` +
    `VERSION:2.0\r\n` +
    `X-WR-CALNAME:${calendarName}\r\n` +
    `X-WR-CALDESC:${calendarName}\r\n` +
    `X-WR-TIMEZONE:Asia/Tokyo\r\n` +
    `PRODID:${calendarName}\r\n` +
    `${ievents}` +
    `END:VCALENDAR\r\n`;
  console.log(ical);
  return ContentService.createTextOutput(ical);
}

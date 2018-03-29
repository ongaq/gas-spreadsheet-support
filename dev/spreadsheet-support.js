const moment = Moment.load();
const READ_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('READ_SPREADSHEET_ID');
const WRITE_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('WRITE_SPREADSHEET_ID');
const READ_SHEET_NAME = PropertiesService.getScriptProperties().getProperty('READ_SHEET_NAME');

// AM9〜10時にwriteToEXILEが毎日実行される
// つまりそれまでに体重を量ること。

// WithingがIFTTT経由で作ったスプレッドシートの最後に追記されたデータを取得
function getTodayMyBodyScale(){
	const sheet = SpreadsheetApp.openById(READ_SPREADSHEET_ID);
	const lastRow = sheet.getLastRow();
	const today = sheet.getActiveSheet().getRange(`A${lastRow}:E${lastRow}`).getValues();
	const data = {
		date: null,
		weight: null,
		muscle: null,
		fatQuant: null,
		fatPercent: null
	};
	Object.keys(data).forEach((key, pos) => {
		data[key] = today[0][pos];
	});
	data.date = moment(`${data.date.split(',')[0]}, 2018`).format('M/D');

	return data;
}
// EXILEに書き込む
function writeToEXILE(){
	const spread = SpreadsheetApp.openById(WRITE_SPREADSHEET_ID);
	const sheet = spread.getSheetByName(READ_SHEET_NAME);
	const myTodayData = getTodayMyBodyScale();
	const startPoint = 29; // AC2
	const exileDayRange = sheet.getRange(2, startPoint, 1, 100); // 29:129 = AC2:DX2
	const getDay = exileDayRange.getValues()[0];
	let columnNumber = 0;

	// getDayをM/D形式に変換し、getTodayMyBodyScaleで
	// 返ってきたdateと同じ値の位置を設定する
	getDay.forEach((key, pos) => {
		getDay[key] = moment(key).format('M/D');

		if (getDay[key] === myTodayData.date) {
			columnNumber = pos;
		}
	});
	// 起点に上で取得した位置を追加し、WithingがIFTTT経由で作った
	// スプレッドシートの最後に追記された日付をアクティベートする
	const todayColumn = sheet.getRange(2, startPoint+columnNumber);
	todayColumn.activate();
	todayColumn.offset(2, 0).setValue(myTodayData.weight); // 体重
	todayColumn.offset(3, 0).setValue(myTodayData.fatPercent); // 体脂肪率
}

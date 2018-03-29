const moment = Moment.load();
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
const today = moment().format('M/D');

// EXILEをチェックし記入漏れ野郎を探す
function checkToEXILE(){
	const spread = SpreadsheetApp.openById(SPREADSHEET_ID);
	const sheets = spread.getSheets();
	const sheetsName = [];
	const startPoint = 29; // AC2
	const kusoGuy = [];
	let columnNumber = 0;

	for (let i=0, len=sheets.length; i < len; i++) {
		const name = sheets[i].getSheetName();
		const sheet = spread.getSheetByName(name);
		
		if (columnNumber === 0) {
			const exileDayRange = sheet.getRange(2, startPoint, 1, 100); // 29:129 = AC2:DX2
			const getDay = exileDayRange.getValues()[0];
			
			// getDayをM/D形式に変換し、todayと同じ値の位置を設定する
			getDay.forEach((key, pos) => {
				getDay[key] = moment(key).format('M/D');
				
				if (getDay[key] === today) {
					columnNumber = pos;
				}
			});
		}
		// 起点に上で取得した位置を追加し、今日の日付をアクティベートし
		// 体重、体脂肪率の記入漏れを確認する
		let todayColumn = sheet.getRange(2, startPoint + columnNumber);
		todayColumn.activate();
		const checkWeight = String(todayColumn.offset(2, 0).getValue()); // 体重
		const checkFat = String(todayColumn.offset(3, 0).getValue()); // 体脂肪率

		// 体重、体脂肪率どちらかを記入漏れしていたらクソ男認定
		if (checkWeight.length === 0 || checkFat.length === 0) {
			kusoGuy[i] = name;
		}
	}
	return kusoGuy;
}
// LINE APIに送る
function sendHttpPost(message){
	const options = {
		'method': 'post',
		'payload': `message=${message}`,
		'headers': {
			'Authorization': `Bearer ${LINE_TOKEN}`
		}
	};
	UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}
// 時間がきたら処理するやつ
function processMainProgramWithTime(){
	const name = checkToEXILE();
	if (name.length) {
		const message = '記録漏れのデブを発見しました。多分'+name.join('と')+'です。';
		sendHttpPost(message);
	}
}
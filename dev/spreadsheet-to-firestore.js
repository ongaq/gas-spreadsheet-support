const moment = Moment.load();
const WRITE_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('WRITE_SPREADSHEET_ID');
const FIREBASE = {
	NAME: PropertiesService.getScriptProperties().getProperty('FIREBASE_PROJECT_NAME'),
	KEY: PropertiesService.getScriptProperties().getProperty('FIREBASE_API_KEY')
};
const myTodayData = moment().format('M/D');
const weekday = moment().day(); // 曜日

// シート名を配列で返す
function getSheetsName(spread){
	const sheetList = spread.getSheets();
	const sheetName = [];

	sheetList.forEach(sheet => sheetName.push(sheet.getName()));

	return sheetName;
}
// 今日の日付の列位置を取得する
function getTodayColumnPosition(data){
	let columnNumber = 0;

	data.forEach((key, pos) => {
		if (moment(key).format('M/D') === myTodayData) {
			columnNumber = pos;
		}
	});

	return columnNumber;
}
// シート名の漢字をローマ字に変換する
function convertIntoRoman(sheetName){
	let name = 'other';

	switch (sheetName) {
		case '吉田': name = 'yoshida'; break;
		case '平田': name = 'hirata'; break;
		case '丸山': name = 'maruyama'; break;
		default: break;
	}

	return name;
}
// 全日程のEXILEデータを読み取る（初回のみ実行）
function readTotalScheduleEXILE(){
	const spread = SpreadsheetApp.openById(WRITE_SPREADSHEET_ID);
	const sheetNames = getSheetsName(spread);
	const startPoint = 29; // AC2
	let columnNumber = 0;

	for (let i=0, len=sheetNames.length; i < len; i++) {
		const sheetName = sheetNames[i];
		const sheet = spread.getSheetByName(sheetName);
		const exileData = sheet.getRange(2, startPoint, 4, 100); // 日付、曜日、体重、体脂肪率を100日分
		const data = exileData.getValues();
		const name = convertIntoRoman(sheetName);

		// 今日の日付の列位置を取得する
		if (columnNumber === 0) {
			columnNumber = getTodayColumnPosition(data[0]);
		}

		for (let j=0; j <= columnNumber; j++) {
			const payload = preparePayload(data, j);
			Utilities.sleep(j*100);
			writeFirestore(name, payload.date, payload.payload);
		}
	}
}
// 本日のEXILEデータを読み取る（一日ごとに定期的に実行）
function readTodaysEXILE(){
	const spread = SpreadsheetApp.openById(WRITE_SPREADSHEET_ID);
	const sheetNames = getSheetsName(spread);
	const startPoint = 29; // AC2
	let columnNumber = 0;
	let payload = null;
	
	// 土日は書かない奴が多いので走査無し
	if (weekday === 0 || weekday === 6) return;

	for (let i=0, len=sheetNames.length; i < len; i++) {
		const sheetName = sheetNames[i];
		const sheet = spread.getSheetByName(sheetName);
		const exileData = sheet.getRange(2, startPoint, 4, 100); // 日付、曜日、体重、体脂肪率を100日分
		const data = exileData.getValues();
		const name = convertIntoRoman(sheetName);
		const writeProcess = (data, time, num) => {
			payload = preparePayload(data, num);
			Utilities.sleep(time*100);
			writeFirestore(name, payload.date, payload.payload);
		};

		// 今日の日付の列位置を取得する
		if (columnNumber === 0) {
			columnNumber = getTodayColumnPosition(data[0]);
		}
		
		// 月曜日は土日分をまとめて送信する
		if (weekday === 1) {
			for (let j=0; j < 3; j++) {
				writeProcess(data, j, (columnNumber-j));
			}
		} else {
			writeProcess(data, i, columnNumber);
		}
	}
}
// DBに書き込む用に値を整形
function preparePayload(data, num){
	const date = moment(data[0][num]).format('YYYY-MM-DD');
	const payload = {
		'fields': {
			'weight': { 'stringValue': String(data[2][num]) },
			'fat': { 'stringValue': String(data[3][num]) }
		}
	};

	return { date, payload };
}
// Firestoreに書き込む
function writeFirestore(name, day, payload){
	const documents = `users/${name}/${day}`;
	const url = `https://firestore.googleapis.com/v1beta1/projects/${FIREBASE.NAME}/databases/(default)/documents/${documents}?key=${FIREBASE.KEY}`;
	const options = {
		'method': 'POST',
		'contentType': 'application/json',
		'payload': JSON.stringify(payload),
		'muteHttpExceptions': true
	};

	UrlFetchApp.fetch(url, options);
}
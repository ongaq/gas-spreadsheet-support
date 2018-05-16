'use strict';

var moment = Moment.load();
var WRITE_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('WRITE_SPREADSHEET_ID');
var FIREBASE = {
	NAME: PropertiesService.getScriptProperties().getProperty('FIREBASE_PROJECT_NAME'),
	KEY: PropertiesService.getScriptProperties().getProperty('FIREBASE_API_KEY')
};
var myTodayData = moment().format('M/D');
var weekday = moment().day(); // 曜日

// シート名を配列で返す
function getSheetsName(spread) {
	var sheetList = spread.getSheets();
	var sheetName = [];

	sheetList.forEach(function (sheet) {
		return sheetName.push(sheet.getName());
	});

	return sheetName;
}
// 今日の日付の列位置を取得する
function getTodayColumnPosition(data) {
	var columnNumber = 0;

	data.forEach(function (key, pos) {
		if (moment(key).format('M/D') === myTodayData) {
			columnNumber = pos;
		}
	});

	return columnNumber;
}
// シート名の漢字をローマ字に変換する
function convertIntoRoman(sheetName) {
	var name = 'other';

	switch (sheetName) {
		case '吉田':
			name = 'yoshida';break;
		case '平田':
			name = 'hirata';break;
		case '丸山':
			name = 'maruyama';break;
		default:
			break;
	}

	return name;
}
// 全日程のEXILEデータを読み取る（初回のみ実行）
function readTotalScheduleEXILE() {
	var spread = SpreadsheetApp.openById(WRITE_SPREADSHEET_ID);
	var sheetNames = getSheetsName(spread);
	var startPoint = 29; // AC2
	var columnNumber = 0;

	for (var i = 0, len = sheetNames.length; i < len; i++) {
		var sheetName = sheetNames[i];
		var sheet = spread.getSheetByName(sheetName);
		var exileData = sheet.getRange(2, startPoint, 4, 100); // 日付、曜日、体重、体脂肪率を100日分
		var data = exileData.getValues();
		var name = convertIntoRoman(sheetName);

		// 今日の日付の列位置を取得する
		if (columnNumber === 0) {
			columnNumber = getTodayColumnPosition(data[0]);
		}

		for (var j = 0; j <= columnNumber; j++) {
			var payload = preparePayload(data, j);
			Utilities.sleep(j * 100);
			writeFirestore(name, payload.date, payload.payload);
		}
	}
}
// 本日のEXILEデータを読み取る（一日ごとに定期的に実行）
function readTodaysEXILE() {
	var spread = SpreadsheetApp.openById(WRITE_SPREADSHEET_ID);
	var sheetNames = getSheetsName(spread);
	var startPoint = 29; // AC2
	var columnNumber = 0;
	var payload = null;

	// 土日は書かない奴が多いので走査無し
	if (weekday === 0 || weekday === 6) return;

	var _loop = function _loop(i, len) {
		var sheetName = sheetNames[i];
		var sheet = spread.getSheetByName(sheetName);
		var exileData = sheet.getRange(2, startPoint, 4, 100); // 日付、曜日、体重、体脂肪率を100日分
		var data = exileData.getValues();
		var name = convertIntoRoman(sheetName);
		var writeProcess = function writeProcess(data, time, num) {
			payload = preparePayload(data, num);
			Utilities.sleep(time * 100);
			writeFirestore(name, payload.date, payload.payload);
		};

		// 今日の日付の列位置を取得する
		if (columnNumber === 0) {
			columnNumber = getTodayColumnPosition(data[0]);
		}

		// 月曜日は土日分をまとめて送信する
		if (weekday === 1) {
			for (var j = 0; j < 3; j++) {
				writeProcess(data, j, columnNumber - j);
			}
		} else {
			writeProcess(data, i, columnNumber);
		}
	};

	for (var i = 0, len = sheetNames.length; i < len; i++) {
		_loop(i, len);
	}
}
// DBに書き込む用に値を整形
function preparePayload(data, num) {
	var date = moment(data[0][num]).format('YYYY-MM-DD');
	var payload = {
		'fields': {
			'weight': { 'stringValue': String(data[2][num]) },
			'fat': { 'stringValue': String(data[3][num]) }
		}
	};

	return { date: date, payload: payload };
}
// Firestoreに書き込む
function writeFirestore(name, day, payload) {
	var documents = 'users/' + name + '/' + day;
	var url = 'https://firestore.googleapis.com/v1beta1/projects/' + FIREBASE.NAME + '/databases/(default)/documents/' + documents + '?key=' + FIREBASE.KEY;
	var options = {
		'method': 'POST',
		'contentType': 'application/json',
		'payload': JSON.stringify(payload),
		'muteHttpExceptions': true
	};
	var response = UrlFetchApp.fetch(url);
	var result = JSON.parse(JSON.stringify(response.getContentText()));

	// 本日分の書き込みがFirebase hosting側から無ければPOSTする
	if (!Object.keys(result).length) {
		return UrlFetchApp.fetch(url, options);
	}
}
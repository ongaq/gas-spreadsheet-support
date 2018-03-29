'use strict';

var moment = Moment.load();
var READ_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('READ_SPREADSHEET_ID');
var WRITE_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('WRITE_SPREADSHEET_ID');
var READ_SHEET_NAME = PropertiesService.getScriptProperties().getProperty('READ_SHEET_NAME');

// AM9〜10時にwriteToEXILEが毎日実行される
// つまりそれまでに体重を量ること。

// WithingがIFTTT経由で作ったスプレッドシートの最後に追記されたデータを取得
function getTodayMyBodyScale() {
	var sheet = SpreadsheetApp.openById(READ_SPREADSHEET_ID);
	var lastRow = sheet.getLastRow();
	var today = sheet.getActiveSheet().getRange('A' + lastRow + ':E' + lastRow).getValues();
	var data = {
		date: null,
		weight: null,
		muscle: null,
		fatQuant: null,
		fatPercent: null
	};
	Object.keys(data).forEach(function (key, pos) {
		data[key] = today[0][pos];
	});
	data.date = moment(data.date.split(',')[0] + ', 2018').format('M/D');

	return data;
}
// EXILEに書き込む
function writeToEXILE() {
	var spread = SpreadsheetApp.openById(WRITE_SPREADSHEET_ID);
	var sheet = spread.getSheetByName(READ_SHEET_NAME);
	var myTodayData = getTodayMyBodyScale();
	var startPoint = 29; // AC2
	var exileDayRange = sheet.getRange(2, startPoint, 1, 100); // 29:129 = AC2:DX2
	var getDay = exileDayRange.getValues()[0];
	var columnNumber = 0;

	// getDayをM/D形式に変換し、getTodayMyBodyScaleで
	// 返ってきたdateと同じ値の位置を設定する
	getDay.forEach(function (key, pos) {
		getDay[key] = moment(key).format('M/D');

		if (getDay[key] === myTodayData.date) {
			columnNumber = pos;
		}
	});
	// 起点に上で取得した位置を追加し、WithingがIFTTT経由で作った
	// スプレッドシートの最後に追記された日付をアクティベートする
	var todayColumn = sheet.getRange(2, startPoint + columnNumber);
	todayColumn.activate();
	todayColumn.offset(2, 0).setValue(myTodayData.weight); // 体重
	todayColumn.offset(3, 0).setValue(myTodayData.fatPercent); // 体脂肪率
}
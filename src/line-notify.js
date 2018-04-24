'use strict';

var moment = Moment.load();
var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
var LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
var today = moment().format('M/D');
var GWIsDoNotNotify = ['4/28', '4/29', '4/30', '5/1', '5/2', '5/3', '5/4', '5/5', '5/6'];
var weekday = moment().day(); // 曜日

// EXILEをチェックし記入漏れ野郎を探す
function checkToEXILE() {
	var spread = SpreadsheetApp.openById(SPREADSHEET_ID);
	var sheets = spread.getSheets();
	var sheetsName = [];
	var startPoint = 29; // AC2
	var kusoGuy = [];
	var columnNumber = 0;

	for (var i = 0, len = sheets.length; i < len; i++) {
		var name = sheets[i].getSheetName();
		var sheet = spread.getSheetByName(name);

		if (columnNumber === 0) {
			(function () {
				var exileDayRange = sheet.getRange(2, startPoint, 1, 100); // 29:129 = AC2:DX2
				var getDay = exileDayRange.getValues()[0];

				// getDayをM/D形式に変換し、todayと同じ値の位置を設定する
				getDay.forEach(function (key, pos) {
					getDay[key] = moment(key).format('M/D');

					if (getDay[key] === today) {
						columnNumber = pos;
					}
				});
			})();
		}
		// 起点に上で取得した位置を追加し、今日の日付をアクティベートし
		// 体重、体脂肪率の記入漏れを確認する
		var todayColumn = sheet.getRange(2, startPoint + columnNumber);
		todayColumn.activate();
		var checkWeight = String(todayColumn.offset(2, 0).getValue()); // 体重
		var checkFat = String(todayColumn.offset(3, 0).getValue()); // 体脂肪率

		// 体重、体脂肪率どちらかを記入漏れしていたらクソ男認定
		if (checkWeight.length === 0 || checkFat.length === 0) {
			kusoGuy.push(name);
		}
	}
	return kusoGuy;
}
// LINE APIに送る
function sendHttpPost(message) {
	var options = {
		'method': 'post',
		'payload': 'message=' + message,
		'headers': {
			'Authorization': 'Bearer ' + LINE_TOKEN
		}
	};
	UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}
// 時間がきたら処理するやつ
function processMainProgramWithTime() {
	// GW、または土日なら処理を中止
	if (GWIsDoNotNotify.indexOf(today) !== -1 || weekday === 0 || weekday === 6) return;

	var name = checkToEXILE();
	if (name.length) {
		var message = '記録漏れのデブを発見しました。多分' + name.join('と') + 'です。';
		sendHttpPost(message);
	}
}
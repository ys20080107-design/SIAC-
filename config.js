/*
 * きづきの診断 — 設定
 * sheetUrl: Google Apps Script の「ウェブアプリ」URL（https://script.google.com/macros/s/…/exec）。
 *           空のままなら送信は行わず、各端末で Excel 保存だけが使えます。手順は README.md を参照。
 * group:    既定のグループ名（イベント名など）。URLの ?g=名前 が優先されます。
 */
window.KIZUKI_CONFIG = {
  sheetUrl: '',
  group: ''
};

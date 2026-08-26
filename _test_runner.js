const XLSX = require("./node_modules/xlsx");
global.XLSX = XLSX;
const logic = require("./_test_pure_logic.js");
// pure_logic.js has no module.exports, so use vm to grab its scope instead.

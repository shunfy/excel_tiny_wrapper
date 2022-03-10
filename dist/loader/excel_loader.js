"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelLoader = void 0;
const xlsx = require("xlsx");
const utils = require("../utils");
class ExcelLoader {
    constructor(worksheet, _sheetName, _fileName) {
        this.worksheet = worksheet;
        this._sheetName = _sheetName;
        this._fileName = _fileName;
    }
    getData(c, r) {
        const cell = xlsx.utils.encode_cell({ c, r });
        return this.worksheet[cell];
    }
    getRange() {
        if (this.worksheet['!ref'] == undefined) {
            utils.debug(`- Pass Sheet "${this.sheetName}" : Sheet is empty`);
            return;
        }
        return xlsx.utils.decode_range(this.worksheet['!ref']);
    }
    get sheetName() { return this._sheetName; }
    get fileName() { return this._fileName; }
}
exports.ExcelLoader = ExcelLoader;
//# sourceMappingURL=excel_loader.js.map
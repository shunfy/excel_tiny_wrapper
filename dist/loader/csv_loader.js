"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CSVLoader = void 0;
const csv_parse = require("csv-parse");
const fs = require("fs-extra-promise");
const path = require("path");
const utils = require("../utils");
const lodash_1 = require("lodash");
class CSVLoader {
    constructor(_sheetName, _fileName, records, range) {
        this._sheetName = _sheetName;
        this._fileName = _fileName;
        this.records = records;
        this.range = range;
    }
    static load(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            const sheetName = path.parse(fileName).name;
            let buffer = yield fs.readFileAsync(fileName);
            return new Promise((resolve, reject) => {
                csv_parse.parse(buffer, { autoParse: false, autoParseDate: false, cast: false, }, (err, records, info) => {
                    let success = err == undefined && lodash_1.isArray(records) && records.length > 0;
                    if (success) {
                        utils.AddIgnoreSheet(sheetName);
                        let csv = new CSVLoader(sheetName, fileName, records, {
                            s: {
                                c: 0, r: 0
                            }, e: {
                                c: records[0].length, r: records.length
                            }
                        });
                        resolve(csv);
                    }
                    else {
                        reject(`parse csv file: ${fileName} failure.${err === null || err === void 0 ? void 0 : err.message}`);
                    }
                });
            });
        });
    }
    getData(c, r) {
        if (this.records == undefined || this.records[r] == undefined || this.records[r][c] == undefined)
            return undefined;
        return { w: this.records[r][c], t: 's' };
    }
    getRange() {
        return this.range;
    }
    get sheetName() { return this._sheetName; }
    get fileName() { return this._sheetName; }
}
exports.CSVLoader = CSVLoader;
//# sourceMappingURL=csv_loader.js.map
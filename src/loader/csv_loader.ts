import csv_parse = require('csv-parse');
import fs = require('fs-extra-promise');
import path = require('path');
import { CellObject, Range } from 'xlsx/types';
import { IDataLoader } from './idata_loader';
import utils = require('../utils');
import xlsx = require('xlsx');
import { isArray } from 'lodash';

export class CSVLoader implements IDataLoader {
	private constructor(private _sheetName: string, private _fileName: string, private records: any | undefined, private range: xlsx.Range) {
	}

	public static async load(fileName: string): Promise<CSVLoader> {
		const sheetName = path.parse(fileName).name;
		let buffer = await fs.readFileAsync(fileName);
		return new Promise<CSVLoader>((resolve, reject) => {
			csv_parse.parse(buffer, { autoParse: false, autoParseDate: false, cast: false, },
				(err: csv_parse.CsvError | undefined, records: any | undefined, info: csv_parse.Info) => {
					let success = err == undefined && isArray(records) && records.length > 0;
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
					} else {
						reject(`parse csv file: ${fileName} failure.${err?.message}`);
					}
				});
		});
	}

	public getData(c: number, r: number): CellObject | undefined {
		if (this.records == undefined || this.records[r] == undefined || this.records[r][c] == undefined) return undefined;
		return { w: this.records[r][c], t: 's' };
	}

	public getRange(): Range | undefined {
		return this.range;
	}

	public get sheetName(): string { return this._sheetName; }
	public get fileName(): string { return this._sheetName; }
}
import xlsx = require('xlsx');
import utils = require('../utils');
import { IDataLoader } from './idata_loader';

export class ExcelLoader implements IDataLoader {
	constructor(private worksheet: xlsx.WorkSheet, private _sheetName: string, private _fileName: string) {
	}

	getData(c: number, r: number): xlsx.CellObject | undefined {
		const cell = xlsx.utils.encode_cell({ c, r });
		return this.worksheet[cell];
	}

	getRange(): xlsx.Range | undefined {
		if (this.worksheet['!ref'] == undefined) {
			utils.debug(`- Pass Sheet "${this.sheetName}" : Sheet is empty`);
			return;
		}
		return xlsx.utils.decode_range(<string>this.worksheet['!ref']);
	}

	public get sheetName(): string { return this._sheetName; }
	public get fileName(): string { return this._fileName; }
}
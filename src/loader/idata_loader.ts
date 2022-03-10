import xlsx = require('xlsx');

export interface IDataLoader {
	// 获取数据
	getData(c: number, r: number): xlsx.CellObject | undefined;

	// 获取表格范围
	getRange(): xlsx.Range | undefined;

	// 表格名称
	sheetName: string;
	// 文件名称
	fileName: string;
}
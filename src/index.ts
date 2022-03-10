#!/usr/bin/env node

import * as fs from 'fs';
import { execute } from './works';
import * as utils from './utils';
import * as config from './config';


////////////////////////////////////////////////////////////////////////////////
function printHelp() {
	console.log(false, `${process.argv[0]} ${process.argv[1]} <config file path(optional)> <type_extens_checker file path>`);
}

async function main() {
	try {
		const configPath = process.argv.length >= 3 ? process.argv[2] : undefined;
		const type_extens_checker_Path = process.argv.length >= 4 ? process.argv[3] : undefined;
		if (configPath == undefined || configPath == '-h' || configPath == '--h' || configPath == '/h' || configPath == '/?') {
			printHelp();
			return;
		} else if (!fs.existsSync(configPath)) {
			printHelp();
			return;
		}
		if (utils.StrNotEmpty(type_extens_checker_Path)) {
			try {
				utils.yellow(`---------: ${type_extens_checker_Path}`);
				config.setTypeCheckerJSFilePath(type_extens_checker_Path);
				utils.yellow("-------------------------");
			} catch (ex) {
				utils.red(`${ex}`);
				printHelp();
				return;
			}
		}
		if (!config.InitGlobalConfig(configPath)) {
			utils.exception(`Init Global Config "${configPath}" Failure.`);
			return;
		}
		await execute();
		console.log('--------------------------------------------------------------------');
	} catch (ex) {
		utils.exception(`${ex}`);
		process.exit(utils.E_ERROR_LEVEL.EXECUTE_FAILURE);
	}
}

// main entry
main();

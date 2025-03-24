/**
 * © 2025 nagaoka.aya
 */

import * as vscode from 'vscode';
import { converse } from './converse';
import path from 'path';

/**
 * 拡張機能がアクティブ化されたときに呼び出される関数
 * @param context - 拡張機能のコンテキスト
 */
export function activate(context: vscode.ExtensionContext) {
	const disposable = vscode.commands.registerCommand('excel-reader.excelToMarkdown', async (...commandArgs) => {

		const args = extractExcelUris(commandArgs);

		if (args.length == 0) {
			vscode.window.showErrorMessage('Excelファイルを選択してください');
			return;
		}

		vscode.window.withProgress({ location: vscode.ProgressLocation.Window, title: "ExcelからMarkdownに変換中" }, async progress => {
			for (let i = 0; i < args.length; i++) {
				await processExcelFile(args[i]);
			}
		});

	});

	context.subscriptions.push(disposable);
}

/**
 * 拡張機能が非アクティブ化されたときに呼び出される関数
 */
export function deactivate() { }

/**
 * Excelファイルを処理してMarkdownに変換する
 * @param uri - 処理するExcelファイルのURI
 * @returns Promise<void>
 */
async function processExcelFile(uri: vscode.Uri): Promise<void> {
	let excelPath = uri.path;
	if (excelPath.startsWith('/')) {
		excelPath = excelPath.substring(1);
	}

	vscode.window.showInformationMessage(excelPath + "を変換します");
	const config = vscode.workspace.getConfiguration('excelToMarkdown');

	const scriptPath = config.get('scriptPath') as string;
	const result = await executeScriptInTerminal(scriptPath, excelPath);

	if (result.exitCode == 0) {
		const configValue = config.get('ignoreSheets') as string[];
		const outputFolder = config.get('outputFolder') as string;
		const extendedPrompt = config.get('extendedPrompt') as string;
		// Extract the filename from the Excel path and remove the extension
		const excelFileName = path.basename(excelPath).split(".").slice(0, -1).join(".")
		const outputFolderPath = outputFolder + "/" + excelFileName;
		await converse(result.consoleLog, configValue, outputFolderPath, extendedPrompt);
	} else {
		vscode.window.showErrorMessage('Failed to run the script');
	}
}

/**
 * コマンド引数からExcelファイルのURIを抽出する
 * @param commandArgs - コマンド引数の配列
 * @returns ExcelファイルのURIの配列
 */
function extractExcelUris(commandArgs: any[]): vscode.Uri[] {
	const args: vscode.Uri[] = [];

	for (let i = 0; i < commandArgs.length; i++) {
		if (commandArgs[i] instanceof vscode.Uri) {
			const uri = commandArgs[i] as vscode.Uri;
			const filePath = uri.fsPath;
			if (path.extname(filePath).toLowerCase() === '.xlsx') {
				args.push(uri);
			}
		}
	}

	return args;
}

/**
 * ターミナルでスクリプトを実行する
 * @param pathToScript - 実行するスクリプトのパス
 * @param pathToExcel - 処理するExcelファイルのパス
 * @returns 終了コードとコンソールログを含むオブジェクト
 */
async function executeScriptInTerminal(pathToScript: string, pathToExcel: string): Promise<{ exitCode: number | undefined, consoleLog: string }> {
	try {
		const testRunTerminal = vscode.window.createTerminal('Excel Read Terminal');
		testRunTerminal.show();

		return new Promise((resolve, reject) => {
			const dispose = vscode.window.onDidChangeTerminalShellIntegration(async ({ terminal }) => {
				if (terminal === testRunTerminal) {
					dispose.dispose();
					if (terminal.shellIntegration) {
						const buildResult = await runScript(terminal.shellIntegration, pathToScript + " '" + pathToExcel + "'");
						resolve(buildResult);
					} else {
						vscode.window.showErrorMessage('Shell integration is not available.');
						reject({
							exitCode: undefined,
							consoleLog: ''
						});
					}
				}
			});
		});

	} catch (error: any) {
		vscode.window.showErrorMessage(`Failed to run tests: ${error.message}`);
	}

	return {
		exitCode: undefined,
		consoleLog: ''
	};
}

/**
 * ターミナルでコマンドを実行し、結果を取得する
 * @param terminal - ターミナルシェル統合オブジェクト
 * @param command - 実行するコマンド
 * @returns 終了コードとコンソールログを含むオブジェクト
 */
async function runScript(terminal: vscode.TerminalShellIntegration, command: string): Promise<{ exitCode: number | undefined, consoleLog: string }> {

	try {
		return new Promise((resolve, reject) => {
			const execution = terminal.executeCommand(command);
			const stream = execution.read();
			let consoleLog = '';

			const didEndDispose = vscode.window.onDidEndTerminalShellExecution(async event => {
				if (event.execution === execution) {
					didEndDispose.dispose();
					for await (const data of stream) {
						consoleLog += decodeANSICode(data);
					}
					resolve({
						exitCode: event.exitCode,
						consoleLog
					});
				}
			});

		});

	} catch (error: any) {
		vscode.window.showErrorMessage(`Failed to run ${command}: ${error.message}`);
	}

	return {
		exitCode: undefined,
		consoleLog: ''
	};
}

/**
 * ANSIエスケープコードをデコードしてプレーンテキストに変換する
 * 参考：https://github.com/chalk/ansi-regex
 * @param text - デコードするANSIエスケープコードを含むテキスト
 * @returns デコード後のプレーンテキスト
 */
function decodeANSICode(text: string): string {
	const ST = '(?:\\u0007|\\u001B\\u005C|\\u009C)';
	const pattern = [
		`[\\u001B\\u009B][[\\]()#;?]*(?:(?:(?:(?:;[-a-zA-Z\\d\\/#&.:=?%@~_]+)*|[a-zA-Z\\d]+(?:;[-a-zA-Z\\d\\/#&.:=?%@~_]*)*)?${ST})`,
		'(?:(?:\\d{1,4}(?:;\\d{0,4})*)?[\\dA-PR-TZcf-nq-uy=><~]))',
	].join('|');

	const reg = new RegExp(pattern, 'g');
	const decodedText = text.replace(reg, '');
	return decodedText;
}

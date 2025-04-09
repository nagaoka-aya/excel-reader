/**
 * © 2025 nagaoka.aya
 */
import * as vscode from 'vscode';

// AIがコードを生成するためのプロンプト
const CONVERSE_PROMPT = `
<excel>
{EXCEL_DATA}
</excel>
excelタグで与えられた情報は、システム設計書のエクセルファイルを変換したものです。
1列目が行番号、2列目が列番号、3列目がそのセルの内容になっています。
内容を解析し、Markdownに変換してください。その際、インデントや文章を忠実に書き起こしてください。表で表現されていると思われる部分は表形式で出力してください。

{extendedPrompt}

`
    + 'Return only the markdown text.'
    + 'Do not include markdown "```" or "```html" at the start or end.'

/**
 * Excelデータを受け取り、AIを使用してMarkdownに変換して保存する
 * @param excelData 処理対象のExcelデータ（文字列形式）
 * @param ignoreSheets 無視するシート名の配列
 * @param outputFolderPath 出力先フォルダパス
 * @param extendedPrompt AIへの追加指示（プロンプト拡張）
 */
export async function converse(excelData: string, ignoreSheets: string[], outputFolderPath: string, extendedPrompt: string, modelFamily: string) {

    let [model] = await vscode.lm.selectChatModels({
        vendor: 'copilot',
        family: modelFamily
    });

    const sheets = splitSheetData(excelData, ignoreSheets);
    for (const sheet of sheets) {
        const response = await sendExcelToMarkdownConversionRequest(model, sheet.contents, extendedPrompt)
        // response を新規ファイルに保存する
        const fileName = `${sheet.sheetName}.md`;
        const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
        if (workspaceFolder) {
            const filePath = vscode.Uri.joinPath(workspaceFolder.uri, outputFolderPath, fileName);
            if (response) {
                await vscode.workspace.fs.writeFile(
                    filePath,
                    Buffer.from(response, 'utf8')
                );
                vscode.window.showInformationMessage(`Saved to ${fileName}`);
            }
        }
    }
}

/**
 * Excelデータを各シートごとに分割する
 * @param excelData 処理対象のExcelデータ（文字列形式）
 * @param ignoreSheets 無視するシート名の配列
 * @returns シート名とその内容を含むオブジェクトの配列
 */
function splitSheetData(excelData: string, ignoreSheets: string[]): { sheetName: string, contents: string[] }[] {
    // エクセルデータを行ごとに分割する
    const excelRows = excelData.split(/\r?\n/);
    let sheetName = "";
    let contents: string[] = [];
    const sheets = [];
    for (let i = 0; i < excelRows.length; i++) {
        if (excelRows[i].startsWith('Sheet_Start')) {
            sheetName = excelRows[i].split('=')[1];
            contents = [];
        } else if (excelRows[i].startsWith('Sheet_End')) {
            if (!ignoreSheets.includes(sheetName)) {
                sheets.push({ sheetName, contents });
            }
        } else {
            contents.push(excelRows[i])
        }
    }
    return sheets
}

/**
 * Excelデータを受け取り、AIモデルを使用してマークダウン形式に変換するリクエストを送信する
 * @param model 使用する言語モデル
 * @param contents Excelの内容（行の配列）
 * @param extendedPrompt 追加のプロンプト指示
 * @returns 変換されたマークダウンテキスト
 */
async function sendExcelToMarkdownConversionRequest(model: vscode.LanguageModelChat, contents: string[], extendedPrompt: string): Promise<string> {

    const contentsRow = contents.join("\n")
    const prompt = CONVERSE_PROMPT
        .replace("{EXCEL_DATA}", contentsRow)
        .replace("{extendedPrompt}", extendedPrompt);
    const messages = [
        vscode.LanguageModelChatMessage.User(prompt)
    ];

    if (model) {
        let chatResponse = await model.sendRequest(
            messages,
            {},
            new vscode.CancellationTokenSource().token
        );
        return await parseChatResponse(chatResponse);
    }
    return "";
}

/**
 * AIからのレスポンスを解析して文字列として返す
 * @param chatResponse AIモデルからのレスポンス
 * @returns 結合された完全なレスポンステキスト
 */
async function parseChatResponse(
    chatResponse: vscode.LanguageModelChatResponse,
): Promise<string> {
    let accumulatedResponse = '';

    for await (const fragment of chatResponse.text) {
        accumulatedResponse += fragment;
    }

    return accumulatedResponse;
}
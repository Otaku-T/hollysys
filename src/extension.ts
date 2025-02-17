// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { XMLParser } from 'fast-xml-parser';
import { XMLBuilder } from 'fast-xml-parser';

// 定义 XmlContent 类型
interface XmlContent {
    typeContent: string[];        // 点类型
    idContent: string[];        // idContent点ID号
    positionContent: string[];  // idContent点坐标
    textContent: string[];      // idContent点名
    inputidxContent: string[][];  //储存OUT类型的输入ID
}
// 定义 ExcelContent 类型
interface ExcelContent {
    sheetName: string[];        // 储存工作表名称
    jsonData: string[][][];  //储存多个工作表内容
}
// 设置解析器选项
const parserOptions = {
    ignoreAttributes: false,  // 不忽略属性
    parseNodeValue: true,     // 解析节点值
    parseAttributeValue: true, // 解析属性值
    attributeNamePrefix: "@_", // 属性名称前缀
    textNodeName: "#text",    // 文本节点名称
    attrNodeName: "@_attr",   // 属性节点名称
    cdataPropName: "#cdata",  // CDATA 节点名称
    cdataPositionChar: "\\c", // CDATA 位置字符
    format: true,             // 格式化输出
    trimValues: true,         // 去除值的前后空格
    ignoreNameSpace: false,   // 不忽略命名空间
    parseTrueNumberOnly: true, // 只解析真正的数字
    arrayMode: false,         // 数组模式
    stopNodes: ["parse-me-as-string"], // 停止解析的节点
    emptyTagPlaceholder: null, // 空标签占位符
};
// 设置生成参数
const builderOptions = {
    format: true,             // 格式化输出
    indentBy: '    ',           // 缩进字符
    newline: '\r\n',          // 行尾符，设置为 CRLF
    suppressEmptyNode: false, // 不抑制空节点
    suppressBooleanAttributes: false, // 不抑制布尔属性
    writeSelfClosingTag: true, // 写自闭合标签
    cdataPropName: '#cdata',  // CDATA 节点名称
    cdataPositionChar: '\\c', // CDATA 位置字符
    textNodeName: '#text',    // 文本节点名称
    attrNodeName: '@_',       // 属性节点名称
    ignoreAttributes: false,  // 不忽略属性
    suppressRoot: true,       // 抑制根节点
    declareProcIns: true,     // 声明处理指令
    procInsName: 'xml',       // 处理指令名称
    procInsTarget: 'xml',     // 处理指令目标
    procInsAttributes: {},    // 处理指令属性
    writeBOM: false,          // 不写 BOM
    encodeSpecialCharacters: true, // 编码特殊字符
    escapeValue: true,        // 转义值
    escapeAttrValue: true     // 转义属性值
};

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	// 使用控制台输出诊断信息（console.log）和错误（console.error）
	// 此行代码仅在扩展激活时执行一次
	console.log('恭喜，您的扩展 "hollysys" 已经激活！');
	// 命令已在 package.json 文件中定义
    // 注册指令hollysys，"新建hollysys"
	const disposable1 = vscode.commands.registerCommand('hollysys.hollysys', () => {
        // 每次命令被执行时，此处的代码将被运行
        try {
            // 获取当前工作区的根目录
            const workspaceFolder = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
            if (!workspaceFolder) {
                vscode.window.showErrorMessage('没有打开的工作区');
                return;
            }

            // 定义要创建的文件夹路径
            const folderPath1 = path.join(workspaceFolder, 'POU替换输出');
            const folderPath2 = path.join(workspaceFolder, 'POU替换输入');
            const folderPath3 = path.join(workspaceFolder, 'ST替换输入');
            const folderPath4 = path.join(workspaceFolder, 'ST替换输出');
            const folderPath5 = path.join(workspaceFolder, '典型回路输出');
            const folderPath6 = path.join(workspaceFolder, '典型回路输入');
            const folderPath7 = path.join(workspaceFolder, 'POU点名统计');
            const folderPath8 = path.join(workspaceFolder, 'ST顺控');
            // 创建文件夹
            fs.mkdirSync(folderPath1, { recursive: true });
            fs.mkdirSync(folderPath2, { recursive: true });
            fs.mkdirSync(folderPath3, { recursive: true });
            fs.mkdirSync(folderPath4, { recursive: true });
            fs.mkdirSync(folderPath5, { recursive: true });
            fs.mkdirSync(folderPath6, { recursive: true });
            fs.mkdirSync(folderPath7, { recursive: true });
            fs.mkdirSync(folderPath8, { recursive: true });
            // 生成ST .xlsx 文件
            const workbook = XLSX.utils.book_new();
            const worksheetData = [
                ['步号', 'S1', 'S2', 'S3', 'S4'],
                ['分支跳转1', , 'S100', 'S100'],
                ['分支跳转2'],
                ['分支跳转3']
            ];
            const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
            XLSX.utils.book_append_sheet(workbook, worksheet, '顺控');

            const filePath = path.join(workspaceFolder, 'ST框架.xlsx');
            XLSX.writeFile(workbook, filePath);

            // 向用户显示一个消息框
            vscode.window.showInformationMessage('工程已成功创建！');
        } catch (error) {
			const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`创建文件夹时出错: ${err.message}`);
        }
    });
    // 注册指令hollysysExcel，"更新excel"
    const disposable2 = vscode.commands.registerCommand('hollysys.hollysysExcel', () => {
        // 每次命令被执行时，此处的代码将被运行
        try {
            // 获取当前工作区的根目录
            const workspaceFolder = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
            if (!workspaceFolder) {
                vscode.window.showErrorMessage('没有打开的工作区');
                return;
            }
            
            generateExcelFilesPOU(workspaceFolder);
            generateExcelFilesPID(workspaceFolder);
            generateExcelFilesST(workspaceFolder);
			
        } catch (error) {
			const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`创建文件夹时出错: ${err.message}`);
        }
    }); 
    // 注册指令hollysysST, "更新ST变量表"
    let disposable3 = vscode.commands.registerCommand('hollysys.hollysysST', () => {
        const editor = vscode.window.activeTextEditor;
        if (editor) {
            console.log(`当前文件语言: ${editor.document.languageId}`);
            if (editor.document.languageId === 'st') {
                vscode.window.showInformationMessage('已更新ST变量表');
            } else {
                vscode.window.showWarningMessage('当前文件不是 ST 文件。');
            }
        } else {
            vscode.window.showWarningMessage('没有打开的编辑器。');
        }
    });
    // 注册指令hollysysPOU, ""替换POU""
    let disposable4 = vscode.commands.registerCommand('hollysys.hollysysPOU', () => {
        // 每次命令被执行时，此处的代码将被运行
        try {
            //console.log('开始执行命令');
            // 获取当前工作区的根目录
            const workspaceFolder = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
            if (!workspaceFolder) {
                vscode.window.showErrorMessage('没有打开的工作区');
                return;
            }
            // 获取当前工作区路径 点名替换.xlsx
            const folderPath1 = path.join(workspaceFolder, '点名替换.xlsx');
            const Exceldata = readExcelFile(folderPath1);   // 调用函数读取Excel文件
            
            // 获取当前工作区路径POU替换输入下的文件夹
            const folderPath2 = path.join(workspaceFolder, 'POU替换输入');
            const folderPath3 = path.join(workspaceFolder, 'POU替换输出');
            const files =  getFilesInDirectory(folderPath2);
            for (let i = 0; i < files.length; i++) {
                // 获取文件名,绝对路径
                const folderPathXML = path.join(folderPath2, files[i]);
                // 调用函数XML解析函数
                const xmlContent = getTextFromXml(folderPathXML);
                 // 检查 xmlContent 是否为 null
                if (xmlContent && xmlContent.textContent ) {
                    //和EXCEL表格工作表的第二行是否为 null
                    if (Exceldata?.jsonData[i][1] && Exceldata?.jsonData[i][1] !== null){
                        console.log(`第${i+1}个文件有数据`);
                        //一个模板多个替换
                        for (let k = 1; k < Exceldata?.jsonData[i][1].length; k++) {
                            // 第二个循环替换点名
                            for (let j = 0; j < xmlContent.textContent.length; j++) {
                                if (xmlContent.textContent[j] === Exceldata?.jsonData[i][j + 1][k-1]) {
                                    if (Exceldata?.jsonData[i][j + 1][k] !==''){
                                        //console.log('替换',Exceldata?.jsonData[i][j + 1][k]);
                                        xmlContent.textContent[j] = Exceldata?.jsonData[i][j + 1][k];
                                    } else {
                                        console.log('点名为空不执行');
                                    }
                                } else {
                                    console.log('点名为空不执行');
                                    //vscode.window.showInformationMessage('EXCEL数据与XML文件点名不匹配,请重新生成点名表');
                                }
                            }
                            // 将更改后jsonData内容写入文件，返回新的json对象
                            const newJson = updateTextInXml(folderPathXML, xmlContent);
                            //修改生成后的文件名称
                            newJson.pou.name = `${newJson.pou.name}${k}`;
                            // 将更改后jsonData内容写入文件
                            const folderPathOut = path.join(folderPath3, `${k}${files[i]}`);
                            //console.log('文件路径',folderPathOut);
                            generateXmlFile (folderPathOut, newJson);
                        }
                    } else {
                        console.log(`第${i+1}个文件没有数据，请检查点名表`);
                        vscode.window.showErrorMessage(`第${i+1}个文件没有数据，请检查点名表`);
                    }
                } else {
                    vscode.window.showErrorMessage(`XML 文件解析失败: ${files[i]}`);
                }
            }
            console.log('已生成替换POU');
            vscode.window.showInformationMessage('已生成替换POU');
        } catch (error) {
			const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`生成替换POU出错: ${err.message}`);
        }
         
    });
    // 注册指令hollysysPID, ""生成回路""
    let disposable5 = vscode.commands.registerCommand('hollysys.hollysysPID', () => {
        try {
            // 获取当前工作区的根目录
            const workspaceFolder = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
            if (!workspaceFolder) {
                vscode.window.showErrorMessage('没有打开的工作区');
                return;
            }
            // 获取当前工作区路径 点名替换.xlsx
            const folderPath1 = path.join(workspaceFolder, '典型回路.xlsx');
            // 获取当前工作区路径POU替换输入下的文件夹
            const folderPath2 = path.join(workspaceFolder, '典型回路输入');
            const folderPath3 = path.join(workspaceFolder, '典型回路输出');
            const files =  getFilesInDirectory(folderPath2);
            const Exceldata  = readExcelFile(folderPath1);   // 调用函数读取Excel文件
            if (Exceldata) {
                const newJsonxml =  excelToXmlContent (Exceldata);  // 调用函数将Excel数据转换为XML内容
                //console.log('回路个数',newJsonxml.length);
                for (let i = 0; i < newJsonxml.length; i++) {
                    if (newJsonxml[i].length > 0){
                        // 获取文件名,绝对路径
                        //console.log('poU个数',newJsonxml[i].length);
                        const folderPathXML = path.join(folderPath2, files[i]);
                        for (let j = 0; j < newJsonxml[i].length; j++) {
                            const json = addTextInXml(folderPathXML,newJsonxml[i][j]);
                            //修改生成后的文件名称
                            json.pou.name = `${json.pou.name}${j}`;
                            // 将更改后jsonData内容写入文件
                            const folderPathOut = path.join(folderPath3, `${j}${files[i]}`);
                            console.log('文件路径',folderPathOut);
                            generateXmlFile (folderPathOut, json);
                        }
                    } else {
                        console.log('不生成回路');
                    }
                }
            } else {
                vscode.window.showErrorMessage('读取 Excel 文件失败，请检查文件是否存在且格式正确。');
            }
            vscode.window.showInformationMessage('已生成回路');
        } catch (error) {
            const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`生成典型回路出错: ${err.message}`);
        }
    });
    // 注册指令hollysysSTExcel, "生成ST顺控"
    let disposable6 = vscode.commands.registerCommand('hollysys.hollysysSTExcel', () => {
        
        vscode.window.showInformationMessage('生成ST顺控');
    });
    // 注册指令hollysysSTPOU, "替换ST"
    let disposable7 = vscode.commands.registerCommand('hollysys.hollysysSTPOU', () => {
        
        vscode.window.showInformationMessage('已生成替换ST');
    });
    // 注册指令hollysysPOUExcel, "更新POU变量表"
    let disposable8 = vscode.commands.registerCommand('hollysys.hollysysPOUExcel', () => {
        
        vscode.window.showInformationMessage('已更新POU变量表');
    });
	// 将注册的命令添加到上下文的 subscriptions 数组中，以确保在扩展停用时正确清理
	context.subscriptions.push(disposable1, disposable2, disposable3, disposable4, disposable5, disposable6, disposable7, disposable8);
    // 获取目录下的所有文件，返回文件名数组
    function getFilesInDirectory(directoryPath: string): string[] {
        try {
            // 读取目录内容
            const files = fs.readdirSync(directoryPath);
            // 返回文件名数组
            return files;
        } catch (error) {
            const err = error as Error; // 类型断言
            throw new Error(`读取目录时出错: ${err.message}`);
        }
    }
    //创建新pou替换excel文件
    function generateExcelFilesPOU(workspaceFolder: string): void {
        try {
            // 获取当前工作区路径POU替换输入下的文件夹
            const folderPath = path.join(workspaceFolder, 'POU替换输入');
            const files = getFilesInDirectory(folderPath);
            let index = 0;  // 索引
            const workbook = XLSX.utils.book_new();  // 创建新的工作簿
            
            for (const file of files) {
                // 获取文件名,绝对路径
                const folderPathXML = path.join(folderPath, files[index]);
                // 调用函数XML解析函数
                const XmlContent = getTextFromXml(folderPathXML);
                // 获取XML文件中的点名数组内容
                const textContent = XmlContent?.textContent || [];
                // 生成 点名替换.xlsx 文件
                const worksheetData = [       // 工作表表头
                    ['原点名', '替换点名']
                ];
                // 假设替换点名为空字符串
                const newRows = textContent.map(originalName => [originalName, '']);
                // 拼接数组
                worksheetData.push(...newRows);
                const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);   // 将数据转换为工作表
                //console.log(`文件夹下XML文件名: ${file}`);
                XLSX.utils.book_append_sheet(workbook, worksheet, file);    // 将工作表添加到工作簿中
                index++;         // 更新索引
            }
    
            const filePath = path.join(workspaceFolder, '点名替换.xlsx');  // 获取文件路径
            XLSX.writeFile(workbook, filePath);                            // 将工作簿写入文件
    
            // 向用户显示一个消息框
            vscode.window.showInformationMessage('点名替换EXCEL已成功创建！');
        } catch (error) {
            const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`创建新pou替换excel文件出错: ${err.message}`);
        }
    }
    //创建新典型回路excel文件
    function generateExcelFilesPID(workspaceFolder: string): void {
        try {
            // 获取当前工作区路径POU替换输入下的文件夹
            const folderPath = path.join(workspaceFolder, '典型回路输入');
            const files = getFilesInDirectory(folderPath);
            let index = 0;  // 索引
            const workbook = XLSX.utils.book_new();  // 创建新的工作簿
            for (const file of files) {
                // 获取文件名,绝对路径
                const folderPathXML = path.join(folderPath, files[index]);
                // 调用函数XML解析函数
                const XmlContent = getTextFromXml(folderPathXML);
                // 处理解析数据中的二维数组
                // 将二维数组转换为一维数组，每个元素是子数组的字符串形式
                const flattenedInputidxContent: string[] = (XmlContent?.inputidxContent || []).map(subArray => subArray.join(', '));
                // 生成 点名替换.xlsx 文件
                const worksheetData:String[][] = [];       // 工作表表头
                worksheetData.push(XmlContent?.typeContent || []);
                worksheetData.push(XmlContent?.idContent || []);
                worksheetData.push(XmlContent?.positionContent || []);
                worksheetData.push(flattenedInputidxContent);
                worksheetData.push(XmlContent?.textContent || []);
                const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);   // 将数据转换为工作表
                console.log(`文件夹下XML文件名: ${file}`);
                XLSX.utils.book_append_sheet(workbook, worksheet, file);    // 将工作表添加到工作簿中
                index++;         // 更新索引
            }
            const filePath = path.join(workspaceFolder, '典型回路.xlsx');   // 获取文件路径
            XLSX.writeFile(workbook, filePath);                             // 将工作簿写入文件
    
            // 向用户显示一个消息框
            vscode.window.showInformationMessage('典型回路EXCEL已成功创建！');
        } catch (error) {
            const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`创建新典型回路excel文件出错: ${err.message}`);
        }
    }
    //创建新ST替换excel文件
    function generateExcelFilesST(workspaceFolder: string): void {
        try {
            // 生成 典型回路.xlsx 文件
            const workbook2 = XLSX.utils.book_new();
            const worksheetData2 = [
                ['原ST点名', '替换ST点名']
            ];
            const worksheet2 = XLSX.utils.aoa_to_sheet(worksheetData2);
            XLSX.utils.book_append_sheet(workbook2, worksheet2, 'Sheet1');
    
            const filePath2 = path.join(workspaceFolder, 'ST替换.xlsx');
            XLSX.writeFile(workbook2, filePath2);
    
            // 向用户显示一个消息框
            vscode.window.showInformationMessage('ST替换EXCEL已成功创建！');
        } catch (error) {
            const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`创建文件夹时出错: ${err.message}`);
        }
    }
    //读取 XML 文件中的 <text> 标签内容
    function getTextFromXml(filePath: string): XmlContent  | null {
        try {
            // 读取 XML 文件内容
            const xmlContent = fs.readFileSync(filePath, 'latin1');
            // 解析 XML
            const parser = new XMLParser(parserOptions);
            const json = parser.parse(xmlContent);
            //console.log('读取XML',JSON.stringify(json, null, 2));
            // 检查 json.pou.cfc 是否存在
            if (!json.pou || !json.pou.cfc || !Array.isArray(json.pou.cfc.element)) {
                vscode.window.showErrorMessage('XML 文件结构不正确，缺少必要的标签');
                return null;
            }
            // 统计 POU.XML文件中有多少个element对象
            const elementCount = json.pou.cfc.element.length;
            let typeContent: string[] = [];  // 初始化为空数组
            let idContent: string[] = [];  // 初始化为空数组
            let positionContent: string[] = [];  // 初始化为空数组
            let textContent: string[] = [];  // 初始化为空数组
            let inputidxContent: string[][] = [];  // 初始化为空数组
            // 提取 <text> 标签的内容
            for (let i = 0; i < elementCount; i++) {
                typeContent.push(json.pou.cfc.element[i]['@_type'] || '');  // 使用 push 方法将字符串添加到数组中
                idContent.push(json.pou.cfc.element[i].id || '');  // 使用 push 方法将字符串添加到数组中
                // 判断 element 中是否有 text 标签
                const hasText = json.pou.cfc.element[i].text !== undefined;
                textContent.push(hasText ? json.pou.cfc.element[i].text : json.pou.cfc.element[i].AT_type);

                if (json.pou.cfc.element[i]['@_type'] === 'input') {
                    positionContent.push(json.pou.cfc.element[i].AT_position || '');  // 使用 push 方法将字符串添加到数组中
                    inputidxContent.push([]);
                } else if (json.pou.cfc.element[i]['@_type'] === 'output') {
                    positionContent.push(json.pou.cfc.element[i].position || '');  // 使用 push 方法将字符串添加到数组中
                    inputidxContent.push([json.pou.cfc.element[i].Inputid || '']);
                } else if (json.pou.cfc.element[i]['@_type'] === 'box') {
                    positionContent.push(json.pou.cfc.element[i].AT_position || '');  // 使用 push 方法将字符串添加到数组中
                    const inputCount = json.pou.cfc.element[i].input ? json.pou.cfc.element[i].input.length : 0;
                    inputidxContent.push([]);  // 确保 inputidxContent[i] 是一个数组
                    for (let j = 0; j < inputCount; j++) {
                        inputidxContent[i].push(json.pou.cfc.element[i].input[j]['@_inputid'] || 0);
                   }
                } else {
                    positionContent.push(json.pou.cfc.element[i].position || '');  // 使用 push 方法将字符串添加到数组中
                    inputidxContent.push([]);
                }
            }
            return { typeContent, idContent, positionContent, textContent, inputidxContent };
        } catch (error) {
            const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`读取 XML 文件出错: ${err.message}`);
            return null;
        }
    }
    // 修改 XML 文件内容并返回修改后的 JSON 对象
    function updateTextInXml(filePath: string, newJson: XmlContent):  any  {
        try {
            const xmlContent = fs.readFileSync(filePath, 'latin1');      // 读取 XML 文件内容
            // 解析 XML
            const parser = new XMLParser(parserOptions);
            const json = parser.parse(xmlContent);

            if (!json.pou || !json.pou.cfc || !Array.isArray(json.pou.cfc.element)) {
                vscode.window.showErrorMessage('XML 文件结构不正确，缺少必要的标签');
                return json;
            }

            const elementCount = json.pou.cfc.element.length;
            // 遍历元素集合，为每个元素设置或更新其属性
            for (let i = 0; i < elementCount; i++) {
                // 设置元素的id属性
                json.pou.cfc.element[i].id = newJson.idContent[i];
                // 根据条件更新元素的text或AT_type属性
                if (json.pou.cfc.element[i].text !== undefined) {
                    json.pou.cfc.element[i].text = newJson.textContent[i];
                } else {
                    json.pou.cfc.element[i].AT_type = newJson.textContent[i];
                }
                // 根据元素类型更新位置相关属性
                if (json.pou.cfc.element[i]['@_type'] === 'input') {
                    json.pou.cfc.element[i].AT_position = newJson.positionContent[i];  
                } else if (json.pou.cfc.element[i]['@_type'] === 'output') {
                    json.pou.cfc.element[i].position = newJson.positionContent[i];  
                    json.pou.cfc.element[i].Inputid = newJson.inputidxContent[i][0];  
                } else if (json.pou.cfc.element[i]['@_type'] === 'box') {
                    json.pou.cfc.element[i].AT_position = newJson.positionContent[i];  
                    // 对于box类型元素，更新其所有输入的inputid属性
                    const inputCount = json.pou.cfc.element[i].input ? json.pou.cfc.element[i].input.length : 0;
                    for (let j = 0; j < inputCount; j++) {
                        json.pou.cfc.element[i].input[j]['@_inputid'] = newJson.inputidxContent[i][j];  
                    }
                } else {
                    json.pou.cfc.element[i].position = newJson.positionContent[i];  
                }
            }
            return  json;
        } catch (error) {
            const err = error as Error;
            vscode.window.showErrorMessage(`修改 XML 文件出错: ${err.message}`);
            return  null;
        }
    }
     // 修改 XML 文件内容并返回修改后的 JSON 对象
     function addTextInXml(filePath: string, newJson: XmlContent):  any  {
        try {
            const xmlContent = fs.readFileSync(filePath, 'latin1');      // 读取 XML 文件内容
            // 解析 XML
            const parser = new XMLParser(parserOptions);
            const json = parser.parse(xmlContent);

            if (!json.pou || !json.pou.cfc || !Array.isArray(json.pou.cfc.element)) {
                vscode.window.showErrorMessage('XML 文件结构不正确，缺少必要的标签');
                return json;
            }
            //console.log('恭喜成功调用添加数据', json.pou.cfc.element);
            //计算一个POU中PID回路的个数
            const oldelementCount = json.pou.cfc.element.length;   //替换前的变量个数
            const pidCount = newJson.idContent.length /oldelementCount;
            //console.log('回路个数', pidCount);
            //在原POU文件内添加新的回路
            let elementtxt= [];
            for (let m = 1; m < pidCount; m++) {  //本身有一组回路，
                 elementtxt.push(JSON.parse(JSON.stringify(json.pou.cfc.element)));
            }
            //console.log('添加数据', elementtxt);
            for (let m = 1; m < pidCount; m++) {  //本身有一组回路，
                for (let n = 0; n < oldelementCount; n++) {
                    json.pou.cfc.element.push(elementtxt[m-1][n]);
                }
                //console.log('回路内元素个数', json.pou.cfc.element.length);
            }
            
            //  遍历元素集合，为每个元素设置或更新其属性
            const newelementCount = json.pou.cfc.element.length;
            for (let i = 0; i < newelementCount; i++) {
                // 设置元素的id属性
                json.pou.cfc.element[i].id = newJson.idContent[i];
                // 根据条件更新元素的text或AT_type属性
                if (json.pou.cfc.element[i].text !== undefined) {
                    json.pou.cfc.element[i].text = newJson.textContent[i];
                } else {
                    json.pou.cfc.element[i].AT_type = newJson.textContent[i];
                }
                // 根据元素类型更新位置相关属性
                if (json.pou.cfc.element[i]['@_type'] === 'input') {
                    json.pou.cfc.element[i].AT_position = newJson.positionContent[i];  
                } else if (json.pou.cfc.element[i]['@_type'] === 'output') {
                    json.pou.cfc.element[i].position = newJson.positionContent[i];  
                    json.pou.cfc.element[i].Inputid = newJson.inputidxContent[i][0];  
                } else if (json.pou.cfc.element[i]['@_type'] === 'box') {
                    json.pou.cfc.element[i].AT_position = newJson.positionContent[i];  
                    // 对于box类型元素，更新其所有输入的inputid属性
                    const inputCount = json.pou.cfc.element[i].input ? json.pou.cfc.element[i].input.length : 0;
                    for (let j = 0; j < inputCount; j++) {
                        json.pou.cfc.element[i].input[j]['@_inputid'] = newJson.inputidxContent[i][j];  
                    }
                } else {
                    json.pou.cfc.element[i].position = newJson.positionContent[i];  
                }
            }
            //console.log('pou内容',json.pou.cfc.element);
            return  json;
        } catch (error) {
            const err = error as Error;
            vscode.window.showErrorMessage(`XML文件添加回路出错: ${err.message}`);
            return  null;
        }
    }
    // 读取 Excel 文件内容并返回三维数组
    function readExcelFile(filePath: string): ExcelContent | null {
        try {
            // 同步读取文件内容
            const data = fs.readFileSync(filePath);  // 使用同步方法读取文件
            // 解析 Excel 文件
            const workbook = XLSX.read(data, { type: 'buffer' });
            // 获取所有工作表的名称
            const sheetName: string[] = workbook.SheetNames;
            const jsonData: string[][][] = [];  // 使用 const 声明并初始化
            for (let i = 0; i < sheetName.length; i++) {
                // 获取工作表数据
                const worksheet = workbook.Sheets[sheetName[i]];
                // 将工作表数据转换为二维数组
                //jsonData.push([]);
                const sheetData: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as string[][];
                jsonData.push(sheetData);
                //console.log(jsonData);    
            }
            
            return { sheetName, jsonData };
        } catch (error) {
            const err = error as Error;
            vscode.window.showErrorMessage(`读取 Excel 文件出错: ${err.message}`);
            return null;
        }
    }
    // 将 Excel 内容转换为典型回路 XML 内容
    function excelToXmlContent(excel: ExcelContent ): any {
        try {
            let newJson: XmlContent[][] = [];
            console.log('恭成功调用数据分析');
            for (let i = 0; i < excel.jsonData.length; i++) {
                let xml = 0;  //同一典型回路要创建几个POU
                let index = 0;  //同一POU下有几个典型回路
                // 获取当前工作表的前四行数据，前四行为常数
                const sheetid = excel.jsonData[i][1];
                const idlength = sheetid.length;                //获取ID长度
                const maxid = Math.max(...sheetid.map(Number)); //取最大ID然后累加
                const sheetposit = excel.jsonData[i][2];
                // 计算X,Y坐标的Y的最大值
                let maxy = -Infinity;
                if (sheetposit && sheetposit.length > 0) {
                    for (let n = 0; n < sheetposit.length; n++) {   // 遍历数组
                        if (typeof sheetposit[n] === 'string') {
                            const parts = sheetposit[n].split(',');
                            const numberAfterComma = parseInt(parts[1], 10); // 转换为数字
                            // 比较并记录最大值
                            if (numberAfterComma > maxy) {
                                maxy = numberAfterComma + 5;  //预留5个像素
                            }
                        } else {
                            console.warn(`sheetposit[${n}] is not a string`);
                        }
                    }
                } else {
                    console.warn('sheetposit is empty or undefined');
                }
                const sheetinputidx = unflattenInputidxContent(excel.jsonData[i][3]);
                //开始数据替换计算,从6开始
                //console.log('表格行数',excel.jsonData[i].length);
                if (excel.jsonData[i].length > 5) {
                    for (let j = 5; j < excel.jsonData[i].length; j++) {
                        if (excel.jsonData[i][j][0] !== '' && excel.jsonData[i][j][0] !== null && excel.jsonData[i][j][0] !== undefined ) {
                            if (!newJson[i]) {
                                newJson[i] = [];
                            }
                            if (!newJson[i][xml]) {
                                newJson[i][xml] = {
                                    typeContent: [],
                                    idContent: [],
                                    positionContent: [],
                                    textContent: [],
                                    inputidxContent: []
                                };
                            }
                            //添加回路类型
                            newJson[i][xml].typeContent.push(...excel.jsonData[i][0]);
                            //添加回路ID
                            newJson[i][xml].idContent.push(...sheetid.map(item => item + (maxid*index)));
                            //添加坐标
                            newJson[i][xml].positionContent.push(...sheetposit.map(item => `${item.split(',')[0]},${parseInt(item.split(',')[1]) + (maxy*index)}`));
                            //添加输入引脚的Idx
                            //newJson[i][xml].inputidxContent.push([]);
                            for (let x = 0; x < sheetinputidx.length; x++) {
                                for (let y = 0; y < sheetinputidx[x].length; y++) {
                                    if (!newJson[i][xml].inputidxContent[x+(idlength*index)]) {
                                        newJson[i][xml].inputidxContent[x+(idlength*index)] = [];
                                    }
                                    if (sheetinputidx[x][y] !== '0' && sheetinputidx[x][y] !== '') {
                                        newJson[i][xml].inputidxContent[x+(idlength*index)].push((parseInt(sheetinputidx[x][y]) + (maxid * index)).toString());
                                    } else {
                                        if (sheetinputidx[x][y] === '0') {
                                            newJson[i][xml].inputidxContent[x+(idlength*index)].push('0');
                                        } else {
                                            newJson[i][xml].inputidxContent[x+(idlength*index)].push('');
                                        }
                                    }
                                } 
                            }
                            //添加点名
                            newJson[i][xml].textContent.push(...excel.jsonData[i][j]);
                            index += 1;     //索引加一
                        } else {
                            //console.log('创建新POU');
                            xml += 1;
                            index = 0;
                        }
                    }
                } else {
                    if (!newJson[i]) {
                        newJson[i] = [];
                    }
                    console.log('excel数据不足');
                }
            }
            console.log('excel生成的json文件',newJson);
            return newJson;
        } catch (error) {
            const err = error as Error;
            vscode.window.showErrorMessage(`典型回路数据分析出错: ${err.message}`);
            return null;
        }
    }
    // 将EXCEL输入框的输入id字符串转换为二维数组，用于典型回路
    function unflattenInputidxContent(flattenedInputidxContent: string[]): string[][] {
        if (!flattenedInputidxContent || flattenedInputidxContent.length === 0) {
            return [];
        }
        return flattenedInputidxContent.map(str => str.split(', ').map(item => item.trim()));
    }
    // 定义生成 XML 文件的函数
    function generateXmlFile(filePath: string, json: any): void {
        try {
            // 修改 @version 属性
            if (json['?xml'] && json['?xml']['@_version']) {
                json['?xml']['@_version'] = "1.0";
            } else {
                vscode.window.showErrorMessage('XML 文件中未找到 @version 属性');
            }
            //console.log('新生成的',JSON.stringify(json, null, 2));
            // 创建 XMLBuilder 实例，并配置生成 XML 的选项
            const builder = new XMLBuilder(builderOptions);
            // 将 JSON 对象转换为 XML 字符串
            const xmlOutput = builder.build(json);
            // 将生成的 XML 字符串写入文件
            fs.writeFileSync(filePath, xmlOutput,'latin1');
            // 向用户显示一个消息框
            vscode.window.showInformationMessage('XML 文件已成功生成！');
        } catch (error) {
            const err = error as Error; // 类型断言
            vscode.window.showErrorMessage(`生成 XML 文件时出错: ${err.message}`);
        }
    }
}

// This method is called when your extension is deactivated
export function deactivate() {}

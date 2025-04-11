const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 配置参数
const inputDir = './merge';    // Excel文件所在目录
const outputFile = './merge/海南-已出席-No1-10000条.xlsx'; // 输出文件路径

async function mergeExcelFiles() {
    // 创建新工作簿
    const newWorkbook = XLSX.utils.book_new();
    
    // 读取目录中的所有文件
    const files = fs.readdirSync(inputDir).filter(file => 
        ['.xlsx', '.xls'].includes(path.extname(file).toLowerCase())
    );

    let mergedData = [];
    let isFirstFile = true;

    // 遍历处理每个Excel文件
    for (const file of files) {
        const filePath = path.join(inputDir, file);
        const workbook = XLSX.readFile(filePath);
        
        // 获取第一个工作表的数据
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // 处理表头和数据
        if (data.length > 0) {
            if (isFirstFile) {
                // 保留第一个文件的表头
                mergedData = mergedData.concat(data);
                isFirstFile = false;
            } else {
                // 跳过后续文件的表头
                mergedData = mergedData.concat(data.slice(1));
            }
        }
    }

    // 将合并后的数据转换为工作表
    const newWorksheet = XLSX.utils.aoa_to_sheet(mergedData);
    
    // 添加工作表并保存文件
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'MergedData');
    XLSX.writeFile(newWorkbook, outputFile);
    
    console.log(`合并完成，共合并 ${files.length} 个文件，输出文件：${outputFile}`);
}

// 执行合并
mergeExcelFiles().catch(console.error);
const XLSX = require('xlsx');

/**
 * 删除Excel中指定列包含*的整行数据
 * @param {string} inputPath 输入文件路径
 * @param {string} outputPath 输出文件路径
 * @param {string} column 要检查的列字母（如'B'）
 */

const inputPath = "重庆-未上课-No1-batch8.xlsx"
const column = "D"
function deleteRowsWithAsterisk(inputPath, outputPath, column) {
    try {
        // 1. 读取Excel文件
        const workbook = XLSX.readFile(inputPath);
        const sheetName = workbook.SheetNames[0]; // 获取第一个工作表
        const worksheet = workbook.Sheets[sheetName];

        // 2. 将工作表数据转换为数组格式（保留空单元格）
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,      // 获取数组格式的数据
            defval: null,   // 保留空单元格为null
            raw: true      // 获取原始值（不进行格式转换）
        });

        // 3. 计算列索引（A=0, B=1, C=2,...）
        const colIndex = column.toUpperCase().charCodeAt(0) - 65;
        if (colIndex < 0 || colIndex > 25) {
            throw new Error('列字母必须是A-Z之间的字符');
        }

        // 4. 筛选数据 - 保留不包含*的行
        const filteredData = jsonData.filter((row, index) => {
            // 第一行通常是表头，默认保留
            if (index === 0) return true;
            
            const cellValue = row[colIndex];
            
            // 检查单元格值是否包含*
            if (cellValue && typeof cellValue === 'string' && cellValue.includes('*')) {
                return false; // 包含*的行将被过滤掉
            }
            return true; // 保留其他行
        });

        // 5. 将过滤后的数据转换回工作表
        const newWorksheet = XLSX.utils.aoa_to_sheet(filteredData);

        // 6. 保留原始工作表的特殊属性（如冻结窗格、列宽等）
        Object.keys(worksheet)
            .filter(key => key.startsWith('!')) // 特殊属性以!开头
            .forEach(key => {
                newWorksheet[key] = worksheet[key];
            });

        // 7. 更新工作簿中的工作表
        workbook.Sheets[sheetName] = newWorksheet;

        // 8. 写入新文件
        XLSX.writeFile(workbook, outputPath);
        
        console.log(`成功删除${column}列包含*的整行数据，结果已保存到: ${outputPath}`);
    } catch (error) {
        console.error('处理文件时出错:', error.message);
    }
}

// 使用示例,新文件output同一个文件路径
deleteRowsWithAsterisk(inputPath, inputPath, column);
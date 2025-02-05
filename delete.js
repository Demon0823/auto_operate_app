const xlsx = require("xlsx")
const fs = require("fs")
const path = require("path")

// 删除指定范围的行，并保留其余数据
// 删除控制台添加的index数量(精确到错误的数据index)
const endRow = 51
const fileName = "batch_6.xlsx"
function removeRows(inputFilePath, outputFilePath) {
  const startRow = 2
  if (!fs.existsSync(inputFilePath)) {
    console.error("输入文件不存在:", inputFilePath)
    return
  }

  // 读取 Excel 文件
  const workbook = xlsx.readFile(inputFilePath)
  const sheet = workbook.Sheets[workbook.SheetNames[0]] // 获取第一个工作表

  // 转换为二维数组
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 }) // 包括表头在内的所有行

  if (data.length === 0) {
    console.error("Excel 文件为空或没有数据")
    return
  }

  // 删除指定范围的行（从第 startRow 到第 endRow）
  const newData = [...data.slice(0, startRow - 1), ...data.slice(endRow)]

  // 将数据写入新的工作表
  const newSheet = xlsx.utils.aoa_to_sheet(newData) // 转为工作表
  const newWorkbook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(newWorkbook, newSheet, workbook.SheetNames[0]) // 使用原工作表名

  // 写回文件
  xlsx.writeFile(newWorkbook, outputFilePath)
  console.log(`成功删除: ${endRow - startRow + 1} 条`)
  console.log(`待处理剩余: ${newData.length - 1} 条`)
  console.log(`待处理数据: ${newData[1]}`)
}

// 示例文件路径
const inputFilePath = path.resolve(__dirname, "input", fileName) // 输入的 Excel 文件路径
const outputFilePath = path.resolve(__dirname, "input", fileName) // 输出文件路径

// 调用函数
removeRows(inputFilePath, outputFilePath)

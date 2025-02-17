const puppeteer = require("puppeteer-extra")
const StealthPlugin = require("puppeteer-extra-plugin-stealth")
const fs = require("fs")
const path = require("path")
const xlsx = require("xlsx")

// 使用隐身模式插件
// 通过插件隐藏 navigator.webdriver、调整浏览器特征等，成功绕过大多数常见的反爬虫检测。
puppeteer.use(StealthPlugin())

const sheetName = "Sheet1"
let inputFilePath = path.resolve(__dirname, "input", "batch_4.xlsx") // 输入的 Excel 文件路径
const outputFilePath = path.resolve(__dirname, "output", "江苏-已出席-No1-batch4.xlsx") // 输出的 Excel 文件路径

// 延时操作
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms))
}

// 读取 Excel 数据
function readExcel(filePath) {
  const workbook = xlsx.readFile(filePath)
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]
  return {
    rows: xlsx.utils.sheet_to_json(sheet),
    workbook,
    sheetName,
  }
}

// 写回 Excel 数据
function writeExcel(newData) {
  console.log("*****************************")
  console.log("数据处理结束，开始写入excel")
  let workbook
  let sheet = []

  if (fs.existsSync(outputFilePath)) {
    // 加载已有文件
    const existingBuffer = fs.readFileSync(outputFilePath)
    workbook = xlsx.read(existingBuffer, { type: "buffer" })
    console.log("已加载现有工作簿")

    // 获取指定工作表
    sheet = workbook.Sheets[sheetName]
    if (!sheet) {
      // 如果工作表不存在，创建新的
      sheet = xlsx.utils.json_to_sheet([])
      xlsx.utils.book_append_sheet(workbook, sheet, sheetName)
    }
  } else {
    // 如果文件不存在，创建新的工作簿和工作表
    workbook = xlsx.utils.book_new()
    sheet = xlsx.utils.json_to_sheet([])
    xlsx.utils.book_append_sheet(workbook, sheet, sheetName)
    console.log("新建了一个工作簿和工作表")
  }

  // 获取已有数据（如果有）
  const existingData = xlsx.utils.sheet_to_json(sheet)

  // 合并新数据
  const combinedData = [...existingData, ...newData]

  // 使用新的数据更新工作表
  const updatedSheet = xlsx.utils.json_to_sheet(combinedData)
  workbook.Sheets[sheetName] = updatedSheet

  // 保存更新后的工作簿
  xlsx.writeFile(workbook, outputFilePath)
  console.log(`本次新增数据：${newData.length}条`)
  console.log(`已成功处理数据：${combinedData.length}条`)
  if (combinedData.length >= 100000) {
    console.log("*****************************")
    console.log("******  已经 10w 条了  *******")
    console.log("*****************************")
  }
}

// 自动化查询并获取电话信息
async function performAutomation(page, selectors) {
  // 读取 Excel 数据
  const { rows } = readExcel(inputFilePath)
  let outputData = [] // for循环遍历成功的数据
  let index = 0
  for (const row of rows) {
    try {
      index = index + 1
      const queryId = row["ID/电话"] // 假设 Excel 中有列 "ID"
      const name = row["姓名"] || "暂无姓名" // 假设 Excel 中有列 "ID"
      if (!queryId) continue // 跳过空 ID

      console.log(`*****  第 ${index} 条数据  *****`)
      console.log(`0.开始查询第name: ${name}`)
      console.log(`1.开始查询第ID: ${queryId}`)

      // 输入查询 ID
      await page.type(selectors.inputField, "") // 替换为实际选择器
      await delay(500) // 延时 1 秒
      await page.type(selectors.inputField, queryId) // 替换为实际选择器
      await delay(1000) // 延时 1 秒
      /***********  查看id详情  ***********/
      try {
        await page.waitForSelector(
          "body .el-autocomplete-suggestion__list li:first-child",
          {
            hidden: false,
          }
        )
      } catch (error) {
        await delay(3000) // 延时 2 秒
        await page.waitForSelector(
          "body .el-autocomplete-suggestion__list li:first-child",
          {
            hidden: false,
          }
        )
        await delay(3000) // 延时 2 秒
      }

      const element = await page.$(
        "body .el-autocomplete-suggestion__list li:first-child"
      )
      const dialog = await page.$("body .u-list-detail-card-user_info")

      try {
        if (element) {
          await page.evaluate(el => el.click(), element) // 使用 evaluate 保证操作的是最新节点
          console.log("2.id点击查询成功")
        } else {
          console.error("未找到id下拉目标元素")
        }
      } catch (error) {
        console.error("操作失败:", error.message)
      }
      await delay(1000) // 延时 1 秒
      /***********  查看电话详情  ***********/
      try {
        if (!dialog) {
          await page.waitForSelector("body .u-list-detail-card-user_info", {
            hidden: false,
            timeout: 10000,
          })
        }
      } catch (error) {
        console.log("查看电话详情error......")
      }
      // 转介绍也有icon-view,所以定位需要精准

      const dtl = await page.$(
        "body .u-list-detail-card-user_info  > div:nth-child(2) .line .el-icon-view"
      )
      if (dtl) {
        await page.evaluate(el => el.click(), dtl) // 使用 evaluate 保证操作的是最新节点
        console.log("3.查看电话详情成功")
        await delay(1500) // 延时 1 秒
        try {
          const spanValue = await page.evaluate(() => {
            const target = document.querySelector(
              "body .u-list-detail-card-user_info > div:nth-child(2) .line .el-icon-view"
            ) // 替换为实际选择器
            const sibling = target?.previousElementSibling
            return sibling?.tagName === "SPAN"
              ? sibling.textContent.trim()
              : null // 确保是 <span>
          })
          // 分割字符串并获取电话号码
          const phoneNumber = spanValue.split("|")[1]?.trim()
          console.log("4.电话号码为:", phoneNumber)
          row["电话"] = phoneNumber || "查询失败" // 将结果写入行数据
        } catch (error) {
          console.log("icon-view匹配不精准")
          row["电话"] = "暂无电话" // 将结果写入行数据
        } finally {
          // 模拟按下 Esc 键触发查询
          await page.keyboard.press("Escape")
          await delay(1000) // 延时 500 毫秒
        }
      } else {
        console.error("未找到目标元素")
        row["电话"] = "暂无电话" // 将结果写入行数据
        await page.type(selectors.inputField, "")
        continue
      }
      outputData.push(row) // 成功的数据存入内存data中
      if (index % 50 === 0) {
        await writeExcel(outputData)
        outputData = []
        await page.reload()
        await delay(5000) // 延时 500 毫秒
        continue
      }
    } catch (error) {
      console.log("查询失败,优先写入现有数据")
      // row["电话"] = "暂无电话2"
      // outputData.push(row)
      await writeExcel(outputData)
      outputData = []
      await delay(5000)
      console.log("---------- reload start ---------")
      await page.reload()
      await delay(5000)
      console.log("---------- reload end ---------")
      // try {
      //   await delay(5000) // 延时 500 毫秒
      //   if (dialog) {
      //     // 模拟按下 Esc 键触发查询
      //     await page.keyboard.press("Escape")
      //     await delay(1000) // 延时 500 毫秒
      //   }
      //   if (!page.isClosed()) {
      //     // 仅在页面未关闭时操作
      //     await page.reload()
      //   } else {
      //     console.error("页面已关闭或分离，无法执行操作")
      //   }
      // } catch (handleError) {
      //   console.log("最外层catch内部逻辑有误")
      // }
      // console.log("--------reload成功--------")
      // await delay(5000) // 延时 500 毫秒
      // continue
    }
  }
  if (index % 50 !== 0) {
    console.log("循环执行结束，不满50条数据")
    await writeExcel(outputData)
    outputData = []
    await page.reload()
    await delay(5000) // 延时 500 毫秒
  }
  // await writeExcel(rows)
}

/******************************** 核心自动化脚本 start  *********************************************/
;(async () => {
  const browser = await puppeteer.launch({
    headless: false, // 设置为 false 可观察扫码操作
    defaultViewport: null, // 保持浏览器窗口大小一致
  })
  const page = await browser.newPage()
  // 设置正常的 User-Agent
  await page.setUserAgent(
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
  )

  const loginUrl =
    // 1.销售
    // "https://cc.vipthink.cn/#/self_manage/query_practice_list_by_staff_id" // 替换为实际扫码登录页面的地址
    // 2.班主任
    "https://cc.vipthink.cn/#/assessment_center/report?admin_group_list=-3&attend_class=-1&date_type=finish_time&equipment_type=-1&opt_date_type=3&order=desc&page_count=10&page_num=1&sex=-1&sort=create_time&start_time=2025-01-01&status=-1" // 替换为实际扫码登录页面的地址
  const cookiesFilePath = path.resolve(__dirname, "cookies.json") // 保存 cookies 的文件路径

  // 检查并加载已保存的登录状态
  if (fs.existsSync(cookiesFilePath)) {
    const cookies = JSON.parse(fs.readFileSync(cookiesFilePath, "utf-8"))
    await page.setCookie(...cookies)
    console.log("加载已保存的登录状态...")
  } else {
    console.log("未找到登录状态，准备扫码登录...")
  }

  // 打开登录页面
  await page.goto(loginUrl, { waitUntil: "networkidle2" })

  // 检测是否需要扫码登录
  const isLoginRequired = await page.evaluate(() => {
    // 判断页面是否显示二维码（根据实际情况修改选择器）
    return document.querySelector(".wx-pic") !== null
  })

  if (isLoginRequired) {
    console.log("请扫码登录...")
    // 等待用户扫码并完成登录
    await page.waitForNavigation({ waitUntil: "networkidle2", timeout: 0 })

    // 登录成功后保存 cookies
    const cookies = await page.cookies()
    fs.writeFileSync(cookiesFilePath, JSON.stringify(cookies, null, 2))
    console.log("登录成功，cookies 已保存。")
  } else {
    console.log("检测到已登录，无需扫码。")
  }

  // 登录后访问目标页面
  const targetUrl =
    // 1.销售
    // "https://cc.vipthink.cn/#/self_manage/query_practice_list_by_staff_id"
    // 2.班主任
    "https://cc.vipthink.cn/#/assessment_center/report?admin_group_list=-3&attend_class=-1&date_type=finish_time&equipment_type=-1&opt_date_type=3&order=desc&page_count=10&page_num=1&sex=-1&sort=create_time&start_time=2025-01-01&status=-1" // 替换为实际目标页面地址
  await page.goto(targetUrl, { waitUntil: "networkidle2" })
  // 检测 navigator.webdriver
  const isWebdriver = await page.evaluate(() => navigator.webdriver)
  console.log("navigator.webdriver:", isWebdriver) // 如果是 true，则可能被检测
  if (isWebdriver) return
  console.log("已进入目标页面。")

  // 执行目标页面操作
  // 示例：获取页面标题
  const pageTitle = await page.title()
  console.log("页面标题：", pageTitle)
  // 页面选择器配置
  const selectors = {
    inputField: ".avatar-container > div.el-autocomplete input", // 替换为查询输入框的选择器
    detailButton: ".detail-btn", // 替换为详情按钮的选择器
    phoneField: ".phone-number", // 替换为详情页中电话号码字段的选择器
  }
  // 执行自动化操作
  console.time("myFunctionTime") // 开始计时
  await performAutomation(page, selectors)
  console.timeEnd("myFunctionTime") // 开始计时

  // 防止无人监控掉线自动执行下一批,不需要监控断开exec执行,注释掉
  async function autoContinue(fileName) {
    inputFilePath = path.resolve(__dirname, "input", fileName) // 输入的 Excel 文件路径
    const { rows } = readExcel(inputFilePath)
    const successNum = rows.length
    if (successNum % 5000 === 0) {
      // 执行自动化操作
      console.time("myFunctionTime") // 开始计时
      await performAutomation(page, selectors)
      console.timeEnd("myFunctionTime") // 开始计时
    }
  }
  // await autoContinue("batch_8.xlsx")
})()
// excel结束ctrl F查找电话号码是否有 ***，是否有暂无电话， 是否有查询失败
/******************************** 核心自动化脚本 end  *********************************************/

const puppeteer = require("puppeteer-extra")
const StealthPlugin = require("puppeteer-extra-plugin-stealth")
const fs = require("fs")
const path = require("path")
const xlsx = require("xlsx")

// 使用隐身模式插件
// 通过插件隐藏 navigator.webdriver、调整浏览器特征等，成功绕过大多数常见的反爬虫检测。
puppeteer.use(StealthPlugin())
// 延时操作
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms))
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
  const element = await page.$("body .el-icon-search")
  for (var i = 0; i < 24; i++) {
    console.log(i)
    await page.evaluate(el => el.click(), element)
    await delay(1000*60*110) // 延时 110 分钟
  }
  console.timeEnd("myFunctionTime") // 开始计时
})()
/******************************** 核心自动化脚本 end  *********************************************/

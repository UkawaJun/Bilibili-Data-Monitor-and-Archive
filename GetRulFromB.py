import csv
import re
import time
from Support import *

# 配置信息
TARGET_URL = "https://space.bilibili.com/23191782/upload/video"
CSV_FILE = "BID.csv"
BASE_URL = "https://www.bilibili.com/video/"

def OpenChrome(DZYBaseUrl):
    DEBPrint("-个人化配置文件初始化完成")
    DEBPrint("正在检查 ChromeDriver 版本并准备启动...")
    
    Rule = Options()
    Rule.add_argument('--no-sandbox')
    Rule.add_argument('--log-level=3')
    Rule.add_experimental_option(name="detach", value=True)
    
    # ... (此处保留你原有的驱动下载逻辑) ...
    try:
        driver_path = ChromeDriverManager().install()
        service = Service(driver_path)
        chrome = webdriver.Chrome(service=service, options=Rule)
        chrome.get(DZYBaseUrl)
        DEBPrint(f"浏览器已成功启动")
        return chrome
    except Exception as e:
        DEBPrint(f"启动失败: {e}")
        return None

# 初始化浏览器
chrome = OpenChrome(TARGET_URL)

# 如果你的页面需要手动登录或处理验证码，可以在这里停顿一下
DEBPrint("请在浏览器中完成登录（如果需要），然后在控制台按回车开始爬取...")
input()

# 初始化 CSV 文件，写入表头
with open(CSV_FILE, mode='w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(["视频标题", "完整链接"])

all_bv_count = 0
page_num = 1

while True:
    DEBPrint(f"正在抓取第 {page_num} 页...")
    
    # 1. 确保页面加载完成，等待视频卡片出现
    try:
        WebDriverWait(chrome, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.bili-cover-card"))
        )
    except:
        DEBPrint("未能加载视频列表，可能已到末尾或网络异常")
        break

    # 2. 提取当前页所有的视频链接
    video_elements = chrome.find_elements(By.CSS_SELECTOR, "a.bili-cover-card")
    
    current_page_data = []
    for el in video_elements:
        href = el.get_attribute("href")
        # B站卡片结构通常包含标题，我们尝试获取一下
        try:
            # 这里的 title 获取取决于具体页面结构，通常在 img 的 alt 或 标题 div 里
            title = el.find_element(By.XPATH, "../../..//h3").text # 示例路径，需根据实际微调
        except:
            title = "未知标题"

        match = re.search(r'(BV[a-zA-Z0-9]+)', href)
        if match:
            bv_id = match.group(1)
            full_link = BASE_URL + bv_id
            current_page_data.append([title, full_link])

    # 3. 写入 CSV（实时追加）
    if current_page_data:
        with open(CSV_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(current_page_data)
        
        all_bv_count += len(current_page_data)
        DEBPrint(f"第 {page_num} 页抓取完成，本页发现 {len(current_page_data)} 条，总计 {all_bv_count} 条")
    else:
        DEBPrint("本页未找到视频，停止翻页")
        break

    # 4. 尝试点击“下一页”
    try:
        # 定位下一页按钮
        next_btn = chrome.find_element(By.XPATH, "//button[contains(text(), '下一页')]")
        
        # 检查按钮是否被禁用 (B站最后一页按钮通常会添加 disabled 属性或特定 class)
        is_disabled = next_btn.get_attribute("disabled") or "disabled" in next_btn.get_attribute("class")
        
        if is_disabled:
            DEBPrint("检测到已经是最后一页，程序结束。")
            break
            
        # 滚动到按钮并点击
        chrome.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
        time.sleep(1) 
        next_btn.click()
        
        page_num += 1
        time.sleep(3) # 等待页面跳转和渲染，时间建议设置长一点防止被封或加载不到
        
    except Exception as e:
        DEBPrint(f"无法找到下一页按钮或点击失败，爬取结束。")
        break

DEBPrint(f"恭喜你，程序顺利结束！总共抓取了 {all_bv_count} 条数据，已保存至 {CSV_FILE}")
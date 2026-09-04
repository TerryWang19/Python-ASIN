import sys
print("当前 Python 路径:", sys.executable)

import re
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ========== 配置区 ==========
INPUT_EXCEL_PATH = "/Users/terry1984/Desktop/Pathon-ASIN/M-AdhesiveHooks.xlsx"
OUTPUT_EXCEL_PATH = "/Users/terry1984/Desktop/Pathon-ASIN/amazon_results.xlsx"

DEBUG = False                 # 价格调试
DEBUG_SALES = False           # 销量调试
HEADLESS = True               # 无头模式，若不稳定可改为 False
SAVE_HTML_ON_FAIL = False     # 失败时是否保存 HTML（调试时可设为 True）
MAX_RETRIES = 2               # 页面加载失败重试次数

# ========== 反检测 User-Agent 池 ==========
USER_AGENTS = [
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.5 Safari/605.1.15",
]


def get_driver():
    chrome_options = Options()
    if HEADLESS:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    user_agent = random.choice(USER_AGENTS)
    chrome_options.add_argument(f"user-agent={user_agent}")
    chrome_options.add_argument("accept-language=en-US,en;q=0.9")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver


def is_valid_price(text):
    if not text:
        return None
    text = text.strip()
    match = re.match(r'^\$([\d,]+(?:\.\d+)?)', text)
    if match:
        return match.group(1)
    return None


def clean_brand(raw_brand):
    if not raw_brand or raw_brand == "N/A":
        return "N/A"
    brand = raw_brand.strip()
    brand = re.sub(r'^Visit\s+the\s+', '', brand, flags=re.IGNORECASE)
    brand = re.sub(r'\s+Store$', '', brand, flags=re.IGNORECASE)
    return brand if brand else "N/A"


def extract_price(driver, debug=False):
    """从多个可能包含价格的容器中提取最可能的当前售价"""
    price = "N/A"
    candidate_prices = []

    container_selectors = [
        "div#corePrice_feature_div",
        "div#apex_desktop",
        "div#apex_offerDisplay_desktop",
        "div#price_inside_buybox",
        "div#desktop_buybox",
        "div#buybox",
        "div#buyNewInner",
        "div#usedBuySection",
        "div#variation_price_div",
        "div#tmmSwatches",
    ]

    for container_sel in container_selectors:
        try:
            container = driver.find_element(By.CSS_SELECTOR, container_sel)
            elems = container.find_elements(By.XPATH, ".//span[contains(@class,'a-price') or contains(@class,'a-offscreen')]")
            for elem in elems:
                try:
                    text = elem.text.strip() or elem.get_attribute("textContent").strip()
                    valid = is_valid_price(text)
                    if valid:
                        parent_text = elem.find_element(By.XPATH, "..").text.lower()
                        if any(kw in parent_text for kw in ["list price", "was", "typical", "r.r.p"]):
                            continue
                        candidate_prices.append((f"${valid}", parent_text, elem))
                except:
                    continue
        except:
            continue

    if not candidate_prices:
        # 如果容器内没找到，页面级搜索（但仍过滤）
        all_elems = driver.find_elements(By.XPATH, "//*")
        for elem in all_elems:
            try:
                text = driver.execute_script("return arguments[0].childNodes[0]?.nodeValue || arguments[0].textContent;", elem)
                text = text.strip()
                valid = is_valid_price(text)
                if valid:
                    parent_text = elem.find_element(By.XPATH, "..").text.lower()
                    if any(kw in parent_text for kw in ["list price", "was", "typical", "r.r.p"]):
                        continue
                    candidate_prices.append((f"${valid}", parent_text, elem))
            except:
                continue

    if debug:
        print(f"\n[DEBUG] 找到 {len(candidate_prices)} 个候选价格：")
        for price_text, parent, elem in candidate_prices:
            size_attr = elem.get_attribute("data-a-size") if elem.tag_name == "span" and "a-price" in elem.get_attribute("class") else ""
            print(f"  price={price_text}, size={size_attr}, parent={parent[:80]}")

    if candidate_prices:
        # 选择优先级：折扣价 > XL 尺寸 > 最小值
        discount_prices = [p for p in candidate_prices if any(kw in p[1] for kw in ['%', 'deal', 'coupon', 'off'])]
        if discount_prices:
            min_price = min(discount_prices, key=lambda x: float(x[0].replace('$','').replace(',','')))
            price = min_price[0]
            if debug:
                print(f"[DEBUG] 选择折扣价: {price}")
        else:
            xl_prices = [p for p in candidate_prices if p[2].get_attribute("data-a-size") == "xl"]
            if xl_prices:
                price = xl_prices[0][0]
                if debug:
                    print(f"[DEBUG] 选择 XL 价格: {price}")
            else:
                min_price = min(candidate_prices, key=lambda x: float(x[0].replace('$','').replace(',','')))
                price = min_price[0]
                if debug:
                    print(f"[DEBUG] 选择最低价: {price}")

    return price


def extract_rating(driver):
    rating = "N/A"
    selectors = [
        ("css", "span.a-icon-alt"),
        ("css", "#acrPopover"),
        ("css", "#averageCustomerReviews span.a-icon-alt"),
    ]
    for method, selector in selectors:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, selector) if method == "css" else driver.find_element(By.XPATH, selector)
            text = elem.text.strip()
            if not text:
                text = elem.get_attribute("title") or ""
            match = re.search(r'(\d+\.?\d*)', text)
            if match:
                rating = match.group(1)
                break
        except:
            continue
    return rating


def extract_reviews(driver):
    reviews = "N/A"
    selectors = [
        "#acrCustomerReviewText",
        "div#averageCustomerReviews span.a-size-base",
        "#acrCustomerReviewLink",
    ]
    for sel in selectors:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, sel)
            text = elem.text.strip()
            match = re.search(r'([\d,]+)', text)
            if match:
                reviews = match.group(1)
                break
        except:
            continue
    return reviews


def extract_brand(driver):
    brand_raw = "N/A"
    selectors = [
        "a#bylineInfo",
        "#bylineInfo",
        "#brand",
        "a#brand",
        "#productOverview_feature_div tr:first-child td.a-size-base",
        "#detailBullets_feature_div li:first-child span.a-text-bold",
    ]
    for sel in selectors:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, sel)
            text = elem.text.strip()
            if text:
                brand_raw = text
                break
        except:
            continue
    return clean_brand(brand_raw)


def normalize_sales(sales_text):
    """
    将销量文本标准化为整数（字符串形式）。
    例如 "10K+ bought in past month" -> "10000"
         "500+ bought in past month" -> "500"
         "2M+ bought" -> "2000000"
    如果无法解析，返回 None
    """
    if not sales_text or sales_text == "N/A":
        return None
    text = sales_text.lower()
    # 提取数字部分，可能带 K/M/+/
    match = re.search(r'([\d,\.]+)\s*([km]?)\s*\+?', text)
    if not match:
        return None
    number_str = match.group(1).replace(',', '')
    suffix = match.group(2).lower()
    try:
        number = float(number_str)
    except:
        return None
    if suffix == 'k':
        number *= 1000
    elif suffix == 'm':
        number *= 1000000
    # 去掉小数（销量通常是整数）
    return str(int(number))


def is_out_of_stock(driver):
    """检测页面是否显示断货"""
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
        if "currently unavailable" in body_text or "out of stock" in body_text or "unavailable" in body_text:
            return True
    except:
        pass
    return False


def extract_sales_raw(driver):
    """提取原始的销量文本，若断货返回特殊标记"""
    if is_out_of_stock(driver):
        return "OUT_OF_STOCK"

    sales = None
    # 已知销量元素
    sales_selectors = [
        "#social-proofing-faceout-title-tk_bought",
        "#social-proofing-faceout-title-tk_bought + span",
        "#social-proofing-faceout-title",
        "span.a-size-small.a-color-secondary",
        "div#social-proofing-faceout-title-tk_bought",
    ]
    for sel in sales_selectors:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, sel)
            sales_text = elem.text.strip()
            if "bought" in sales_text.lower():
                sales = sales_text
                break
        except:
            continue

    if not sales:
        try:
            sales_elem = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'bought in past month')]"))
            )
            sales = sales_elem.text.strip()
        except:
            pass

    if not sales:
        try:
            elems = driver.find_elements(By.XPATH, "//*[contains(text(), 'bought')]")
            for elem in elems:
                text = elem.text.strip()
                if "past month" in text or "month" in text:
                    sales = text
                    break
        except:
            pass

    if not sales:
        # 从整个页面文本正则提取
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            patterns = [
                r'([\d,\.KkM]+)\+?\s*bought\s+in\s+past\s+month',
                r'([\d,\.KkM]+)\s*bought\s+in\s+past\s+month',
                r'([\d,\.KkM]+)\+?\s*bought\s+last\s+month',
                r'([\d,\.KkM]+)\s*bought\s+last\s+month',
            ]
            for pattern in patterns:
                match = re.search(pattern, body_text, re.IGNORECASE)
                if match:
                    sales = match.group(0)
                    break
        except:
            pass

    return sales if sales else None


def scrape_asin(driver, asin):
    """抓取单个 ASIN 的所有数据，总是返回一个字典（即使失败）"""
    url = f"https://www.amazon.com/dp/{asin}"
    print(f"\n===== 开始抓取 ASIN: {asin} =====")

    # 默认结果
    result = {
        "asin": asin,
        "brand": "N/A",
        "price": "N/A",
        "rating": "N/A",
        "reviews": "N/A",
        "sales": "25",  # 默认销量
    }

    # 先访问首页设置 cookie
    driver.get("https://www.amazon.com")
    driver.add_cookie({"name": "i18n-prefs", "value": "USD", "domain": ".amazon.com"})
    driver.add_cookie({"name": "lc-main", "value": "en_US", "domain": ".amazon.com"})
    time.sleep(1)

    page_loaded = False
    for attempt in range(MAX_RETRIES + 1):
        print(f"[{asin}] 尝试加载页面，第 {attempt + 1} 次")
        driver.get(url)
        time.sleep(random.uniform(4, 7))
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

        current_url = driver.current_url
        page_title = driver.title
        print(f"[{asin}] URL: {current_url}")
        print(f"[{asin}] 标题: {page_title}")

        if "amazon.com/dp/" not in current_url and "amazon.com/gp/product/" not in current_url:
            print(f"[{asin}] ⚠️ 页面被重定向，URL异常")
            if attempt < MAX_RETRIES:
                time.sleep(random.uniform(5, 10))
                continue
            else:
                break

        if any(kw in current_url.lower() for kw in ["captcha", "robot", "validate"]):
            print(f"[{asin}] ⚠️ 疑似被反爬拦截")
            if attempt < MAX_RETRIES:
                time.sleep(random.uniform(5, 10))
                continue
            else:
                break

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#productTitle"))
            )
            print(f"[{asin}] ✅ 商品标题已出现，页面加载成功")
            page_loaded = True
            break
        except:
            print(f"[{asin}] ❌ 商品标题未出现")
            if attempt < MAX_RETRIES:
                time.sleep(random.uniform(5, 10))
                continue
            else:
                print(f"[{asin}] 重试耗尽，跳过数据提取")
                if SAVE_HTML_ON_FAIL:
                    with open(f"debug_{asin}_no_elements.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    print(f"[{asin}] 已保存页面源码到 debug_{asin}_no_elements.html")
                break

    if not page_loaded:
        print(f"[{asin}] 未能加载页面，返回默认结果")
        return result

    # 提取品牌
    result["brand"] = extract_brand(driver)
    print(f"[{asin}] 品牌: {result['brand']}")

    # 提取价格前滚动到顶部，确保 Buy Box 可见
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    result["price"] = extract_price(driver, debug=DEBUG)
    print(f"[{asin}] 价格: {result['price']}")

    # 评分
    result["rating"] = extract_rating(driver)
    print(f"[{asin}] 评分: {result['rating']}")

    # 评论数
    result["reviews"] = extract_reviews(driver)
    print(f"[{asin}] 评论数: {result['reviews']}")

    # 销量
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)

    raw_sales = extract_sales_raw(driver)
    if raw_sales == "OUT_OF_STOCK":
        result["sales"] = "0"
        print(f"[{asin}] 销量: 0 (断货)")
    else:
        if raw_sales:
            normalized = normalize_sales(raw_sales)
            if normalized:
                result["sales"] = normalized
                print(f"[{asin}] 销量: {normalized} (原始: {raw_sales})")
            else:
                result["sales"] = "25"  # 解析失败，使用默认
                print(f"[{asin}] 销量解析失败，使用默认值 25")
        else:
            result["sales"] = "25"
            print(f"[{asin}] 未获取到销量，使用默认值 25")

    print(f"[{asin}] 提取完成：{result}")
    return result


def main():
    import os
    import re

    print(f"正在从 {INPUT_EXCEL_PATH} 读取 ASIN...")
    if not os.path.exists(INPUT_EXCEL_PATH):
        print(f"❌ 文件不存在：{INPUT_EXCEL_PATH}")
        return

    try:
        df = pd.read_excel(INPUT_EXCEL_PATH, header=None, dtype=str, engine='openpyxl')
        print(f"✅ Excel 读取成功，数据形状：{df.shape}")

        # 提取第一列所有非空值
        raw_values = df.iloc[:, 0].dropna().astype(str).str.strip()
        print(f"第一列非空单元格数：{len(raw_values)}")
        print("前5个原始值：", raw_values.head(5).tolist())

        # 使用正则过滤：只保留 10 位大写字母和数字组成的 ASIN
        asin_pattern = re.compile(r'^[A-Z0-9]{10}$')
        asins = [val for val in raw_values if asin_pattern.match(val)]

        # 如果有的 ASIN 是小写，可以转为大写后再匹配
        # asins = [val.upper() for val in raw_values if asin_pattern.match(val.upper())]

        print(f"过滤后有效 ASIN 数量：{len(asins)}")
        if len(asins) == 0:
            print("⚠️ 未提取到任何有效 ASIN，请检查 Excel 第一列内容是否只包含 ASIN 代码（如 B07KFDML8G）")
            return
        else:
            print("前3个有效 ASIN：", asins[:3])
    except Exception as e:
        print(f"❌ 读取 Excel 失败：{e}")
        return

    if not asins:
        print("未读取到任何 ASIN，程序退出")
        return

    # 继续原有抓取流程...
    driver = get_driver()
    results = []
    try:
        for idx, asin in enumerate(asins, 1):
            print(f"\n处理进度: {idx}/{len(asins)}")
            data = scrape_asin(driver, asin)
            results.append(data)
            time.sleep(random.uniform(6, 10))
    finally:
        driver.quit()

    # 保存结果
    if results:
        result_df = pd.DataFrame(results, columns=["asin", "brand", "price", "rating", "reviews", "sales"])
        result_df.to_excel(OUTPUT_EXCEL_PATH, index=False, engine="openpyxl")
        print(f"\n✅ 结果已保存到 {OUTPUT_EXCEL_PATH}")
    else:
        print("\n⚠️ 没有成功提取到任何数据，请检查调试输出")

if __name__ == "__main__":
    main()
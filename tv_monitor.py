import os
import re
import csv
import json
import pytz
import queue
import random
import shutil
import logging
import subprocess
import time
from io import StringIO
from pathlib import Path
from datetime import datetime
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor

import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# ===================== 日志 =====================
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# ===================== 路径配置 =====================
SCRIPT_DIR  = Path(__file__).parent
INPUT_FILE  = SCRIPT_DIR / "input" / "it_xiaomi_TV.xlsx"
OUTPUT_DIR  = SCRIPT_DIR / "output"

COUNTRY_CODE = "it"


def load_local_env():
    for env_file in (SCRIPT_DIR.parent / ".env", SCRIPT_DIR / ".env"):
        if not env_file.exists():
            continue
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            os.environ.setdefault(key, value)


load_local_env()

# ===================== 飞书配置 =====================
FEISHU_APP_ID     = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")
FEISHU_CHAT_IDS = [
    chat_id.strip()
    for chat_id in os.environ.get("FEISHU_CHAT_IDS", "").split(",")
    if chat_id.strip()
]
NO_FEISHU = os.environ.get("NO_FEISHU", "").lower() in ("1", "true", "yes", "on")

FEISHU_PRICE_SHEET_TOKEN = os.environ.get("FEISHU_PRICE_SHEET_TOKEN", "DLJMs1Q2Pht7xmtPxmPcmKhCnjd")
FEISHU_PRICE_SHEET_ID = os.environ.get("FEISHU_PRICE_SHEET_ID", "300dee")
FEISHU_PRICE_SHEET_AS = os.environ.get("FEISHU_PRICE_SHEET_AS", "user")
FEISHU_PRICE_FIRST_PRICE_COLUMN = os.environ.get("FEISHU_PRICE_FIRST_PRICE_COLUMN", "I")

# ===================== 爬虫配置 =====================
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]

BROWSER_CONFIG = {
    "headless": True,
    "window_size": "1920,1080",
    "disable_images": True,
    "page_load_strategy": "eager",
}

SCRAPER_CONFIG = {
    "max_workers": 1,
    "max_retries": 2,
    "restart_every": 3,
    "url_column": 3,
    "price_column": 6,
    "timeout": 30,
}

# ===================== 飞书 API =====================
def get_feishu_token() -> str:
    if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
        raise RuntimeError("Missing FEISHU_APP_ID or FEISHU_APP_SECRET environment variable")
    resp = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
        timeout=15,
    )
    return resp.json()["tenant_access_token"]


def send_feishu_message(text: str):
    if NO_FEISHU:
        logger.info("NO_FEISHU=1, skip Feishu text message")
        print(text)
        return

    if not FEISHU_CHAT_IDS:
        raise RuntimeError("Missing FEISHU_CHAT_IDS environment variable")
    token = get_feishu_token()
    for chat_id in FEISHU_CHAT_IDS:
        payload = {
            "receive_id": chat_id,
            "msg_type": "text",
            "content": json.dumps({"text": text}, ensure_ascii=False),
        }
        resp = requests.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=15,
        )
        if resp.json().get("code") != 0:
            logger.error(f"飞书文本消息发送失败({chat_id}): {resp.json()}")


def send_feishu_report(results: list[dict], title: str, elapsed_minutes: float):
    """发送富文本价格报告到飞书"""
    if NO_FEISHU:
        logger.info("NO_FEISHU=1, print report instead of sending Feishu")
        print(title)
        for r in results:
            print(
                " | ".join(
                    clean_cell(r.get(key))
                    for key in ("name", "price", "stock", "seller", "delivery", "basis_price", "deal_tag", "url")
                )
            )
        print(f"Elapsed: {elapsed_minutes:.1f} minutes")
        return

    if not FEISHU_CHAT_IDS:
        raise RuntimeError("Missing FEISHU_CHAT_IDS environment variable")

    post_content = []

    def text_tag(text: str, bold: bool = False) -> dict:
        tag = {"tag": "text", "text": text}
        if bold:
            tag["style"] = ["bold"]
        return tag

    def is_amazon_seller(seller: str) -> bool:
        merchant = seller.split("|", 1)[0].strip()
        if ":" in merchant:
            merchant = merchant.split(":", 1)[1].strip()
        if "：" in merchant:
            merchant = merchant.split("：", 1)[1].strip()
        return merchant.lower() == "amazon"

    def seller_name(seller: str) -> str:
        merchant = seller.split("|", 1)[0].strip()
        if ":" in merchant:
            merchant = merchant.split(":", 1)[1].strip()
        if "：" in merchant:
            merchant = merchant.split("：", 1)[1].strip()
        return merchant or seller

    in_stock_count = 0
    out_of_stock_count = 0

    for r in results:
        name   = r.get("name") or "商品"
        price  = clean_cell(r.get("price"))
        stock  = clean_cell(r.get("stock"))
        seller = clean_cell(r.get("seller"))
        url    = clean_cell(r.get("url"))
        delivery = clean_cell(r.get("delivery"))
        overtime = clean_cell(r.get("overtime"))
        basis_price = clean_cell(r.get("basis_price"))
        deal_tag = clean_cell(r.get("deal_tag"))

        is_out = stock in ("缺货", "应该无库存")
        stock_icon = "❌" if is_out else "✅"

        if is_out:
            out_of_stock_count += 1
            price_str = "缺货"
        else:
            in_stock_count += 1
            price_str = price if price and price != "\\" else "价格获取失败"

        buybox_stolen = bool(seller and seller not in ("未知", "") and not is_amazon_seller(seller))
        is_overtime = bool(overtime and overtime != "未超时")
        is_missing_basis = basis_price == "无basis price"

        row = [
            text_tag(f"{stock_icon} "),
            {"tag": "a",    "text": name, "href": url},
            text_tag(f"   {price_str}"),
        ]
        if buybox_stolen:
            row.append(text_tag(f"  [Buybox 被抢: {seller_name(seller)}]", bold=True))
        if delivery:
            row.append(text_tag(f"  {delivery}", bold=is_overtime))
        if basis_price:
            row.append(text_tag(f"  {basis_price}", bold=is_missing_basis))
        if deal_tag:
            row.append(text_tag(f"  [DT: {deal_tag}]"))
        post_content.append(row)

    # 汇总行
    total = len(results)
    summary = (
        f"共 {total} 条 | 有货 {in_stock_count} | 缺货 {out_of_stock_count}"
        f" | 耗时 {elapsed_minutes:.1f} 分钟"
    )
    post_content.append([{"tag": "text", "text": summary}])

    token = get_feishu_token()
    ok_count = 0
    for chat_id in FEISHU_CHAT_IDS:
        payload = {
            "receive_id": chat_id,
            "msg_type": "post",
            "content": json.dumps(
                {"zh_cn": {"title": title, "content": post_content}},
                ensure_ascii=False,
            ),
        }
        resp = requests.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=15,
        )
        if resp.json().get("code") != 0:
            logger.error(f"飞书富文本发送失败({chat_id}): {resp.json()}")
        else:
            ok_count += 1
    logger.info(f"飞书报告发送成功 {ok_count}/{len(FEISHU_CHAT_IDS)}")

# ===================== 数据处理 =====================
def load_input_data():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"找不到输入文件：{INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
    logger.info(f"已加载 {INPUT_FILE.name}，共 {len(df)} 行")
    return df


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    price_col = SCRAPER_CONFIG["price_column"]
    required_cols = price_col + 7
    for i in range(len(df.columns), required_cols):
        df[f"Column_{i}"] = pd.Series("", index=df.index, dtype=object)
    # 确保写入列是 object 类型，避免写字符串时的 FutureWarning
    for col in df.columns[price_col:]:
        df[col] = df[col].astype(object)
    return df


def clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_url(value) -> str:
    text = clean_cell(value)
    if not text:
        return ""
    if text.startswith("//"):
        return "https:" + text
    if text.startswith("www.amazon.") or text.startswith("amazon."):
        return "https://" + text
    return text


def build_results_from_df(df: pd.DataFrame) -> list[dict]:
    """从爬取完毕的 DataFrame 提取结构化结果供飞书报告使用"""
    url_col      = SCRAPER_CONFIG["url_column"]
    price_col    = SCRAPER_CONFIG["price_column"]
    seller_col   = price_col + 1
    stock_col    = price_col + 2
    delivery_col = price_col + 3
    overtime_col = price_col + 4
    basis_col    = price_col + 5
    deal_col     = price_col + 6

    results = []
    for _, row in df.iterrows():
        url = normalize_url(row.iloc[url_col]) if url_col < len(row) else ""
        if not url.startswith("http"):
            continue

        # 产品名称：取 model + size 列（列1=Model, 列2=Size）
        model = clean_cell(row.iloc[1]) if 1 < len(row) else ""
        size  = clean_cell(row.iloc[2]) if 2 < len(row) else ""
        name  = f"{model} {size}".strip() or "小米产品"

        results.append({
            "name":     name,
            "url":      url,
            "price":    clean_cell(row.iloc[price_col])    if price_col    < len(row) else "",
            "seller":   clean_cell(row.iloc[seller_col])   if seller_col   < len(row) else "",
            "stock":    clean_cell(row.iloc[stock_col])    if stock_col    < len(row) else "",
            "delivery": clean_cell(row.iloc[delivery_col]) if delivery_col < len(row) else "",
            "overtime": clean_cell(row.iloc[overtime_col]) if overtime_col < len(row) else "",
            "basis_price": clean_cell(row.iloc[basis_col]) if basis_col    < len(row) else "",
            "deal_tag": clean_cell(row.iloc[deal_col])      if deal_col     < len(row) else "",
        })
    return results


def extract_asin(url: str) -> str:
    match = re.search(r"/(?:dp|gp/product)/([A-Z0-9]{10})", url)
    return match.group(1) if match else ""


def parse_price_amount(price_text: str) -> tuple[str, float | None]:
    text = clean_cell(price_text)
    if not text or text == "\\":
        return "", None
    currency = "EUR" if "€" in text or "EUR" in text.upper() else ""
    number_text = re.sub(r"[^0-9,\.]", "", text)
    if not number_text:
        return currency, None
    number_text = normalize_price_number(number_text)
    try:
        return currency, float(number_text)
    except ValueError:
        return currency, None


def normalize_price_number(number_text: str) -> str:
    comma_pos = number_text.rfind(",")
    dot_pos = number_text.rfind(".")
    if comma_pos >= 0 and dot_pos >= 0:
        if dot_pos > comma_pos:
            return number_text.replace(",", "")
        return number_text.replace(".", "").replace(",", ".")
    if comma_pos >= 0:
        tail_len = len(number_text) - comma_pos - 1
        return number_text.replace(",", ".") if tail_len == 2 else number_text.replace(",", "")
    if dot_pos >= 0:
        tail_len = len(number_text) - dot_pos - 1
        return number_text if tail_len == 2 else number_text.replace(".", "")
    return number_text


def build_model_group(model_name: str) -> str:
    text = clean_cell(model_name).upper().replace(" ", "")
    if not text:
        return ""
    if text.startswith("FPRO"):
        size = re.search(r"(\d+)", text)
        return f"FPro{size.group(1)}" if size else "FPro"
    if text.startswith("F"):
        size = re.search(r"(\d+)", text)
        return f"F{size.group(1)}" if size else "F"
    if text.startswith("MAX"):
        size = re.search(r"(\d+)", text)
        return f"Max{size.group(1)}" if size else "Max"
    if text.startswith("O32"):
        return "O32/SMPro"
    if text.startswith("N32"):
        return "N32"
    if text.startswith("P37M"):
        return "P37M"
    if text.startswith("P37"):
        return "P37"
    return clean_cell(model_name).replace(" ", "")


def export_price_outputs(df: pd.DataFrame, captured_at: datetime) -> tuple[Path, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    date_str = captured_at.strftime("%Y-%m-%d")
    prices_dir = OUTPUT_DIR / "prices"
    day_dir = OUTPUT_DIR / "prices" / f"date={date_str}"
    prices_dir.mkdir(parents=True, exist_ok=True)
    day_dir.mkdir(parents=True, exist_ok=True)

    excel_path = OUTPUT_DIR / "latest.xlsx"
    price_csv_path = day_dir / "prices.csv"
    cumulative_csv_path = prices_dir / "prices.csv"

    df.to_excel(excel_path, index=False)

    price_col    = SCRAPER_CONFIG["price_column"]
    seller_col   = price_col + 1
    stock_col    = price_col + 2
    delivery_col = price_col + 3
    overtime_col = price_col + 4
    basis_col    = price_col + 5
    deal_col     = price_col + 6
    url_col      = SCRAPER_CONFIG["url_column"]

    rows = []
    for _, row in df.iterrows():
        url = normalize_url(row.iloc[url_col]) if url_col < len(row) else ""
        if not url.startswith("http"):
            continue
        model_name = clean_cell(row.get("model", "")) or " ".join(
            part for part in [clean_cell(row.get("Model", "")), clean_cell(row.get("Size", ""))] if part
        )
        price_text = clean_cell(row.iloc[price_col]) if price_col < len(row) else ""
        currency, price_amount = parse_price_amount(price_text)
        rows.append(
            {
                "date": date_str,
                "captured_at": captured_at.isoformat(timespec="seconds"),
                "marketplace": "DE",
                "country_code": COUNTRY_CODE.upper(),
                "brand": clean_cell(row.get("brand", "")),
                "model_group": build_model_group(model_name),
                "model_name": model_name,
                "asin": extract_asin(url),
                "url": url,
                "price_text": price_text,
                "price_amount": price_amount,
                "currency": currency,
                "seller": clean_cell(row.iloc[seller_col]) if seller_col < len(row) else "",
                "stock": clean_cell(row.iloc[stock_col]) if stock_col < len(row) else "",
                "delivery": clean_cell(row.iloc[delivery_col]) if delivery_col < len(row) else "",
                "overtime": clean_cell(row.iloc[overtime_col]) if overtime_col < len(row) else "",
                "basis_price": clean_cell(row.iloc[basis_col]) if basis_col < len(row) else "",
                "deal_tag": clean_cell(row.iloc[deal_col]) if deal_col < len(row) else "",
            }
        )

    prices_df = pd.DataFrame(rows)
    prices_df.to_csv(price_csv_path, index=False, encoding="utf-8-sig")
    update_cumulative_prices(cumulative_csv_path, prices_df, date_str)
    (day_dir / "prices.done").write_text(f"captured_at={captured_at.isoformat(timespec='seconds')}\n", encoding="utf-8")
    (prices_dir / "prices.done").write_text(f"captured_at={captured_at.isoformat(timespec='seconds')}\n", encoding="utf-8")
    return excel_path, cumulative_csv_path


def update_cumulative_prices(path: Path, today_df: pd.DataFrame, date_str: str) -> None:
    if path.exists():
        existing_df = pd.read_csv(path, encoding="utf-8-sig")
        combined_df = pd.concat([existing_df, today_df], ignore_index=True)
    else:
        combined_df = today_df
    combined_df = combined_df.sort_values(["date", "captured_at", "model_group", "model_name", "asin"]).reset_index(drop=True)
    combined_df.to_csv(path, index=False, encoding="utf-8-sig")


def write_prices_to_feishu_sheet(df: pd.DataFrame, captured_at: datetime) -> None:
    if not FEISHU_PRICE_SHEET_TOKEN or not FEISHU_PRICE_SHEET_ID:
        logger.info("Skip Feishu price sheet write: token or sheet id is not configured")
        return

    lark_cli = shutil.which("lark-cli") or shutil.which("lark-cli.cmd")
    if not lark_cli:
        raise RuntimeError("lark-cli not found; cannot write price calendar sheet")

    workbook = run_lark_cli_json(
        lark_cli,
        [
            "sheets",
            "+workbook-info",
            "--as",
            FEISHU_PRICE_SHEET_AS,
            "--spreadsheet-token",
            FEISHU_PRICE_SHEET_TOKEN,
            "--format",
            "json",
        ],
    )
    sheet_info = next(
        (sheet for sheet in workbook.get("data", {}).get("sheets", []) if sheet.get("sheet_id") == FEISHU_PRICE_SHEET_ID),
        None,
    )
    if not sheet_info:
        raise RuntimeError(f"sheet_id {FEISHU_PRICE_SHEET_ID} not found in workbook")

    row_count = int(sheet_info.get("row_count") or 200)
    col_count = int(sheet_info.get("column_count") or 21)
    last_col = column_number_to_letter(max(col_count, column_letter_to_number(FEISHU_PRICE_FIRST_PRICE_COLUMN)))
    table = run_lark_cli_json(
        lark_cli,
        [
            "sheets",
            "+csv-get",
            "--as",
            FEISHU_PRICE_SHEET_AS,
            "--spreadsheet-token",
            FEISHU_PRICE_SHEET_TOKEN,
            "--sheet-id",
            FEISHU_PRICE_SHEET_ID,
            "--range",
            f"A2:{last_col}{row_count}",
            "--format",
            "json",
        ],
    )

    rows = parse_annotated_csv(table.get("data", {}).get("annotated_csv", ""))
    if not rows:
        raise RuntimeError("price calendar sheet returned no rows")

    header_row = rows[0]["values"]
    first_price_index = column_letter_to_number(FEISHU_PRICE_FIRST_PRICE_COLUMN) - 1
    next_col_index = next_price_column_index(header_row, first_price_index)
    next_col = column_number_to_letter(next_col_index + 1)

    price_by_asin = build_price_by_asin(df)
    data_rows = [row for row in rows[1:] if len(row["values"]) >= 4 and clean_cell(row["values"][3])]
    if not data_rows:
        raise RuntimeError("price calendar sheet has no ASIN rows to update")

    max_sheet_row = max(row["row_number"] for row in data_rows)
    values_by_row = {row["row_number"]: price_by_asin.get(clean_cell(row["values"][3]), "") for row in data_rows}
    capture_label = format_capture_label(captured_at)
    column_values = [capture_label] + [values_by_row.get(row_number, "") for row_number in range(3, max_sheet_row + 1)]
    csv_payload = make_single_column_csv(column_values)

    result = run_lark_cli_json(
        lark_cli,
        [
            "sheets",
            "+csv-put",
            "--as",
            FEISHU_PRICE_SHEET_AS,
            "--spreadsheet-token",
            FEISHU_PRICE_SHEET_TOKEN,
            "--sheet-id",
            FEISHU_PRICE_SHEET_ID,
            "--start-cell",
            f"{next_col}2",
            "--csv",
            csv_payload,
            "--format",
            "json",
        ],
    )
    verify = run_lark_cli_json(
        lark_cli,
        [
            "sheets",
            "+csv-get",
            "--as",
            FEISHU_PRICE_SHEET_AS,
            "--spreadsheet-token",
            FEISHU_PRICE_SHEET_TOKEN,
            "--sheet-id",
            FEISHU_PRICE_SHEET_ID,
            "--range",
            f"{next_col}2:{next_col}2",
            "--format",
            "json",
        ],
    )
    if capture_label not in verify.get("data", {}).get("annotated_csv", ""):
        raise RuntimeError(f"Feishu price sheet write verification failed for {next_col}2")
    updated = result.get("data", {}).get("updated_cells_count", 0)
    logger.info(f"Wrote Feishu price sheet column {next_col} with {updated} cells")


def run_lark_cli_json(lark_cli: str, args: list[str]) -> dict:
    proc = subprocess.run([lark_cli, *args], capture_output=True, text=True, encoding="utf-8", timeout=120)
    if proc.returncode != 0:
        raise RuntimeError(f"lark-cli failed ({proc.returncode}): {proc.stderr or proc.stdout}")
    return json.loads(proc.stdout)


def parse_annotated_csv(annotated_csv: str) -> list[dict]:
    rows = []
    for line in annotated_csv.splitlines():
        match = re.match(r"^\[row=(\d+)\]\s?(.*)$", line)
        if not match:
            continue
        values = next(csv.reader([match.group(2)]))
        rows.append({"row_number": int(match.group(1)), "values": values})
    return rows


def build_price_by_asin(df: pd.DataFrame) -> dict[str, object]:
    url_col = SCRAPER_CONFIG["url_column"]
    price_col = SCRAPER_CONFIG["price_column"]
    stock_col = price_col + 2
    price_by_asin = {}
    for _, row in df.iterrows():
        url = normalize_url(row.iloc[url_col]) if url_col < len(row) else ""
        asin = extract_asin(url)
        if not asin:
            continue
        price_text = clean_cell(row.iloc[price_col]) if price_col < len(row) else ""
        currency, price_amount = parse_price_amount(price_text)
        if price_amount is not None:
            price_by_asin[asin] = format_price_for_sheet(price_amount, currency)
            continue
        stock = clean_cell(row.iloc[stock_col]) if stock_col < len(row) else ""
        price_by_asin[asin] = stock or price_text
    return price_by_asin


def format_price_for_sheet(price_amount: float, currency: str) -> str:
    prefix = "€" if currency == "EUR" else ""
    return f"{prefix}{price_amount:,.2f}"


def next_price_column_index(header_row: list[str], first_price_index: int) -> int:
    last_seen = first_price_index - 1
    for index in range(first_price_index, len(header_row)):
        if clean_cell(header_row[index]):
            last_seen = index
    return last_seen + 1


def make_single_column_csv(values: list[object]) -> str:
    buffer = StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    for value in values:
        writer.writerow([value])
    return buffer.getvalue().rstrip("\n")


def format_capture_label(captured_at: datetime) -> str:
    return f"{captured_at.year}/{captured_at.month}/{captured_at.day} {captured_at.hour}:{captured_at.minute:02d}"


def column_letter_to_number(column: str) -> int:
    number = 0
    for char in column.upper():
        if not ("A" <= char <= "Z"):
            continue
        number = number * 26 + ord(char) - ord("A") + 1
    return number


def column_number_to_letter(number: int) -> str:
    letters = []
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))

# ===================== 浏览器 =====================
def get_chrome_options() -> Options:
    opts = Options()
    if BROWSER_CONFIG["headless"]:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument(f"--window-size={BROWSER_CONFIG['window_size']}")
    opts.add_argument("--disable-notifications")
    opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-default-apps")
    opts.add_argument("--disable-sync")
    opts.add_argument("--metrics-recording-only")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--no-proxy-server")
    opts.add_argument("--remote-debugging-port=0")
    if BROWSER_CONFIG["disable_images"]:
        opts.add_argument("--disable-images")
        opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.page_load_strategy = BROWSER_CONFIG["page_load_strategy"]
    return opts


def create_driver() -> webdriver.Chrome:
    # Selenium 4.6+ 内置 Selenium Manager，自动下载匹配的 chromedriver
    last_error = None
    for attempt in range(1, 4):
        try:
            driver = webdriver.Chrome(options=get_chrome_options())
            driver.set_page_load_timeout(SCRAPER_CONFIG["timeout"])
            driver.set_script_timeout(10)
            return driver
        except Exception as exc:
            last_error = exc
            logger.warning(f"Chrome 启动失败，第 {attempt}/3 次: {exc}")
            time.sleep(5 * attempt)
    raise last_error

# ===================== Browser warm-up =====================
def warm_browser_session(driver, country_code: str) -> bool:
    logger.info("Using fresh Chrome session")
    return True

# ===================== 页面信息提取 =====================
def is_error_page(driver) -> bool:
    try:
        title = driver.title.lower()
        body  = driver.find_element(By.TAG_NAME, "body").text.lower()
        return (
            "503" in title or "reison" in title
            or "nous sommes désolés" in body
            or "sorry" in body
            or "something went wrong" in body
        )
    except Exception:
        return True


def get_stock_info(driver) -> str:
    try:
        out_kws = [
            "unavailable", "currently unavailable", "out of stock",
            "temporarily out of stock", "no stock", "no offers available",
            "no featured offers available", "ausverkauft", "nicht verfügbar",
            "derzeit nicht verfügbar", "nicht auf lager", "niet op voorraad",
            "non disponibile", "no disponible", "de momento agotado",
        ]
        in_kws = [
            "in stock", "available at checkout", "auf lager", "verfügbar",
            "en stock", "disponibile", "disponible",
        ]

        # 只有二手区块、没有新品购买按钮 → 缺货
        if driver.find_elements(By.CSS_SELECTOR, "#usedBuySection"):
            if not driver.find_elements(By.CSS_SELECTOR, "#newAccordionRow, #newBuyBoxAccordion"):
                return "缺货"

        # No featured offers / No offers → 缺货
        for selector in (
            "#fod-cx-message-with-learn-more",
            "#outOfStock",
            "#availability",
        ):
            for elem in driver.find_elements(By.CSS_SELECTOR, selector):
                text = (elem.text or "").strip().lower()
                if text and any(kw in text for kw in out_kws):
                    return "缺货"

        # Amazon 新版 offer 面板有时没有传统 add-to-cart id。
        # 只要能读到 Sold by / Shipper Seller，且页面显示可购买状态，就算有货。
        seller = get_seller_info(driver)
        availability_text = get_availability_text(driver).lower()
        if seller and seller != "未知" and any(kw in availability_text for kw in in_kws):
            return "有货"

        # 有明确购买按钮才算有货。价格/RRP 存在不代表可买。
        if has_purchase_button(driver):
            return "有货"

        # 兜底再检查 availability：明确缺货直接缺货，但即便写着 in stock，
        # 没有购买按钮也不算可买。
        avail_elems = driver.find_elements(By.CSS_SELECTOR, "#availability")
        if avail_elems:
            avail = (avail_elems[0].text or "").strip().lower()
            if any(kw in avail for kw in out_kws):
                return "缺货"
            if avail:
                return "应该无库存"

        # 老 buy-box 容器存在但无购买按钮时，不再算有货。
        if driver.find_elements(By.CSS_SELECTOR, "#buy-box"):
            return "应该无库存"

        return "应该无库存"
    except Exception:
        return "未知"


def get_availability_text(driver) -> str:
    texts = []
    for selector in (
        "#availability",
        "#outOfStock",
        "#availabilityInsideBuyBox_feature_div",
        "#desktop_buybox",
        "#buybox",
        "#buy-box",
    ):
        for elem in driver.find_elements(By.CSS_SELECTOR, selector):
            text = (elem.text or "").strip()
            if not text:
                try:
                    text = (driver.execute_script("return arguments[0].textContent;", elem) or "").strip()
                except Exception:
                    text = ""
            text = re.sub(r"\s+", " ", text)
            if text:
                texts.append(text)
    return " | ".join(texts)


def has_purchase_button(driver) -> bool:
    try:
        purchase_kws = (
            "add to basket",
            "add to cart",
            "buy now",
            "in den einkaufswagen",
            "jetzt kaufen",
            "zum warenkorb hinzufügen",
        )
        reject_kws = (
            "add to list",
            "join prime",
            "see all buying options",
            "other sellers",
        )

        for selector in (
            "#add-to-cart-button",
            "#buy-now-button",
            "input[name='submit.add-to-cart']",
            "input[name='submit.buy-now']",
            "#one-click-button",
            "input.a-button-input[type='submit']",
            "button.a-button-input[type='submit']",
        ):
            for elem in driver.find_elements(By.CSS_SELECTOR, selector):
                disabled = elem.get_attribute("disabled") or elem.get_attribute("aria-disabled")
                if not elem.is_displayed() or not elem.is_enabled() or str(disabled).lower() in ("true", "disabled"):
                    continue

                text_parts = [
                    elem.get_attribute("id"),
                    elem.get_attribute("name"),
                    elem.get_attribute("value"),
                    elem.get_attribute("aria-label"),
                    elem.text,
                ]
                labelled_by = elem.get_attribute("aria-labelledby")
                if labelled_by:
                    for label_id in labelled_by.split():
                        labels = driver.find_elements(By.ID, label_id)
                        if labels:
                            text_parts.append(labels[0].text)

                try:
                    parent_text = driver.execute_script(
                        "return arguments[0].closest('span.a-button, div.a-button, form, div')?.innerText || '';",
                        elem,
                    )
                    text_parts.append(parent_text)
                except Exception:
                    pass

                text = " ".join(part for part in text_parts if part).lower()
                if any(kw in text for kw in reject_kws):
                    continue
                if selector != "input.a-button-input[type='submit']" and selector != "button.a-button-input[type='submit']":
                    return True
                if any(kw in text for kw in purchase_kws):
                    return True
        return False
    except Exception:
        return False


def get_seller_info(driver) -> str:
    """
    返回格式：「售: <卖家> | 发: <发货方>」
    发货方即 Shipping from（Amazon 仓库 or 第三方）
    """
    try:
        def block_text(block_id: str) -> str:
            # 先尝试 offer-display-feature-text-message（桌面端）
            elems = driver.find_elements(
                By.CSS_SELECTOR, f"#{block_id} span.offer-display-feature-text-message"
            )
            if elems:
                return (elems[0].text or "").strip()
            # 回退：直接取块内所有文字
            elems2 = driver.find_elements(By.CSS_SELECTOR, f"#{block_id}")
            return (elems2[0].text or "").strip() if elems2 else ""

        merchant  = block_text("merchantInfoFeature_feature_div")   # 卖家（Sold by）
        fulfiller = block_text("fulfillerInfoFeature_feature_div")   # 发货方（Ships from）

        def css_text(selector: str) -> str:
            for elem in driver.find_elements(By.CSS_SELECTOR, selector):
                text = (elem.text or "").strip()
                if not text:
                    text = (driver.execute_script("return arguments[0].textContent;", elem) or "").strip()
                text = re.sub(r"\s+", " ", text)
                if text:
                    return text
            return ""

        if not merchant:
            merchant = (
                css_text("#merchantInfoFeature_feature_div div.odf-truncation-popover > span.offer-display-feature-text-message")
                or css_text("#merchantInfoFeature_feature_div span.offer-display-feature-text-message")
                or css_text("div.odf-truncation-popover > span.offer-display-feature-text-message")
                or css_text("#sellerProfileTriggerId")
                or css_text("#merchant-info a")
                or css_text("#merchant-info")
            )
        if not fulfiller:
            fulfiller = (
                css_text("#fulfillerInfoFeature_feature_div div.odf-truncation-popover > span.offer-display-feature-text-message")
                or css_text("#fulfillerInfoFeature_feature_div span.offer-display-feature-text-message")
            )

        parts = []
        if merchant:
            parts.append(f"售:{merchant}")
        if fulfiller:
            parts.append(f"发:{fulfiller}")
        return " | ".join(parts) if parts else "未知"
    except Exception:
        return "未知"


def get_basis_price_info(driver) -> str:
    """Return RRP/Was price information from span.basisPrice."""
    elems = driver.find_elements(By.CSS_SELECTOR, "span.basisPrice")
    if not elems:
        return "无basis price"

    texts = []
    for elem in elems:
        text = (elem.text or "").strip()
        if not text:
            text = (driver.execute_script("return arguments[0].textContent;", elem) or "").strip()
        text = re.sub(r"\s+", " ", text)
        if text:
            texts.append(text)

    if not texts:
        return "无basis price"

    raw = " | ".join(texts)
    amount_match = re.search(r"(?:€|EUR)\s*[\d.,]+|[\d.,]+\s*(?:€|EUR)", raw, re.I)
    amount = amount_match.group(0).strip() if amount_match else ""
    raw_lower = raw.lower()

    if "rrp" in raw_lower:
        label = "RRP"
    elif "was price" in raw_lower or re.search(r"\bwas\b", raw_lower):
        label = "Was price"
    else:
        label = "basis price"

    return f"{label} {amount}".strip()


def get_deal_tag_info(driver) -> str:
    elems = driver.find_elements(By.CSS_SELECTOR, "span#dealBadgeSupportingText > span")
    for elem in elems:
        text = (elem.text or "").strip()
        if not text:
            text = (driver.execute_script("return arguments[0].textContent;", elem) or "").strip()
        text = re.sub(r"\s+", " ", text)
        if text:
            return text
    return ""


def get_delivery_info(driver, wait=None) -> str:
    if wait is None:
        wait = WebDriverWait(driver, 8)

    # 方法1：data-csa-c-delivery-time 属性（最准确）
    try:
        elems = driver.find_elements(By.CSS_SELECTOR, "span[data-csa-c-delivery-time]")
        for elem in elems:
            val = (elem.get_attribute("data-csa-c-delivery-time") or "").strip()
            if val:
                return val
            txt = (elem.text or "").strip()
            if txt:
                return txt
    except Exception:
        pass

    # 方法2：#mir-layout-DELIVERY_BLOCK 块（含英文/德文日期行）
    try:
        block = driver.find_element(By.CSS_SELECTOR, "#mir-layout-DELIVERY_BLOCK")
        text = (block.text or "").strip()
        if text:
            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
            en_months = ["January","February","March","April","May","June",
                         "July","August","September","October","November","December",
                         "Jan","Feb","Mar","Apr","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
            de_months = ["Januar","Februar","März","April","Mai","Juni",
                         "Juli","August","September","Oktober","November","Dezember"]
            weekdays  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday",
                         "Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"]
            all_kws = [m.lower() for m in en_months + de_months + weekdays]
            for line in lines:
                if any(kw in line.lower() for kw in all_kws):
                    return line
            # 没匹配到月份/星期，返回第一行有意义的文字
            return lines[0] if lines else "配送时间获取失败"
    except Exception:
        pass

    # 方法3：#delivery-message / #ddm-deliver-by（部分页面布局）
    for sel in ("#delivery-message", "#ddm-deliver-by", "#deliveryBlockMessage"):
        try:
            elem = driver.find_element(By.CSS_SELECTOR, sel)
            txt = (elem.text or "").strip()
            if txt:
                return txt
        except Exception:
            continue

    return "配送时间获取失败"


def parse_delivery_date(date_text: str):
    try:
        if re.search(r'\d{2}:\d{2}\s*-\s*\d{2}:\d{2}', date_text):
            match = re.search(r'(\d+)\s+([A-Za-z]+)', date_text)
        elif " - " in date_text and not re.search(r'\d{2}:\d{2}', date_text):
            end_date = date_text.split(" - ")[1]
            match = re.search(r'(\d+)\s+([A-Za-z]+)', end_date)
        else:
            match = re.search(r'(\d+)\s+([A-Za-z]+)', date_text)
        if not match:
            return None
        day, month_str = int(match.group(1)), match.group(2).replace(',', '').strip()
        months = {
            'January':1,'February':2,'March':3,'April':4,
            'May':5,'June':6,'July':7,'August':8,
            'September':9,'October':10,'November':11,'December':12,
        }
        month = next(
            (num for name, num in months.items() if name.lower().startswith(month_str.lower())),
            None
        )
        if month is None:
            return None
        year = datetime.now().year
        if month < datetime.now().month:
            year += 1
        return datetime(year, month, day)
    except Exception:
        return None


def calculate_delivery_status(delivery_date, country_code: str):
    if not delivery_date:
        return "无法解析日期", "无配送信息"
    tz_map = {"fr":"Europe/Paris","it":"Europe/Rome","uk":"Europe/London",
              "es":"Europe/Madrid","de":"Europe/Berlin"}
    local_tz = pytz.timezone(tz_map.get(country_code, "Europe/Berlin"))
    today = datetime.now(local_tz).replace(tzinfo=None)
    days_diff = (delivery_date - today).days + 1
    overtime = f"超时{days_diff - 7}天" if days_diff > 7 else "未超时"
    return f"{days_diff}天", overtime

# ===================== 核心爬虫 =====================
class AmazonScraper:
    def __init__(self, df: pd.DataFrame, country_code: str):
        self.df = df
        self.country_code = country_code
        self.retry_queue = queue.Queue()
        self.results: dict = {}

    PRICE_CONTAINERS = [
        "#corePrice_feature_div",
        "#corePriceDisplay_desktop_feature_div",
        "#apex_desktop_newAccordionRow",
        "#price_inside_buybox",
    ]

    @staticmethod
    def _clean_whole(raw: str) -> str:
        # span.a-price-whole 内部有子元素 <span class="a-price-decimal">.</span>
        # Selenium .text 会把子元素文字换行追加，例如 "449\n."，需要去掉
        return raw.replace("\n", "").strip().rstrip(".,").strip()

    def extract_price(self, driver, index) -> str:
        # 方法1：直接读 textContent（最干净，不含子元素换行）
        for sel in self.PRICE_CONTAINERS:
            try:
                container = driver.find_element(By.CSS_SELECTOR, sel)
                currency = container.find_element(By.CSS_SELECTOR, "span.a-price-symbol").text.strip()
                wholes    = container.find_elements(By.CSS_SELECTOR, "span.a-price-whole")
                fractions = container.find_elements(By.CSS_SELECTOR, "span.a-price-fraction")
                for w, f in zip(wholes, fractions):
                    # 用 textContent 而非 .text，避免子元素换行问题
                    whole_raw = driver.execute_script("return arguments[0].textContent;", w)
                    whole = self._clean_whole(whole_raw)
                    if whole and whole.replace(",", "").replace(".", "").isdigit():
                        price = f"{currency}{whole}.{f.text.strip()}"
                        logger.info(f"URL {index}: 价格 {price}")
                        return price
            except Exception:
                continue
        # 方法2：回退全页扫描
        try:
            currency  = driver.find_element(By.CSS_SELECTOR, "span.a-price-symbol").text.strip()
            wholes    = driver.find_elements(By.CSS_SELECTOR, "span.a-price-whole")
            fractions = driver.find_elements(By.CSS_SELECTOR, "span.a-price-fraction")
            for w, f in zip(wholes, fractions):
                whole_raw = driver.execute_script("return arguments[0].textContent;", w)
                whole = self._clean_whole(whole_raw)
                if whole and whole.replace(",", "").replace(".", "").isdigit():
                    price = f"{currency}{whole}.{f.text.strip()}"
                    logger.info(f"URL {index}: 价格 {price}")
                    return price
        except Exception:
            pass
        logger.warning(f"URL {index}: 未找到价格")
        return "\\"

    def process_single_url(self, driver, thread_id, index, url, retry_count):
        results = {}
        pc = SCRAPER_CONFIG["price_column"]
        seller_col   = pc + 1
        stock_col    = pc + 2
        delivery_col = pc + 3
        overtime_col = pc + 4
        basis_col    = pc + 5
        deal_col     = pc + 6

        try:
            try:
                logger.info(f"URL {index}: start {url}")
                driver.get(url)
                if is_error_page(driver):
                    driver.refresh()
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.TAG_NAME, "body"))
                        )
                    except TimeoutException:
                        if retry_count < SCRAPER_CONFIG["max_retries"]:
                            self.retry_queue.put((index, url, retry_count + 1))
                        else:
                            self.df.iloc[index, pc] = "访问失败"
                        return results

                stock = get_stock_info(driver)
                if stock in ("缺货", "应该无库存"):
                    results[index] = "\\"
                    self.df.iloc[index, pc]           = "\\"
                    self.df.iloc[index, seller_col]   = ""
                    self.df.iloc[index, stock_col]    = stock
                    self.df.iloc[index, delivery_col] = ""
                    self.df.iloc[index, overtime_col] = ""
                    self.df.iloc[index, basis_col]    = get_basis_price_info(driver)
                    self.df.iloc[index, deal_col]     = get_deal_tag_info(driver)
                    return results

                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "span.a-price-whole"))
                )
                price = self.extract_price(driver, index)
                results[index] = price
                self.df.iloc[index, pc]        = price
                self.df.iloc[index, stock_col] = stock

                try:
                    self.df.iloc[index, seller_col] = get_seller_info(driver)
                except Exception:
                    self.df.iloc[index, seller_col] = ""

                try:
                    raw = get_delivery_info(driver)
                    date = parse_delivery_date(raw) if isinstance(raw, str) else None
                    d_status, overtime = calculate_delivery_status(date, self.country_code)
                    self.df.iloc[index, delivery_col] = d_status
                    self.df.iloc[index, overtime_col] = overtime
                except Exception:
                    self.df.iloc[index, delivery_col] = ""
                    self.df.iloc[index, overtime_col] = ""

                try:
                    self.df.iloc[index, basis_col] = get_basis_price_info(driver)
                except Exception:
                    self.df.iloc[index, basis_col] = "无basis price"

                try:
                    self.df.iloc[index, deal_col] = get_deal_tag_info(driver)
                except Exception:
                    self.df.iloc[index, deal_col] = ""

            except TimeoutException:
                if retry_count < SCRAPER_CONFIG["max_retries"]:
                    self.retry_queue.put((index, url, retry_count + 1))
                else:
                    self.df.iloc[index, pc] = "\\"
                    self.df.iloc[index, basis_col] = ""
                    self.df.iloc[index, deal_col] = ""
            except Exception as e:
                logger.warning(f"线程 {thread_id}, URL {index}: {e}")
                if retry_count < SCRAPER_CONFIG["max_retries"]:
                    self.retry_queue.put((index, url, retry_count + 1))
                else:
                    self.df.iloc[index, pc] = "\\"
                    self.df.iloc[index, basis_col] = ""
                    self.df.iloc[index, deal_col] = ""
        finally:
            pass
        return results

    def process_urls_batch(self, thread_data):
        thread_id, urls_batch, _ = thread_data
        results = {}
        restart_every = max(1, SCRAPER_CONFIG.get("restart_every", 3))

        for offset in range(0, len(urls_batch), restart_every):
            chunk = urls_batch[offset:offset + restart_every]
            driver = None
            try:
                driver = create_driver()
                if not warm_browser_session(driver, self.country_code):
                    logger.error(f"线程 {thread_id}: 浏览器启动失败，跳过本组 URL")
                    continue
                for index, url, retry_count in chunk:
                    results.update(self.process_single_url(driver, thread_id, index, url, retry_count))
            finally:
                if driver:
                    driver.quit()
        return results

    def run(self):
        url_col = SCRAPER_CONFIG["url_column"]
        pc      = SCRAPER_CONFIG["price_column"]

        all_tasks = [
            (i, normalize_url(row.iloc[url_col]), 0)
            for i, row in self.df.iterrows()
            if normalize_url(row.iloc[url_col]).startswith("http")
        ]

        max_workers = SCRAPER_CONFIG["max_workers"]
        batch_size  = max(1, len(all_tasks) // max_workers + (1 if len(all_tasks) % max_workers else 0))
        batches = [
            (i // batch_size, all_tasks[i:i + batch_size], None)
            for i in range(0, len(all_tasks), batch_size)
        ]

        if max_workers == 1:
            pending = all_tasks
            round_no = 0
            while pending:
                self.retry_queue = queue.Queue()
                batch_result = self.process_urls_batch((round_no, pending, None))
                self.results.update(batch_result)
                for idx, res in batch_result.items():
                    self.df.iloc[idx, pc] = res
                logger.info(f"已处理 {len(self.results)} 个 URL")

                retry_batch = []
                while not self.retry_queue.empty():
                    retry_batch.append(self.retry_queue.get())
                pending = retry_batch
                round_no += 1
            logger.info(f"全部 URL 处理完成，共 {len(self.results)} 条")
            return self.results

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(self.process_urls_batch, b): b for b in batches}

            while futures:
                done, _ = concurrent.futures.wait(
                    futures, timeout=None, return_when=concurrent.futures.FIRST_COMPLETED
                )
                for future in done:
                    batch = futures.pop(future)
                    try:
                        batch_result = future.result()
                        self.results.update(batch_result)
                        for idx, res in batch_result.items():
                            self.df.iloc[idx, pc] = res
                        logger.info(f"已处理 {len(self.results)} 个 URL")
                    except Exception as e:
                        logger.error(f"批次处理失败: {e}")

                retry_batch = []
                while not self.retry_queue.empty():
                    retry_batch.append(self.retry_queue.get())
                if retry_batch:
                    retry_future = executor.submit(
                        self.process_urls_batch, (len(futures) + 999, retry_batch, None)
                    )
                    futures[retry_future] = (len(futures) + 999, retry_batch, None)

        logger.info(f"全部 URL 处理完成，共 {len(self.results)} 条")
        return self.results

# ===================== 主流程 =====================
def main():
    start_time = time.time()
    now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    logger.info("Using fresh Chrome session")
    logger.info(f"===== 小米意大利 TV 巡店开始 {now_str} =====")

    df = load_input_data()
    df = prepare_dataframe(df)

    scraper = AmazonScraper(df, COUNTRY_CODE)
    scraped = scraper.run()
    if not scraped:
        raise RuntimeError("本次没有成功抓取任何 URL，请检查 Selenium/Chrome 是否启动失败或服务器内存是否不足")

    elapsed = (time.time() - start_time) / 60
    results = build_results_from_df(scraper.df)
    captured_at = datetime.now()
    excel_path, price_csv_path = export_price_outputs(scraper.df, captured_at)
    logger.info(f"Saved monitor workbook: {excel_path}")
    logger.info(f"Saved standard price CSV: {price_csv_path}")
    try:
        write_prices_to_feishu_sheet(scraper.df, captured_at)
    except Exception:
        logger.exception("Failed to write Feishu price calendar sheet")

    title = f"意大利 TV 巡店 {datetime.now().strftime('%Y/%m/%d %H:%M')}"
    send_feishu_report(results, title, elapsed)

    logger.info(f"===== 全部完成，耗时 {elapsed:.1f} 分钟 =====")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        logger.exception("Xiaomi monitor failed")
        try:
            send_feishu_message(f"小米意大利 TV 巡店运行失败：{exc}")
        except Exception:
            logger.exception("Failed to send Feishu failure message")
        raise


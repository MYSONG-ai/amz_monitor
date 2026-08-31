# TV Store Monitor

意大利 Amazon TV 巡店脚本。使用 Selenium 直接启动一个新的 Chrome 会话，不导入 Chrome profile，不读取 cookie。

## Files

- `tv_monitor.py`: 主脚本
- `input/de_xiaomi_TV.xlsx`: TV 商品输入表
- `.env`: 飞书配置，本地或服务器手动创建，不提交到 GitHub

## Setup

```bash
sudo apt update
sudo apt install -y git python3 python3-venv python3-pip wget curl unzip
wget -q https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt install -y ./google-chrome-stable_current_amd64.deb

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

For small VPS instances, add swap before running Chrome:

```bash
sudo fallocate -l 2G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile
echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab
```

## Test Without Feishu

```bash
source .venv/bin/activate
NO_FEISHU=1 python3 tv_monitor.py
```

## Run With Feishu

Create `.env`:

```env
FEISHU_APP_ID=your_app_id
FEISHU_APP_SECRET=your_app_secret
FEISHU_CHAT_IDS=your_chat_id
FEISHU_PRICE_SHEET_TOKEN=DLJMs1Q2Pht7xmtPxmPcmKhCnjd
FEISHU_PRICE_SHEET_ID=300dee
FEISHU_PRICE_SHEET_AS=user
```

Then run:

```bash
source .venv/bin/activate
python3 tv_monitor.py
```

## Price CSV Output

Each successful run now writes local output files:

```text
output/latest.xlsx
output/prices/prices.csv
output/prices/prices.done
output/prices/date=YYYY-MM-DD/prices.csv
output/prices/date=YYYY-MM-DD/prices.done
```

`output/prices/prices.csv` is the cumulative standard file for the sales forecast project. It has `date` and `captured_at` columns, and repeated runs on the same date are appended instead of overwritten. Important columns:

```text
date,captured_at,marketplace,model_group,model_name,asin,url,price_text,price_amount,currency,seller,stock,delivery,basis_price,deal_tag
```

On the local sales forecast machine, sync from the Germany server with:

```powershell
.\scripts\sync_price_from_germany.ps1 `
  -Date YYYY-MM-DD `
  -HostName 3.65.207.109 `
  -UserName ubuntu `
  -RemotePriceCsv /home/ubuntu/amz_monitor/output/prices/prices.csv
```

## Feishu Price Calendar Write

After each successful monitor run, the script also tries to write the latest prices to the Feishu price calendar sheet:

```text
https://rcne7vk00ucv.feishu.cn/wiki/EPcHwidX7iitmikzGjBcxl5WnDg
```

It matches rows by ASIN in column D, finds the next price snapshot column from column I onward, writes the capture time in row 2, and writes prices from row 3 down. If Lark authorization or network access fails, the monitor logs the error but still keeps the local CSV/XLSX outputs and Feishu chat report.

The server needs `lark-cli` installed and authorized for the configured identity. For the default `FEISHU_PRICE_SHEET_AS=user`, run the monitor once and, if it reports a missing sheet scope, run the exact `lark-cli auth login --scope "..."` command shown in the error on the server.

```bash
cd /home/ubuntu/amz_monitor
source .venv/bin/activate
python3 tv_monitor.py
```

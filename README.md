# TV Store Monitor

德国 Amazon TV 巡店脚本。使用 Selenium 直接启动一个新的 Chrome 会话，不导入 Chrome profile，不读取 cookie。

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
```

Then run:

```bash
source .venv/bin/activate
python3 tv_monitor.py
```

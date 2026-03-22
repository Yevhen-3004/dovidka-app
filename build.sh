#!/bin/bash
# Встановлення залежностей Python
pip install -r requirements.txt

# Встановлення шрифтів DejaVu (підтримка кирилиці)
apt-get install -y fonts-dejavu-core 2>/dev/null || true
fc-cache -f 2>/dev/null || true

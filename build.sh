#!/bin/bash
pip install -r requirements.txt
apt-get install -y fonts-dejavu-core libreoffice 2>/dev/null || true
fc-cache -f 2>/dev/null || true

#!/bin/bash
set -e

echo "===================================================================="
echo " Setup & Deploy Milkrun System di Hostinger VPS (rzlabs.tech)"
echo "===================================================================="

# 1. Update package system
echo "[1/6] Mengupdate paket sistem Linux..."
sudo apt update && sudo apt upgrade -y

# 2. Install dependensi inti (Python3, Venv, Nginx, Certbot, Git)
echo "[2/6] Menginstall Python3, Virtualenv, Nginx, dan Certbot..."
sudo apt install -y python3 python3-pip python3-venv nginx certbot python3-certbot-nginx git curl ufw

# 3. Setup direktori aplikasi /var/www/milkrun
echo "[3/6] Menyiapkan direktori aplikasi di /var/www/milkrun..."
sudo mkdir -p /var/www/milkrun
sudo chown -R $USER:$USER /var/www/milkrun

# Copy file jika dijalankan dari dalam folder project yang di-upload
CURRENT_DIR=$(pwd)
if [ "$CURRENT_DIR" != "/var/www/milkrun" ]; then
    cp -r * /var/www/milkrun/
fi

cd /var/www/milkrun

# 4. Setup Python Virtual Environment
echo "[4/6] Mengkonfigurasi Python virtual environment & dependensi..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

# 5. Konfigurasi Systemd Service (Uvicorn Daemon)
echo "[5/6] Mengkonfigurasi systemd service (milkrun.service)..."
sudo cp milkrun.service /etc/systemd/system/milkrun.service
sudo systemctl daemon-reload
sudo systemctl enable milkrun
sudo systemctl restart milkrun

# 6. Konfigurasi Nginx Reverse Proxy
echo "[6/6] Mengkonfigurasi Nginx untuk rzlabs.tech..."
sudo cp nginx_rzlabs.tech.conf /etc/nginx/sites-available/rzlabs.tech
sudo ln -sf /etc/nginx/sites-available/rzlabs.tech /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default

sudo nginx -t
sudo systemctl restart nginx

# Firewall Setup
sudo ufw allow OpenSSH
sudo ufw allow 'Nginx Full'
echo "y" | sudo ufw enable || true

echo ""
echo "===================================================================="
echo " [BERHASIL] Backend FastAPI & Nginx telah aktif di VPS!"
echo " Akses via IP Server atau http://rzlabs.tech"
echo ""
echo " Langkah Terakhir (Aktivasi HTTPS / SSL Gratis Let's Encrypt):"
echo " Jalankan perintah:"
echo "   sudo certbot --nginx -d rzlabs.tech -d www.rzlabs.tech"
echo "===================================================================="

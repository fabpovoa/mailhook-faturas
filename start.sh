#!/bin/bash
echo "🚀 Subindo mailhook..."

# Inicia ngrok em background
ngrok http 8000 --log=stdout > /tmp/ngrok.log &
sleep 3

# Pega a URL do ngrok
NGROK_URL=$(curl -s http://localhost:4040/api/tunnels | python3 -c "import sys,json; print(json.load(sys.stdin)['tunnels'][0]['public_url'])")
echo "🌐 ngrok URL: $NGROK_URL"
echo "🔗 Webhook Make: $NGROK_URL/webhook/make?token=15d4512b881a2dfee2aa4ff3e2fb8171"

# Inicia o servidor
uvicorn main:app --port 8000

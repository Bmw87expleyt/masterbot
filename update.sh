#!/bin/bash
cd /root/masterbot
git fetch origin main
LOCAL=$(git rev-parse HEAD)
REMOTE=$(git rev-parse origin/main)
if [ "$LOCAL" != "$REMOTE" ]; then
    git pull origin main
    screen -S bot -X quit
    screen -d -m -S bot bash -c "source /root/masterbot/venv/bin/activate && python3 /root/masterbot/bot.py"
fi

#!/bin/bash
cd /Users/jenny/lemon || exit 1

PREV_YM=$(date -v-1m +%Y%m)

/usr/bin/python3 預收.py "$PREV_YM" >> /Users/jenny/lemon/cron.log 2>&1
sleep 900
/usr/bin/python3 儲值金結算.py >> /Users/jenny/lemon/cron.log 2>&1
sleep 300
/usr/bin/python3 儲值金預收.py >> /Users/jenny/lemon/cron.log 2>&1
sleep 65100
/usr/bin/python3 上下半月訂單.py 1 >> /Users/jenny/lemon/cron.log 2>&1

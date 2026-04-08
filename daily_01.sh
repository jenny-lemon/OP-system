#!/bin/bash

export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
cd /Users/jenny/lemon || exit 1

echo "===== daily_01 start $(date) =====" >> /Users/jenny/lemon/cron.log

/usr/bin/python3 排班統計表.py >> /Users/jenny/lemon/cron.log 2>&1
echo "----- 排班統計表 finished $(date) -----" >> /Users/jenny/lemon/cron.log

sleep 120

/usr/bin/python3 專員班表.py >> /Users/jenny/lemon/cron.log 2>&1
echo "----- 專員班表 finished $(date) -----" >> /Users/jenny/lemon/cron.log

echo "===== daily_01 end $(date) =====" >> /Users/jenny/lemon/cron.log

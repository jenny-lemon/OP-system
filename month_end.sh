#!/bin/bash
cd /Users/jenny/lemon || exit 1

TODAY_YM=$(date +%Y%m)
NEXT_MONTH=$(date -v+1d +%m)
THIS_MONTH=$(date +%m)

if [ "$THIS_MONTH" != "$NEXT_MONTH" ]; then
  /usr/bin/python3 上下半月訂單.py 2 >> /Users/jenny/lemon/cron.log 2>&1
  sleep 900
  /usr/bin/python3 已退款.py "$TODAY_YM" >> /Users/jenny/lemon/cron.log 2>&1
else
  echo "Not month end: $(date)" >> /Users/jenny/lemon/cron.log
fi

#!/bin/bash
cd /Users/jenny/lemon || exit 1

/usr/bin/python3 當月次月訂單.py >> /Users/jenny/lemon/cron.log 2>&1
sleep 120
/usr/bin/python3 專員系統個資.py >> /Users/jenny/lemon/cron.log 2>&1

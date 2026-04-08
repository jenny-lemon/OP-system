#!/bin/bash
cd /Users/jenny/lemon || exit 1
/usr/bin/python3 業績報表.py >> /Users/jenny/lemon/cron.log 2>&1

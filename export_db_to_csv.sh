#!/bin/bash

sqlite3 -header -csv data/teamstracker.db "select * from presence_log;" > presence_logs.csv
sqlite3 -header -csv data/teamstracker.db "select * from ooo_log;" > ooo_logs.csv

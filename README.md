# NBU-rate-VBA
VBA script for receiving currencies rate in Excel from NBU (National Bank of Ukraine) API

Usage: \
NBU_RATE(currency_code; date)\
NBU_RATE_NAME(currency_name;date)

Data:\
currency_code - code of currency, default value = 840 (USD);\
carrency_name - name of currency, default = USD; \
data - date, default value = Today()

Any value can be skiped, if you are skiping first argument, it have to be empty. For example: NBU_RATE_NAME(;"01.11.2018") will return official rate of USD to Ukrainian hryvna in 01.11.2018 (28,118832).

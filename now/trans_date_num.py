###数字を日付に変換(手動でデータを修正する際に使用)###

import sys
from datetime import datetime, timedelta

args = sys.argv

print(datetime(1899, 12, 30) + timedelta(days=int(args[1])))
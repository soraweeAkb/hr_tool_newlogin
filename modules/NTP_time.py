import ntplib
import datetime
import pytz
import time

class NTP_DateTime():
    def __init__(self):
        self.time_server='pool.ntp.org'

    def get_datetime(self):
        client=ntplib.NTPClient()
        while True:
            try:
                res=client.request(self.time_server)
            except:
                time.sleep(0.5)
            else:
                break

        ts=res.tx_time
        t=datetime.datetime.fromtimestamp(ts,pytz.timezone('Asia/Bangkok'))
        return t

#NTP=NTP_DateTime()
#t=NTP.get_datetime()
#b=datetime.datetime(2020, 11, 1, 12, 34, 41, 907292,pytz.timezone('Asia/Bangkok'))

#print((t-b).seconds)
#_date=t.strftime('%d/%m/%Y')
#_time=t.strftime('%H:%M:%S')

#print(_date)
#print(_time)
#print(t)


import pymysql

db=pymysql.connect(host='210.1.31.3',
                   user='hr',
                   port=3306,
                   passwd='gwP6xTsA',
                   db='akaganeHR',)

cursor=db.cursor()
#sql="""CREATE TABLE leave_request (SERIAL int(20) primary key auto_increment, USER_ID char(20), USER_NAME char(50), TYPE char(50), APPLY_DTTM datetime, APPLY_DT char(20),START_DT date, START_LEN char(10), END_DT date, END_LEN char(10), SINGLE_DAY tinyint, DURING float, REMARKS text, LEADER tinyint, DM tinyint, MD tinyint, HR tinyint, CURRENT_TO char(20), CURRENT_PO char(20)) character set utf8"""
#sql="""CREATE TABLE time_card (SERIAL char(20) primary key, USER_ID char(10), CLOCK_IN datetime, CLOCK_OUT datetime,OUT_1 datetime, IN_1 datetime, OUT_2 datetime, IN_2 datetime, DAY_LAG tinyint) character set utf8"""
#sql="""DROP TABLE IF EXISTS leave_request"""
#sql="""DROP TABLE IF EXISTS time_card"""

#sql="""CREATE TABLE ot_request (SERIAL int(20) primary key auto_increment, USER_ID char(20), USER_NAME char(50), APPLY_DTTM datetime, APPLY_DT char(20), OT_DT date, START_TM time, END_TM time, DURING float, LEADER tinyint, DM tinyint, MD tinyint, HR tinyint, REMARKS text, CURRENT_TO char(20), CURRENT_PO char(20)) character set utf8"""
#sql="""DROP TABLE IF EXISTS ot_request"""

#sql="""CREATE TABLE book_meeting_room (SERIAL int(20) primary key auto_increment, APPLY_DTTM datetime, MEETING_DT date, START_TM time, END_TM time, USER_ID char(20), USER_NAME char(50), DIVISION char(10), CONTENTS text) character set utf8"""
#sql="""DROP TABLE IF EXISTS  book_meeting_room"""

#sql="""CREATE TABLE apply_late (SERIAL int(20) primary key auto_increment, USER_ID char(20), USER_NAME char(50), APPLY_DTTM datetime, LATE_DT date, CLOCKIN_TM time, REMARKS text, LEADER char(15), DM char(15),  HR char(15), CURRENT_TO char(20), CURRENT_PO char(20)) character set utf8"""

#sql="""CREATE TABLE forget_record (SERIAL int(20) primary key auto_increment, USER_ID char(20), USER_NAME char(50), REQUEST_DTTM datetime, CLOCK_DT date, CLOCK_IN datetime, CLOCK_OUT datetime, OUT_1 datetime, IN_1 datetime, OUT_2 datetime, IN_2 datetime, REMARKS text, LEADER char(15), DM char(15), HR char(15), CURRENT_TO char(20), CURRENT_PO char(20)) character set utf8"""
#sql="""DROP TABLE IF EXISTS  forget_record"""

#sql="""CREATE TABLE calendar (DATE date, WEEKDAY int(10), IF_WORK char(10), REMARKS char(20)) character set utf8"""
#sql="""DROP TABLE IF EXISTS  calendar"""

#cursor.execute(sql)
cursor.close()
db.close()


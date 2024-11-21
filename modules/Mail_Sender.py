import yagmail
import threading

class MailSender():
    def __init__(self):
        self.yag_server= yagmail.SMTP(user='akaganethai.it@gmail.com', password='lrnnclztdlsjuxzg', host='smtp.gmail.com',port='465')

    def send_request_mail(self, email_add, receiver_name, sender_name, mode):
        if mode=='leave':
            email_title=f'Notification: Leave request from {sender_name}'
            email_content=f'Dear {receiver_name},\n\n' \
                          f'{sender_name} has sent a request for taking leave, please approve or decline the request on AKT HR online system.\n\n' \
                          f'มีคำร้องขออนุมัติลาจาก{sender_name} กรุณาทำการอนุมัติ หรือ ไม่อนุมัติ ผ่านระบบHRออนไลน์ของAKTค่ะ\n\n' \
                          f'{sender_name}から休暇申請が届きました。AKT人事オンラインシステムで承認または拒否をしてください。\n\n' \
                          f'Sent from AKT HR online system.'

        elif mode=='ot':
            email_title = f'Notification: OT request from {sender_name}'
            email_content = f'Dear {receiver_name},\n\n' \
                            f'{sender_name} has sent a request for OT work, please approve or decline the request on AKT HR online system.\n\n' \
                            f'มีคำร้องขออนุมัติทำงานล่วงเวลาจาก{sender_name} กรุณาทำการอนุมัติ หรือ ไม่อนุมัติ ผ่านระบบHRออนไลน์ของAKTค่ะ\n\n' \
                            f'{sender_name}から残業申請が届きました。AKT人事オンラインシステムで承認または拒否をしてください。\n\n' \
                            f'Sent from AKT HR online system.'

        elif mode=='forget':
            email_title = f'Notification: Time card data recording request from {sender_name}'
            email_content =f'Dear {receiver_name},\n\n' \
                           f'{sender_name} forgot clock-in/out and has sent a request for updating the time card data, please approve or decline the request on AKT HR online system.\n\n' \
                           f'มีคำร้องขอแก้ไขเวลาclock-in/outจาก{sender_name} เนื่องจากลืมบันทึกเวลาเข้า/ออกงาน กรุณาทำการอนุมัติ หรือ ไม่อนุมัติ ผ่านระบบHRออนไลน์ของAKTค่ะ\n\n' \
                           f'{sender_name}は出勤・退勤のタイムカードを打刻し忘れたため、タイムカード時間の追加登録申請が届きました。AKT人事オンラインシステムで承認または拒否をしてください。\n\n' \
                           f'Sent from AKT HR online system.'

        else:  #elif mode=='late':
            email_title = f'Notification: Late clock-in request from {sender_name}'
            email_content = f'Dear {receiver_name},\n\n' \
                            f'{sender_name} has sent a request for late clock-in, please approve or decline the request on AKT HR online system.\n\n' \
                            f'มีคำร้องขออนุมัติเข้างานสายจาก{sender_name} กรุณาทำการอนุมัติ หรือ ไม่อนุมัติ ผ่านระบบHRออนไลน์ของAKTค่ะ\n\n' \
                            f'{sender_name}から遅刻出勤許可の申請が届きました。AKT人事オンラインシステムで承認または拒否をしてください。\n\n' \
                            f'Sent from AKT HR online system.'

        self.email_to = [f'{email_add}', ]
        self.email_title = email_title
        self.email_content = email_content

        th=threading.Thread(target=self.launching)
        th.start()
        th.join()
        #self.yag_server.send(to=self.email_to, subject=self.email_title, contents=self.email_content)
        #self.yag_server.close()

    def send_approved_mail(self, email_add, receiver_name, mode):
        if mode=='leave':
            email_title='Leave request approved'
            email_content=f'Dear {receiver_name},\n\n' \
                          f'Your leave request has been approved by all of the sections, please check AKT HR online system for further details.\n\n' \
                          f'คำร้องขออนุมัติการลาของท่านได้รับการอนุมัติแล้วค่ะ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                          f'あなたの休暇申請は各部門に承認されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                          f'Sent from AKT HR online system.'

        elif mode=='ot':
            email_title = 'OT request approved'
            email_content = f'Dear {receiver_name},\n\n' \
                            f'Your OT request has been approved by all of the sections, please check AKT HR online system for further details.\n\n' \
                            f'คำร้องขออนุมัติทำงานล่วงเวลาของท่านได้รับการอนุมัติแล้วค่ะ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                            f'あなたの残業申請は各部門に承認されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                            f'Sent from AKT HR online system.'

        elif mode=='forget':
            email_title = 'Time card data-recording request approved'
            email_content =f'Dear {receiver_name},\n\n' \
                           f'Your time card data-recording request has been approved by all of the sections, and the time card database has also been updated automatically, please check AKT HR online system for further details.\n\n' \
                           f'คำร้องขออนุมัติแก้ไขเวลาclock-in/outของท่านได้รับการอนุมัติแล้วค่ะ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                           f'タイムカードの出勤・退勤時間の追加登録申請は既に各部門に承認され、タイムカードのデータベースは自動更新されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                           f'Sent from AKT HR online system.'

        else:  #elif mode=='late':
            email_title = 'Late clock-in request approved'
            email_content = f'Dear {receiver_name},\n\n' \
                            f'Your late clock-in request has been approved by all of the sections, please check AKT HR online system for further details.\n\n' \
                            f'คำร้องขออนุมัติเข้างานสายของท่านได้รับการอนุมัติแล้วค่ะ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                            f'あなたの遅刻出勤許可の申請は各部門に承認されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                            f'Sent from AKT HR online system.'

        self.email_to = [f'{email_add}', ]
        self.email_title = email_title
        self.email_content = email_content

        th=threading.Thread(target=self.launching)
        th.start()
        th.join()

    def send_declined_mail(self, email_add, receiver_name, mode):
        if mode=='leave':
            email_title='Leave request declined'
            email_content=f'Dear {receiver_name},\n\n' \
                          f'Sorry, your leave request has been declined, please check AKT HR online system for further details.\n\n' \
                          f'ขอโทษค่ะ คำร้องขออนุมัติการลาของท่านไม่ได้รับการอนุมัติ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                          f'申し訳ございませんが、あなたの休暇申請は拒否されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                          f'Sent from AKT HR online system.'

        elif mode=='ot':
            email_title = 'OT request declined'
            email_content = f'Dear {receiver_name},\n\n' \
                            f'Sorry, your OT request has been declined, please check AKT HR online system for further details.\n\n' \
                            f'ขอโทษค่ะ คำร้องขออนุมัติทำงานล่วงเวลาของท่านไม่ได้รับการอนุมัติ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                            f'申し訳ございませんが、あなたの残業申請は拒否されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                            f'Sent from AKT HR online system.'

        elif mode=='forget':
            email_title = 'Time card data-recording request declined'
            email_content =f'Dear {receiver_name},\n\n' \
                           f'Sorry, your time card data-recording request has been declined, please check AKT HR online system for further details.\n\n' \
                           f'ขอโทษค่ะ คำร้องขออนุมัติแก้ไขเวลาclock-in/outของท่านไม่ได้รับการอนุมัติ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                           f'申し訳ございませんが、タイムカードの出勤・退勤時間の追加登録申請は拒否されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                           f'Sent from AKT HR online system.'

        else:  #elif mode=='late':
            email_title = 'Late clock-in request declined'
            email_content = f'Dear {receiver_name},\n\n' \
                            f'Sorry, your late clock-in request has been declined, please check AKT HR online system for further details.\n\n' \
                            f'ขอโทษค่ะ คำร้องขออนุมัติเข้างานสายของท่านไม่ได้รับการอนุมัติ โปรดดูรายละเอียดเพิ่มเติมได้ที่ระบบHRออนไลน์ของAKTค่ะ\n\n' \
                            f'申し訳ございませんが、あなたの遅刻出勤許可の申請は拒否されました。詳細はAKT人事オンラインシステムにてご確認ください。\n\n' \
                            f'Sent from AKT HR online system.'

        self.email_to = [f'{email_add}', ]
        self.email_title = email_title
        self.email_content = email_content

        th=threading.Thread(target=self.launching)
        th.start()
        th.join()

    def launching(self):
        self.yag_server.send(to=self.email_to, subject=self.email_title, contents=self.email_content)
        #self.yag_server.close()
        print('Launched!!!!!')
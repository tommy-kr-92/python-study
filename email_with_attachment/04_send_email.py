import os

from libs.email_sender_with_attachment import EmailSender


if __name__ == '__main__':
    es = EmailSender('jyhuh8775@naver.com',
                 os.getenv('MY_NAVER_PASSWORD'),
                 manager_name='김미령',
                 template_filename='templates/email_template_1.html')
    es.send_all_emails('email_list.xlsx')

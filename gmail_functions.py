'''
password = 'yapzqipmdivkmyvb'
'''

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import streamlit as st


@st.cache
def send_email(send_from, gmail_app_pw, send_to, content):
    """
    Sends email based on sender's email, gmail app password, receiver's email, and the email's content
    :param send_from: sender_email
    :param gmail_app_pw: gmail_app_password
    :param send_to: receiver_email
    :param content: content.
    :return:
    """
    with smtplib.SMTP(host='smtp.gmail.com', port=587) as smtp:
        try:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(send_from, gmail_app_pw)
            status = smtp.sendmail(send_from, send_to, content.as_string())
            if status == {}:
                print("寄送郵件完成!")
            else:
                print("寄送郵件失敗")
        except Exception as e:
            print("Error message: ", e)


@st.cache
def send_test_email2(some_subjectline, some_html, send_from, send_to, gmail_app_password, receiver_discount):
    """
    Send test email to sender itself

    :param some_subjectline:
    :param some_html:
    :param send_from: sender
    :param send_to: receiver (sender)
    :param gmail_app_password: pw
    :return:
    """
    print("send_test_email2")
    content = MIMEMultipart()
    content["From"] = send_from
    receiver_name = 'LA Lakers'
    content["To"] = receiver_name
    content["Cc"] = send_from
    # receiver_discount = '50%'
    content["Subject"] = some_subjectline.format(receiver_discount)
    txt_html = some_html.format(receiver_name, receiver_discount)

    content.attach(MIMEText(txt_html, 'html', 'utf-8'))

    send_email(send_from=send_from,
               gmail_app_pw=gmail_app_password,
               send_to=send_to,
               content=content)


def main():
    sender_email = input("請輸入寄件者Gmail: ")
    password = input("請輸入Gmail 應用程式16位數密碼: ")
    password = "yapzqipmdivkmyvb"

    print("End of main() ")


if __name__ == "__main__":
    main()

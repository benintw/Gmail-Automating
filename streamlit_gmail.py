
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import docx2txt
from stqdm import stqdm
from gmail_functions import send_email, send_test_email2


def main():
    st.title("Materials - Gmail 發送自動化")
    login_pw = st.sidebar.text_input("輸入密碼", type='password')
    if login_pw == st.secrets["login_pw"]:
        st.markdown("-" * 30)
        st.header("1. 上傳 Excel 清單")
        
        st.warning("The first three column names must be in the following order AND be in EXACT words")
        st.markdown("- Name\n- Email\n - Discount")
        st.warning("Remove the columns after [Discount] ")
        if st.checkbox("The column names are in the required form"):
            
            uploaded_excel = st.file_uploader("", type=[".xlsx"], accept_multiple_files=False)
            if uploaded_excel is not None:

                df = pd.read_excel(uploaded_excel)
                df_raw = df.copy()
                df_raw.columns = ["Name", "Email", "Discount"]

                null_values = df_raw.isnull().sum().sum()
                repeated_names = df_raw.duplicated(subset=['Name']).sum()
                repeated_emails = df_raw.duplicated(subset=['Email']).sum()
                final_check = null_values + repeated_names + repeated_emails

                if final_check != 0:
                    if null_values != 0:
                        print("有缺失值, 請重新檢查 Excel 清單")
                        st.error("有缺失值, 請重新檢查 Excel 清單！")

                    if repeated_names != 0:
                        print("名字有重複, 請重新檢查 Excel 清單")
                        st.error("名字有重複, 請重新檢查 Excel 清單！")

                    if repeated_emails != 0:
                        print("Email 有重複, 請重新檢查 Excel 清單")
                        st.error("Email 有重複, 請重新檢查 Excel 清單！")

                    st.markdown("修改後請重新上傳 Excel 清單")
                else:
                    st.success("Excel 無缺失值")
                    st.success("Excel 無重複值")
                    # st.markdown("-" * 30)
                    print("Excel 無缺失值也無重複值, 可進行下一步")
                    st.subheader("Excel 內容:")

                    ### 清冊長度
                    length_of_receiver_list = len(df_raw["Name"])
                    print(f"總寄送數量: {length_of_receiver_list} 人")

                    ### 轉換50% 從 0.5 至 '50%'
                    df_raw["Discount"] = df_raw["Discount"].apply(str)
                    df_raw["Discount"] = df_raw["Discount"].replace(["0.5"], "50%")

                    st.dataframe(df_raw)
                    st.write(f"總寄送數量: {length_of_receiver_list} 人")

                    ### User input
                    st.markdown("-" * 30)
                    st.header("2. Gmail 登入")
                    sender_email = st.text_input("請輸入寄件者Gmail: ", value="@gmail.com")
                    cc = []
                    cc_email = st.text_input("請輸入副本收件人 email: ")
                    cc.append(cc_email)
                    password = st.text_input("請輸入Gmail 應用程式16位數密碼: ", type='password', help='不是Gmail密碼')
                    correct_password = st.secrets["pw"]

                    if not sender_email.endswith("@gmail.com") or password != correct_password:
                        if not sender_email.endswith("@gmail.com"):
                            st.error("這個不是 Gmail.")
                        if password != correct_password:
                            st.error("請輸入正確密碼")
                    else:
                        st.success("###### 密碼正確")

                        # test variables
                        receiver_name = "firstName_lastName"
                        receiver_discount = "test%"

                        st.markdown("-" * 30)
                        st.header("3. 檢視 Word Template")

                        doc_file = st.file_uploader("Upload Word Template",
                                                    type=['docx'])
                        if doc_file is not None:
                            raw_text = docx2txt.process(doc_file)
                            st.subheader("3a) 檢視 Word 模板內容:")
                            st.write(raw_text)
                            st.subheader("3b) 修改並確認以下內容：")

                            st.markdown("修改會用到以下 html 語法:\n\n")
                            code0 = """
                                <p>段落</p>\n<b>粗體</b>\n<i>斜體</i>\n<a href="超連結" target='_blank'>對超連結的描述</a>
                            """
                            st.code(code0)
                            subject_line = st.text_area("把 Subject 複製貼上到這: ", value='edit the subject line here')
                            st.markdown("Subject Line Preview: ")
                            components.html(subject_line, height=60, scrolling=True)
                            if "{" not in subject_line:
                                st.error("記得把變數改成\t{ }")

                            info = """
                            變數用\t{ }\t代替,\t例如：\n
                            第一個\t{ }\t用來代替收件人名稱,\n
                            第二個\t{ }\t代替收件人折扣
                            """
                            st.info(info)
                            txt_html = st.text_area("內文內容：",
                                                    value=raw_text,
                                                    height=700)

                            remove_subject = st.checkbox("確認移除內文中的 Subject Line")
                            if st.checkbox("內文預覽(html) 與 Subject") and remove_subject:
                                st.header("4. 預覽郵件內文: ")
                                st.markdown("##### Subject Line: ")
                                components.html(subject_line.format(receiver_discount), height=60, scrolling=True)
                                st.markdown("##### email 內文: ")
                                components.html(txt_html.format(receiver_name, receiver_discount), height=600,
                                                scrolling=True)

                            if st.button(f"寄 '測試郵件' 至{sender_email}"):
                                with st.spinner(f"Sending test email to {sender_email}"):
                                    send_test_email2(some_subjectline=subject_line.format(receiver_discount),
                                                    some_html=txt_html,
                                                    send_from=sender_email,
                                                    send_to=sender_email,
                                                    gmail_app_password=password,
                                                    receiver_discount=receiver_discount)

                                st.success(f"測試郵件已寄送至 {sender_email} !")

                            st.markdown("-" * 30)
                            st.header("5. 準備傳送郵件")
                            if st.checkbox("我已確認過 '測試郵件'內容正常"):
                                st.info("請再次確認 word 模板內容並在下面打勾(兩個都要)")
                                proofread = st.checkbox("已確認模板內容沒問題")
                                proceed = st.checkbox("準備前往傳送郵件")
                                if proofread and proceed:
                                    st.success("可以開始傳送")
                                    if st.button("傳送郵件至清單內聯絡人"):

                                        for idx in stqdm(range(length_of_receiver_list), desc='郵件傳送中: '):
                                            receiver_name = df_raw.iloc[idx, 0]
                                            receiver_email = df_raw.iloc[idx, 1]
                                            receiver_discount = df_raw.iloc[idx, 2]
                                            content = MIMEMultipart()
                                            content["To"] = receiver_name
                                            content["Cc"] = cc
                                            content["Subject"] = subject_line.format(receiver_discount)

                                            content.attach(MIMEText(txt_html.format(receiver_name, receiver_discount),
                                                                    'html', 'utf-8'))
                                            toAddress = [receiver_email] + cc
                                            send_email(send_from=sender_email,
                                                    gmail_app_pw=password,
                                                    send_to=toAddress,
                                                    content=content)

                                        st.write("寄送郵件完成!")
                                        print(f"已傳送給 {idx + 1} 位收件者")
                                        st.write(f"已傳送給 {idx + 1} 位收件者")

                        st.markdown("-" * 100)
                        st.markdown("by BC")
    else:
        st.sidebar.write("請輸入正確密碼")


if __name__ == "__main__":
    main()

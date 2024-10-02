import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from docx import Document
from stqdm import stqdm
import smtplib

# Utility functions
# added a comment Jun 27 2024


def send_email(send_from, gmail_app_pw, send_to, content):
    """Sends an email using SMTP."""
    with smtplib.SMTP(host="smtp.gmail.com", port=587) as smtp:
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


def send_test_to_ben(
    some_subjectline,
    some_html,
    send_from,
    send_to,
    gmail_app_password,
    receiver_discount,
):
    """Sends a test email to Ben."""
    print("send_test_to_ben")
    content = MIMEMultipart()
    content["From"] = send_from
    receiver_name = "Test Email"
    content["To"] = receiver_name
    content["Cc"] = send_from
    content["Subject"] = some_subjectline.format(receiver_discount)
    txt_html = some_html.format(receiver_name, receiver_discount)

    content.attach(MIMEText(txt_html, "html", "utf-8"))

    send_email(
        send_from=send_from,
        gmail_app_pw=gmail_app_password,
        send_to=send_to,
        content=content,
    )


def read_docx(file):
    """Reads a .docx file and returns the text content."""
    doc = Document(file)
    return "\n".join(para.text for para in doc.paragraphs)


def replace_dynamic_fields(template, data):
    for key, value in data.items():
        template = template.replace(f"{{{{ {key} }}}}", str(value))
    return template


def main():
    st.title("Materials - Gmail 發送自動化")
    login_pw = st.sidebar.text_input("輸入密碼", type="password")

    st.markdown("-" * 30)
    st.header("1. 上傳 Excel 清單")

    if st.checkbox("Ready????"):
        uploaded_excel = st.file_uploader(
            "", type=[".xlsx"], accept_multiple_files=False
        )

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            df.columns = df.columns.str.strip().str.lower()
            df.columns = df.columns.str.replace("email", "Email", case=False)
            df.columns = df.columns.str.replace("discount", "Discount", case=False)
            df.columns = df.columns.str.replace("name", "Name", case=False)

            revised_df = df[["Name", "Email", "Discount"]].copy()

            null_values = revised_df.isnull().sum().sum()
            repeated_emails = revised_df.duplicated(subset=["Email"]).sum()
            final_check = null_values + repeated_emails

            if final_check != 0:
                if null_values != 0:
                    st.error("有缺失值, 請重新檢查 Excel 清單！")

                if repeated_emails != 0:
                    st.error("Email 有重複, 請重新檢查 Excel 清單！")

                st.markdown("修改後請重新上傳 Excel 清單")
            else:
                st.success("Excel 無缺失值")
                st.success("Excel 無重複值")
                st.success("Excel檔案 符合條件")

                num_receivers = len(revised_df["Name"])
                revised_df["Discount"] = (
                    revised_df["Discount"]
                    .apply(str)
                    .replace({"0.5": "50%", "1.0": "100%", "0.2": "20%"})
                )

                st.dataframe(revised_df)
                st.write(f"總寄送數量: {num_receivers} 人")
                
                st.markdown("-" * 30)
                values = st.slider("要到和包含第幾row?", min_value=0, max_value=num_receivers, value=(0, num_receivers),step=1)
                min_value = int(values[0])
                max_value = int(values[-1])
                special_revised_df = revised_df.iloc[min_value:max_value + 1, :]
                st.dataframe(special_revised_df)
                special_num_receivers = len(special_revised_df["Name"])
                
                st.write(f"special總寄送數量: {special_num_receivers} 人")

                st.markdown("-" * 30)
                st.header("2. Gmail 登入")
                sender_email = st.text_input("請輸入寄件者Gmail: ", value="@gmail.com")
                cc_email = st.text_input("請輸入副本收件人 email: ")
                password = "hhadvckutoyuidyh"
                # correct_password = st.secret["pw"]


                if True:
                    st.success("###### 密碼正確")

                    st.markdown("-" * 30)
                    st.header("3. 檢視 Word Template")

                    doc_file = st.file_uploader("Upload Word Template", type=["docx"])
                    if doc_file is not None:
                        raw_text = read_docx(doc_file)
                        st.subheader("3a) 檢視 Word 模板內容:")
                        st.write(raw_text)
                        st.subheader("3b) 修改並確認以下內容：")

                        st.markdown("修改會用到以下 html 語法:\n\n")
                        st.code(
                            "<p>段落</p>\n<b>粗體</b>\n<i>斜體</i>\n<a href='超連結' target='_blank'>對超連結的描述</a>\n<br> 換行"
                        )
                        subject_line = st.text_area(
                            "把 Subject 複製貼上到這: ",
                            value="edit the subject line here",
                        )
                        # st.markdown("Subject Line Preview: ")
                        # st.write(subject_line)
                        if "{" not in subject_line:
                            st.error("記得把變數改成\t{ }")

                        st.info(
                            "變數用\t{ }\t代替,\t例如：\n第一個\t{ }\t用來代替收件人名稱,\n第二個\t{ }\t代替收件人折扣"
                        )
                        txt_html = st.text_area(
                            "內文內容：", value=raw_text, height=700
                        )

                        if st.checkbox("內文預覽(html) 與 Subject"):
                            ### For checking purposes
                            idx = 0
                            data = {
                                "Name": special_revised_df.iloc[idx]["Name"],
                                "Discount": special_revised_df.iloc[idx][
                                    "Discount"
                                ],  # eg. "50%"
                            }
                            st.header("4. 預覽郵件內文: ")
                            st.markdown("##### Subject Line: ")
                            components.html(
                                subject_line.format(f"{data['Discount']}"),
                                height=60,
                                scrolling=True,
                            )
                            st.markdown("##### Email 內文: ")
                            components.html(
                                txt_html.format(
                                    f"{data['Name']}", f"{data['Discount']}"
                                ),
                                height=600,
                                scrolling=True,
                            )

                        sent_to_ben = "chen.ben968@gmail.com"
                        if st.button(f"寄 '測試郵件' 至{sent_to_ben}"):
                            with st.spinner(f"Sending test email to {sent_to_ben}"):
                                send_test_to_ben(
                                    some_subjectline=subject_line,
                                    some_html=txt_html,
                                    send_from=sender_email,
                                    send_to=sent_to_ben,
                                    gmail_app_password=password,
                                    receiver_discount="100%",
                                )
                            st.success(f"測試郵件已寄送至 {sent_to_ben} !")

                        st.markdown("-" * 30)
                        st.header("5. 準備傳送郵件")
                        if st.checkbox(label="我已確認過 '測試郵件'內容正常"):
                            st.info("請再次確認 word 模板內容並在下面打勾(兩個都要)")
                            proofread = st.checkbox(
                                label="已確認模板內容沒問題",
                            )
                            proceed = st.checkbox(
                                label="準備前往傳送郵件",
                            )
                            if proofread and proceed:
                                st.success("可以開始傳送")
                                

                                
                                if st.button(
                                    f"傳送郵件至清單內 {special_num_receivers} 位聯絡人"
                                ):
                                    for idx in stqdm(
                                        range(special_num_receivers),
                                        desc="郵件傳送中: ",
                                    ):
                                        stop_button = st.button("即時終止")
                                        if stop_button:
                                            break
                                        
                                        receiver_name = special_revised_df.iloc[idx]["Name"]
                                        receiver_email = special_revised_df.iloc[idx]["Email"]
                                        receiver_discount = special_revised_df.iloc[idx]["Discount"]
                                        data = {
                                            "name": receiver_name,
                                            "discount": receiver_discount,
                                        }

                                        personalized_subject = replace_dynamic_fields(
                                            subject_line, data
                                        )
                                        personalized_body = replace_dynamic_fields(
                                            txt_html, data
                                        )

                                        content = MIMEMultipart()
                                        content["To"] = receiver_name
                                        content["Cc"] = cc_email

                                        # content["Subject"] = subject_line.format(
                                        #     receiver_discount
                                        # )
                                        content["Subject"] = personalized_subject
                                        content.attach(
                                            MIMEText(personalized_body, "html", "utf-8")
                                        )

                                        send_email(
                                            send_from=sender_email,
                                            gmail_app_pw=password,
                                            send_to=[receiver_email] + [cc_email],
                                            content=content,
                                        )
                                        st.write(
                                            f"已寄給 Name: {receiver_name} | Email: {receiver_email} | Discount: {receiver_discount}"
                                        )
                                    print(f"已傳送給 {idx + 1} 位收件者")
                                    st.write(f"已傳送給 {idx + 1} 位收件者")

                        st.markdown("-" * 100)
                        st.markdown("by BC")


if __name__ == "__main__":
    main()

# Gmail-Automating with Python


![擷取](https://user-images.githubusercontent.com/104064349/218375067-0dbe12ec-1ecf-4452-89a1-059045e2169e.PNG)

---
## Description
This streamlit app automates the process of sending invitation emails to a list of people, with the purpose of inviting them to contribute to a journal. The app takes a formatted excel file and an invitation email template as inputs. After checking for any duplicates or missing information, the user can edit the content of the email and send a test email to themselves before proceeding. The final step is to press the "send to inviting authors" button to send the email to each recipient.

---
## Why did i write this app?
During my time as a research assistant for this civil engineering professor at NTUT, he had to send out 
paper invitations to other researchers/professors to contribute to the journal. The number of contactee usually exceeds 100. It used to take his students two or three afternoons to finish this task. 

The only differences between these emails are "receiver_name" and "receiver_discount" in the subject line and in the body of the email. This is a sign for automating...

Once I was aware of this task, i immediately took the initiative to write a email automating app to speed up this highly repetitive process. 

---
## How to use it?
_The input_:
- A formatted excel (.xslx)
- A invitation email template (.docx)

Here are the step-by-step instructions for using this app:

1. Enter the password required to access the app (specifically for the professor you are working with).
2. Check that the columns of the uploaded excel file match the required format.
3. Upload a formatted excel file containing the list of contacts.
4. The app will check for any duplicated or missing names and email addresses. If everything is in order, the list will be displayed in a streamlit dataframe format.
5. Enter the sender's email, cc's email, and the 16-digit Gmail app password.
6. Upload the invitation template and customize the content as desired using simple HTML tags.
7. Optionally, send a test email to yourself to ensure that everything is correct.
8. Review everything again and check two checkboxes before sending mass emails.
9. Press the "send to inviting authors" button to complete the process.

_The output_
- Complete gmails sent to each inviting author.


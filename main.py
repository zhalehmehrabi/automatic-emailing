import win32com.client


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0  # size of the new email
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Testing Mail'
    newmail.To = 'mahshid.shiri@mail.polimi.it'
    # newmail.CC = 'xyz@example.com'
    newmail.Body = 'Hello, this is a test email.'

    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)

    # To display the mail before sending it
    newmail.Display()

    newmail.Send()


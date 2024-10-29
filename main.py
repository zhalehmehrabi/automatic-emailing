import win32com.client
import pandas as pd
import os

def read_excel():
    df = pd.read_excel("list.xlsx")
    return df

def get_files(folder_path):
  files = []
  for entry in os.listdir(folder_path):
    if os.path.isfile(os.path.join(folder_path, entry)):
        absolute_path = os.path.abspath(folder_path+"/"+entry)
        absolute_path.replace("/", "\\")
        files.append(absolute_path)
  return files

def send_emails():
    ol = win32com.client.Dispatch("outlook.application")
    df = read_excel()
    olmailitem = 0x0  # size of the new email
    for row in df.itertuples():
        if row.send == 'yes':
            newmail = ol.CreateItem(olmailitem)
            newmail.Subject = row.subject
            newmail.To = row.email
            # newmail.CC = 'xyz@example.com'
            newmail.Body = row.body
            # files = get_files("Attachments/" + row.email)
            # for attach in files:
            #     newmail.Attachments.Add(attach)
            CV_file = get_files("Attachments/CV")
            for attach in CV_file:
                newmail.Attachments.Add(attach)
            # newmail.Display()

            newmail.Send()
            print(f"Email to {row.email} is sent")

def change_email_text():
    new_text = "I hope you are doing well. My name is Amirhossein, and I am currently a post-graduate research assistant, supported by a research scholarship. Before this, I completed my Master's at Politecnico di Milano, graduating with a final score of 107/110.\n\nI am writing to express my interest in potential PhD opportunities within your research group. I understand that admission to EPFL is determined by a committee, but I would greatly appreciate knowing if there might be openings in your lab, and whether I could potentially list you as a prospective supervisor. My research interests lie in applying reinforcement learning to robotics, and I am very enthusiastic about the possibility of contributing to work in these fields under your guidance.\n\nI have attached my CV for your consideration, and I would be happy to provide any additional information you may need.\n\nThank you for your time, and I look forward to the possibility of discussing this further.\n\nBest regards,\nMahshid"
    df = read_excel()
    for index, row in df.iterrows():
        if index > 2:
            break
        text = row['body']

        # Check if the text starts with "Dear Professor"
        if text.lower().startswith("dear professor"):
            # Find the greeting and separate it from the rest of the text
            end_of_greeting = text.find('\n')
            greeting = text[:end_of_greeting] if end_of_greeting != -1 else text
            updated_text = f"{greeting}\n\n{new_text}"

            # Update the 'body' column in the DataFrame
            df.at[index, 'body'] = updated_text

    # Save the modified DataFrame back to Excel
    df.to_excel('emails_updated.xlsx', index=False)  # replace with your preferred output path


if __name__ == '__main__':
    send_emails()
    # change_email_text()
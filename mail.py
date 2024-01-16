import win32com.client as client
import os
import pandas as pd
my_ol_app = client.Dispatch('Outlook.Application')  #.GetNameSpace('MAPI')
# this part reads mail messages from the inbox mr.Ahmed
ol = my_ol_app.GetNameSpace('MAPI')
my_inbox = ol.GetDefaultFolder(6).Folders('Automated')
my_ib_items = my_inbox.Items
attach_path = r'D:\2023\jan'
for Msg in my_ib_items:
    if 'Hakiem_WZ_Daily' in str(Msg.Subject) and '2023-01' in str(Msg.ReceivedTime):          #'3G Top Worst'
        print(Msg.SenderName)
        print(Msg.Subject)
        print(str(Msg.ReceivedTime))
        for attach in Msg.Attachments:
            file_name = str(attach.FileName)
            #attach.SaveAsFile(os.path.join(attach_path, file_name))
#this part intends to send email messages mr.Ahmed
mail_item = my_ol_app.CreateItem(0)
mail_item.Subject = 'Python Test'
mail_item.BodyFormat = 1
mail_item.Body = 'Hello Mr.Ahmed, This bot runs to test python script\n' \
                 'Best Regards,\n' \
                 'Ahmed ELsayed\n' \
                 'NPO Engineer\n' \
                 'Mobile: +201020667812\n' \
                 'NOKIA'
xl_path = r'D:\2023\jan\10_1\Hakiem_V4__RNC50_daily-11847-2023_01_10-08_00_00__94.xlsx'
df = pd.read_excel(xl_path,sheet_name='MAIL',usecols=[0,1,2,3,4,5,6,7])
df = df[df['Site Name'].str.len() > 0]
tabl_html = df.to_html(table_id='table1')
tabl_html = tabl_html.replace('class="dataframe" id="table1"','class="styled-table" id="table1"')
print(tabl_html)

with open(r'C:\Users\user\Desktop\table style.txt','r') as file:
    table_style = file.read()
template = open(r'C:\Users\user\Desktop\template.txt', 'r').read()
signature = open(r'C:\Users\user\Desktop\sig.txt','r').read()
template_format = template.format('Support team')
final_html_file = table_style+template_format+tabl_html+signature
with open(r'C:\Users\user\Desktop\final.html','w') as final:
    final.write(final_html_file)
print(final)
mail_item.HTMLBody = final_html_file
mail_item.To = 'asayed9883@gmail.com; ahmedelsayed98@hotmail.com'
mail_item.CC = 'asayed9883@gmail.com'
mail_item.Attachments.Add(r'D:\2023\jan\2_1\scr_adjust.xlsx')
mail_item.Display()
#mail_item.Send()
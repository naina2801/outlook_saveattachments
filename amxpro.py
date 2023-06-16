# from win32com.client import Dispatch
# import os
# import glob
# path = r'C:\Users\Unaina\PycharmProjects\pandas1\assesment\outlook_attachments'
# ol =  Dispatch("Outlook.Application").GetNamespace("MAPI")
# inbox = ol.GetDefaultFolder(6)
# messages = inbox.Items
# for message in messages:
# 	for attachment in message.Attachments:
# 		attachment.SaveAsFile(os.path.join(path, str(attachment)))
# all_files = os.listdir(path)
# csv_files = glob.glob(os.path.join(path,"*.csv"))
# for cs in csv_files:
#     print(cs)

from win32com.client import Dispatch
import os
import glob
import psycopg2
path = r'C:\Users\Unaina\PycharmProjects\pandas1\assesment\outlook_attachments'
ol =  Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = ol.GetDefaultFolder(6)
messages = inbox.Items



class Amx_pro():
	def f_save(self):
		for message in messages:
			for attachment in message.Attachments:
				 attachment.SaveAsFile(os.path.join(path, str(attachment)))
		all_files = os.listdir(path)

		csv_files = glob.glob(os.path.join(path, "*.csv"))
		print(csv_files)


sa =Amx_pro()
sa.f_save()



conn = psycopg2.connect(host="localhost",
                        database="postgres",
                        user="postgres",
                        password="280120",
                        port=5432)

cursor = conn.cursor()
cursor.execute("SELECT * FROM tbl_warehouses")
rows = cursor.fetchall()
for r in rows:
    print(r)
conn.commit()
cursor.close()
conn.close()
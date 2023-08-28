import os
import shutil
from smb.SMBConnection import SMBConnection

current_path = os.getcwd()
parent_path = os.path.abspath(os.path.join(current_path, os.pardir))
#\\ITPC051911\mail_group
server_name = "ITPC059100"
server_ip = "ITPC059100"
share_name = "mail_group"
user_name = "Mail_Group_Updater"
password = "Mail_Group_Updater"
conn = SMBConnection(user_name, password, "localhost", server_name, use_ntlm_v2=True)
conn.connect(server_ip)
client = conn

# Copia il contenuto della cartella di origine in un file temporaneo
source_folder = "update"
destination_folder = parent_path

file_list = conn.listPath(share_name, source_folder)
for file in file_list:
    if not file.isDirectory:
        source_file_path = f"{source_folder}/{file.filename}"
        destination_file_path = f"{parent_path}/{file.filename}"
        with open(destination_file_path, 'wb') as f:
            conn.retrieveFile(share_name, source_file_path, f)

# Chiudi la connessione
conn.close()

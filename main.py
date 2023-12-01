from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd

def authenticate_drive(client_secrets_path):
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()

    drive = GoogleDrive(gauth)
    return drive

def get_file_ids_in_folder(drive, folder_name):
    # Retrieve the file IDs in the specified folder
    folder_query = f"title = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
    folder_list = drive.ListFile({'q': folder_query}).GetList()

    if folder_list:
        folder_id = folder_list[0]['id']
        file_query = f"'{folder_id}' in parents"
        file_list = drive.ListFile({'q': file_query}).GetList()
        return [{'FolderName': folder_name, 'FileID': file['id']} for file in file_list]
    else:
        print(f"Folder '{folder_name}' not found.")
        return []

def enable_sharing(drive, file_id):
    file = drive.CreateFile({'id': file_id})
    file.InsertPermission({
        'type': 'anyone',
        'role': 'reader',
    })

def create_excel_log(log_data, excel_filename='log_file.xlsx'):
    # Create a DataFrame from the log data
    df = pd.DataFrame(log_data)

    # Export the DataFrame to an Excel file
    df.to_excel(excel_filename, index=False)
    print(f"Log file '{excel_filename}' created successfully.")

if __name__ == "__main__":
    client_secrets_path = 'E:/Google Drive/Test'
    drive = authenticate_drive(client_secrets_path)

    # Replace the list below with the names of your folders
    folder_names = ['Fol1', 'Fol2','Fol3','Fol4']

    log_data = []

    for folder_name in folder_names:
        file_ids = get_file_ids_in_folder(drive, folder_name)

        if file_ids:
            for data in file_ids:
                enable_sharing(drive, data['FileID'])
                log_data.append(data)

    create_excel_log(log_data)

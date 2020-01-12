#!/usr/bin/python
from __future__ import print_function
import pprint
import six
import httplib2
from googleapiclient.discovery import build
import googleapiclient.http
import oauth2client.client

# OAuth 2.0 scope that will be authorized.
# Check https://developers.google.com/drive/scopes for all available scopes.
OAUTH2_SCOPE = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.metadata']

# Location of the client secrets.
CLIENT_SECRETS = 'client_secrets.json'
XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


# create a g drive service
def create_service():
    # Perform OAuth2.0 authorization flow.
    flow = oauth2client.client.flow_from_clientsecrets(CLIENT_SECRETS, OAUTH2_SCOPE)
    flow.redirect_uri = oauth2client.client.OOB_CALLBACK_URN
    authorize_url = flow.step1_get_authorize_url()
    print('Go to the following link in your browser: ' + authorize_url)
    # `six` library supports Python2 and Python3 without redefining builtin input()
    code = six.moves.input('Enter verification code: ').strip()
    credentials = flow.step2_exchange(code)

    # Create an authorized Drive API client.
    http = httplib2.Http()
    credentials.authorize(http)
    drive_service = build('drive', 'v2', http=http)

    return drive_service


# get all files with that mimetype
def list_files_mimetype(drive_service, mimetype):
    # get all the files with the specified mimetype
    response = drive_service.files().list(q="mimeType='{}'".format(mimetype),
                                          spaces='drive').execute()
    # pprint.pprint(response)
    return response


# download a file given the id
def download_file_by_id(id_file, drive_service):
    result = drive_service.files().get(fileId=id_file).execute()
    download_url = result['downloadUrl']
    file_name = "{}_test.xlsx".format(result['originalFilename'])
    resp, content = drive_service._http.request(download_url)
    if resp.status == 200:
        print('Status: %s' % resp)
        fo = open(file_name, "wb")
        fo.write(content)
        fo.close()
    else:
        print('An error occurred: %s' % resp)

    return result, file_name



# upload and convert file
def upload_convert_xlsx_g_sheet(file_to_upload, file_name_on_g_drive, description, drive_service):
    # Insert a file. Files are comprised of contents and metadata.
    # The body contains the metadata for the file.
    file_metadata = {'title': file_name_on_g_drive, 'description': description}
    # MediaFileUpload abstracts uploading file contents from a file on disk.
    media = googleapiclient.http.MediaFileUpload(file_to_upload)
    # IMPORTANT #
    # the  convert=True will make the mimetype like the google ones mimetype='application/vnd.google-apps.spreadsheet')
    new_file = drive_service.files().insert(body=file_metadata, convert=True, media_body=media, fields='id').execute()
    # Perform the request and print the result. TO CHECK
    # pprint.pprint(new_file)


def main():
    already_read_name_list = []

    # create a g drive service
    drive_service = create_service()
    # list all files of a specific mimetype
    mimetype_files = list_files_mimetype(drive_service=drive_service, mimetype=XLSX_MIMETYPE)

    for item in mimetype_files['items']:

        curr_filename = item['originalFilename']
        if curr_filename not in already_read_name_list:
            curr_id = item['id']
            id_file = curr_id
            # downlaod the file
            file_metadata, file_to_upload = download_file_by_id(id_file=id_file, drive_service=drive_service)

            file_name_on_g_drive = file_to_upload.split(".xlsx")[0]
            description = 'description'
            upload_convert_xlsx_g_sheet(file_to_upload=file_to_upload, file_name_on_g_drive=file_name_on_g_drive,
                                        description=description, drive_service=drive_service)


main()

# xlsx_conversion_gWorksheet_driveApi
Many libraries are defined for the excel manipulation on google drive, but not for the Microsoft office worksheet uploaded on G Drive

This Python Script would like to help people who need to automatically convert their microsoft sheet into GDrive Sheet, in order to process them easily.

The script is divided in four main function:

1)create_service()
It is needed to create the google drive service object. 
From that Object is possible to call all the usefull functions.

2)list_files_mimetype(drive_service, mimetype)
It lists all the files present in "your" google drive for the specified mimetype.
Here you can find the list of mimetype https://developers.google.com/drive/api/v3/mime-types

3)download_file_by_id(id_file, drive_service)
It downloads the file specified by the id_file.
In this case we impose the xlsx file type, but it is not necessary.

4)upload_convert_xlsx_g_sheet(file_to_upload, file_name_on_g_drive, description, drive_service)
Here the file is update on "your" drive.
During the insert function is very important the param "convert=True" that can permit the conversion from the microsoft office mimetype to the google drive worksheet mimetype

CAVEAT : you should have your creedential in a file "google_secret.json", in this case, of your google application.
Here you can find more info about the google api credentials https://developers.google.com/adwords/api/docs/guides/authentication.

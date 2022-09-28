from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
siteurl = sharepointpath
subsite = sharepoint_subdir1

from_path = desktoppath
to_path = sharepoint_subdir2

to_name = filename
usr = usr
passw = pasw

ctx = ClientContext(siteurl+subsite).with_credentials(UserCredential(usr, passw))
print('Connection done')

with open(from_path, 'rb') as content_file:
    file_content = content_file.read()
print('File found')
file = ctx.web.get_folder_by_server_relative_url(subsite+to_path).upload_file(to_name, file_content).execute_query()
print('Upload done')

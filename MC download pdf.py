url = 'https://btgplc.mastercontrol.com/btgplc/restapi/v1/document/NEP4PQY5YZBTLHZ5BL/publishedMainFile'
mcurl ='https://btgplc.mastercontrol.com/btgplc/restapi/identity/authentication/login?action=login'
username = 'cvieites'
password = 'Celtic33'
outputfile = 'carlos5.pdf'
from pathlib import Path
import msvcrt
import win32timezone
import requests
from requests import get
from requests_ntlm import HttpNtlmAuth

s = requests.Session()
login_data =  {'username':username, 'password':password}
s.post(mcurl, login_data)
#logged in! cookies saved for future requests
response = s.get(url)
filename = Path(outputfile)
filename.write_bytes(response.content)

# Note: Use Google Colab to run this program and split the code into blocks. This is not mandatory. This is how I use it. So...
# Import Libraries
from oauth2client.client import OAuth2WebServerFlow
from googleapiclient.discovery import build
import httplib2
from google.oauth2 import service_account
import requests
from googleapiclient.errors import HttpError

# Authentication Code
CLIENT_ID = "{update this}"
CLIENT_SECRET = "{update this}"

OAUTH_SCOPE = "https://www.googleapis.com/auth/webmasters.readonly"
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'

flow = OAuth2WebServerFlow(CLIENT_ID, CLIENT_SECRET, OAUTH_SCOPE, REDIRECT_URI)

authorize_url = flow.step1_get_authorize_url()
print("Go to the following link in your browser:", authorize_url)

auth_code = input("Enter Your Authorization Code: ")

credentials = flow.step2_exchange(auth_code)
print("The Credentials are Generated Successfully")

http = httplib2.Http()

creds = credentials.authorize(http)

webmasters_service = build('searchconsole', 'v1', http=creds)


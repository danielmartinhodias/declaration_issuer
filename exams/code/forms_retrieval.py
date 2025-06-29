from google.oauth2 import service_account
from googleapiclient.discovery import build


def get_answers(form_number):
  print ("entered get-answers")
  SCOPES = ["https://www.googleapis.com/auth/forms.responses.readonly"]
  SERVICE_ACCOUNT_FILE = '/Users/danielmdias/docs/_fun_time/00_madrid_prescriber/exams/code/data/required_files/forms_acess_key.json'

  # Authenticate using service account
  creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

  # Build the service object
  service = build('forms', 'v1', credentials=creds)

  # Prints the responses of your specified form:
  form_id = form_number
  result = service.forms().responses().list(formId=form_id).execute()
  print ("results", result)
  return(result)

from __future__ import print_function
import time
import os
import base64
from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools
from email.mime.text import MIMEText

# Fill-in IDs of your Docs template & any Sheets data source
DOCS_FILE_ID = '12Rfj6MCJACYwb64vH0eI79okH-ajSkgoMTbcIJr8sjo'
SHEETS_FILE_ID = '1cjnH5p2ihx86ZOMMZjJoUD5Xukv_DuFFbgopStNxnzI'

# authorization constants
CLIENT_ID_FILE = 'credentials.json'
TOKEN_STORE_FILE = 'token.json'
SCOPES = (  # iterable or space-delimited string
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://mail.google.com/',
)
DISCOVERY_DOC = 'https://docs.googleapis.com/$discovery/rest?version=v1'

# application constants
SOURCES = ('text', 'sheets')
SOURCE = 'sheets'  # Choose one of the data SOURCES
COLUMNS = ['name', 'email']
TEXT_SOURCE_DATA = ('')
COPIED_DOC_ID = ""
USER_EMAIL = ""


def get_http_client():
    """Uses project credentials in CLIENT_ID_FILE along with requested OAuth2
        scopes for authorization, and caches API tokens in TOKEN_STORE_FILE.
    """
    store = file.Storage(TOKEN_STORE_FILE)
    creds = store.get()
    if os.path.exists(CLIENT_ID_FILE):
        print('yes, the client secrets file exists')
    else:
        print('the client secrets file does not exist!')
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_ID_FILE, SCOPES)
        creds = tools.run_flow(flow, store)
    return creds.authorize(Http())


# service endpoints to Google APIs
HTTP = get_http_client()
DRIVE = discovery.build('drive', 'v3', http=HTTP)
DOCS = discovery.build('docs', 'v1', http=HTTP)
SHEETS = discovery.build('sheets', 'v4', http=HTTP)


def get_data(source):
    """Gets mail merge data from chosen data source.
    """
    if source not in {'sheets', 'text'}:
        raise ValueError(
            'ERROR: unsupported source %r; choose from %r' % (source, SOURCES))
    return SAFE_DISPATCH[source]()


def _get_text_data():
    """(private) Returns plain text data; can alter to read from CSV file.
    """
    return TEXT_SOURCE_DATA


def _get_sheets_data(service=SHEETS):
    """(private) Returns data from Google Sheets source. It gets all rows of
        'Sheet1' (the default Sheet in a new spreadsheet), but drops the first
        (header) row. Use any desired data range (in standard A1 notation).
    """
    return service.spreadsheets().values().get(
        spreadsheetId=SHEETS_FILE_ID,
        range='Sheet1').execute().get('values')[1:]  # skip header row


# data source dispatch table [better alternative vs. eval()]
SAFE_DISPATCH = {k: globals().get('_get_%s_data' % k) for k in SOURCES}


def _copy_template(tmpl_id, source, service):
    """(private) Copies letter template document using Drive API then
        returns file ID of (new) copy.
    """
    body = {'name': 'Merged form letter (%s)' % source}
    return service.files().copy(
        body=body, fileId=tmpl_id, fields='id').execute().get('id')


def merge_template(tmpl_id, source, service):
    """Copies template document and merges data into newly-minted copy then
        returns its file ID.
    """
    # copy template and set context data struct for merging template values
    copy_id = _copy_template(tmpl_id, source, service)
    context = merge.iteritems() if hasattr({}, 'iteritems') else merge.items()
    global COPIED_DOC_ID
    COPIED_DOC_ID = copy_id
    print(COPIED_DOC_ID)

    # "search & replace" API requests for mail merge substitutions
    reqs = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{%s}}' % key.upper(),  # {{VARS}} are uppercase
                    'matchCase': True,
                },
                'replaceText': value,
            }
        } for key, value in context
    ]

    # send requests to Docs API to do actual merge
    DOCS.documents().batchUpdate(
        body={
            'requests': reqs
        }, documentId=copy_id, fields='').execute()
    return copy_id


#------------Extract the contents of document-------------#
SCOPE = 'https://www.googleapis.com/auth/documents.readonly'
DISCOVERY_DOC = 'https://docs.googleapis.com/$discovery/rest?version=v1'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth 2.0 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    store = file.Storage('token.json')
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPE)
        credentials = tools.run_flow(flow, store)
    return credentials


def read_paragraph_element(element):
    """Returns the text in the given ParagraphElement.

        Args:
            element: a ParagraphElement from a Google Doc.
    """
    text_run = element.get('textRun')
    if not text_run:
        return ''
    return text_run.get('content')


def read_strucutural_elements(elements):
    """Recurses through a list of Structural Elements to read a document's text where text may be
        in nested elements.

        Args:
            elements: a list of Structural Elements.
    """
    text = ''
    for value in elements:
        if 'paragraph' in value:
            elements = value.get('paragraph').get('elements')
            for elem in elements:
                text += read_paragraph_element(elem)
        elif 'table' in value:
            # The text in table cells are in nested Structural Elements and tables may be
            # nested.
            table = value.get('table')
            for row in table.get('tableRows'):
                cells = row.get('tableCells')
                for cell in cells:
                    text += read_strucutural_elements(cell.get('content'))
        elif 'tableOfContents' in value:
            # The text in the TOC is also in a Structural Element.
            toc = value.get('tableOfContents')
            text += read_strucutural_elements(toc.get('content'))
    return text


def main():
    """Uses the Docs API to print out the text of a document."""
    credentials = get_credentials()
    http = credentials.authorize(Http())
    import pdb
    # pdb.set_trace()
    print(DOCUMENT_ID)

    docs_service = discovery.build(
        'docs', 'v1', http=http, discoveryServiceUrl=DISCOVERY_DOC)

    doc = docs_service.documents().get(documentId=DOCUMENT_ID).execute()
    print(docs_service.documents().get(documentId=DOCUMENT_ID).execute())

    doc_content = doc.get('body').get('content')
    print(read_strucutural_elements(doc_content))
    return read_strucutural_elements(doc_content)


#-------------Extraction completed---------------#


#------------Create Email Message----------------#
def create_message(sender, to, subject, message_text):
    """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
    b64_string = b64_bytes.decode()
    body = {'raw': b64_string}
    return body


#--------------Create Email Message with Attachment-------------#
def create_message_with_attachment(sender, to, subject, message_text, file):
    """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(file, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(file, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(file, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(file, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(file)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_string())}


#-------------Sending the Message--------------#
def send_message(service, user_id, message):
    """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except e:
        print('An error occurred')


if __name__ == '__main__':
    # fill-in your data to merge into document template variables
    merge = {
        # sender data

        # - - - - - - - - - - - - - - - - - - - - - - - - - -
        # recipient data (supplied by 'text' or 'sheets' data source)
        'name': None,
        'email': None,

        # - - - - - - - - - - - - - - - - - - - - - - - - - -

        # - - - - - - - - - - - - - - - - - - - - - - - - - -
    }

    # get row data, then loop through & process each form letter
    data = get_data(SOURCE)  # get data from data source
    for i, row in enumerate(data):
        merge.update(dict(zip(COLUMNS, row)))
        USER_EMAIL = merge["email"]
        print('Merged letter %d: docs.google.com/document/d/%s/edit' %
              (i + 1, merge_template(DOCS_FILE_ID, SOURCE, DRIVE)))
    print(COPIED_DOC_ID)
    DOCUMENT_ID = COPIED_DOC_ID

    doc_content = main()

    new_msg = create_message('paritosh.yadav58@gail.com', USER_EMAIL, "TEST 1",
                             doc_content)
    #----loading the gmail service------#
    store = file.Storage(TOKEN_STORE_FILE)
    creds = store.get()
    if os.path.exists(CLIENT_ID_FILE):
        print('yes, the client secrets file exists')
    else:
        print('the client secrets file does not exist!')
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_ID_FILE, SCOPES)
        creds = tools.run_flow(flow, store)

    service = discovery.build('gmail', 'v1', credentials=creds)

    send_message(service,'me', new_msg)

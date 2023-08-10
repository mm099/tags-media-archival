# Imports for Google Drive API ###############################################
from __future__ import print_function

import os.path
import io

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload

# If modifying scopes, delete the token.json file.
SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/drive.metadata']
##############################################################################

# Imports ###################################################
import time, sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.common.by import By
import pyautogui
import shutil
from pathlib import Path
import subprocess
import re
import math
import pyperclip
from getopt import getopt

DELAYSHORT = 1
DELAY = 5
CWD = str(Path.cwd())
OUTPUTDIR = '/files/'
# TODO: move this to more appropriate location ###############
# todo: make setting up the settings file automatic
settingsFile = open(CWD + '/settings.txt', 'r')
for line in settingsFile:
    line = line.split(' ')
    if line[0] == 'spreadsheetId':
        SPREADSHEET = line[2].replace('\n', '')
    elif line[0] == 'spreadsheetTestId':
        SPREADSHEET_TEST = line[2].replace('\n', '')
    elif line[0] == 'driveId':
        SHARED_DRIVE = line[2].replace('\n', '')
settingsFile.close()
##############################################################
GOFULLPAGE = 'fdpohaocaechififmbbbbbknoalclacl'
# note: sheets API counts rows starting from zero, but I want the user to
# enter the row numbers as they appear in a spreadsheet, which counts from one
# startRow = int(sys.argv[1])-1 if len(sys.argv)-1 >= 1 else 1
# startRow = 1
FIRST_ROW = 2
DEFAULT_TIMEOUT = 2000
SCREENSHOTMAXLENGTH = 23040
dlDirSetup = False

HELP_MSG = 'Usage: ' + \
    'python mat.py [mode] [flag] [rows]\n\n' + \
    'Modes are:\n\n' + \
    '-u, --upload\n' + \
    'Read the spreadsheet row-by-row and perform archival.\n\n' + \
    '-v, --validate\n' + \
    'Inspect the drive and check that all PNGs, PDFs, and MHTMLs are valid and present. Output results to validator_out.txt.\n\n' + \
    '-r, --repair\n' + \
    'Use validator_out.txt to repair uploads. Drive files will be deleted, re-saved, and re-uploaded.\n\n' + \
    'Flags are:\n\n' + \
    '--start=<start>\n' + \
    'What spreadsheet row to start on. (Counting from 1 just as in Google Sheets.) \n\n' + \
    '--timeout=<timeout>\n' + \
    'How many rows to work on.\n\n' + \
    '--redo\n' + \
    'Run the uploader but in "repair" mode. Basically, delete existing files and re-save and re-upload them.\n\n' + \
    'You can manually list rows to work on. E.g.\n' + \
    '\tpython mat.py --upload 245 610\n' + \
    'will only work on rows 245 and 610 specifically.\nBy default, the tool starts from the second row of the spreadsheet and timeouts after 2000 entries.'


VALIDATOR_FILE_PATH = CWD + '/validator_out.txt'
INVALID_PNG_MSG = '"Invalid PNG file"'
INVALID_PDF_MSG = '"Invalid PDF file"'
INVALID_MHTML_MSG = '"Invalid MHTML file"'
MISSING_PNG_MSG = '"Missing PNG file"'
MISSING_PDF_MSG = '"Missing PDF file"'
MISSING_MHTML_MSG = '"Missing MHTML file"'
SEPARATOR = '     '
SEPARATOR_SUB = '||'

# install extensions
options = webdriver.ChromeOptions()
options.add_extension('./UblockOrigin.crx')
options.add_extension('./GoFullPage.crx')
options.add_argument('--start-maximized')
options.add_argument('--disable-notifications')
# options.add_argument('--force-dark-mode=true')
# https://github.com/angular/protractor/issues/5347
# options.add_argument('--disable-site-isolation-trials')
# block Chrome popups https://www.browserstack.com/docs/automate/selenium/handle-permission-pop-ups#python
# options.add_experimental_option('excludeSwitches', ['disable-popup-blocking'])
# set download directory https://stackoverflow.com/a/19024814
prefs = {
    # pin GoFullPage
    'extensions.pinned_extensions': [GOFULLPAGE],
    # grant GoFullPage download permission
    'extensions.settings.' + GOFULLPAGE + '.runtime_granted_permissions':
        {"api":["downloads","webNavigation"],"explicit_host":[],"manifest_permissions":[],"scriptable_host":[]},
    # assign GoFullPage hotkey
    'extensions.settings.' + GOFULLPAGE + '.commands':
        {"_execute_browser_action":{"suggested_key":"Alt+Shift+P","was_assigned":"true"}},
    # set default download directory
    'download.default_directory': CWD + OUTPUTDIR
}
options.add_experimental_option('prefs', prefs)
##############################################################################

def closeProgram():
    sys.exit()

def getItem(list, index):
    try:
        return list[index]
    except:
        return ''

def bringWindowToFront(driver, altTab=True):
    if not altTab:
        position = driver.get_window_position()
        driver.minimize_window()
        driver.set_window_position(position.get('x'), position.get('y'))
        driver.maximize_window()
        time.sleep(DELAYSHORT)
    else:
        pyautogui.hotkey('alt','tab')
        time.sleep(DELAYSHORT)
        pyautogui.hotkey('alt','tab')
        time.sleep(DELAYSHORT)

def checkCapsLock():
    capslock = subprocess.getoutput('xset q | grep Caps')[21:24]
    while capslock == 'on ':
        pyautogui.confirm(text='Error: caps lock is ON. Turn it OFF.', title='', buttons=['OK'])
        time.sleep(1)
        capslock = subprocess.getoutput('xset q | grep Caps')[21:24] # NOTE: may not work on windows

def proceedPrompt(msg='Proceed?',yesPrompt=True,yesText='Yes',skipPrompt=False):
    buttons=['Wait','Quit']
    okToProceed = 'Wait'
    if yesPrompt:
        buttons.insert(0, yesText)
    if skipPrompt:
        buttons.insert(2, 'Skip')
    while okToProceed == 'Wait':
        time.sleep(5)
        okToProceed = pyautogui.confirm(text=msg, title='', buttons=buttons)
        if okToProceed == 'Quit':
            print('Aborted.')
            closeProgram()
        elif okToProceed == 'Skip':
            return False
    time.sleep(1)
    return True

def getDocLength(driver):
    # (Note that some webpages, e.g. ici.radio-canada, return zero for all of the below commands)
    SH = int(driver.execute_script('return document.body.scrollHeight') or 0)
    OH = int(driver.execute_script('return document.body.offsetHeight') or 0)
    CH = int(driver.execute_script('return document.body.clientHeight') or 0)
    eSH = int(driver.execute_script('document.documentElement.scrollHeight') or 0)
    eOH = int(driver.execute_script('document.documentElement.offsetHeight') or 0)
    eCH = int(driver.execute_script('document.documentElement.clientHeight') or 0)
    return max(SH,OH,CH,eSH,eOH,eCH)

def closeAllTabsExceptFirst(driver):
    # close all tabs except the first
    # https://www.geeksforgeeks.org/opening-and-closing-tabs-using-selenium/
    for jj in range(len(driver.window_handles)-1, 0, -1):
        driver.switch_to.window(driver.window_handles[jj])
        driver.close()
        driver.switch_to.window(driver.window_handles[jj-1])
        time.sleep(DELAYSHORT)

def preparePage(driver, docLength, entryNumber):
    # scroll to bottom of page incrementally to ensure rendered properly
    interval = 4
    for j in range(interval):
        driver.execute_script(f'window.scrollBy(0, {docLength} / {interval});')
        time.sleep(DELAYSHORT)
    driver.execute_script('window.scrollTo(0, 0)')
    time.sleep(DELAYSHORT)

    bringWindowToFront(driver, altTab=False)
    proceedVal = proceedPrompt(f'Entry {entryNumber}.\n\n' +
        'Take a moment to visually inspect the page.' + \
        ' Check that it looks okay to save. Close popups, etc. Click "Continue" when finished.',
        yesText='Continue',
        skipPrompt=True)
    return proceedVal

def saveSingleScreenshot(driver, fileName):
    global dlDirSetup

    time.sleep(DELAYSHORT)
    if len(driver.window_handles) == 1:
        pyautogui.hotkey('shift', 't')

    bringWindowToFront(driver)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(DELAYSHORT)
    # navigate to correct directory on first time
    if not dlDirSetup:
        pyautogui.typewrite(CWD + OUTPUTDIR)
        time.sleep(DELAYSHORT)
        pyautogui.hotkey('enter')
        time.sleep(DELAYSHORT)
        dlDirSetup = True
    # pyautogui.typewrite(fileName + ('-' + str(i) if N > 1 else '') + '.png')
    pyautogui.typewrite(fileName + '.png')
    time.sleep(DELAYSHORT)
    pyautogui.hotkey('enter')
    time.sleep(DELAYSHORT)
    driver.close()

def saveScreenshot(driver, docLength, iterationCount, entryFilename):
    # If the page is too long, GoFullPage will split screenshot into multiple pages.
    # Each page has max length of 23040px.
    # Determine how many sceenshots are required.
    # (But note that some webpages, e.g. radio-canada, return zero for all of the below
    # commands, so the user has to enter the number of screenshots manually.)

    # https://stackoverflow.com/a/3930320
    numScreenshots = math.ceil(docLength / SCREENSHOTMAXLENGTH)
    """
    proceedPrompt(f'document.body.scrollHeight = {SH}\n' +
                    f'document.body.offsetHeight = {OH}\n' +
                    f'document.body.clientHeight = {CH}\n' +
                    f'documentElement.scrollHeight = {eSH}\n' + 
                    f'documentElement.offsetHeight = {eOH}\n' + 
                    f'documentElement.clientHeight = {eCH}\n' +
                    f'Number of screenshots required = {numScreenshots}')
    """

    # Now take the screenshot ########################################
    bringWindowToFront(driver,altTab=False)

    # set up GoFullPage on first iteration
    if iterationCount == 0:
        # open extension options
        driver.switch_to.new_window('tab')
        driver.get('chrome-extension://' + GOFULLPAGE + '/options.html')
        time.sleep(DELAYSHORT)

        # enable "save as" check box
        driver.find_element(By.XPATH,'//input[@type="checkbox"][@name="save_as"]').click()
        time.sleep(DELAYSHORT)

        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(DELAYSHORT)

    checkCapsLock()
    driver.execute_script('window.scrollTo(0, 0)')
    time.sleep(DELAYSHORT)
    pyautogui.hotkey('alt', 'shift', 'p')
    time.sleep(DELAYSHORT)

    tabsInitial = driver.window_handles

    # wait for GoFullPage to finish
    while driver.window_handles == tabsInitial:
        # print(driver.window_handles)
        time.sleep(DELAYSHORT)
    time.sleep(DELAY)

    if numScreenshots == 1:
        # GoFullPage popup
        driver.switch_to.window(driver.window_handles[-1])
        # click download button
        try:
            driver.find_element(By.XPATH,'//a[@id="btn-download"]').click()
            time.sleep(DELAYSHORT)
            saveSingleScreenshot(driver,entryFilename)
        except:
            try:
                # If download button can't be found, it' because GoFullPage has a popup
                # ("Try annotating, cropping, and doing more with your screenshots!")
                # that you need to close first
                driver.find_element(By.XPATH,
                                    '//*[@id="modal-wrapper"]/div/div/div/a[@class="close close-topright"]'
                                    ).click()
                driver.find_element(By.XPATH,'//a[@id="btn-download"]').click()
                time.sleep(DELAYSHORT)
                saveSingleScreenshot(driver,entryFilename)
            except:
                # if it still can't be found, tell the user to take it manually.
                pyperclip.copy(entryFilename + '.png')
                proceedPrompt('Cannot take/save screenshot. Do it manually.' + 
                                ' The correct name has been copied to the clipboard.' +
                                ' Close the GoFullPage tab afterward and ensure the window is maximized.',
                                yesText='Continue')
        
        driver.switch_to.window(driver.window_handles[0])
    elif numScreenshots < 1:
        pyperclip.copy(entryFilename)
        promptMsg = 'Could not detect the number of screenshots. ' + \
                    'Save them all manually.' + \
                    ' The correct prefix has been copied to the clipboard.' + \
                    ' Close the GoFullPage tab afterward and ensure the window is maximized.' + \
                    ' When finished, enter how many screenshots were saved.'
        numScreenshots = None
        while numScreenshots is None:
            numScreenshots = pyautogui.prompt(text=promptMsg)
            # check that input is non-negative integer
            try:
                numScreenshots = int(numScreenshots)
                if numScreenshots < 0:
                    numScreenshots = None
            except:
                numScreenshots = None
    else:
        pyperclip.copy(entryFilename)
        proceedPrompt('Multiple screenshots. Save them all manually.' + 
                        ' The correct prefix has been copied to the clipboard.' +
                        ' Close the GoFullPage tab afterward and ensure the window is maximized.',
                        yesText='Continue')
        
    return numScreenshots

def savePDF(driver, iterationCount, entryFilename):
    global dlDirSetup

    # print to PDF
    bringWindowToFront(driver,altTab=False)

    try:
        driver.execute_script('window.scrollTo(0, 0)')
        time.sleep(DELAYSHORT)
        checkCapsLock()
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(DELAY)
        if iterationCount == 0: # set up printing on first iteration
            proceedPrompt('Ensure that the print dialog is set to "Save to PDF".',
                          yesText='Continue')
        time.sleep(DELAY)
        pyautogui.hotkey('enter')
        time.sleep(DELAY)
        bringWindowToFront(driver)

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(DELAYSHORT)
        if not dlDirSetup:
            pyautogui.typewrite(CWD + OUTPUTDIR)
            time.sleep(DELAYSHORT)
            pyautogui.hotkey('enter')
            time.sleep(DELAYSHORT)
            dlDirSetup = True
        pyautogui.typewrite(entryFilename + '.pdf')
        time.sleep(DELAYSHORT)
        pyautogui.hotkey('enter')
        time.sleep(DELAY)
    except:
        # if it still can't be found, tell the user to take it manually.
        pyperclip.copy(entryFilename + '.pdf')
        proceedPrompt('Cannot print page. Do it manually.' + 
                        ' The correct name has been copied to the clipboard.' +
                        ' Afterward, ensure the window is maximized.',
                        yesText='Continue')
        
    driver.switch_to.window(driver.window_handles[0])

def saveMHTML(driver, iterationCount, entryFilename):
    global dlDirSetup

    bringWindowToFront(driver,altTab=False)

    driver.execute_script('window.scrollTo(0, 0)')
    time.sleep(DELAYSHORT)
    checkCapsLock()
    pyautogui.hotkey('ctrl', 's')
    time.sleep(DELAYSHORT)
    if iterationCount == 0:
        proceedPrompt('Ensure that the dropdown is set to "Webpage, Single File".')
    try:
        bringWindowToFront(driver)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(DELAYSHORT)
        if not dlDirSetup:
            pyautogui.typewrite(CWD + OUTPUTDIR)
            time.sleep(DELAYSHORT)
            pyautogui.hotkey('enter')
            time.sleep(DELAYSHORT)
            dlDirSetup = True
        pyautogui.typewrite(entryFilename + '.mhtml')
        time.sleep(DELAYSHORT)
        pyautogui.hotkey('enter')
        time.sleep(DELAY)

        entryFilepath = CWD + OUTPUTDIR + entryFilename
        if not downloadSuccessful(entryFilepath + '.mhtml'):
            closeProgram()
    except:
        pyperclip.copy(entryFilename + '.mhtml')
        proceedPrompt('Cannot save page. Do it manually.' + 
                        ' The correct name has been copied to the clipboard.' +
                        ' Close the GoFullPage tab afterward and ensure the window is maximized.',
                        yesText='Continue')

# https://stackoverflow.com/a/62319275
def downloadSuccessful(filename, timeout=600):
    end_time = time.time() + timeout
    while not os.path.exists(filename):
        time.sleep(1)
        if time.time() > end_time:
            print('Download timed out after ' + timeout + ' seconds: ' + filename)
            return False

    if os.path.exists(filename):
        print('Download successful: ' + filename)
        return True

def searchFileByName(service, name, parentID, testMode):
    if testMode:
        results = service.files().list(
            q='name="' + name + '" and mimeType!="application/vnd.google-apps.folder"' + 
            'and parents in "' + parentID + '" and trashed=false'
        ).execute()
    else:
        results = service.files().list(
            q='name="' + name + '" and mimeType!="application/vnd.google-apps.folder"' + 
            'and parents in "' + parentID + '" and trashed=false',
            driveId = SHARED_DRIVE,
            corpora='drive',
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
    
    searchResult = results.get('files', [])

    if not searchResult:
        return None
    else:
        # if there are multiple files of the same name (there shouldn't be),
        # return the ID of the first one
        return searchResult[0].get('id')

def searchFileByID(service, fileID, testMode):
    try:
        return service.files().get(fileId = fileID,
                                   supportsAllDrives = (not testMode)).execute()
    except HttpError as error:
        return None

def downloadFile(service, fileID, verbose=True):
    request = service.files().get_media(fileId=fileID)
    file = io.BytesIO()
    downloader = MediaIoBaseDownload(file, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if verbose:
            print(f'Download {int(status.progress() * 100)}.')
    return file.getvalue()

def downloadFileChunk(service, fileID, start, length):
    # https://stackoverflow.com/a/59764650
    if length <= 0:
        return None
    request = service.files().get_media(fileId=fileID)
    request.headers['Range'] = f'bytes={start}-{start+length-1}'
    file = io.BytesIO(request.execute())
    return file.getvalue()

def uploadFile(service, name, parentID, testMode, path=CWD+OUTPUTDIR, mimetype=None, verbose=True):
    # check if the file already exists first
    if searchFileByName(service, name, parentID, testMode):
        proceedVal = proceedPrompt('A file named ' + name + ' has already been uploaded. Proceed?',
                                   skipPrompt=True)
        if not proceedVal:
            if verbose:
                print(f'Did not upload "{name}".')
            return False

    # also check that file exists locally
    while not os.path.exists(path+name):
        pyperclip.copy(name)
        proceedVal = proceedPrompt(f'File "{name}" was not found in the "./files" folder. ' +
                                   'Check that all local files are correctly named. ' +
                                   'The correct filename has been copied to the clipboard.',
                                   yesText='Continue')

    file_metadata = {
        'name': name,
        'parents': [parentID]
    }
    media = MediaFileUpload(
        path + name,
        mimetype=mimetype,
        resumable=True,
    )
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id',
        supportsAllDrives=True
    ).execute()

    if verbose:
        print(f'Uploaded file: "{name}" ({file.get("id")}).')

    return True

def trashRemoteFile(service, fileID, testMode, verbose=True):
    # check if the file exists
    # print(f'testMode = {testMode}')
    """
    if not searchFileByID(service, fileID, testMode):
        proceedVal = proceedPrompt(f'File with ID {fileID} does not exist.',
                                   yesPrompt=False,
                                   skipPrompt=True)
        if not proceedVal:
            if verbose:
                print(f'Did not trash file with ID "{fileID}".')
            return False
    """
    # print(f'ID={fileID}')

    # https://stackoverflow.com/a/56470194
    service.files().update(fileId = fileID,
                           supportsAllDrives = True,
                           body = {'trashed': True}).execute()

    if verbose:
        print(f'Trashed file with ID {fileID}.')

    return True

def deleteLocalFile(name):
    if os.path.isfile(name):
        os.remove(name)
    else:
        print("Error: file not found: %s" % name)
        closeProgram()

def numFilesInFolder(path):
    return len(os.listdir(path))

def openFolder(service, name, testMode, parentID=None, create=False, verbose=True):
    strBegin = 'name="' + name + '" and mimeType="application/vnd.google-apps.folder"'
    strEnd = ' and trashed=false'
    if parentID:
        strMid = ' and parents in "' + parentID + '"'
        qStr = strBegin + strMid + strEnd
    else:
        qStr = strBegin + strEnd

    if testMode:
        results = service.files().list(q=qStr).execute()
    else:
        results = service.files().list(
            q=qStr,
            driveId = SHARED_DRIVE,
            corpora='drive',
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
    folder = results.get('files', [])

    # If the folder doesn't exist, or the user doesn't care if it does, create it
    if create or not folder:
        proceedPrompt(f'Create Drive folder "{name}"?')
        
        file_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parentID:
            file_metadata['parents'] = [parentID]

        file = service.files().create(
            body=file_metadata,
            fields='id',
            supportsAllDrives=True
        ).execute()
        folder = {'name': name, 'id': file['id']}
        if verbose:
            print(u'Created Drive folder "{0}" ({1})'.format(folder['name'], folder['id']))
    else:
        assert(len(folder) == 1)
        folder = folder[0]
        if verbose:
            print(u'Opened Drive folder "{0}" ({1})'.format(folder['name'], folder['id']))

    return folder

def uploadFolder(service, name, parentID, testMode, create=False, verbose=True):
    
    folder = openFolder(service, name, testMode, parentID=parentID, create=create, verbose=verbose)

    dirPath = CWD + OUTPUTDIR + name
    if dirPath[-1] != '/':
        dirPath = dirPath + '/'

    count = 1
    for filename in os.listdir(dirPath):
        f = os.path.join(dirPath, filename)
        if os.path.isfile(f):
            uploadFile(service, str(filename).split('/')[-1], folder['id'], testMode, path=dirPath)
            if verbose:
                print(str(count) + '/' + str(numFilesInFolder(dirPath)) + ' successfully uploaded.')
        count += 1

def deleteLocalFolder(dirPath):
    # proceedPrompt('About to delete folder ' + CWD + OUTPUTDIR + name)
    try:
        shutil.rmtree(dirPath)
    except OSError as e:
        print("Error: %s - %s." % (e.filename, e.strerror))
        closeProgram()

def emptyLocalFolder(dirPath):
    for filename in os.listdir(dirPath):
        os.remove(os.path.join(dirPath, filename))

def setupWorkFolder():
    dirPath = CWD + OUTPUTDIR
    if not os.path.exists(dirPath):
        os.mkdir(dirPath)
    else:
        if len(os.listdir(dirPath)) > 0:
            proceedPrompt('The "./files" directory must be emptied.' + 
                          ' Backup any important files that are in there now.' +
                          ' Then click "Continue" to proceed with deletion.',
                          yesText='Continue')
            emptyLocalFolder(dirPath)

"""
def parseDate(sheet, row):
    dateValid = False
    # parse the date
    # date format can be YYYY-DD-MM or MM/DD/YYYY
    while not dateValid:
        dRe = re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})/(\d{1,2})/(\d{4})')
        dReS = dRe.search(sheet[1, row])
        if dReS:
            dReSG = dReS.groups()
            if dReSG[0] and dReSG[1] and dReSG[2] and not dReSG[3] and not dReSG[4] and not dReSG[5]:
                entryYear = dReSG[0]
                entryMonth = dReSG[1].zfill(2)
                dateValid = True
            elif not dReSG[0] and not dReSG[1] and not dReSG[2] and dReSG[3] and dReSG[4] and dReSG[5]:
                entryYear = dReSG[5]
                entryMonth = dReSG[3].zfill(2)
                dateValid = True
        if not dateValid:
            proceedVal = proceedPrompt('Date of spreadsheet entry ' + str(row) +
                                        ' not formatted correctly. ',
                                        yesPrompt='Continue',skipPrompt=True)
            if not proceedVal:
                return None
        else:
            return entryMonth, entryYear
"""

def initialize():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

# Step 0: open Google Drive API ##########################################3###
def uploader(creds, start=FIRST_ROW, timeout=DEFAULT_TIMEOUT, testMode=True, prompts=True, redo=False, rowsToProcess=[]):
    global dlDirSetup

    try:
        # Call the Drive v3 API
        driveService = build('drive', 'v3', credentials=creds)
        # Call the Sheets v4 API
        sheetsService = build('sheets', 'v4', credentials=creds)

        # start Chrome driver
        driver = webdriver.Chrome(options=options) 

        # Step 1: open the spreadsheet #######################################

        # get spreadsheet data
        sheet = sheetsService.spreadsheets()
        # in Test Mode, only operate on personal drive, not shared drive
        sheetProperties = sheet.get(spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                                    fields = 'sheets.properties').execute()
        rowCount = sheetProperties['sheets'][0]['properties']['gridProperties']['rowCount']
        colCount = sheetProperties['sheets'][0]['properties']['gridProperties']['columnCount']

        result = sheet.values().get(spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                                    range = f'Sheet1!R1C1:R{rowCount}C{colCount}').execute()
        sheetValues = result.get('values', [])

        if not sheetValues:
            print('No data found.')
            closeProgram()

        # Step 3: loop through rows of the sheet #############################
        # note: I'm counting from one, but GoogleSheets API counts from zero
        iterationCount = 0
        dlDirSetup = False

        if rowsToProcess == []:
            loopRange = range(start, rowCount+1)
        else:
            loopRange = rowsToProcess

        for i in loopRange:
            # get corresponding row of spreadsheet
            currentRow = sheetValues[i-1]

            # make sure work directory is empty
            setupWorkFolder()

            # entryDate = parseDate(sheet,i)
            # if not entryDate:
            #     time.sleep(DELAYSHORT)
            #     continue
            # entryMonth = entryDate[0]
            # entryYear = entryDate[1]

            # row should at least have a filename
            if len(currentRow) < 19:
                continue

            # get row data
            entryNumber = i
            entryURL = currentRow[5]
            entryMediaType = currentRow[7]
            entryTopic = currentRow[8]
            entryFilename = currentRow[18]
            entryFilepath = CWD + OUTPUTDIR + entryFilename
            entryArchived = getItem(currentRow, 19)

            if not redo and \
            (entryTopic == '' or entryTopic == ' ' or entryArchived.lower() == 'yes' \
            or entryFilename == '' or entryFilename == ' ' or entryURL[-4:] == '.pdf' \
            or entryMediaType == 'Radio' or entryMediaType == 'Television' \
            or entryMediaType == 'Social Media'):
                continue

            # TODO: assert file is not empty string
            # do this in repairer() also
            # TODO: i believe parsing should be independent of the filename formatting, fix this
            entryDate = entryFilename.split(' - ')[0].split('-')
            entryYear = entryDate[0]
            entryMonth = entryDate[1]
            
            indent = '  '
            print('========================================================================')
            print('Current file:')
            curFileMsg = indent + f'File {entryNumber} (Iteration {iterationCount}) \n' + \
                indent + f'YYYY-DD: {entryYear}-{entryMonth}\n' + \
                indent + f'URL: {entryURL}\n' + \
                indent + f'Topic: {entryTopic}\n' + \
                indent + f'Filename: {entryFilename}\n' + \
                indent + f'Archived: {entryArchived}'
            print(curFileMsg)
            print('------------------------------------------------------------------------')

            if prompts:
                proceedVal = proceedPrompt(msg=f'About to work on file {entryNumber} ("{entryFilename}").\nProceed?',
                                           skipPrompt=True)
                if not proceedVal:
                    time.sleep(DELAYSHORT)
                    continue

            # Step 5: Find/create the corresponding year folder ##############
            yearFolder = openFolder(driveService, entryYear, testMode,
                                    parentID = None if testMode else SHARED_DRIVE)
            monthFolder = openFolder(driveService, entryMonth, testMode, yearFolder['id'])

            # Step 4: Visit the website ######################################
            # https://stackoverflow.com/a/69137491

            # open page with selenium
            driver.get(entryURL)
            time.sleep(DELAYSHORT)

            # close all tabs except the first
            closeAllTabsExceptFirst(driver)

            # determine length of document
            docLength = getDocLength(driver)

            # copy filename to clipboard just in case we need to past it
            pyperclip.copy(entryFilename)

            # visually prepare page for archival, or skip if user indicates
            proceedVal = preparePage(driver, docLength, entryNumber)
            if not proceedVal: # user wants to skip
                time.sleep(DELAYSHORT)
                continue

            # save screenshot
            numScreenshots = saveScreenshot(driver, docLength, iterationCount, entryFilename)

            # save PDF
            savePDF(driver, iterationCount, entryFilename)

            # save to MHTML
            saveMHTML(driver, iterationCount, entryFilename)

            # Step 6: Upload files ###########################################

            # if redoing, delete old files
            if redo:
                pngID = []
                completePngID = searchFileByName(driveService, entryFilename + '.png', monthFolder['id'], testMode)
                if not completePngID:
                    # can't find PNG file, may be split up into multiple screenshot so check
                    j = 1
                    partialPngID = searchFileByName(driveService, entryFilename + f'-{j}.png', monthFolder['id'], testMode)
                    while partialPngID is not None:
                        pngID.append(partialPngID)
                        j += 1
                        partialPngID = searchFileByName(driveService, entryFilename + f'-{j}.png',
                                                        monthFolder['id'],
                                                        testMode)   
                else:
                    pngID.append(completePngID)

                for ik in range(len(pngID)):
                    if prompts:
                        proceedPrompt(f'About to delete file {entryFilename}' +
                                      (f'-{ik}' if len(pngID)>1 else '') +
                                      f'.png ({pngID[ik]}).')
                    trashRemoteFile(driveService, pngID[ik], testMode)

                # delete pdf
                pdfID = searchFileByName(driveService, entryFilename + '.pdf', monthFolder['id'], testMode)
                if prompts:
                    proceedPrompt(f'About to delete file {entryFilename}.pdf ({pdfID}).')
                if pdfID is not None:
                    trashRemoteFile(driveService, pdfID, testMode)

                # delete mhtml
                mhtmlID = searchFileByName(driveService, entryFilename + '.mhtml', monthFolder['id'], testMode)
                if prompts:
                    proceedPrompt(f'About to delete file {entryFilename}.mhtml ({mhtmlID}).')
                if mhtmlID is not None:
                    trashRemoteFile(driveService, mhtmlID, testMode)

            ##################################################################

            if prompts:
                proceedPrompt('About to upload files to Drive folder ' + yearFolder['name'] +
                              '/' + monthFolder['name'] + '.\n\nProceed?')

            if numScreenshots == 1:
                if uploadFile(driveService, entryFilename + '.png', monthFolder['id'], testMode):
                    deleteLocalFile(entryFilepath + '.png')
            else:
                for jjj in range(1, numScreenshots+1):
                    if uploadFile(driveService, entryFilename + '-' + str(jjj) + '.png', monthFolder['id'], testMode):
                        deleteLocalFile(entryFilepath + '-' + str(jjj) + '.png')

            if uploadFile(driveService, entryFilename + '.pdf', monthFolder['id'], testMode):
                deleteLocalFile(entryFilepath + '.pdf')

            if uploadFile(driveService, entryFilename + '.mhtml', monthFolder['id'], testMode):
                deleteLocalFile(entryFilepath + '.mhtml')

            ##################################################################

            # update spreadsheet
            sheetsService.spreadsheets().values().update(
                spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                range = f'T{entryNumber}',
                valueInputOption = 'RAW',
                body = {
                    'values': [['Yes']]
                }).execute()
            
            iterationCount += 1

            if iterationCount > timeout-1:
                closeProgram()

    except HttpError as error:
        # Handle errors from drive API.
        print(f'An error occurred: {error}')

def validator(creds,start,timeout=DEFAULT_TIMEOUT,testMode=False, rowsToProcess=[]):

    validatorOutputFile = open(VALIDATOR_FILE_PATH, 'a')
    
    try:
        # Call the Drive v3 API
        driveService = build('drive', 'v3', credentials=creds)
        # Call the Sheets v4 API
        sheetsService = build('sheets', 'v4', credentials=creds)

        iterationCount = 0
        
        # get spreadsheet data
        sheet = sheetsService.spreadsheets()
        sheetProperties = sheet.get(spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                                    fields = 'sheets.properties').execute()
        rowCount = sheetProperties['sheets'][0]['properties']['gridProperties']['rowCount']
        colCount = sheetProperties['sheets'][0]['properties']['gridProperties']['columnCount']
        # print(f'rows = {rowCount}, cols = {colCount}')

        result = sheet.values().get(spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                                    range = f'Sheet1!R1C1:R{rowCount}C{colCount}').execute()
        sheetValues = result.get('values', [])

        if not sheetValues:
            print('No data found.')
            closeProgram()

        if rowsToProcess == []:
            loopRange = range(start, rowCount+1)
        else:
            loopRange = rowsToProcess

        for i in loopRange:
            # get corresponding row of spreadsheet
            currentRow = sheetValues[i-1]

            # only validate already-archived entries
            if len(currentRow) < 20 or currentRow[19].lower() != 'yes':
                continue

            # get row data
            entryNumber = i
            entryURL = currentRow[5]
            entryMediaType = currentRow[7]
            entryTopic = currentRow[8]
            entryFilename = currentRow[18]
            entryFilepath = CWD + OUTPUTDIR + entryFilename
            entryArchived = currentRow[19]

            entryDate = entryFilename.split(' - ')[0].split('-')
            entryYear = entryDate[0]
            entryMonth = entryDate[1]

            yearFolder = openFolder(driveService, entryYear, testMode,
                                    parentID = None if testMode else SHARED_DRIVE,
                                    verbose = False)
            monthFolder = openFolder(driveService, entryMonth, testMode, yearFolder['id'], verbose=False)

            # output buffer
            # note: in the output file, rows are counted starting from one
            outputLine = [str(entryNumber)]

            # check if PNG file(s) exist
            pngID = []
            completePngID = searchFileByName(driveService, entryFilename + '.png', monthFolder['id'], testMode)
            if not completePngID:
                # can't find PNG file, may be split up into multiple screenshot so check
                partialPngID = searchFileByName(driveService, entryFilename + '-1.png', monthFolder['id'], testMode)
                if not partialPngID:
                    outputLine.append(MISSING_PNG_MSG)
                else:
                    j = 1
                    while partialPngID is not None:
                        pngID.append(partialPngID)
                        j += 1
                        partialPngID = searchFileByName(driveService, entryFilename + f'-{j}.png',
                                                        monthFolder['id'],
                                                        testMode)   
                    outputLine.append('"Multiple PNG files"')
            else:
                pngID.append(completePngID)

            # check if PDF files exist
            pdfID = searchFileByName(driveService, entryFilename + '.pdf', monthFolder['id'], testMode)
            if not pdfID:
                outputLine.append(MISSING_PDF_MSG)

            # check if MHTML files exist
            mhtmlID = searchFileByName(driveService, entryFilename + '.mhtml', monthFolder['id'], testMode)
            if not mhtmlID:
                outputLine.append(MISSING_MHTML_MSG)

            # check if PNG file(s) is (are) valid
            if len(pngID) > 0:
                for k in range(len(pngID)):
                    fileBytes = downloadFileChunk(driveService,pngID[k],0,4)
                    if fileBytes != b'\x89PNG':
                        outputLine.append(SEPARATOR_SUB.join([INVALID_PNG_MSG, pngID[k]]))
            
            # check if PDF files are valid
            if pdfID:
                fileBytes = downloadFileChunk(driveService,pdfID,0,4)
                if fileBytes != b'%PDF':
                    outputLine.append(SEPARATOR_SUB.join([INVALID_PDF_MSG, pdfID]))
            
            # check if MHTML files are valid
            if mhtmlID:
                fileBytes = downloadFileChunk(driveService,mhtmlID,0,4)
                if fileBytes != b'From':
                    outputLine.append(SEPARATOR_SUB.join([INVALID_MHTML_MSG, mhtmlID]))

            print(f'Processed entry {entryNumber} - {entryFilename}.')

            # write buffer to file
            if len(outputLine) > 1:
                validatorOutputFile.write('\n' + SEPARATOR.join(outputLine))
            outputLine.clear()
            validatorOutputFile.flush()

            iterationCount += 1

            if iterationCount > timeout-1:
                break

    except HttpError as error:
        # Handle errors from drive API.
        print(f'An error occurred: {error}')

    validatorOutputFile.close()

def repairer(creds, prompts=True, testMode=False):
    global dlDirSetup

    validatorOutputFile = open(VALIDATOR_FILE_PATH, 'r')

    try:
        # Call the Drive v3 API
        driveService = build('drive', 'v3', credentials=creds)
        # Call the Sheets v4 API
        sheetsService = build('sheets', 'v4', credentials=creds)

        # start Chrome driver
        driver = webdriver.Chrome(options=options) 
        
        sheet = sheetsService.spreadsheets()
        sheetProperties = sheet.get(spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                                    fields='sheets.properties').execute()
        rowCount = sheetProperties['sheets'][0]['properties']['gridProperties']['rowCount']
        colCount = sheetProperties['sheets'][0]['properties']['gridProperties']['columnCount']

        result = sheet.values().get(spreadsheetId = SPREADSHEET_TEST if testMode else SPREADSHEET,
                                    range=f'Sheet1!R1C1:R{rowCount}C{colCount}').execute()
        sheetValues = result.get('values', [])

        if not sheetValues:
            print('No data found.')
            closeProgram()

        dlDirSetup = False
        iterationCount = 0

        for line in validatorOutputFile:
            words = line.split(sep=SEPARATOR)
            # skip improperly formatted lines
            if not words[0].isnumeric():
                continue

            entryNumber = int(words[0])
            numErrors = len(words) - 1

            # determine which errors to fix
            # problem: this assumes that there is only one invalid PNG
            # TODO: fix
            missingPNG = False
            missingPDF = False
            missingMHTML = False
            invalidPNG = False
            invalidPDF = False
            invalidMHTML = False
            pngID = None
            pdfID = None
            mhtmlID = None

            for j in range(1, len(words)):
                # https://stackoverflow.com/a/13298933
                # remove newline character from text
                currentError = words[j].replace('\n', '').split(sep=SEPARATOR_SUB)
                if currentError[0] == MISSING_PNG_MSG:
                    missingPNG = True
                elif currentError[0] == MISSING_PDF_MSG:
                    missingPDF = True
                elif currentError[0] == MISSING_MHTML_MSG:
                    missingMHTML = True
                elif currentError[0] == INVALID_PNG_MSG:
                    invalidPNG = True
                    pngID = currentError[1]
                elif currentError[0] == INVALID_PDF_MSG:
                    invalidPDF = True
                    pdfID = currentError[1]
                elif currentError[0] == INVALID_MHTML_MSG:
                    invalidMHTML = True
                    mhtmlID = currentError[1]

            # if there's no error, skip
            if not missingPNG and not missingPDF and not missingMHTML and \
            not invalidPNG and not invalidPDF and not invalidMHTML:
                continue

            # get corresponding row of spreadsheet
            currentRow = sheetValues[entryNumber-1]

            # make sure work directory is empty
            setupWorkFolder()

            # get spreadsheet data
            entryURL = currentRow[5]
            entryMediaType = currentRow[7]
            entryTopic = currentRow[8]
            entryFilename = currentRow[18]
            entryFilepath = CWD + OUTPUTDIR + entryFilename
            entryArchived = currentRow[19]

            entryDate = entryFilename.split(' - ')[0].split('-')
            entryYear = entryDate[0]
            entryMonth = entryDate[1]
            
            indent = '  '
            print('========================================================================')
            print('Current file:')
            curFileMsg = indent + f'File {entryNumber} (Iteration {iterationCount}) \n' + \
                indent + f'YYYY-DD: {entryYear}-{entryMonth}\n' + \
                indent + f'URL: {entryURL}\n' + \
                indent + f'Topic: {entryTopic}\n' + \
                indent + f'Filename: {entryFilename}\n' + \
                indent + f'Archived: {entryArchived}'
            print(curFileMsg)
            print('------------------------------------------------------------------------')

            if prompts:
                proceedVal = proceedPrompt(msg=f'About to work on file {entryNumber} ("{entryFilename}").\nProceed?',
                                           skipPrompt=True)
                if not proceedVal:
                    time.sleep(DELAYSHORT)
                    continue

            # Step 5: Find/create the corresponding year folder ##############
            yearFolder = openFolder(driveService, entryYear, testMode,
                                    parentID = None if testMode else SHARED_DRIVE)        
            monthFolder = openFolder(driveService, entryMonth, testMode, parentID=yearFolder['id'])

            # Step 4: Visit the website ######################################
            # https://stackoverflow.com/a/69137491

            # open page with selenium
            driver.get(entryURL)
            time.sleep(DELAYSHORT)

            # close all tabs except the first
            closeAllTabsExceptFirst(driver)

            # determine length of document
            docLength = getDocLength(driver)

            # copy filename to clipboard just in case we need to past it
            pyperclip.copy(entryFilename)

            # visually prepare page for archival, or skip if user indicates
            proceedVal = preparePage(driver, docLength, entryNumber)
            if not proceedVal: # user wants to skip
                time.sleep(DELAYSHORT)
                continue

            # save and upload screenshot
            if invalidPNG:
                if prompts:
                    proceedPrompt(f'About to delete file {entryFilename}.png ({pngID}).')
                trashRemoteFile(driveService, pngID, testMode)
            if invalidPNG or missingPNG:
                numScreenshots = saveScreenshot(driver, docLength, iterationCount, entryFilename)
                proceedPrompt(f'About to upload file {entryFilename}.png.')
                if numScreenshots == 1:
                    if uploadFile(driveService, entryFilename + '.png', monthFolder['id'], testMode):
                        deleteLocalFile(entryFilepath + '.png')
                else:
                    for jjj in range(1, numScreenshots+1):
                        if uploadFile(driveService, f'{entryFilename}-{jjj}.png', monthFolder['id'], testMode):
                            deleteLocalFile(f'{entryFilepath}-{jjj}.png')

            # save and upload PDF
            if invalidPDF:
                if prompts:
                    proceedPrompt(f'About to delete file {entryFilename}.pdf ({pdfID}).')
                trashRemoteFile(driveService, pdfID, testMode)
            if invalidPDF or missingPDF:
                savePDF(driver, iterationCount, entryFilename)
                proceedPrompt(f'About to upload file {entryFilename}.pdf.')
                if uploadFile(driveService, entryFilename + '.pdf', monthFolder['id'], testMode):
                    deleteLocalFile(entryFilepath + '.pdf')

            # save and upload MHTML
            if invalidMHTML:
                if prompts:
                    proceedPrompt(f'About to delete file {entryFilename}.mhtml ({mhtmlID}).')
                trashRemoteFile(driveService, mhtmlID, testMode)
            if invalidMHTML or missingMHTML:
                saveMHTML(driver, iterationCount, entryFilename)
                proceedPrompt(f'About to upload file {entryFilename}.mhtml.')
                if uploadFile(driveService, entryFilename + '.mhtml', monthFolder['id'], testMode):
                    deleteLocalFile(entryFilepath + '.mhtml')

            # Step 6: Upload files ###########################################

            ##################################################################

            iterationCount += 1
    
    except HttpError as error:
        # Handle errors
        print(f'An error occurred: {error}')

    validatorOutputFile.close()

if __name__ == '__main__':
    # https://pymotw.com/2/getopt/
    try:
        opts, args = getopt(sys.argv[1:], 'uvr', ['upload', 'validate', 'repair', 'help', 'start=', 'timeout=', 'redo'])
    except:
        print('Error 1')
        closeProgram()
    
    upload = False
    uploadRedo = False
    validate = False
    repair = False
    help = False
    startRow = FIRST_ROW
    timeout = DEFAULT_TIMEOUT
    rows = []

    for opt, arg in opts:
        if opt in ('-u', '--upload'):
            upload = True
            validate = False
            repair = False
            help = False
        elif opt in ('-v', '--validate'):
            upload = False
            validate = True
            repair = False
            help = False
        elif opt in ('-r', '--repair'):
            upload = False
            validate = False
            repair = True
            help = False
        elif opt in ('--help'):
            upload = False
            validate = False
            repair = False
            help = True
        elif opt in ('--start'):
            startRow = max(int(arg), 2)
        elif opt in ('--timeout'):
            timeout = max(int(arg), 1)
        elif opt in ('--redo'):
            uploadRedo = True

    for num in args:
        try:
            rows.append(int(num))
        except:
            print('Error 2')
            closeProgram()

    creds = initialize()

    if upload:
        uploader(creds, startRow, timeout, testMode=False, prompts=uploadRedo, redo=uploadRedo, rowsToProcess=rows)
    elif validate:
        validator(creds, startRow, timeout, testMode=False, rowsToProcess=rows)
    elif repair:
        repairer(creds)
    elif help:
        print(HELP_MSG)












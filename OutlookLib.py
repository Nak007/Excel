'''
Available methods are the followings:
[1] ReadMail
[2] SendMail
[3] SaveAttachments

Authors: Danusorn Sitdhirasdr <danusorn.si@gmail.com>
versionadded:: 30-07-2023

'''
import win32com.client, os, numpy as np
from datetime import datetime, timedelta
from tqdm.notebook import tqdm_notebook
import ipywidgets as widgets
from IPython.display import display

__all__ = ["ReadMail",
           "SendMail",
           "SaveAttachments"]

def ReadMail(path, stop=None, start=None, days=100, sort=True):
    
    '''
    Read and extract Microsoft Outlook mails from designated folder.
    
    Parameters
    ----------
    path : str
        A valid Microsoft Outlook folder path e.g. "xxx@gmail\Inbox". 
    
    stop : str, default=None
        A stop date (format = "%d/%m/%Y %H:%M:%S"). If None, it uses 
        the current date i.e. datetime.now().
        
    start : str, default=None
        A start date (format = "%d/%m/%Y %H:%M:%S"). If None, it uses 
        the following formula : start = stop - timedelta(days).
        
    days : int, default=100
        Duration in days from start to stop dates. This is relevant 
        when start date is not defined (None).
        
    sort : bool, default=True
        If True, the 0th index will start from the most recent mail,
        otherwise the oldest.
        
    References
    ----------
    [1] https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem
    
    Returns
    -------
    mails : dict
        For each key, it contains dictionary, whose keys (properties) are 
        as follows:
        - "SenderName"   : Name of the sender
        - "To"           : List of recipients
        - "CC"           : List of carbon copy (CC) names
        - "Subject"      : Mail subject
        - "Body"         : The clear-text body of the Outlook item
        - "HTMLBody"     : HTML body of the specified item
        - "ReceivedTime" : Date & time at which the item was received
        - "CreationTime" : Creation time
        - "SentOn"       : Date & time at which the item was sent
        - "Attachments"  : Attached files.
    
    '''
    # Initialize parameters
    outlook = win32com.client.Dispatch("Outlook.Application")
    outlook = outlook.GetNamespace("MAPI")
    
    # Attributes (mailitem)
    keys = ["SenderName", "To", "CC", "Subject", "Body", 
            "HTMLBody", "ReceivedTime", "CreationTime", 
            "SentOn"]
 
    # Get subfolders
    subfolders = path.split("\\")
    folder = outlook.Folders[subfolders[0]]
    for name in subfolders[1:]:
        folder = folder.folders[name]

    # Query by start and stop dates
    start, finish = ValidateDate(stop, start, max(days,1))
    dt0 = start  - timedelta(minutes=1)
    dt1 = finish + timedelta(minutes=1)
    query = [f"[ReceivedTime] >= '{dt0:%d/%m/%Y %H:%M}'",
             f"[ReceivedTime] <= '{dt1:%d/%m/%Y %H:%M}'"]
    items = folder.items.Restrict(" AND ".join(query))
    items.Sort("[ReceivedTime]", sort)
    
    # Initialize widgets
    t = widgets.HTMLMath(value='')
    display(widgets.HBox([t]))
    t.value = f" Number of emails : {items.Count:,.0f}"

    # Loop through all mail items
    mails = dict()
    for item in tqdm_notebook(items):
        received = ToDatetime(item.ReceivedTime)
        if (received>=start) & (received<=finish):
            mails[len(mails)] = ExtractContent(item, keys)
        
    return mails

def ExtractContent(item, keys):
    '''Extract content from mail'''
    content = dict()
    for key in keys:
        value = getattr(item, key, None)
        if isinstance(value, datetime):
            content[key] = value.strftime("%d/%m/%Y %H:%M:%S")
        else: content[key] = value
        
    # Add attached files
    if item.Attachments.Count>0:
        files = dict()
        for a in item.Attachments:
            try: files[a.FileName] = a
            except: pass
    else: files = None
    content["Attachments"] = files 
    return content

def ToDatetime(date):
    '''Convert to datetime format'''
    if isinstance(date, datetime):
        format = "%d/%m/%Y %H:%M:%S"
        return datetime.strptime(date.strftime(format),format)
    else: return None

def ValidateDate(stop, start=None, days=100):
    '''Validate start and stop dates''' 
    format = "%d/%m/%Y %H:%M:%S"
    stop = datetime.now() if stop is None else \
    datetime.strptime(stop, format)
    start = stop - timedelta(days=days) if start \
    is None else datetime.strptime(start, format)
    return start, stop

def SendMail(htmlbody, recipients, cc=None, subject=None, 
             attachments=None, display=True, kwargs=None):
    
    '''
    Send email via Microsoft Outlook.
    
    Parameters
    ----------
    htmlbody : str
        htmlbody must be in HTML format.
    
    recipients : list of str
        List of valid recipients' email address e.g. ["xx@gmail.com"].
    
    cc : list of str, default=None
        List of valid email addresses to be copied ("CC"). 
    
    subject : str, default=None
        Subject of mail. If None, "No Subject" is used instead.
    
    attahcments : list of path object, default=None
        List must contain valid file paths. The string must be a full
        path to a local file e.g. ["C://path//example.xlsx"].

    display : bool, default=True
        If True, it displays the mail, otherwise send.
    
    kwargs : dict, default=None
        kwargs are used to specify other properties of mail e.g.
        {"Importance":1, "Sensitivity":0}. 
    
    '''
    # Initialize parameters
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0x0)
    
    # Add Subject
    if not isinstance(subject, str):
        mail.Subject = "No Subject"
    else: mail.Subject = subject
    
    # Add "TO", and "CC" recipients
    if isinstance(recipients, list): 
        mail.To = ";".join(recipients)
    if isinstance(cc, list): mail.CC = ";".join(cc)
    mail.HTMLBody = htmlbody
    
    # Attach files. 
    if isinstance(attachments, list):
        for path in attachments:
            if os.path.exists(path):
                mail.Attachments.Add(path)
    
    # Add other attributes
    if isinstance(kwargs, dict):
        for key,value in kwargs.items():
            setattr(mail, key, value)
    
    # Display or Send mail
    if display: mail.Display()
    else: mail.Send()

def SaveAttachments(items, folder=None):
    
    '''
    Save attachments to designated folder.
    
    Parameters
    ----------
    items : mail items
        An item collection representing the mail items in a folder.
    
    folder : path-like, default=None
        Folder path. If None, it uses a current working directory
        i.e. os.getcwd().
        
    Returns
    -------
    SavedFiles : list of str
        A list of saved file paths.
        
    '''
    if folder is None: folder = os.getcwd()
    if os.path.exists(folder)==False:
        raise ValueError(f"<{folder}> does not exist.")
    
    SavedFiles = []
    if items.Attachments.Count>0:
        for item in items.Attachments:
            filepath = os.path.join(folder, item.FileName)
            item.SaveAsFile(filepath)
            SavedFiles.append(filepath)
    return SavedFiles
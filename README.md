#  🤖Automating Download and Organization of E-mail Attachments in Outlook

This is a Python script that automates the process of filtering, downloading and sorting email attachments in Outlook.

## 📋Description

This script uses the `pywin32` library to interact with Outlook and download email attachments that match a specific subject. Attachments are saved to a destination folder and renamed based on the email subject and sender.

## 📝Requirements

- Python 3.x
- `pywin32` library

## 🔧Settings

1. Clone this repository to your local system.
2. Install the `pywin32` library if it is not already installed:

3. Open the `SaveAttachment.py` file and configure options according to your needs, such as your email address, search folder and path to save attachments.

## Usage
1. Make sure Outlook is open and authenticated with the correct email account.
2. Run the `SaveAttachment.py` file.

## Functionalities
The automated process includes:

1. Subject Filtering: The script scans emails in the specified folder (eg Inbox), filters out those having the desired subject ("Document Request").
2. Download PDF Attachments: The script automatically downloads attachments in PDF format from identified emails.
3. Auto Sort: Downloaded attachments are renamed based on subject and sender and then saved in the designated destination folder. This makes it easy to find the documents you need.
   
## Working

<a href="https://s11.gifyu.com/images/ScyAv.gif" title="Demonstration">
  <img src="https://s11.gifyu.com/images/ScyAv.gif" alt="Demonstration" height="400px">
</a>

## IDE Used
Project developed in:
<div align="center">
  <a href="https://i.ibb.co/VvHCbPg/1-k-Ig3-dwee-DFVGCQBUNWc-Fw.png" title="Jupyter Notebook">
    <img src="https://i.ibb.co/VvHCbPg/1-k-Ig3-dwee-DFVGCQBUNWc-Fw.png" alt="Jupyter Notebook" height="50px">
  </a>
  <a href="https://www.python.org/" title="Python">
    <img src="https://www.python.org/static/community_logos/python-logo-inkscape.svg" alt="Python" height="50px">
  </a>
</div>

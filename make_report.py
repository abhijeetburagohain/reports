from api import *
from helper_functions import *
from first_page_report_helper_functions import *
from first_page_report import first_page
from facebook_report import facebook
from instagram_report import instagram
from youtube_ import youtube
from last_page_report import last_page
from title import insert_centered_front_page
import time

start=time.time()
###Inputs
original_doc_id = 'original_doc_id-->and existing template of the first page of the report'
folder_id="folder_id where the report is to be created"
#################


print("Starting the first page insertion")
document_id=first_page(original_doc_id,folder_id)
time.sleep(1)

print("Starting the facebook insertion")
facebook(document_id)
time.sleep(2)

print("Starting the instagram insertion")
instagram(document_id)
time.sleep(1)

print("Starting the youtube insertion")
youtube(document_id)
time.sleep(1)

print("Starting the last page insertion")
last_page(document_id)

print("Adding the Front Page")

title="name of the report Weekly Report -"
insert_centered_front_page(docs_service, document_id, title)

print("The Report is Generated.")
print(f"Document_id = {document_id}")

end=time.time()
print(f"time taken:{end-start} secs")

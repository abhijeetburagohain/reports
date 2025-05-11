from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from itertools import islice
import re
import time

# Define the required scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive'
]

# Function to connect to Google Sheets, Docs, and Drive APIs using service account
def connect_to_google_services(credentials_file):
    # Load the service account credentials
    creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)

    sheets_service = build('sheets', 'v4', credentials=creds)  # Connect to Google Sheets API
    docs_service = build('docs', 'v1', credentials=creds)  # Connect to Google Docs API    
    drive_service = build('drive', 'v3', credentials=creds)  # Connect to Google Drive API

    return sheets_service, docs_service, drive_service

####################### Authenticating with the APIs ##################################
# Path to your service account credentials file
credentials_file = r"path to your credentials file" 

# Make the required connections
sheets_service, docs_service, drive_service = connect_to_google_services(credentials_file)
############################################################################################

############################ API CALL FUNCTIONS   #######################################

def create_document_in_folder(folder_id,docs_service,drive_service):
    # Step 1: Create a new document
    document = docs_service.documents().create(body={"title": "Name of the Report"}).execute()
    doc_id = document.get("documentId")

    # Step 2: Move the document to the specified folder
    drive_service.files().update(
        fileId=doc_id,
        addParents=folder_id,
        removeParents='root',  # Move from root to the target folder
        fields='id, parents'
    ).execute()

    # Step 3: Set page size and margins for A4 with 0.5 inch margins
    page_setup_request = {
        'updateDocumentStyle': {
            'documentStyle': {
                'pageSize': {
                    'width': {
                        'magnitude': 8.27* 72,
                        'unit': 'PT'
                    },
                    'height': {
                        'magnitude': 11.69* 72,
                        'unit': 'PT'
                    }
                },
                'marginTop': {
                    'magnitude': 0.3*72,
                    'unit': 'PT'
                },
                'marginBottom': {
                    'magnitude': 0.3*72,
                    'unit': 'PT'
                },
                'marginLeft': {
                    'magnitude': 0.3*72,
                    'unit': 'PT'
                },
                'marginRight': {
                    'magnitude': 0.3*72,
                    'unit': 'PT'
                }
            },
            'fields': 'pageSize,marginTop,marginBottom,marginLeft,marginRight'
        }
    }
    
    # Execute the batchUpdate to set up the page
    docs_service.documents().batchUpdate(
        documentId=doc_id, body={'requests': [page_setup_request]}
    ).execute()

    return doc_id

def copy_document(drive_service, original_doc_id, new_title, folder_id=None):
    """
    Copy a document and optionally move it to a specific folder.

    :param drive_service: The authenticated Drive API service.
    :param original_doc_id: The ID of the document to copy.
    :param new_title: The title of the new copied document.
    :param folder_id: The ID of the folder where the document should be copied (optional).
    :return: The ID of the copied document.
    """
    # Create the request to copy the document
    file_metadata = {'name': new_title}
    
    # If a folder_id is provided, add it to the metadata
    if folder_id:
        file_metadata['parents'] = [folder_id]

    # Make the copy of the document
    copied_file = drive_service.files().copy(
        fileId=original_doc_id,
        body=file_metadata
    ).execute()

    # Return the new file's ID
    return copied_file['id']


def insert_styled_paragraph(docs_service, document_id, text):
    """
    Inserts a paragraph with specific styling: #073763 text color, 14 font size, center-aligned, and bold.

    :param docs_service: Google Docs API service instance.
    :param document_id: The ID of the document.
    :param text: The text to insert as a paragraph.
    """
    document = docs_service.documents().get(documentId=document_id).execute()
    start_index = document['body']['content'][1]['startIndex']
    # print(start_index)
    requests = [
        # Insert the text as a new paragraph
        {
            'insertText': {
                'location': {'index': start_index},  # Insert at the beginning; modify as needed
                'text': text + "\n"
            }
        },
        # Apply styling to the inserted text
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': 1 + len(text)
                },
                'textStyle': {
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 0.03,
                                'green': 0.22,
                                'blue': 0.39
                            }
                        }
                    },
                    'fontSize': {
                        'magnitude': 14,
                        'unit': 'PT'
                    },
                    'bold': True
                },
                'fields': 'foregroundColor,fontSize,bold'
            }
        },
        # Center-align the paragraph
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': 1 + len(text)
                },
                'paragraphStyle': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        }
    ]

    # Execute the batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Styled paragraph inserted successfully.")



### Inserting Headers and Footers Function
def insert_header_footer(docs_service, document_id, header_text, footer_text):
    try:
        # Step 1: Create header and footer
        requests = [
            {'createHeader': {'type': 'DEFAULT'}},
            {'createFooter': {'type': 'DEFAULT'}}
        ]
        
        # Execute the initial batch update to create header and footer
        response = docs_service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests}
        ).execute()
        
        # Extract the headerId and footerId from the response
        header_id = response['replies'][0]['createHeader']['headerId']
        footer_id = response['replies'][1]['createFooter']['footerId']
        
        # Step 2: Add text to the header and footer
        requests = [
            # Insert text into the header with right alignment
            {
                'insertText': {
                    'location': {'segmentId': header_id, 'index': 0},
                    'text': header_text
                }
            },
            {
                'updateTextStyle': {
                    'range': {'segmentId': header_id, 'startIndex': 0, 'endIndex': len(header_text)},
                    'textStyle': {
                        'italic': True,
                        'fontSize': {'magnitude': 8, 'unit': 'PT'},
                        'foregroundColor': {'color': {'rgbColor': {'red': 0, 'green': 0, 'blue': 0}}}
                    },
                    'fields': 'italic,fontSize,foregroundColor'
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {'segmentId': header_id, 'startIndex': 0, 'endIndex': len(header_text)},
                    'paragraphStyle': {'alignment': 'END'},
                    'fields': 'alignment'
                }
            },
            # Insert text into the footer with left alignment
            {
                'insertText': {
                    'location': {'segmentId': footer_id, 'index': 0},
                    'text': footer_text
                }
            },
            {
                'updateTextStyle': {
                    'range': {'segmentId': footer_id, 'startIndex': 0, 'endIndex': len(footer_text)},
                    'textStyle': {
                        'fontSize': {'magnitude': 8, 'unit': 'PT'},
                        'foregroundColor': {'color': {'rgbColor': {'red': 0.6, 'green': 0.6, 'blue': 0.6}}}
                    },
                    'fields': 'fontSize,foregroundColor'
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {'segmentId': footer_id, 'startIndex': 0, 'endIndex': len(footer_text)},
                    'paragraphStyle': {'alignment': 'START'},
                    'fields': 'alignment'
                }
            }
        ]
        
        # Execute the batch update to insert text and apply styles
        docs_service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests}
        ).execute()
        
    except HttpError as error:
        print(f"An error occurred: {error}")


def insert_numbered_texts_after_table(docs_service, document_id, texts, start_index_after_table):
    """
    Inserts a paragraph and numbered texts after the third table in a Google Doc.

    :param docs_service: Google Docs API service instance.
    :param document_id: Document ID.
    :param texts: List of texts to be inserted as numbered items.
    :param start_index_after_table: Index after which to start inserting the texts.
    """
    requests = []
    requests.append({
        'insertText': {
            'location': {'index': start_index_after_table},
            'text': "\n"  # Adding extra newline for spacing
        }
    })
    start_index_after_table+=1
    # Insert "Observation" as a paragraph
    requests.append({
        'insertText': {
            'location': {'index': start_index_after_table},
            'text': "4. Observation\n"  # Adding extra newline for spacing
        }
    })
    requests.append({
        'updateTextStyle': {
            'range': {
                'startIndex': start_index_after_table,  # Start at the beginning of the title
                'endIndex': start_index_after_table +len("4. Observation")+1  # End at the end of the title
            },
            'textStyle': {
                'bold': True,
                'foregroundColor': {
                    'color': {
                         'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                    }
                }
            },
            'fields': 'bold,foregroundColor'
        }
    })
    start_index_after_table += len("4. Observation\n")+1  # Update index for next insertion

    # Insert the texts as manually numbered list items with indentation
    for i, text in enumerate(texts, start=1):
        numbered_text = f"{i}. {text}\n"  # Prepend manual numbering
        requests.append({
            'insertText': {
                'location': {'index': start_index_after_table},
                'text': numbered_text
            }
        })

        # Apply font size 12 to the entire numbered text
        requests.append({
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index_after_table,
                    'endIndex': start_index_after_table + len(numbered_text)
                },
                'textStyle': {'fontSize': {'magnitude': 12, 'unit': 'PT'}},
                'fields': 'fontSize'
            }
        })

        # Add left indent to match the bullet style
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': start_index_after_table,
                    'endIndex': start_index_after_table + len(numbered_text)
                },
                'paragraphStyle': {
                    'indentStart': {
                        'magnitude': 18,  # Adjust this value as needed
                        'unit': 'PT'
                    },
                    'indentFirstLine': {
                        'magnitude': 18,  # Adjust this value to match indentStart
                        'unit': 'PT'
                    }
                },
                'fields': 'indentStart,indentFirstLine'
            }
        })

        # Update the start index for the next insertion
        start_index_after_table += len(numbered_text)

    # Execute batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Inserted numbered texts and 'Observation' successfully.")


def insert_numbered_texts_after_table_2(docs_service, document_id, texts, start_index_after_table):
    """
    Inserts a paragraph and numbered texts after the third table in a Google Doc.

    :param docs_service: Google Docs API service instance.
    :param document_id: Document ID.
    :param texts: List of texts to be inserted as numbered items.
    :param start_index_after_table: Index after which to start inserting the texts.
    """
    requests = []
    requests.append({
        'insertText': {
            'location': {'index': start_index_after_table},
            'text': "\n"  # Adding extra newline for spacing
        }
    })
    start_index_after_table += 1
    # Insert "Observation" as a paragraph
    requests.append({
        'insertText': {
            'location': {'index': start_index_after_table},
            'text': "4. Observation\n"  # Adding extra newline for spacing
        }
    })
    requests.append({
        'updateTextStyle': {
            'range': {
                'startIndex': start_index_after_table,  # Start at the beginning of the title
                'endIndex': start_index_after_table + len("4. Observation") + 1  # End at the end of the title
            },
            'textStyle': {
                'bold': True,
                'foregroundColor': {
                    'color': {
                        'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                    }
                }
            },
            'fields': 'bold,foregroundColor'
        }
    })
    start_index_after_table += len("4. Observation\n")  # Update index for next insertion

    # Apply PT Sans font and line spacing to the "Observation" paragraph
    requests.append({
        'updateTextStyle': {
            'range': {
                'startIndex': start_index_after_table - len("4. Observation\n") - 1,
                'endIndex': start_index_after_table
            },
            'textStyle': {
                'bold': True,
                'weightedFontFamily': {'fontFamily': 'PT Sans'},
                'fontSize': {'magnitude': 12, 'unit': 'PT'}
            },
            'fields': 'bold,weightedFontFamily,fontSize'
        }
    })
    requests.append({
        'updateParagraphStyle': {
            'range': {
                'startIndex': start_index_after_table - len("4. Observation\n") - 1,
                'endIndex': start_index_after_table
            },
            'paragraphStyle': {'lineSpacing': 150},
            'fields': 'lineSpacing'
        }
    })

    # Insert the texts as manually numbered list items with indentation
    for i, text in enumerate(texts, start=1):
        numbered_text = f"{i}. {text}\n"  # Prepend manual numbering
        requests.append({
            'insertText': {
                'location': {'index': start_index_after_table},
                'text': numbered_text
            }
        })

        # Apply font size 12 and PT Sans to the entire numbered text
        requests.append({
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index_after_table,
                    'endIndex': start_index_after_table + len(numbered_text)
                },
                'textStyle': {
                    'fontSize': {'magnitude': 12, 'unit': 'PT'},
                    'weightedFontFamily': {'fontFamily': 'PT Sans'}
                },
                'fields': 'fontSize,weightedFontFamily'
            }
        })

        # Add line spacing and left indent to match the bullet style
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': start_index_after_table,
                    'endIndex': start_index_after_table + len(numbered_text)
                },
                'paragraphStyle': {
                    'indentStart': {'magnitude': 18, 'unit': 'PT'},
                    'indentFirstLine': {'magnitude': 18, 'unit': 'PT'},
                    'lineSpacing': 150  # Set line spacing to 1.5
                },
                'fields': 'indentStart,indentFirstLine,lineSpacing'
            }
        })

        # Update the start index for the next insertion
        start_index_after_table += len(numbered_text)

    # Execute batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Inserted numbered texts and 'Observation' successfully.")


#### Final Code for insertion of static texts.

def format_paragraphs_and_numbered_bullets(docs_service, document_id, table_rows, end_index):
    """
    Insert and format the first two rows as bolded paragraphs, with text before ':' bold.
    Insert the last four rows as numbered list items manually, with font size 12 for all text.
    """
    requests = []
    requests.append({
            'insertText': {
                'location': {'index': end_index},
                'text': "\n"  # Add newline to separate paragraphs
            }
        })
    
    end_index += 1
    
    # Insert the first two rows as paragraphs with specified styles
    for row_text in table_rows[:2]:
        # Insert text
        requests.append({
            'insertText': {
                'location': {'index': end_index},
                'text': row_text + "\n\n"  # Add newline to separate paragraphs
            }
        })
        requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + len(row_text)
                    },
                    'textStyle': {
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'}
                    },
                    'fields': 'fontSize,weightedFontFamily'
                }
            })

        # Apply bold and PT Sans font to text before ":" if present, and set font size and line spacing
        if ":" in row_text:
            colon_index = row_text.index(":")
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + colon_index
                    },
                    'textStyle': {'bold': True,
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'foregroundColor': {
                            'color': {
                                'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'foregroundColor,bold,fontSize,weightedFontFamily'
                }
            })
        
        # Apply font size, font type, and line spacing to the entire paragraph text
        else:
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + len(row_text)
                    },
                    'textStyle': {
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'}
                    },
                    'fields': 'fontSize,weightedFontFamily'
                }
            })

        # Apply line spacing to the paragraph
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': end_index,
                    'endIndex': end_index + len(row_text)
                },
                'paragraphStyle': {
                    'lineSpacing': 150  # Set line spacing to 1.5
                },
                'fields': 'lineSpacing'
            }
        })

        # Update end_index for each insertion
        end_index += len(row_text) + 2  # Including the two newline characters
    
    # Insert the last four rows as manually numbered list items with font settings
    for i, row_text in enumerate(table_rows[2:], start=1):
        numbered_text = f"{i}. {row_text}\n"  # Prepend manual numbering
        requests.append({
            'insertText': {
                'location': {'index': end_index},
                'text': numbered_text
            }
        })

        # Apply font size, font type, and line spacing to the entire numbered text
        requests.append({
            'updateTextStyle': {
                'range': {
                    'startIndex': end_index,
                    'endIndex': end_index + len(numbered_text)
                },
                'textStyle': {
                    'fontSize': {'magnitude': 12, 'unit': 'PT'},
                    'weightedFontFamily': {'fontFamily': 'PT Sans'}
                },
                'fields': 'fontSize,weightedFontFamily'
            }
        })
        
        # Apply line spacing and indentation to match the numbered item style
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': end_index,
                    'endIndex': end_index + len(numbered_text)
                },
                'paragraphStyle': {
                    'indentStart': {
                        'magnitude': 18,
                        'unit': 'PT'
                    },
                    'indentFirstLine': {
                        'magnitude': 18,
                        'unit': 'PT'
                    },
                    'lineSpacing': 150  # Set line spacing to 1.5
                },
                'fields': 'indentStart,indentFirstLine,lineSpacing'
            }
        })

        # Apply bold style to text before the colon if a colon exists
        colon_position = numbered_text.find(':')
        if colon_position != -1:
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + colon_position
                    },
                    'textStyle': {'bold': True},
                    'fields': 'bold'
                }
            })
        
        # Update end_index for the next insertion
        end_index += len(numbered_text)

    # Execute batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Formatted paragraphs and manually numbered bullets inserted successfully.")

def format_paragraphs_and_numbered_bullets_youtube(docs_service, document_id, table_rows, end_index):
    """
    Insert and format the first two rows as bolded paragraphs, with text before ':' bold.
    Insert the last four rows as numbered list items manually, with font size 12 for all text.
    """
    requests = []
    requests.append({
            'insertText': {
                'location': {'index': end_index},
                'text': "\n"  # Add newline to separate paragraphs
            }
        })
    
    end_index += 1
    
    # Insert the first two rows as paragraphs with specified styles
    for row_text in table_rows:
        # Insert text
        requests.append({
            'insertText': {
                'location': {'index': end_index},
                'text': row_text + "\n\n"  # Add newline to separate paragraphs
            }
        })

        # Apply bold and PT Sans font to text before ":" if present, and set font size and line spacing
        if ":" in row_text:
            colon_index = row_text.index(":")
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + colon_index
                    },
                    'textStyle': {'bold': True,
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'foregroundColor': {
                            'color': {
                                'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'foregroundColor,bold,fontSize,weightedFontFamily'
                }
            })
        
        # Apply font size, font type, and line spacing to the entire paragraph text
        else:
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + len(row_text)
                    },
                    'textStyle': {
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'}
                    },
                    'fields': 'fontSize,weightedFontFamily'
                }
            })

        # Apply line spacing to the paragraph
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': end_index,
                    'endIndex': end_index + len(row_text)
                },
                'paragraphStyle': {
                    'lineSpacing': 150  # Set line spacing to 1.5
                },
                'fields': 'lineSpacing'
            }
        })

        # Update end_index for each insertion
        end_index += len(row_text) + 2  # Including the two newline characters

    # Execute batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Formatted paragraphs and manually numbered bullets inserted successfully.")


def insert_styled_paragraph(docs_service, document_id, text,index):
    """
    Inserts a paragraph with specific styling: #073763 text color, 14 font size, center-aligned, and bold.

    :param docs_service: Google Docs API service instance.
    :param document_id: The ID of the document.
    :param text: The text to insert as a paragraph.
    """
    requests = [
        # Insert the text as a new paragraph
        {
            'insertText': {
                'location': {'index': index},  # Insert at the beginning; modify as needed
                'text': text + "\n"
            }
        },
        # Center-align the paragraph
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': index,
                    'endIndex': index + len(text)
                },
                'paragraphStyle': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        },
        # Apply styling to the inserted text
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': index,
                    'endIndex': index + len(text)
                },
                'textStyle': {
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 7 / 255.0,
                                'green': 55 / 255.0,
                                'blue': 99 / 255.0
                            }
                        }
                    },
                    'fontSize': {
                        'magnitude': 14,
                        'unit': 'PT'
                    },
                    'bold': True
                },
                'fields': 'foregroundColor,fontSize,bold'
            }
        }]
    # Execute the batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Styled paragraph inserted successfully.")

def insert_styled_paragraph_index(docs_service, document_id, text,index):
    """
    Inserts a paragraph with specific styling: #073763 text color, 14 font size, center-aligned, and bold.

    :param docs_service: Google Docs API service instance.
    :param document_id: The ID of the document.
    :param text: The text to insert as a paragraph.
    """
    requests = [
        # Insert the text as a new paragraph
        {
            'insertText': {
                'location': {'index': index},  # Insert at the beginning; modify as needed
                'text': text + "\n"
            }
        },
        # Apply styling to the inserted text
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': index,
                    'endIndex': index + len(text)
                },
                'textStyle': {
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 7 / 255.0,
                                'green': 55 / 255.0,
                                'blue': 99 / 255.0
                            }
                        }
                    },
                    'fontSize': {
                        'magnitude': 14,
                        'unit': 'PT'
                    },
                    'bold': True
                },
                'fields': 'foregroundColor,fontSize,bold'
            }
        },
        # Center-align the paragraph
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': index,
                    'endIndex': index + len(text)
                },
                'paragraphStyle': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        }
    ]

    # Execute the batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Styled paragraph inserted successfully.")

### Insertion Index for the paragraph data
def get_insertion_index(docs_service, document_id):
    """
    Retrieve the index to insert a new element in the Google Docs document.
    
    :param docs_service: Google Docs API service instance.
    :param document_id: The ID of the Google Docs document.
    :return: The index where new content can be inserted.
    """
    # Fetch the document
    document = docs_service.documents().get(documentId=document_id).execute()

    # Find the end index of the content
    body_content = document['body']['content']
    
    # Start with the last element's endIndex
    if body_content:
        last_element = body_content[-1]
        insertion_index = last_element['endIndex']  # Get the endIndex of the last element
    else:
        insertion_index = 1  # Set to 1 to insert at the beginning if the document is empty
    
    return insertion_index

#### Code for finding the Last valid index
def find_last_valid_index(docs_service, document_id):
    """
    Finds the last valid index in the document to insert content after.
    """
    # Get the document content
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content', [])

    # Find the end index of the last element in the document
    if content:
        last_element = content[-1]
        if 'endIndex' in last_element:
            return last_element['endIndex'] - 1  # Subtract 1 to avoid going out of bounds

    return None

### Code for insertion of Data
def insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows,insert_index):
    # Step 1: Insert an empty table with the correct dimensions
    num_columns = len(headers)
    num_rows = len(table_rows) + 1  # +1 for the header row
    
    insert_table_request = {
        'insertTable': {
            'rows': num_rows,
            'columns': num_columns,
            'location': {
                'index': insert_index  # Insert table at the start (you can modify the index)
            }
        }
    }
    
    # Send the table insertion request
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [insert_table_request]}).execute()

    # Step 2: Get the document content to find the start index of the table
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content')

    # Step 3: Find the indices of the paragraphs inside the table cells
    table_cells = []
    table_found = False
    print("Hello")
    for element in content:
        if 'table' in element and 'startIndex' in element and element['startIndex'] >= insert_index:
            table_found = True  # We've located the newly inserted table
            for row in element['table']['tableRows']:
                for cell in row['tableCells']:
                    if 'content' in cell and len(cell['content']) > 0:
                        cell_content = cell['content'][0]
                        if 'paragraph' in cell_content and 'elements' in cell_content['paragraph']:
                            elements = cell_content['paragraph']['elements']
                            if len(elements) > 0 and 'startIndex' in elements[0]:
                                start_index = elements[0]['startIndex']
                                table_cells.append(start_index)
            break  # Once we've found the table and collected the indices, we can stop

    if not table_found:
        raise Exception(f"Could not find the table inserted at index {insert_index}")
    
    # Step 4: Prepare the requests to insert the headers and data
    populate_requests = []
    current_offset = 0  # Tracks how much text has been inserted so far to adjust future insertions

    # Insert headers
    for col_index, header in enumerate(headers):
        # Insert text at the corresponding paragraph index inside the table cell
        if not header:  # Ensure header is not None or empty
            header = ' '  # Insert a space if it's empty
        start_index = table_cells[col_index] + current_offset  # Adjust index with current offset
        populate_requests.append({
            'insertText': {
                'text': header,
                'location': {
                    'index': start_index  # Insert header text at the correct index
                }
            }
        })
        populate_requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': start_index + len(header)
                    },
                    'textStyle': {
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'fontSize': {
                            'magnitude': 12,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'weightedFontFamily,fontSize'
                }
            })
        current_offset += len(header)  # Update the offset by the length of the inserted 
    # Insert data rows
    for row_index, row in enumerate(table_rows):
        for col_index, cell_value in enumerate(row):
            # Handle empty or None values by replacing them with a space
            if not cell_value:
                cell_value = ' '
            
            # URL pattern for detecting links
            url_pattern = re.compile(r'(https?://[^\s]+)')

            # Calculate the correct index for each data cell
            table_cell_index = num_columns + row_index  * num_columns + col_index
            start_index = table_cells[table_cell_index] + current_offset  # Adjust with current offset

            populate_requests.append({
                'insertText': {
                    'text': str(cell_value),
                    'location': {
                        'index': start_index  # Insert data text at the correct index
                    }
                }
            })
            populate_requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': start_index + len(str(cell_value))
                    },
                    'textStyle': {
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'fontSize': {
                            'magnitude': 12,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'weightedFontFamily,fontSize'
                }
            })
            if url_pattern.match(cell_value):
                update_text_style_request = {
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': start_index + len(cell_value)
                        },
                        'textStyle': {
                            'link': {
                                'url': cell_value
                            }
                        },
                        'fields': 'link'
                    }
                }
                populate_requests.append(update_text_style_request)
            else:
                pass
            current_offset += len(str(cell_value))  # Update the offset by the length of the inserted text

    # Step 5: Send the request to populate the table
       # Step 5: Send the request to populate the table
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
    print(f"Inserted table from dataframe into Google Docs with ID: {document_id}")

#### Code for inserting the tables from the dictionary
def insert_texts_and_multiple_tables_into_docs(docs_service, document_id, df_dict,instagram=0):
    if instagram==1:
        start=0
        end=4
    else:
        start=0
        end=3

    for title, df in islice(df_dict.items(),start, end):   
        ##Move cursor to the next line
        end_index=get_insertion_index(docs_service, document_id)
        reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
        ###################

        # Find the current last index to insert content after the previous table and title
        insert_index = find_last_valid_index(docs_service, document_id)
        # print(insert_index)
        formatted_title = title.replace("_", " ")
        if title=="Week_on_Week_Change_in_Performance":
            title_request = {
                'insertText': {
                    'location': {'index': insert_index},
                    'text': f"{formatted_title}\n"  # Adding blank lines before and after the title
                }
            }
            # Step 2: Apply formatting (bold and color)
            update_style_request = {
                'updateTextStyle': {
                    'range': {
                        'startIndex': insert_index ,  # Start after the newline
                        'endIndex': insert_index+len(formatted_title)  # End at the end of the title
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {
                            'magnitude': 12,
                            'unit': 'PT'
                        },
                    'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'foregroundColor': {
                            'color': {
                                 'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'bold,foregroundColor,weightedFontFamily,fontSize'
                }
            }
            page_break_request = {
                'insertPageBreak': {
                    'location': {'index': insert_index }
                }
            }
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request,page_break_request]}).execute()
        else:
            title_request = {
                'insertText': {
                    'location': {'index': insert_index},
                    'text': f"{formatted_title}\n"  # Adding blank lines before and after the title
                }
            }
            # Step 2: Apply formatting (bold and color)
            update_style_request = {
                'updateTextStyle': {
                    'range': {
                        'startIndex': insert_index ,  # Start after the newline
                        'endIndex': insert_index+ len(formatted_title)  # End at the end of the title
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {
                            'magnitude': 12,
                            'unit': 'PT'
                        },
                    'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'foregroundColor': {
                            'color': {
                                 'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'bold,foregroundColor,weightedFontFamily,fontSize'
                }
            }
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request]}).execute()

        # Update the insert index after inserting the title
        insert_index = find_last_valid_index(docs_service, document_id)

        # Prepare headers and rows for the table data
        headers = df.columns.tolist()
        table_rows = df.values.tolist()
        
        # Insert the table at the new index after the title
        insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)

### Function to insert the table into the document [Youtube]
def insert_texts_and_multiple_tables_into_docs_youtube(docs_service, document_id, df_dict,start,end):
    """
        df_dict: should be an itertool object, eg: islice(df_dict,start,end)
    """
    for title, df in islice(df_dict.items(),start,end):   
        ##Move cursor to the next line
        end_index=get_insertion_index(docs_service, document_id)
        reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
        ###################

        # Find the current last index to insert content after the previous table and title
        insert_index = find_last_valid_index(docs_service, document_id)
        # print(insert_index)
        formatted_title = title.replace("_", " ")
        title_request = {
                'insertText': {
                    'location': {'index': insert_index},
                    'text': f"{formatted_title}\n"  # Adding blank lines before and after the title
                }
            }
            # Step 2: Apply formatting (bold and color)
        update_style_request = {
                'updateTextStyle': {
                    'range': {
                        'startIndex': insert_index ,  # Start after the newline
                        'endIndex': insert_index+ len(formatted_title)  # End at the end of the title
                    },
                    'textStyle': {
                        'bold': True,
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'fontSize': {
                            'magnitude': 12,
                            'unit': 'PT'
                        },
                        'foregroundColor': {
                            'color': {
                                 'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'bold,foregroundColor,weightedFontFamily,fontSize'
                }
            }
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request]}).execute()
        # Update the insert index after inserting the title
        insert_index = find_last_valid_index(docs_service, document_id)

        # Prepare headers and rows for the table data
        headers = df.columns.tolist()
        table_rows = df.values.tolist()
        
        # Insert the table at the new index after the title
        insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)

def reset_format_and_move_to_next_line(docs_service, document_id, end_index):
    """
    Inserts a new line in the document and resets the formatting to default settings.
    
    :param docs_service: Google Docs API service instance.
    :param document_id: The ID of the Google Docs document.
    :param end_index: The current index in the document to insert a new line.
    """
    requests = []

    # Step 1: Insert a newline to move to the next line
    requests.append({
        'insertText': {
            'location': {'index': end_index},
            'text': "\n"  # Add newline to move the cursor to the next line
        }
    })
    
    # Update end_index after inserting the newline
    end_index += 1

    # Step 2: Reset formatting to default settings
    reset_format_request = {
        'updateParagraphStyle': {
            'range': {
                'startIndex': end_index,
                'endIndex': end_index + 1  # Range includes the newline
            },
            'paragraphStyle': {
                'indentStart': {'magnitude': 0, 'unit': 'PT'},
                'indentFirstLine': {'magnitude': 0, 'unit': 'PT'},
                'spaceAbove': {'magnitude': 0, 'unit': 'PT'},
                'spaceBelow': {'magnitude': 0, 'unit': 'PT'}
            },
            'fields': 'indentStart,indentFirstLine,spaceAbove,spaceBelow'
        }
    }

    requests.append(reset_format_request)

    # Execute batch update
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Cursor moved to the next line and formatting reset to default settings.")

#### Code for inserting the tables from the dictionary
def insert_texts_and_multiple_tables_into_docs_last_5(docs_service, document_id, df_dict,instagram=0,dataframes_reels=None):
    if instagram==1:
        start=5
        end=None
        text="5. name on Instagram:"
    else:
        start=4
        end=None
        text="5. name on Facebook:"
    time.sleep(1)
    insert_index = find_last_valid_index(docs_service, document_id)
    title_request = {
    'insertText': {
        'location': {'index': insert_index},
        'text': f"{text}\n"  # Adding blank lines before and after the title
    }
    }
    update_style_request = {
        'updateTextStyle': {
            'range': {
                'startIndex': insert_index , # +1 # Start after the newline
                'endIndex': insert_index+ len(text)+1  # End at the end of the title
            },
            'textStyle': {
                'bold': True,
                'fontSize': {'magnitude': 12, 'unit': 'PT'},
                'weightedFontFamily': {'fontFamily': 'PT Sans'},
                'foregroundColor': {
                    'color': {
                        'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                    }
                }
            },
            'fields': 'bold,foregroundColor,fontSize,weightedFontFamily'
        }
    }
    page_break_request = {
            'insertPageBreak': {
                'location': {'index': insert_index}
            }
        }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request,page_break_request]}).execute()
    ### Manuel insertion for the 4th table.
    time.sleep(0.5)
    insert_index = find_last_valid_index(docs_service, document_id)
    title_request = {
        'insertText': {
            'location': {'index': insert_index},
            'text': "\nheader_1\n"  # Adding blank lines before and after the title
        }
    }
    update_title_style_request = {
        'updateTextStyle': {
            'range': {
                'startIndex': insert_index ,  # Start after the newline
                'endIndex': insert_index+ len("\nheader_1\n")  # End at the end of the title
            },
            'textStyle': {
                'bold': True,
                'fontSize': {'magnitude': 12, 'unit': 'PT'},
                'weightedFontFamily': {'fontFamily': 'PT Sans'},
                'foregroundColor': {
                    'color': {
                        'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                    }
                }
            },
            'fields': 'bold,foregroundColor,fontSize,weightedFontFamily'
        }
    }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_title_style_request]}).execute()
    insert_index = find_last_valid_index(docs_service, document_id)
    # Prepare headers and rows for the table data
    headers = df_dict["header_1"].columns.tolist()
    table_rows = df_dict["header_1"].values.tolist()
    
    # Insert the table at the new index after the title
    time.sleep(0.5)
    insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)

    if dataframes_reels:
        end_index=get_insertion_index(docs_service, document_id)
        reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
        insert_index = find_last_valid_index(docs_service, document_id)

        headers=dataframes_reels["BJYM"].columns.tolist()
        table_rows=dataframes_reels["BJYM"].values.tolist()
        insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)
    ######################################
        for title, df in islice(df_dict.items(),start, end):
            # Find the current last index to insert content after the previous table and title
            insert_index = find_last_valid_index(docs_service, document_id)
            # print(insert_index)
            formatted_title = title.replace("_", " ")
            title_request = {
                'insertText': {
                    'location': {'index': insert_index},
                    'text': f"{formatted_title}\n"  # Adding blank lines before and after the title
                }
            }
            update_style_request = {
                    'updateTextStyle': {
                    'range': {
                        'startIndex': insert_index ,  # Start after the newline
                        'endIndex': insert_index+ len(formatted_title)+1  # End at the end of the title
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'foregroundColor': {
                            'color': {
                                'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'bold,foregroundColor,fontSize,weightedFontFamily'
                }
            }
            page_break_request = {
                'insertPageBreak': {
                    'location': {'index': find_last_valid_index(docs_service, document_id)}
                }
            }
        
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request,page_break_request]}).execute()
            
            # Update the insert index after inserting the title
            insert_index = find_last_valid_index(docs_service, document_id)

            # Prepare headers and rows for the table data
            headers = df.columns.tolist()
            table_rows = df.values.tolist()
            
            # Insert the table at the new index after the title
            insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)
            time.sleep(1)
            end_index=get_insertion_index(docs_service, document_id)
            reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)

            headers=dataframes_reels[title].columns.tolist()
            table_rows=dataframes_reels[title].values.tolist()
            time.sleep(0.1)
            insert_index = find_last_valid_index(docs_service, document_id)
            time.sleep(0.3)
            insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)
    else:
        for title, df in islice(df_dict.items(),start, end):
            # Find the current last index to insert content after the previous table and title
            insert_index = find_last_valid_index(docs_service, document_id)
            # print(insert_index)
            formatted_title = title.replace("_", " ")
            title_request = {
                'insertText': {
                    'location': {'index': insert_index},
                    'text': f"{formatted_title}\n"  # Adding blank lines before and after the title
                }
            }
            update_style_request = {
                    'updateTextStyle': {
                    'range': {
                        'startIndex': insert_index ,  # Start after the newline
                        'endIndex': insert_index+ len(formatted_title)+1  # End at the end of the title
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {'fontFamily': 'PT Sans'},
                        'foregroundColor': {
                            'color': {
                                'rgbColor': {'red': 0.03, 'green': 0.22, 'blue': 0.39}
                            }
                        }
                    },
                    'fields': 'bold,foregroundColor,fontSize,weightedFontFamily'
                }
            }
            page_break_request = {
                'insertPageBreak': {
                    'location': {'index': find_last_valid_index(docs_service, document_id)}
                }
            }
        
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request,page_break_request]}).execute()
            
            # Update the insert index after inserting the title
            insert_index = find_last_valid_index(docs_service, document_id)

            # Prepare headers and rows for the table data
            headers = df.columns.tolist()
            table_rows = df.values.tolist()
            
            # Insert the table at the new index after the title
            insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows, insert_index)

def inspect_all_tables_content(docs_service, document_id):
    """
    Inspects content of all tables in the document and returns a list of tables with their cell data
    along with each table's starting index.
    """
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content', [])
    
    all_tables = []  # List to hold structured output of all tables
    table_start_indices = []  # List to store start index of each table

    for element in content:
        if 'table' in element:
            table_start_index = element['startIndex']  # Capture the start index of the table
            table_start_indices.append(table_start_index)  # Add to start indices list
            
            table_content = element['table']['tableRows']
            table_data = []  # To hold cell data for each table
            for row in table_content:
                row_data = []
                for cell in row['tableCells']:
                    start_index = cell.get('startIndex')
                    end_index = cell.get('endIndex')
                    text = ""
                    for content_element in cell.get('content', []):
                        if 'paragraph' in content_element:
                            for paragraph_element in content_element['paragraph']['elements']:
                                if 'textRun' in paragraph_element:
                                    text += paragraph_element['textRun']['content']
                    row_data.append({
                        'Start': start_index,
                        'End': end_index,
                        'Text': text.strip()  # Removing extra spaces
                    })
                table_data.append(row_data)
            all_tables.append(table_data)  # Append each table's data

    return all_tables, table_start_indices

def inspect_all_tables_content_2(docs_service, document_id):
    """
    Inspects content of all tables in the document and returns a list of tables with their cell data
    along with each table's starting index.
    """
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content', [])
    
    all_tables = []  # List to hold structured output of all tables
    table_start_indices = []  # List to store start index of each table

    for element in content:
        if 'table' in element:
            table_start_index = element['startIndex']  # Capture the start index of the table
            table_start_indices.append(table_start_index)  # Add to start indices list
            
            table_content = element['table']['tableRows']
            table_data = []  # To hold cell data for each table
            for row in table_content:
                row_data = []
                for cell in row['tableCells']:
                    start_index = cell.get('startIndex')
                    end_index = cell.get('endIndex')
                    text = ""
                    for content_element in cell.get('content', []):
                        if 'paragraph' in content_element:
                            for paragraph_element in content_element['paragraph']['elements']:
                                if 'textRun' in paragraph_element:
                                    text += paragraph_element['textRun']['content']
                    row_data.append({
                        'Start': start_index,
                        'End': end_index,
                        'Text': text ###.strip()  # Removing extra spaces
                    })
                table_data.append(row_data)
            all_tables.append(table_data)  # Append each table's data

    return all_tables, table_start_indices

### merge function
def merge_cells_in_table(docs_service, document_id, table_index, start_row, end_row, start_column, end_column):
    """
    Merges cells in a table within a specified range (rows and columns).
    :param docs_service: Google Docs API service
    :param document_id: ID of the Google Document
    :param table_index: Index of the table in the document
    :param start_row: Starting row index
    :param end_row: Ending row index
    :param start_column: Starting column index
    :param end_column: Ending column index
    """
    merge_request = {
        'mergeTableCells': {
            'tableRange': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': table_index},  # Specify the table index
                    'rowIndex': start_row,
                    'columnIndex': start_column
                },
                'rowSpan': end_row - start_row + 1,  # Row span to merge
                'columnSpan': end_column - start_column + 1  # Column span to merge
            }
        }
    }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [merge_request]}).execute()

def bold_table_headers(docs_service, document_id, all_tables):
    """
    Bolds header rows in each table dynamically.
    """
    requests = []
    
    for table in all_tables[:3]:
        # Define how many rows to consider as header
        num_header_rows = min(1, len(table))  # Example: max 3 header rows, can be adjusted

        for row in range(num_header_rows):  # Loop through header rows
            for cell in table[row]:  # Loop through each cell in the row
                start_index = cell['Start']
                end_index = cell['End']
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })
    for table in all_tables[3:]:
        # Define how many rows to consider as header
        num_header_rows = min(2, len(table))  # Example: max 3 header rows, can be adjusted

        for row in range(num_header_rows):  # Loop through header rows
            for cell in table[row]:  # Loop through each cell in the row
                start_index = cell['Start']
                end_index = cell['End']
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })

    # Execute batch update for bold formatting
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Bold formatting applied to headers.")


def bold_table_headers_insta(docs_service, document_id, all_tables):
    """
    Bolds header rows in each table dynamically.
    """
    requests = []
    
    for table in all_tables[12:16]:
        # Define how many rows to consider as header
        num_header_rows = min(1, len(table))  # Example: max 3 header rows, can be adjusted

        for row in range(num_header_rows):  # Loop through header rows
            for cell in table[row]:  # Loop through each cell in the row
                start_index = cell['Start']
                end_index = cell['End']
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })
    for table in all_tables[16:]:
        # Define how many rows to consider as header
        num_header_rows = min(2, len(table))  # Example: max 3 header rows, can be adjusted

        for row in range(num_header_rows):  # Loop through header rows
            for cell in table[row]:  # Loop through each cell in the row
                start_index = cell['Start']
                end_index = cell['End']
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })

    # Execute batch update for bold formatting
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Bold formatting applied to headers.")

def bold_table_headers_youtube(docs_service, document_id, all_tables):
    """
    Bolds header rows in each table dynamically.
    """
    requests = []
    
    for table in all_tables:
        # Define how many rows to consider as header
        num_header_rows = min(1, len(table))  # Example: max 3 header rows, can be adjusted

        for row in range(num_header_rows):  # Loop through header rows
            for cell in table[row]:  # Loop through each cell in the row
                start_index = cell['Start']
                end_index = cell['End']
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })
    
    # Execute batch update for bold formatting
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Bold formatting applied to headers.")


def center_align_all_table_text(docs_service, document_id, all_tables):
    """
    Center-aligns text in all cells of each table.
    """
    requests = []

    for table in all_tables:
        for row in table:
            for cell in row:
                start_index = cell['Start']
                # print(start_index)
                end_index = cell['End']
                requests.append({
                    'updateParagraphStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'paragraphStyle': {
                            'alignment': 'CENTER'
                        },
                        'fields': 'alignment'
                    }
                })

    # Execute batch update for center alignment
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Center alignment applied to all table cells.")

def set_column_widths(docs_service, document_id, table_start_indices, column_widths_by_table):
    populate_requests = []
    
    # Iterate through each table and set the specified column widths
    for table_index, column_widths in zip(table_start_indices, column_widths_by_table):
        for col_index, width in enumerate(column_widths):
            # Update each column width
            populate_requests.append({
                'updateTableColumnProperties': {
                    'tableStartLocation': {
                        'index': table_index,  # Starting index of the table
                    },
                    'columnIndices': [col_index],  # Index of the column to update
                    'tableColumnProperties': {
                        'width': {
                            'magnitude': width,  # Column width in points
                            'unit': 'PT'
                        },
                        'widthType': 'FIXED_WIDTH'
                    },
                    'fields': 'width,widthType'
                }
            })
    
    # Execute the request in batch
    if populate_requests:
        try:
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
            print("Column widths set successfully.")
        except Exception as e:
            print(f"An error occurred: {e}")

def apply_colors_to_all_table_headers_till_3rd(docs_service, document_id, all_tables, table_start_indices):
    """
    Applies background colors to header rows in each table, leveraging each table's starting index dynamically.
    """
    populate_requests = []
    # colors = [
    #     {"red": 1.0, "green": 0.6, "blue": 0.6},  # Light red for the first row (header)
    #     {"red": 0.6, "green": 0.6, "blue": 1.0},  # Light blue for the second row
    #     {"red": 0.8, "green": 0.8, "blue": 0.8}   # Light gray for the third row
    # ]
    colors = [
    {"red": 1.0, "green": 0.6, "blue": 0.0}  # Color converted from #ff9900
    ]

    for table_index, table in enumerate(all_tables):
        table_start = table_start_indices[table_index]  # Starting index for this table in the document
        num_columns = len(table[0]) if table else 0  # Assume uniform columns per row

        # Loop through first 3 rows to apply colors
        for row_index in range(min(1, len(table))):  # Only up to 3 rows if they exist
            populate_requests.append({
                'updateTableCellStyle': {
                    'tableRange': {
                        'tableCellLocation': {
                            'tableStartLocation': {
                                'index': table_start  # Use dynamic table start index
                            },
                            'rowIndex': row_index,
                            'columnIndex': 0,
                        },
                        'rowSpan': 1,
                        'columnSpan': num_columns
                    },
                    'tableCellStyle': {
                        'backgroundColor': {
                            'color': {
                                'rgbColor': colors[row_index]
                            }
                        }
                    },
                    'fields': 'backgroundColor'
                }
            })

    # Send batch update request
    if populate_requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
        print("Color styling applied to header rows.")

def apply_colors_to_all_table_headers_from_4th(docs_service, document_id, all_tables, table_start_indices):
    """
    Applies background colors to header rows in each table, leveraging each table's starting index dynamically.
    """
    populate_requests = []
    # colors = [
    #     {"red": 1.0, "green": 0.6, "blue": 0.6},  # Light red for the first row (header)
    #     {"red": 0.6, "green": 0.6, "blue": 1.0},  # Light blue for the second row
    #     {"red": 0.8, "green": 0.8, "blue": 0.8}   # Light gray for the third row
    # ]
    colors = [
    {"red": 1.0, "green": 0.6, "blue": 0.0},  # Color converted from #ff9900
    {"red": 1.0, "green": 0.6, "blue": 0.0} 
    ]

    for table_index, table in enumerate(all_tables):
        table_start = table_start_indices[table_index]  # Starting index for this table in the document
        num_columns = len(table[0]) if table else 0  # Assume uniform columns per row

        # Loop through first 3 rows to apply colors
        for row_index in range(min(2, len(table))):  # Only up to 3 rows if they exist
            populate_requests.append({
                'updateTableCellStyle': {
                    'tableRange': {
                        'tableCellLocation': {
                            'tableStartLocation': {
                                'index': table_start  # Use dynamic table start index
                            },
                            'rowIndex': row_index,
                            'columnIndex': 0,
                        },
                        'rowSpan': 1,
                        'columnSpan': num_columns
                    },
                    'tableCellStyle': {
                        'backgroundColor': {
                            'color': {
                                'rgbColor': colors[row_index]
                            }
                        }
                    },
                    'fields': 'backgroundColor'
                }
            })

    # Send batch update request
    if populate_requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
        print("Color styling applied to header rows.")

### column coloring function

def apply_color_to_one_column_in_all_rows(docs_service, document_id, all_tables, table_start_indices, column_to_color):
    """
    Applies background color to a single specified column across all rows in each table.

    Parameters:
    - docs_service: Google Docs API service instance.
    - document_id: The ID of the Google Docs document.
    - all_tables: A list of tables with their cell content.
    - table_start_indices: A list of starting indices for each table.
    - column_to_color: The index of the column (0-based) to color across all rows.
    """
    populate_requests = []
    color = {"red": 1.0, "green": 0.6, "blue": 0.0}  # Color example (e.g., orange #ff9900)

    for table_index, table in enumerate(all_tables):
        table_start = table_start_indices[table_index]  # Starting index for this table in the document
        num_rows = len(table)  # Get total row count for the current table
        # print(table[0])
        # Apply color to the specified column across all rows
        if column_to_color < len(table[0]):  # Ensure column is within bounds
            populate_requests.append({
                'updateTableCellStyle': {
                    'tableRange': {
                        'tableCellLocation': {
                            'tableStartLocation': {
                                'index': table_start  # Use dynamic table start index
                            },
                            'rowIndex': 0,  # Start at the first row
                            'columnIndex': column_to_color  # Target the specified column
                        },
                        'rowSpan': num_rows,  # Apply to all rows
                        'columnSpan': 1  # Single column
                    },
                    'tableCellStyle': {
                        'backgroundColor': {
                            'color': {
                                'rgbColor': color
                            }
                        }
                    },
                    'fields': 'backgroundColor'
                }
            })

    # Send batch update request if there are any requests
    if populate_requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
        print("Color styling applied to specified column across all rows.")


def color_cells_based_on_arrows(docs_service, document_id, table_index):
    # Fetch document content
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body', {}).get('content', [])
    
    # List to store the requests for coloring text
    requests = []

    # Identify the table and table rows
    current_table_index = 0
    for element in content:
        if 'table' in element:
            if current_table_index == table_index:
                table = element['table']
                
                # Iterate over rows, skipping the header row (first row, i.e., index 0)
                for row_index, row in enumerate(table['tableRows']):
                    if row_index == 0:  # Skip the header row
                        continue
                    
                    # Iterate over cells in each row
                    for cell in row['tableCells']:
                        cell_content = cell['content']

                        # Check for paragraph elements inside the cell
                        if cell_content and 'paragraph' in cell_content[0]:
                            elements = cell_content[0]['paragraph']['elements']
                            for element in elements:
                                text_content = element.get('textRun', {}).get('content', '')

                                # Color entire text based on arrow presence
                                if '▲' in text_content:
                                    # Green color for up-arrow
                                    requests.append({
                                        'updateTextStyle': {
                                            'range': {
                                                'startIndex': element['startIndex'],
                                                'endIndex': element['endIndex'],
                                            },
                                            'textStyle': {
                                                'foregroundColor': {
                                                    'color': {
                                                        'rgbColor': {
                                                            'red': 0.0,
                                                            'green': 0.6,
                                                            'blue': 0.0
                                                        }
                                                    }
                                                }
                                            },
                                            'fields': 'foregroundColor'
                                        }
                                    })
                                elif '▼' in text_content:
                                    # Red color for down-arrow
                                    requests.append({
                                        'updateTextStyle': {
                                            'range': {
                                                'startIndex': element['startIndex'],
                                                'endIndex': element['endIndex'],
                                            },
                                            'textStyle': {
                                                'foregroundColor': {
                                                    'color': {
                                                        'rgbColor': {
                                                            'red': 1.0,
                                                            'green': 0.0,
                                                            'blue': 0.0
                                                        }
                                                    }
                                                }
                                            },
                                            'fields': 'foregroundColor'
                                        }
                                    })
            
            # Increment the table index to move to the next table
            current_table_index += 1

    # Batch update the document to apply the formatting requests
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Text styling applied based on arrow signs.")

def color_cells_based_on_texts(docs_service, document_id, table_index):
    # Fetch document content
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body', {}).get('content', [])
    
    # List to store the requests for coloring text
    requests = []

    # Identify the table and table rows
    current_table_index = 0
    for element in content:
        if 'table' in element:
            if current_table_index == table_index:
                table = element['table']
                
                # Iterate over rows, skipping the header row (first row, i.e., index 0)
                for row_index, row in enumerate(table['tableRows']):
                    if row_index == 0:  # Skip the header row
                        continue
                    
                    # Iterate over cells in each row
                    for cell in row['tableCells']:
                        cell_content = cell['content']

                        # Check for paragraph elements inside the cell
                        if cell_content and 'paragraph' in cell_content[0]:
                            elements = cell_content[0]['paragraph']['elements']
                            for element in elements:
                                text_content = element.get('textRun', {}).get('content', '')

                                # Color entire text based on arrow presence
                                if ('Verified' in text_content) or ('Active' in text_content):
                                    # Green color for up-arrow
                                    requests.append({
                                        'updateTextStyle': {
                                            'range': {
                                                'startIndex': element['startIndex'],
                                                'endIndex': element['endIndex'],
                                            },
                                            'textStyle': {
                                                'foregroundColor': {
                                                    'color': {
                                                        'rgbColor': {
                                                            'red': 0.0,
                                                            'green': 0.6,
                                                            'blue': 0.0
                                                        }
                                                    }
                                                }
                                            },
                                            'fields': 'foregroundColor'
                                        }
                                    })
                                elif ('Unverified' in text_content) or ('Inactive' in text_content) :
                                    # Red color for down-arrow
                                    requests.append({
                                        'updateTextStyle': {
                                            'range': {
                                                'startIndex': element['startIndex'],
                                                'endIndex': element['endIndex'],
                                            },
                                            'textStyle': {
                                                'foregroundColor': {
                                                    'color': {
                                                        'rgbColor': {
                                                            'red': 1.0,
                                                            'green': 0.0,
                                                            'blue': 0.0
                                                        }
                                                    }
                                                }
                                            },
                                            'fields': 'foregroundColor'
                                        }
                                    })
            
            # Increment the table index to move to the next table
            current_table_index += 1

    # Batch update the document to apply the formatting requests
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Text styling applied based on Texts.")

def color_cells_based_on_texts_2(docs_service, document_id, table_index):
    # Fetch document content
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body', {}).get('content', [])
    
    # List to store the requests for coloring cells
    requests = []
    current_table_index = 0
    
    for element in content:
        if 'table' in element:
            # Check if we are at the target table by index
            if current_table_index == table_index:
                table = element['table']
                table_start_index=element['startIndex']
                
                # Iterate over rows, skipping the header row (index 0)
                for row_index, row in enumerate(table['tableRows']):
                    if row_index == 0:  # Skip the header row
                        continue

                    # Iterate over each cell in the row
                    for col_index, cell in enumerate(row['tableCells']):
                        cell_content = cell['content']
                        
                        # Check cell content for specific text to set color
                        if cell_content and 'paragraph' in cell_content[0]:
                            elements = cell_content[0]['paragraph']['elements']
                            for element in elements:
                                text_content = element.get('textRun', {}).get('content', '')
                                
                                # Set color based on text in the cell
                                if 'Verified' in text_content or 'Active' in text_content:
                                    # Green color for verified or active status
                                    color = {'red': 0.0, 'green': 0.6, 'blue': 0.0}
                                elif 'Unverified' in text_content or 'Inactive' in text_content:
                                    # Red color for unverified or inactive status
                                    color = {'red': 1.0, 'green': 0.0, 'blue': 0.0}
                                else:
                                    continue  # Skip if no matching condition
                                    
                                print(element["startIndex"])
                                
                                # Create cell style update request
                                requests.append({
                                    'updateTableCellStyle': {
                                        'tableCellStyle': {
                                            'backgroundColor': {'color': {'rgbColor': color}}
                                        },
                                        'tableRange': {
                                            'tableCellLocation': {
                                                'tableStartLocation': {'index': table_start_index},
                                                'rowIndex': row_index,
                                                'columnIndex': col_index
                                            },
                                            'rowSpan': 1,
                                            'columnSpan': 1
                                        },
                                        'fields': 'backgroundColor'
                                    }
                                })
            current_table_index += 1  # Move to the next table

    # Execute batch request to color cells
    if requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Cell coloring applied based on texts.")

def apply_color_to_single_column(docs_service, document_id, table, table_start_index, column_to_color):
    """
    Applies background color to a single specified column across all rows in a specific table.
    
    Parameters:
    - docs_service: Google Docs API service instance.
    - document_id: The ID of the Google Docs document.
    - table: The target table with its cell content.
    - table_start_index: The starting index for the target table.
    - column_to_color: The index of the column (0-based) to color across all rows.
    """
    populate_requests = []
    color = {"red": 1.0, "green": 0.6, "blue": 0.0}  # Color example (e.g., orange #ff9900)
    num_rows = len(table)

    # Apply color to the specified column across all rows
    if column_to_color < len(table[0]):  # Ensure column is within bounds
        populate_requests.append({
            'updateTableCellStyle': {
                'tableRange': {
                    'tableCellLocation': {
                        'tableStartLocation': {
                            'index': table_start_index  # Use the specific table start index
                        },
                        'rowIndex': 0,  # Start at the first row
                        'columnIndex': column_to_color  # Target the specified column
                    },
                    'rowSpan': num_rows,  # Apply to all rows
                    'columnSpan': 1  # Single column
                },
                'tableCellStyle': {
                    'backgroundColor': {
                        'color': {
                            'rgbColor': color
                        }
                    }
                },
                'fields': 'backgroundColor'
            }
        })

    # Send batch update request if there are any requests
    if populate_requests:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
        print("Color styling applied to specified column across all rows.")

## Remove the Extra Space Before Table.
def remove_extra_space_before_table(service, document_id, table_start_index):
    try:
        # Add a delete request to remove one extra newline character before the table
        requests = [
            {
                'deleteContentRange': {
                    'range': {
                        'startIndex': table_start_index - 1,
                        'endIndex': table_start_index
                    }
                }
            }
        ]
        
        # Execute the delete request
        service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Removed extra space before the table.")

    except Exception as e:
        print(f"An error occurred: {e}")

def delete_and_replace_text(service, document_id, table_index, row_index, column_index, new_text):
    try:
        document = service.documents().get(documentId=document_id).execute()
        tables = [content for content in document['body']['content'] if 'table' in content]
        
        # Get the cell content
        cell_content = tables[table_index]['table']['tableRows'][row_index]['tableCells'][column_index]['content']
        # print(cell_content)

        # Collect delete requests
        requests = []
        for element in cell_content:
            if 'paragraph' in element:
                for paragraph_element in element['paragraph']['elements']:
                    text_run = paragraph_element.get('textRun')
                    if text_run and 'content' in text_run:
                        text = text_run['content']
                        print(f"Original text: '{text}'")
                        
                        # Skip newline-only elements
                        if text.strip() == "":
                            continue

                        start_index = paragraph_element['startIndex']
                        end_index = start_index + len(text)

                        # Only add a delete request if start and end are valid
                        if start_index is not None and end_index > start_index:
                            requests.append({
                                'deleteContentRange': {
                                    'range': {
                                        'startIndex': start_index,
                                        'endIndex': end_index
                                    }
                                }
                            })

        # Execute delete requests if there are any
        if requests:
            service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
            print("Deleted text successfully")
        else:
            print("No deletable text found.")

        # Insert new text at the original start index of the cell
        if cell_content:
            insert_index = cell_content[0]['startIndex']
        else:
            insert_index = 1  # Default start if cell is empty

        # Insert the new text
        requests_insert = [{
            'insertText': {
                'location': {
                    'index': insert_index
                },
                'text': new_text
            }
        }]

        service.documents().batchUpdate(documentId=document_id, body={'requests': requests_insert}).execute()
        print("Replaced text successfully")

    except Exception as e:
        print(f"An error occurred: {e}")

def replace_entire_cell_content(service, document_id, original_text, new_text):
    """
    Replaces the entire content of a merged cell by searching for the original text pattern 
    and replacing it with the cleaned new text.
    """
    try:
        # Create a batch update request to replace text in the document
        requests = [
            {
                "replaceAllText": {
                    "containsText": {
                        "text": original_text,  # Original content or pattern in the cell
                        "matchCase": True
                    },
                    "replaceText": new_text  # New content to replace with
                }
            }
        ]
        
        # Execute the batch update request
        service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Text replaced successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")
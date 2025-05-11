from api import *
from datetime import datetime

def insert_centered_front_page(docs_service, document_id, text):

    ## Readin the Footer Data [Date Range]
    SPREADSHEET_ID="data for the title page--> spreadsheet_id"
    start_range="'name_1'!C5"
    end_range="'name_2'!D5"
    start_date=sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=start_range).execute().get('values', [])
    end_date=sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=end_range).execute().get('values', [])

    start_date=datetime.strptime(start_date[0][0], "%d-%b").strftime("%d-%B")
    end_date=datetime.strptime(end_date[0][0], "%d-%b").strftime("%d-%B")

    subtitle_text=f"{start_date} to {end_date} 2024"

    num_line_breaks = 15  # Approximate center with 0.3" margins and 36pt font

    # Text content with line breaks for vertical centering
    centered_text = "\n" * num_line_breaks + text + "\n"

    # Define requests
    requests = [
        # Insert the text with line breaks for vertical centering
        {
            'insertText': {
                'location': {
                    'index': 1  # Start of the document
                },
                'text': centered_text
            }
        },
        # Apply formatting to the text
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': 1 + num_line_breaks,
                    'endIndex': 1 + num_line_breaks + len(text)
                },
                'textStyle': {
                    'weightedFontFamily': {'fontFamily': 'Arial'},
                    'fontSize': {'magnitude': 36, 'unit': 'PT'},
                    'bold': True,
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 28 / 255,
                                'green': 69 / 255,
                                'blue': 135 / 255
                            }
                        }
                    }
                },
                'fields': 'weightedFontFamily,fontSize,foregroundColor,bold'
            }
        },
        # Insert the subtitle text below the title
        {
            'insertText': {
                'location': {
                    'index': 1 + num_line_breaks + len(text) + 1  # +2 for the newlines after title
                },
                'text': subtitle_text
            }
        },
        # Apply formatting to the subtitle text
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': 1 + num_line_breaks + len(text) + 1,
                    'endIndex': 1 + num_line_breaks + len(text) + 1 + len(subtitle_text)
                },
                'textStyle': {
                    'weightedFontFamily': {'fontFamily': 'Arial'},
                    'fontSize': {'magnitude': 24, 'unit': 'PT'},
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 102 / 255,
                                'green': 102 / 255,
                                'blue': 102 / 255
                            }
                        }
                    }
                },
                'fields': 'weightedFontFamily,fontSize,foregroundColor'
            }
        },
        # Set paragraph alignment to left
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1 + num_line_breaks,
                    'endIndex': 1 + num_line_breaks + len(text) +len(subtitle_text)
                },
                'paragraphStyle': {
                    'alignment': 'START'  # 'START' means left-aligned
                },
                'fields': 'alignment'
            }
        },
        # Insert a page break after the front page text
        {
            'insertPageBreak': {
                'location': {
                    'index': 1 + num_line_breaks + len(text) +len(subtitle_text)+1 ## for extra space
                }
            }
        }
    ]

    try:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Centered front page inserted successfully.")
    except HttpError as error:
        print(f"An error occurred: {error}")
    return None

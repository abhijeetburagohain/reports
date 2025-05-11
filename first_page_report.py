from api import *
from helper_functions import *
from first_page_report_helper_functions import *
import time


def first_page(original_doc_id,folder_id):
    # Example: Get spreadsheet data
    SPREADSHEET_ID = 'spreadsheet_id of the dataset'
    RANGE_NAME = "'range of the data'!A:B"

    result = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    rows = result.get('values', [])

    print(f"Retrieved {len(rows)} rows from the spreadsheet.")

    df = pd.DataFrame(result["values"][1:],columns=result["values"][0]+["","",""])
    headers=result["values"][0]+["","",""]
    table_rows=result["values"][1:]

    updated_table_rows = [
        [cell.replace('Twitter (𝕏)', 'Twitter (X)') if isinstance(cell, str) else cell for cell in row]
        for row in table_rows
    ]

    new_text_1=df.iloc[2, 2]
    new_text_2=df.iloc[6, 2]

    ## Readin the Footer Data [Date Range]
    start_range="'range of the data'!C5"
    end_range="'range of the data'!D5"
    start_date=sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=start_range).execute().get('values', [])
    end_date=sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=end_range).execute().get('values', [])

    #### Making the Report 
    title=f"name of the report {start_date[0][0]}-{end_date[0][0]}"
    folder_id = folder_id

    # document_id = create_document_in_folder(folder_id,docs_service,drive_service)
    document_id=copy_document(drive_service, original_doc_id, title, folder_id=folder_id)

    print(f"Document created with ID: {document_id}")

    time.sleep(1)
    # Inserting the headers and footers

    header_text = "Confidential"
    footer_text = f"name of the report: {start_date[0][0]} to {end_date[0][0]} 2024"
    insert_header_footer(docs_service, document_id, header_text, footer_text)

    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)

    insert_index = find_last_valid_index(docs_service, document_id)
    page_break_request = {
            'insertPageBreak': {
                'location': {'index': insert_index}
            }
        }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [page_break_request]}).execute()

    end_index=get_insertion_index(docs_service, document_id)
    text="name of the report"
    insert_styled_paragraph(docs_service, document_id, text,end_index-1)

    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)

    ###Inserting the Texts
    text="1. name of the header"
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
                    'startIndex': insert_index ,  # Start after the newline
                    'endIndex': insert_index+ len(text)+1  # End at the end of the title
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {'magnitude': 12, 'unit': 'PT'},
                    'weightedFontFamily': {'fontFamily': 'PT Sans'},
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                    'red': 7 / 255.0,
                                    'green': 55 / 255.0,
                                    'blue': 99 / 255.0
                                }
                        }
                    }
                },
                'fields': 'bold,fontSize,weightedFontFamily,foregroundColor'
            }
        }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request]}).execute()

    ########################################
    time.sleep(0.5)
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)

    insert_index=get_insertion_index(docs_service, document_id)
    insert_table_and_data_into_docs_first_table(docs_service, document_id, headers, updated_table_rows,insert_index-1)

    table_start_index = get_table_start_index(docs_service, document_id, table_number=1)
    if table_start_index is not None:
        # Merge header cells (1 to 5) and (6 to 9)
        merge_cells_in_table(docs_service, document_id, table_start_index, start_row=0, end_row=0, start_column=0, end_column=5)
        merge_cells_in_table(docs_service, document_id, table_start_index, start_row=0, end_row=0, start_column=6, end_column=9)

        # Merge table rows (2nd and 3rd rows)
        for i in range(6):
            merge_cells_in_table(docs_service, document_id, table_start_index, start_row=1, end_row=2, start_column=i, end_column=i)
            time.sleep(0.3)
        # Merge 4th row's 2nd column with the next rows, keeping the value from the 4th row
        merge_cells_in_table(docs_service, document_id, table_start_index, start_row=3, end_row=6, start_column=2, end_column=2)

        # Merge 8th row's 2nd column with the next rows, keeping the value from the 4th row
        merge_cells_in_table(docs_service, document_id, table_start_index, start_row=7, end_row=9, start_column=2, end_column=2)
    else:
        print("Table not found in the document.")

    cell_output=inspect_table_content(docs_service, document_id)

    bold_table_rows_dynamic(docs_service, document_id, cell_output)

    center_align_table_text(docs_service, document_id, cell_output)
    print("center alligned")

    ## Coloring the background
    colors = [
        {"red": 1.0, "green": 0.6, "blue": 0.0},  # Light red for the first row (header)
        {"red": 1.0, "green": 0.6, "blue": 0.0},  # Light blue for the second row
        {"red": 0.027, "green": 0.216, "blue": 0.388}   # Light gray for the third row
    ]
    num_columns = 10  # Example for a table with 9 columns
    apply_table_colors(docs_service, document_id, num_columns,table_start_index,colors)
    print("background color's applied")

    table_start_index = get_table_start_index(docs_service, document_id, table_number=1)
    time.sleep(0.5)
    adjust_row_height_and_clean_spaces(docs_service, document_id,table_start_index, row_index=2)   ### Hard coded for the 2nd row
    adjust_row_height_and_clean_spaces(docs_service, document_id,table_start_index, row_index=3)   ### Hard coded for the 4th row
    adjust_row_height_and_clean_spaces(docs_service, document_id,table_start_index, row_index=7)   ### Hard coded for the 8th row

    ## Setting Tect Color
    color={
        'red': 1,
        'green': 1,
        'blue': 1
    }
    set_text_color_in_table(docs_service, document_id, table_index=0, row_indices=[1,2], column_indices=[0,1,2,3,4,5,6,7,8,9], color=color)

    color={
        'red': 0,
        'green': 0,
        'blue': 0
    }
    set_text_color_in_table(docs_service, document_id, table_index=0, row_indices=[1], column_indices=[6,7,8,9], color=color)

    ### Removing the Extra Spaces in the First row and the 3rd row of the table.

    ## Removing from the merged row in the 4th row
    i=1
    while i<3:
        delete_and_replace_text(docs_service, document_id, table_index=0, row_index=3, column_index=2,
                                new_text=new_text_1)
        i+=1
        time.sleep(0.3)
    i=1
    while i<3:
        delete_and_replace_text(docs_service, document_id, table_index=0, row_index=7, column_index=2,
                                new_text=new_text_2)
        i+=1
        time.sleep(0.3)
    ## Removing from the merged columns in the 1st row
    i=1
    while i<6:
        delete_and_replace_text(docs_service, document_id, table_index=0, row_index=0, column_index=0,
                                new_text="Summary")
        i+=1
        time.sleep(0.3)
    i=1
    while i<4:
        delete_and_replace_text(docs_service, document_id, table_index=0, row_index=0, column_index=6,
                                new_text="Platform-wise Content Shared")
        i+=1
        time.sleep(0.3)
    ### Remove Extra Space Between Tables
    table_start_index = get_table_start_index(docs_service, document_id, table_number=1)
    remove_extra_space_before_table(docs_service, document_id, table_start_index=table_start_index-1)


    ### Inserting the Italics Texts

    SPREADSHEET_ID = 'spreadsheet_id to the italics sheet'
    RANGE_NAME = "'range name'!A1:A4"

    result_2 = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    rows = result_2.get('values', [])

    print(f"Retrieved {len(rows)} rows from the spreadsheet.")
    df = pd.DataFrame(result_2["values"][1:],columns=result_2["values"][0])
    content=list(df["Italic_Observations"].unique())

    insert_formatted_text_below_table(docs_service, document_id, content)

    #### Inserting the Heading of the 2nd Table
    ###Inserting the Texts
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)

    time.sleep(0.3)
    text="2. header text"
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
                    'startIndex': insert_index ,  # Start after the newline
                    'endIndex': insert_index+ len(text)+1  # End at the end of the title
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {'magnitude': 12, 'unit': 'PT'},
                    'weightedFontFamily': {'fontFamily': 'PT Sans'},
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                    'red': 7 / 255.0,
                                    'green': 55 / 255.0,
                                    'blue': 99 / 255.0
                                }
                        }
                    }
                },
                'fields': 'bold,fontSize,weightedFontFamily,foregroundColor'
            }
        }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [title_request,update_style_request]}).execute()

    time.sleep(0.3)
    ########################################
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)


    ### Inserting the 2nd Table

    SPREADSHEET_ID = 'spreadsheet_id to the data of the 2nd table'
    RANGE_NAME = "'name of the range'!A20:J28"

    result_2 = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    rows = result_2.get('values', [])

    print(f"Retrieved {len(rows)} rows from the spreadsheet.")
    df = pd.DataFrame(result_2["values"][1:],columns=result_2["values"][0]+[""])
    headers=result_2["values"][0]+[""]
    table_rows=result_2["values"][1:]

    start_index_after_table = find_last_valid_index(docs_service, document_id)
    insert_table_and_data_into_docs(docs_service, document_id, headers, table_rows,start_index_after_table)

    #### 2nd Table Formatting
    table_start_index = get_table_start_index(docs_service, document_id, table_number=2)
    # print(table_start_index)

    if table_start_index is not None:
        # Merge table columns (1st row's adjacent column)
        i=2
        while i<10:
            merge_cells_in_table(docs_service, document_id, table_start_index, start_row=0, end_row=0, start_column=i, end_column=i+1)
            i=i+2
            time.sleep(0.3)
        # Merge 1st row's 1st column with the 2nd rows
        merge_cells_in_table(docs_service, document_id, table_start_index, start_row=0, end_row=1, start_column=0, end_column=0)
        merge_cells_in_table(docs_service, document_id, table_start_index, start_row=0, end_row=1, start_column=1, end_column=1)
    else:
        print("Table not found in the document.")

    ## Bolding the cells
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    rows_to_bold = [0, 1]  # Bold all cells in rows 1, 2, and 3
    column_rows_to_bold = {2: [1],3: [1], 4: [1], 5: [1], 6: [1],7: [1],8: [1]}  # Bold only the first cell in rows 4, 5, 6, and 7

    bold_table_rows_dynamic_2(docs_service, document_id, all_tables[1], rows_to_bold, column_rows_to_bold)

    center_align_table_text(docs_service, document_id,  all_tables[1])
    print("center alligned")

    ## Coloring the background

    colors = [
        {"red": 1.0, "green": 0.6, "blue": 0.0},  # Light red for the first row (header)
        {"red": 0.027, "green": 0.216, "blue": 0.388}   # Light gray for the third row
    ]
    num_columns = 10  # Example for a table with 9 columns
    apply_table_colors(docs_service, document_id, num_columns,table_start_index,colors)
    print("background color's applied")

    table_start_index = get_table_start_index(docs_service, document_id, table_number=2)
    adjust_row_height_and_clean_spaces(docs_service, document_id,table_start_index, row_index=0)   ### Hard coded for the 5th row
    adjust_row_height_and_clean_spaces(docs_service, document_id,table_start_index, row_index=1)   ### Hard coded for the 5th row

    ## Setting Tect Color [white]
    color={
        'red': 1,
        'green': 1,
        'blue': 1
    }
    set_text_color_in_table(docs_service, document_id, table_index=1, row_indices=[0,1], column_indices=[0,1,2,3,4,5,6,7,8,9], color=color)

    # black
    color={
        'red': 0,
        'green': 0,
        'blue': 0
    }
    set_text_color_in_table(docs_service, document_id, table_index=1, row_indices=[0], column_indices=[2,3,4,5,6,7,8,9], color=color)

    ## Removing from the merged row in the 4th row
    i=1
    while i<2:
        delete_and_replace_text(docs_service, document_id, table_index=1, row_index=0, column_index=2,
                                new_text="Video")
        i+=1
        time.sleep(0.3)
    i=1
    while i<2:
        delete_and_replace_text(docs_service, document_id, table_index=1, row_index=0, column_index=4,
                                new_text="Image Card")
        i+=1
        time.sleep(0.3)
    i=1
    while i<2:
        delete_and_replace_text(docs_service, document_id, table_index=1, row_index=0, column_index=6,
                                new_text="Reel")
        i+=1
        time.sleep(0.3)
    i=1
    while i<2:
        delete_and_replace_text(docs_service, document_id, table_index=1, row_index=0, column_index=8,
                                new_text="Shorts")
        i+=1
        time.sleep(0.3)
    ### Remove Extra Space Between Tables
    table_start_index = get_table_start_index(docs_service, document_id, table_number=2)
    remove_extra_space_before_table(docs_service, document_id, table_start_index=table_start_index-1)

    # Set Column Width
    # Retrieve all table content and starting indices
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)

    column_widths_by_table = [
        [40,66.6,52.3,80.3,51.3,61.3,51.3,51.3,51.3,50.3],     
        [40,86.6,52.3,60.3,51.3,61.3,51.3,51.3,51.3,50.3]
        ]

    set_column_widths(docs_service, document_id, table_start_indices, column_widths_by_table)
    print("columns width sorted")

    ### Adding the italics below the 2nd table

    SPREADSHEET_ID = 'spreadsheet_id to the italics data'
    RANGE_NAME = "'name of the range'!B1:B2"

    result_2 = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    rows = result_2.get('values', [])

    print(f"Retrieved {len(rows)} rows from the spreadsheet.")
    df = pd.DataFrame(result_2["values"][1:],columns=result_2["values"][0])
    content=list(df["after_second_table"].unique())

    insert_formatted_text_below_table(docs_service, document_id, content)
    print("the italics added")
    print("the first page data insertion done")

    return document_id









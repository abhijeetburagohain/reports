from api import *
from helper_functions import *


def youtube(document_id):
    spreadsheet_id_static = 'spreadsheet_id of the static data'

    ### Fetching the Static data
    RANGE_NAME_STATIC = "'range_name_1'!C2:C9"
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id_static, range=RANGE_NAME_STATIC).execute()
    df = pd.DataFrame(result["values"])
    df = df.dropna(how='all').reset_index(drop=True)
    table_rows_static=df[0].to_list()

    ### Observations
    RANGE_NAME_OBSERVATION = "'range_name_2'!C2:C10"
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id_static, range=RANGE_NAME_OBSERVATION).execute()
    df = pd.DataFrame(result["values"])
    df = df.dropna(how='all').reset_index(drop=True)
    df=df[df[0]!=""]
    table_rows=df[0].to_list()

    ## Fetching the Youtube Tables Data
    spreadsheet_id_instagram = 'spreadsheet_id of the youtube data'

    tab_name="tabname_1"


    ranges_dict = {
        "name_1":f"{tab_name}!B4:F11",
        "name_2": f"{tab_name}!I4:K11",
        }

    # Fetching the data as dataframes
    dataframes = fetch_multiple_ranges_as_dataframes(sheets_service, spreadsheet_id_instagram, ranges_dict)

    ## Applying Page Breaks
    insert_index=get_insertion_index(docs_service, document_id)
    page_break_request = {
                'insertPageBreak': {
                    'location': {'index': insert_index-1}
                }
            }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [page_break_request]}).execute()
    # Inserting the Static Texts

    text="name of the report Youtube"
    insert_index=get_insertion_index(docs_service, document_id)
    insert_styled_paragraph_index(docs_service, document_id, text,insert_index-1)

    ### Static data Insertion
    end_index=get_insertion_index(docs_service, document_id)
    format_paragraphs_and_numbered_bullets_youtube(docs_service, document_id, table_rows_static, end_index-1)

    # 1 table insertion [1st table]
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
    insert_texts_and_multiple_tables_into_docs_youtube(docs_service, document_id, dataframes,0,1)

    # Inserting the Observations
    start_index_after_table = find_last_valid_index(docs_service, document_id)
    insert_numbered_texts_after_table_2(docs_service, document_id, table_rows, start_index_after_table)

    # 1 table insertion [2nd table]
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
    insert_texts_and_multiple_tables_into_docs_youtube(docs_service, document_id, dataframes,1,2)

    ##Table formatting
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    bold_table_headers_youtube(docs_service, document_id, all_tables[30:])

    apply_colors_to_all_table_headers_till_3rd(docs_service, document_id, all_tables[30:], table_start_indices[30:])

    color_cells_based_on_texts_2(docs_service, document_id, table_index=30) 
    center_align_all_table_text(docs_service, document_id, all_tables)

    column_widths_by_table = [
        [100,152,100,100,100],
        [176.5,176.5,176.5]     
    ]

    set_column_widths(docs_service, document_id, table_start_indices[30:], column_widths_by_table)
    return None

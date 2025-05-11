from api import *
from helper_functions import *

def last_page(document_id):
    ## Fetching the Facebook Tables Data
    spreadsheet_id = 'spreadsheet_id of the last page data'

    tab_name="last_page"

    ranges_dict = {
        "name_1": f"{tab_name}!A3:C5",
        "name_2": f"{tab_name}!E3:G5",
        "name_3": f"{tab_name}!I3:K5",
        "name_4": f"{tab_name}!M3:O5",
        "name_5": f"{tab_name}!Q3:S5"
        }

    # Fetching the data as dataframes
    dataframes = fetch_multiple_ranges_as_dataframes(sheets_service, spreadsheet_id, ranges_dict,Instagram=0,last_page=1)

    ## Applying Page Breaks
    insert_index=get_insertion_index(docs_service, document_id)
    page_break_request = {
                'insertPageBreak': {
                    'location': {'index': insert_index-1}
                }
            }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [page_break_request]}).execute()
    # Inserting the Static Texts

    text="name of the header Analysis"
    insert_index=get_insertion_index(docs_service, document_id)
    insert_styled_paragraph_index(docs_service, document_id, text,insert_index-1)

    # 3 table insertion
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
    insert_texts_and_multiple_tables_into_docs_youtube(docs_service, document_id, dataframes,0,len(dataframes))
    print(len(dataframes))
    #merging the 1st column of the third table
    i=32
    while i<37:
        all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
        merge_cells_in_table(docs_service, document_id, table_start_indices[i], start_row=2, end_row=2, start_column=0, end_column=1)
        i=i+1
    
    print("merged_table_cells")

    # Bold text
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    bold_table_headers_youtube(docs_service, document_id, all_tables[32:])

    center_align_all_table_text(docs_service, document_id, all_tables)
    print("center alligned")

    apply_colors_to_all_table_headers_till_3rd(docs_service, document_id, all_tables[32:], table_start_indices[32:])

    column_widths_by_table = [
        [111,331,111],
        [111,331,111],
        [111,331,111],
        [111,331,111],
        [111,331,111]     
    ]

    set_column_widths(docs_service, document_id, table_start_indices[32:], column_widths_by_table)

    ### Removing the Extra Spaces with a new method
    all_tables, table_start_indices = inspect_all_tables_content_2(docs_service, document_id)
    replace_dict={}
    for i in range(32,len(all_tables)):
        ori_text=all_tables[i][2][0]["Text"]
        new_text=ori_text.strip()
        replace_dict[ori_text]=new_text

    for key,value in replace_dict.items():
        original_text=key
        new_text=value
        replace_entire_cell_content(docs_service, document_id, original_text, new_text)
        
    return None

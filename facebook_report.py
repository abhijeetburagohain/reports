from api import *
from helper_functions import *
from itertools import islice
import numpy as np

### FACEBOOK 
def facebook(document_id):
    spreadsheet_id_static = 'static data spresdsheet_id'

    ### Fetching the Static data
    RANGE_NAME_STATIC = "'range_name_1'!A2:A9"
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id_static, range=RANGE_NAME_STATIC).execute()
    df = pd.DataFrame(result["values"])
    df = df.dropna(how='all').reset_index(drop=True)
    table_rows_static=df[0].to_list()

    ### Observations
    RANGE_NAME_OBSERVATION = "'range_name_2'!A2:A10"
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id_static, range=RANGE_NAME_OBSERVATION).execute()
    df = pd.DataFrame(result["values"])
    df = df.dropna(how='all').reset_index(drop=True)
    df=df[df[0]!=""]
    table_rows=df[0].to_list()

    ## Fetching the Facebook Tables Data
    spreadsheet_id_facebook = 'facebook data spresdsheet_id'

    tab_name_1="name_1"
    tab_name_2="name_2"

    ranges_dict = {
        "name_1": f"{tab_name_1}!B8:J15",
        "name_2":f"{tab_name_1}!N8:Q15",
        "name_3": f"{tab_name_1}!T8:X29",
        "name_4": f"{tab_name_1}!AA8:AF30",
        "name_5": f"{tab_name_1}!AI8:AN30",
        "name_6": f"{tab_name_1}!AQ8:AV30",
        "name_7": f"{tab_name_1}!AY8:BD30",
        "name_8": f"{tab_name_2}!Z6:AE26",
        "name_9": f"{tab_name_2}!AH6:AM26",
        "name_10": f"{tab_name_2}!AP6:AU26"
        }

    # Fetching the data as dataframes
    dataframes = fetch_multiple_ranges_as_dataframes(sheets_service, spreadsheet_id_facebook, ranges_dict)
    #### Data Preprocessing
    for key, value in islice(dataframes.items(), 3,None):
        dataframes[key]=dataframes[key][dataframes[key].iloc[:, 1]!=""]

    dataframes["df_1"]["col_1"] = dataframes["df_1"]["col_1"].replace("", np.nan)
    dataframes["df_1"]["col_1"]=dataframes["df_1"]["col_1"].fillna(method="ffill")
    dataframes["df_1"]["col_1"] = dataframes["df_1"]["col_1"].where(dataframes["df_1"].index % 3 == 1, "")
    dataframes["df_1"]=dataframes["df_1"].replace("#N/A","-")
    ### Adding page break
    insert_index = find_last_valid_index(docs_service, document_id)
    page_break_request = {
                'insertPageBreak': {
                    'location': {'index': insert_index}
                }
            }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [page_break_request]}).execute()
    # Inserting the Static Texts
    end_index=get_insertion_index(docs_service, document_id)

    text="header text"
    insert_styled_paragraph(docs_service, document_id, text,end_index-1)

    ### Static data Insertion

    end_index=get_insertion_index(docs_service, document_id)
    format_paragraphs_and_numbered_bullets(docs_service, document_id, table_rows_static,end_index-1)

    insert_texts_and_multiple_tables_into_docs(docs_service, document_id, dataframes)

    # Inserting the Observations

    start_index_after_table = find_last_valid_index(docs_service, document_id)
    insert_numbered_texts_after_table_2(docs_service, document_id, table_rows, start_index_after_table)

    # the remaining table insertion
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)

    insert_texts_and_multiple_tables_into_docs_last_5(docs_service, document_id, dataframes)

    ##Table formatting
    #merging the 1st column of the third table
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    if table_start_indices[4] is not None:
        # Merge table columns (1st row's adjacent column)
        i=1
        while i<=21:
            merge_cells_in_table(docs_service, document_id, table_start_indices[4], start_row=i, end_row=i+2, start_column=0, end_column=0)
            i=i+3

    #merging the first row's of the Morcha Table
    i=5
    while i<12:
        all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
        merge_cells_in_table(docs_service, document_id, table_start_indices[i], start_row=0, end_row=0, start_column=0, end_column=5)
        i+=1
    print("merging the cells.....Done")

    # Bold text
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    bold_table_headers(docs_service, document_id, all_tables[2:])

    # Coloring 
    apply_colors_to_all_table_headers_till_3rd(docs_service, document_id, all_tables[2:5], table_start_indices[2:5])
    apply_colors_to_all_table_headers_from_4th(docs_service, document_id, all_tables[5:], table_start_indices[5:])

    #color cells based on condition

    color_cells_based_on_arrows(docs_service, document_id, table_index=3)  ### 1 for 2nd table

    # Centre allign
    center_align_all_table_text(docs_service, document_id, all_tables[2:])

    # Set Column Width
    # Retrieve all table content and starting indices
    # Define column widths for each table; ensure each inner list corresponds to the correct number of columns

    column_widths_by_table = [
        [62,62,62,61,62,61,61,61,61],     
        [141,137,137,137],
        [111,110,110,110,110],
        [40,175,83,85,85,85],
        [40,175,83,85,85,85],
        [40,175,83,85,85,85],
        [40,175,83,85,85,85],
        [40,175,83,85,85,85],
        [40,175,83,85,85,85],
        [40,175,83,85,85,85]
    ]

    set_column_widths(docs_service, document_id, table_start_indices[2:], column_widths_by_table)

    ### Removing the Extra Spaces with a new method
    all_tables, table_start_indices = inspect_all_tables_content_2(docs_service, document_id)
    replace_dict={}
    for i in range(5,len(all_tables)):
        ori_text=all_tables[i][0][0]["Text"]
        new_text=ori_text.strip()
        replace_dict[ori_text]=new_text

    print(replace_dict)
      
    for key,value in replace_dict.items():
        original_text=key
        new_text=value
        replace_entire_cell_content(docs_service, document_id, original_text, new_text)
    
    print("facebook data insertion done")
    return None














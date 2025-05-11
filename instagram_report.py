from api import *
from helper_functions import *
from itertools import islice
import time
import numpy as np


def instagram(document_id):
    spreadsheet_id_static = 'spreadsheet_id intagram'

    ### Fetching the Static data
    RANGE_NAME_STATIC = "'name of range'!B2:B9"
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id_static, range=RANGE_NAME_STATIC).execute()
    df = pd.DataFrame(result["values"])
    df = df.dropna(how='all').reset_index(drop=True)
    table_rows_static=df[0].to_list()

    ### Observations
    RANGE_NAME_OBSERVATION = "'Observations'!B2:B10"
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id_static, range=RANGE_NAME_OBSERVATION).execute()
    df = pd.DataFrame(result["values"])
    df = df.dropna(how='all').reset_index(drop=True)
    df=df[df[0]!=""]
    table_rows=df[0].to_list()

    ##Reading the Reels Data
    spreadsheet_reels = 'spreadsheet_id of the reels data'
    tab_name="name of the tab"
    ranges_dict = {
        "name_1": f"{tab_name}!BQ5:BX6",
        "name_2": f"{tab_name}!BQ8:BX9",
        "name_3": f"{tab_name}!BQ12:BX13",
        "name_4": f"{tab_name}!BQ17:BX18",
        "name_5": f"{tab_name}!BQ22:BX23",
        "name_6": f"{tab_name}!BQ25:BX26",
        "name_7": f"{tab_name}!BQ29:BX30"
        }
    
    # Fetching the data as dataframes
    dataframes_reels = fetch_multiple_ranges_as_dataframes_reels(sheets_service, spreadsheet_reels, ranges_dict)


    ## Fetching the Facebook Tables Data
    spreadsheet_id_instagram = 'spreadsheet_id_instagram'

    tab_name_1="tab_name_1"
    tab_name_2="tab_name_1"

    ranges_dict = {
        "name_1":f"{tab_name_1}!B5:F12",
        "name_2": f"{tab_name_1}!H5:M12",
        "name_3":f"{tab_name_1}!P5:S12",
        "name_4": f"{tab_name_1}!Z5:AE19",
        "name_5": f"{tab_name_1}!AH5:AM27",
        "name_6": f"{tab_name_1}!AP5:AU27",
        "name_7": f"{tab_name_1}!AX5:BC27",
        "name_8": f"{tab_name_1}!BF5:BK27",
        "name_9": f"{tab_name_2}!AF8:AK28",
        "name_10": f"{tab_name_2}!AN8:AS28",
        "name_11": f"{tab_name_2}!AW8:BB28"
        }

    # Fetching the data as dataframes
    dataframes = fetch_multiple_ranges_as_dataframes(sheets_service, spreadsheet_id_instagram, ranges_dict,Instagram=1)
    dataframes["df_1"].columns
    #### Data Preprocessing
    for key, value in islice(dataframes.items(), 3,None):
        dataframes[key]=dataframes[key][dataframes[key].iloc[:, 1]!=""]

    dataframes["df_1"]["col_name_1"] = dataframes["df_1"]["col_name_1"].replace("", np.nan)
    dataframes["df_1"]["col_name_1"]=dataframes["df_1"]["col_name_1"].fillna(method="ffill")
    dataframes["df_1"]["col_name_1"] = dataframes["df_1"]["col_name_1"].where(dataframes["Performance_of_Image/Video_Published_on_Instagram"].index % 2 == 0, "")


    ## Applying Page Breaks

    insert_index=get_insertion_index(docs_service, document_id)
    page_break_request = {
                'insertPageBreak': {
                    'location': {'index': insert_index-1}
                }
            }
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [page_break_request]}).execute()

    # Inserting the Static Texts

    text="name of the report Instagram"
    insert_index=get_insertion_index(docs_service, document_id)
    insert_styled_paragraph_index(docs_service, document_id, text,insert_index-1)

    ### Static data Insertion
    end_index=get_insertion_index(docs_service, document_id)
    format_paragraphs_and_numbered_bullets(docs_service, document_id, table_rows_static,end_index-1)

    # 3 table insertion
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
    insert_texts_and_multiple_tables_into_docs(docs_service, document_id, dataframes,instagram=1)

    # Inserting the Observations
    start_index_after_table = find_last_valid_index(docs_service, document_id)
    insert_numbered_texts_after_table_2(docs_service, document_id, table_rows, start_index_after_table)

    # the remaining table insertion
    end_index=get_insertion_index(docs_service, document_id)
    reset_format_and_move_to_next_line(docs_service, document_id, end_index-1)
    insert_texts_and_multiple_tables_into_docs_last_5(docs_service, document_id, dataframes,instagram=1,dataframes_reels=dataframes_reels)

    ##Table formatting
    #merging the 1st column of the third table
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    if table_start_indices[15] is not None:
        # Merge table columns (1st row's adjacent column)
        print("table 15 found")
        i=1
        while i<15:
            merge_cells_in_table(docs_service, document_id, table_start_indices[15], start_row=i, end_row=i+1, start_column=0, end_column=0)
            i=i+2
            time.sleep(0.3)

    #merging the first row's of the Morcha Table
    i=16
    while i<29:
        all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
        merge_cells_in_table(docs_service, document_id, table_start_indices[i], start_row=0, end_row=0, start_column=0, end_column=5)
        i+=2
        time.sleep(0.3)
    ## merging the 2nd table in the page.
    i=17
    while i<=29:
        all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
        merge_cells_in_table(docs_service, document_id, table_start_indices[i], start_row=0, end_row=0, start_column=0, end_column=7)
        i+=2
        time.sleep(0.3)

    # Bold text
    all_tables, table_start_indices = inspect_all_tables_content(docs_service, document_id)
    bold_table_headers_insta(docs_service, document_id, all_tables)
    time.sleep(1)
    # Coloring 
    
    apply_colors_to_all_table_headers_till_3rd(docs_service, document_id, all_tables[12:16], table_start_indices[12:16])
    apply_colors_to_all_table_headers_from_4th(docs_service, document_id, all_tables[16:], table_start_indices[16:])

    #color cells based on condition

    color_cells_based_on_arrows(docs_service, document_id, table_index=14)  
    color_cells_based_on_texts_2(docs_service, document_id, table_index=12) 
    time.sleep(0.5)

    column_to_color = 0  #for 1st column 
    apply_color_to_single_column(docs_service, document_id, all_tables[12], table_start_indices[12], column_to_color)
    time.sleep(0.5)

    # Centre allign
    center_align_all_table_text(docs_service, document_id, all_tables[12:])

    # Set Column Width
    # Retrieve all table content and starting indices
    # Define column widths for each table; ensure each inner list corresponds to the correct number of columns

    column_widths_by_table = [
        [100,172,100,100,100],
        [92,92,92,92,92,92],     
        [143,136,136,136],
        [92,92,92,92,92,92],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69],
        [40,175,83,85,85,85],
        [69,69,69,69,69,69,69,69]
    ]

    set_column_widths(docs_service, document_id, table_start_indices[12:], column_widths_by_table)
    print("columns width set")

    ### Removing the Extra Spaces with a new method
    all_tables, table_start_indices = inspect_all_tables_content_2(docs_service, document_id)
    replace_dict={}
    for i in range(16,len(all_tables)):
        ori_text=all_tables[i][0][0]["Text"]
        new_text=ori_text.strip()
        replace_dict[ori_text]=new_text
      
    for key,value in replace_dict.items():
        original_text=key
        new_text=value
        replace_entire_cell_content(docs_service, document_id, original_text, new_text)

    print("instagram data inserted")
    return None


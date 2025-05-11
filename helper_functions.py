import pandas as pd

def fetch_multiple_ranges_as_dataframes(service, spreadsheet_id, ranges_dict,Instagram=0,last_page=None):
    """
    Fetches data from multiple ranges in a Google Spreadsheet and returns a dictionary of dataframes.
    
    :param service: The Google Sheets API service instance.
    :param spreadsheet_id: The ID of the Google Spreadsheet.
    :param ranges_dict: A dictionary where keys are dataframe names and values are the ranges in the sheet.
                        Example: {"df1": "Sheet1!A1:C10", "df2": "Sheet2!B1:D5"}
    :return: A dictionary where keys are dataframe names and values are Pandas dataframes.
    """
    # Extract all the ranges from the dictionary
    ranges = list(ranges_dict.values())
    # print(ranges)
    
    # Perform a batchGet request to get all the ranges in one API call
    result = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges).execute()
    
    # Get the data corresponding to the ranges
    fetched_data = result.get('valueRanges', [])
    
    # Initialize a dictionary to store the dataframes
    dataframes_dict = {}
    
    if Instagram==1:
        print("Hello")
        # Loop through the fetched data and create dataframes
        for i, range_name in enumerate(ranges_dict.keys()):
            # Get the values for this range
            values = fetched_data[i].get('values', [])
            if i>=4:
                column=[f"{range_name} name of the report Instagram"]+['' for _ in range(len(values[0]) - 1)]
                df = pd.DataFrame(values,columns=column)
                dataframes_dict[range_name] = df
            else:
                df = pd.DataFrame(values[1:],columns=values[0])
                dataframes_dict[range_name] = df
    elif last_page==1:
        for i, range_name in enumerate(ranges_dict.keys()):
            # Get the values for this range
            values = fetched_data[i].get('values', [])
            df = pd.DataFrame(values[1:],columns=values[0])
            dataframes_dict[range_name] = df
    else:
        # Loop through the fetched data and create dataframes
        for i, range_name in enumerate(ranges_dict.keys()):
            # Get the values for this range
            values = fetched_data[i].get('values', [])
            if i>=3:
                column=[f"{range_name} name of the report Facebook"]+['' for _ in range(len(values[0]) - 1)]
                df = pd.DataFrame(values,columns=column)
                dataframes_dict[range_name] = df
            else:
                df = pd.DataFrame(values[1:],columns=values[0])
                dataframes_dict[range_name] = df

    return dataframes_dict




def fetch_multiple_ranges_as_dataframes_reels(service, spreadsheet_id, ranges_dict):
    # Extract all the ranges from the dictionary
    ranges = list(ranges_dict.values())
    # print(ranges)
    
    # Perform a batchGet request to get all the ranges in one API call
    result = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges).execute()
    
    # Get the data corresponding to the ranges
    fetched_data = result.get('valueRanges', [])
    
    # Initialize a dictionary to store the dataframes
    dataframes_dict = {}
    print("Hello")
    # Loop through the fetched data and create dataframes
    for i, range_name in enumerate(ranges_dict.keys()):
        # Get the values for this range
        values = fetched_data[i].get('values', [])
        column=[f"{range_name} name of the report Instagram"]+['' for _ in range(len(values[0]) - 1)]
        df = pd.DataFrame(values,columns=column)
        dataframes_dict[range_name] = df
    return dataframes_dict
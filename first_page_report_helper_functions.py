
def insert_table_and_data_into_docs_first_table(docs_service, document_id, headers, table_rows,insert_index):
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
    for element in content:
        if 'table' in element:
            for row in element['table']['tableRows']:
                for cell in row['tableCells']:
                    # Each table cell contains a paragraph, so we get the start index of each
                    paragraph = cell['content'][0]['paragraph']['elements'][0]['startIndex']
                    paragraph_1 = cell['content'][0]['paragraph']['elements'][0]['endIndex']
                    table_cells.append(paragraph)

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
            current_offset += len(str(cell_value))  # Update the offset by the length of the inserted text

    # Step 5: Send the request to populate the table
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
    print(f"Inserted table from dataframe into Google Docs with ID: {document_id}")

def insert_table_and_data_into_docs_first_table_2(docs_service, document_id, headers, table_rows, insert_index):
    # Step 1: Insert an empty table with the correct dimensions
    num_columns = len(headers)
    num_rows = len(table_rows) + 1  # +1 for the header row

    insert_table_request = {
        'insertTable': {
            'rows': num_rows,
            'columns': num_columns,
            'location': {
                'index': insert_index  # Insert table at the specified index
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
    for element in content:
        if 'table' in element:
            for row in element['table']['tableRows']:
                for cell in row['tableCells']:
                    # Each table cell contains a paragraph, so we get the start index of each
                    paragraph = cell['content'][0]['paragraph']['elements'][0]['startIndex']
                    table_cells.append(paragraph)

    # Step 4: Prepare the requests to insert the headers and data
    populate_requests = []

    # Insert headers
    for col_index, header in enumerate(headers):
        # Insert text at the corresponding paragraph index inside the table cell
        header_text = header or ' '  # Ensure header is not None or empty
        start_index = table_cells[col_index]  # Get the exact index for the header cell

        populate_requests.append({
            'insertText': {
                'text': header_text,
                'location': {
                    'index': start_index
                }
            }
        })
        populate_requests.append({
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': start_index + len(header_text)
                },
                'textStyle': {
                    'fontFamily': 'PT Sans',
                    'fontSize': {
                        'magnitude': 12,
                        'unit': 'PT'
                    }
                },
                'fields': 'fontFamily,fontSize'
            }
        })

    # Insert data rows
    for row_index, row in enumerate(table_rows):
        for col_index, cell_value in enumerate(row):
            cell_text = str(cell_value or ' ')  # Handle empty or None values
            # Calculate the correct index for each data cell
            table_cell_index = num_columns + row_index * num_columns + col_index
            start_index = table_cells[table_cell_index]

            populate_requests.append({
                'insertText': {
                    'text': cell_text,
                    'location': {
                        'index': start_index
                    }
                }
            })
            populate_requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': start_index + len(cell_text)
                    },
                    'textStyle': {
                        'fontFamily': 'PT Sans',
                        'fontSize': {
                            'magnitude': 12,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'fontFamily,fontSize'
                }
            })

    # Step 5: Send the request to populate the table
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': populate_requests}).execute()
    print(f"Inserted table from dataframe into Google Docs with ID: {document_id}")


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

def get_table_start_index(docs_service, document_id, table_number=1):
    """
    Retrieves the start index of the specified table in a Google Document.
    :param docs_service: Google Docs API service
    :param document_id: ID of the Google Document
    :param table_number: The table number to search for (default is 1st table)
    :return: The start index of the table if found, else None
    """
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content', [])
    
    table_count = 0
    
    # Loop through the document content to find tables
    for element in content:
        if 'table' in element:
            table_count += 1
            if table_count == table_number:
                # The 'startIndex' key represents the start location of the table
                return element['startIndex']
    
    return None  # If the specified table isn't found

def inspect_table_content(docs_service, document_id):
    """
    Inspects the table content and returns a structured output with startIndex, endIndex, and text for each cell in the document.
    """
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content', [])
    
    cell_output = []  # This will hold the structured output

    for element in content:
        if 'table' in element:
            table_content = element['table']['tableRows']
            for row_index, row in enumerate(table_content):
                row_cells = []  # List to hold cells of the current row
                for cell_index, cell in enumerate(row['tableCells']):
                    start_index = cell.get('startIndex')
                    end_index = cell.get('endIndex')
                    text = ""
                    for content_element in cell.get('content', []):
                        if 'paragraph' in content_element:
                            for paragraph_element in content_element['paragraph']['elements']:
                                if 'textRun' in paragraph_element:
                                    text += paragraph_element['textRun']['content']
                    
                    # Append the cell information as a dictionary
                    row_cells.append({
                        'Start': start_index,
                        'End': end_index,
                        'Text': text.strip()  # Stripping any extra whitespace
                    })
                cell_output.append(row_cells)  # Append the row cells to cell_output

    return cell_output


def bold_table_rows_dynamic(docs_service, document_id, cell_output):
    """
    Bold the text in the first three rows of the table and the first column from rows 4 to 7 dynamically
    based on the provided cell output.
    """
    requests = []

    # Extract indices from the provided cell output
    # Assuming cell_output is a list of dictionaries with 'Start' and 'End' keys
    for row in range(3):  # For the first three rows
        for cell in range(10):  # Assuming 9 cells per row
            start_index = cell_output[row][cell]['Start']
            end_index = cell_output[row][cell]['End']
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

    # Bold the first column from rows 4 to 7
    for row in range(3, 10):  # From row 4 (index 3) to row 7 (index 6)
        start_index = cell_output[row][1]['Start']  # First cell in the row
        end_index = cell_output[row][1]['End']  # First cell in the row
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

    # Execute the batch update
    try:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Bold formatting applied successfully.")
    except Exception as e:
        print(f"Error applying bold formatting: {e}")

def bold_table_rows_dynamic_2(docs_service, document_id, cell_output, rows_to_bold=None, column_rows_to_bold=None):
    """
    Bold the text in specified rows and columns dynamically based on provided cell output.

    :param docs_service: Google Docs API service instance.
    :param document_id: Document ID.
    :param cell_output: List of lists containing dictionaries with 'Start' and 'End' keys for each cell.
    :param rows_to_bold: List of row indices to bold all cells within.
    :param column_rows_to_bold: Dictionary where keys are row indices and values are lists of column indices to bold in those rows.
    """
    requests = []

    # Bold all cells in specified rows
    if rows_to_bold:
        for row in rows_to_bold:
            for cell in range(len(cell_output[row])):  # Iterate over each cell in the row
                start_index = cell_output[row][cell]['Start']
                end_index = cell_output[row][cell]['End']
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

    # Bold specific cells in specified rows and columns
    if column_rows_to_bold:
        for row, columns in column_rows_to_bold.items():
            for col in columns:
                start_index = cell_output[row][col]['Start']
                end_index = cell_output[row][col]['End']
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

    # Execute the batch update
    try:
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Bold formatting applied successfully.")
    except Exception as e:
        print(f"Error applying bold formatting: {e}")



def center_align_table_text(docs_service, document_id, cell_output):
    """
    Center and middle-align cell text in the entire table using previously inspected cell output.
    """
    requests = []

    for row_cells in cell_output:
        for cell in row_cells:
            start_index = cell['Start']
            end_index = cell['End']
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': end_index
                    },
                    'paragraphStyle': {
                        'alignment': 'CENTER'  # Center align text
                    },
                    'fields': 'alignment'
                }
            })

    # Execute the batch update
    if requests:  # Only execute if there are requests
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

def apply_table_colors(docs_service, document_id, num_columns,table_index,colors):
    populate_requests = []
    
    # Loop through first 3 rows to apply different background colors
    for row_index in range(len(colors)):
        populate_requests.append({
            'updateTableCellStyle': {
                'tableRange': {
                    'tableCellLocation': {
                        'tableStartLocation': {
                            'index': table_index  # Make sure to specify the correct table index
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

    # Send the batch update request to apply colors
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': populate_requests}
    ).execute()

    print(f"Applied color styling to the first 3 rows in the table of document: {document_id}")

### Final code for maintaining cell padding.

def adjust_row_height_and_clean_spaces(docs_service, document_id, table_index, row_index):
    """
    Removes extra spaces in cells and adjusts the row height and padding in a table.
    :param docs_service: Google Docs API service instance.
    :param document_id: Document ID.
    :param table_index: The index of the table in the document (0 for the first table).
    :param row_index: The index of the row to adjust (0 for the first row).
    """
    # First, remove extra spaces in the table cells
    # remove_extra_spaces_in_table(docs_service, document_id, table_index)
    
    # Then, adjust the row height
    requests = [{
        'updateTableCellStyle': {
            'tableCellStyle': {
                'paddingTop': {'magnitude': 0, 'unit': 'PT'},
                'paddingBottom': {'magnitude': 0, 'unit': 'PT'},
                'contentAlignment': 'MIDDLE'
            },
            'tableRange': {
                'tableCellLocation': {
                    'tableStartLocation': {
                        'index': table_index
                    },
                    'rowIndex': row_index,
                    'columnIndex': 0,
                },
                'rowSpan': 1,
                'columnSpan': 9
            },
            'fields': 'paddingTop,paddingBottom,contentAlignment'
        }
    }]

    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()
    
    print(f"Adjusted row height for row {row_index} in table at index {table_index}.")

def set_text_color_in_table(service, document_id, table_index, row_indices, column_indices, color="#FFFFFF"):
    """
    Set text color for specific rows and columns in a Google Docs table.
    
    Parameters:
    - service: Authorized Google Docs API service instance.
    - document_id: The ID of the Google Doc.
    - table_index: The index of the table in the document.
    - row_indices: List of row indices where text color should be set.
    - column_indices: List of column indices in each row to apply the text color.
    - color: Hex color code for the text color (default is white).
    """
    # Retrieve the document content to locate the specific table
    document = service.documents().get(documentId=document_id).execute()
    tables = [content for content in document['body']['content'] if 'table' in content]
    
    # Validate the table index
    if table_index >= len(tables):
        print("Table index out of range.")
        return

    table = tables[table_index]['table']
    
    # Prepare requests to change text color
    requests = []
    
    # Iterate through specified rows and columns
    for row_idx in row_indices:
        if row_idx >= len(table['tableRows']):
            print(f"Row index {row_idx} out of range.")
            continue
        
        row = table['tableRows'][row_idx]
        for col_idx in column_indices:
            if col_idx >= len(row['tableCells']):
                print(f"Column index {col_idx} out of range in row {row_idx}.")
                continue
            
            cell_content = row['tableCells'][col_idx]['content']
            
            # Add a request to change the text color for each paragraph in the cell
            for element in cell_content:
                if 'paragraph' in element:
                    for paragraph_element in element['paragraph']['elements']:
                        text_run = paragraph_element.get('textRun')
                        if text_run:
                            start_index = paragraph_element['startIndex']
                            end_index = paragraph_element['endIndex']
                            
                            # Define a request to update the text style to white
                            requests.append({
                                'updateTextStyle': {
                                    'range': {
                                        'startIndex': start_index,
                                        'endIndex': end_index
                                    },
                                    'textStyle': {
                                        'foregroundColor': {
                                            'color': {
                                                'rgbColor': color
                                            }
                                        }
                                    },
                                    'fields': 'foregroundColor'
                                }
                            })
    
    # Execute the batch update to set the text color
    if requests:
        service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print("Text color updated successfully.")
    else:
        print("No requests to process.")

def insert_formatted_text_below_table(docs_service, document_id, text_list):
    """
    Inserts formatted text below the last table in the document.
    :param docs_service: Google Docs API service instance.
    :param document_id: Document ID.
    :param text_list: List of text lines to insert.
    """
    # Fetch the document content
    document = docs_service.documents().get(documentId=document_id).execute()
    
    # Find the end index of the last table
    content = document.get('body').get('content')
    end_index = 0
    for element in content:
        if 'table' in element:
            end_index = element['endIndex']
    
    # Prepare the requests to insert the text and format it
    requests = []
    # Insert an empty paragraph (for space)
    requests.append({
        'insertText': {
            'location': {
                'index': end_index
            },
            'text': '\n'  # This creates a space
        }
    })
    end_index += 1  # Update the end index after inserting the space
    start_indices = []
    for text in text_list:
        requests.append({
            'insertText': {
                'location': {
                    'index': end_index
                },
                'text': text + '\n'
            }
        })
        start_indices.append(end_index)  # Store the start index of each text for styling
        end_index += len(text) + 1  # Update the index for the next line of text
    
    # Add styling requests (italics, font size 10, PT Sans)
    for i, text in enumerate(text_list):
        start_index = start_indices[i]
        end_index = start_index + len(text)
        requests.append({
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                },
                'textStyle': {
                    'italic': True,
                    'fontSize': {
                        'magnitude': 10,
                        'unit': 'PT'
                    },
                    'weightedFontFamily': {
                        'fontFamily': 'PT Sans',
                        'weight': 400
                    }
                },
                'fields': 'italic,fontSize,weightedFontFamily'
            }
        })
    
    # Execute the batch update request
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    print("Formatted text inserted below the table successfully.")

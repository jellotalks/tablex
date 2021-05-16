
import sys
import requests
import pandas as pd
from bs4 import BeautifulSoup
import os
from validators import url as urlv

# Main function. This is intended to be used by individual users and has prompts for users to enter their own information
def main():
    
    print("Bailey Muckel's HTML to xlsx table exporter.")
    
    # Loop through user inputs until user gets it right
    while True:
        url,tablenum,filepath = "",0,""

        # Check if valid url
        url = str(input("Please enter the URL you wish to capture: "))
        if not urlv(url):
            print("URL is not valid")
            continue
        
        # Check if valid integer
        val = input("Please enter which table to capture as a number [0 - x]: ")
        try:
            tablenum = int(val)
        except ValueError:
            print("Table number is not an integer")
            continue

        # Input filepath
        filepath = str(input("Please enter the filepath to export to (including the extension): "))

        break

    print()


    # Connect to the URL, display progress messages
    print("Connecting to URL...",end='',flush=True)
    req = connect_to_url(url)
    print("\rConnected.          ")
    
    # Parse the html, and grab the table the function wants
    print("Parsing data...",end='',flush=True)
    df = extract_table(req,tablenum)
    print("\rTable extracted.")

    # Make unique file name then pass dataframe into the file.
    print("Writing data to file...",end='',flush=True)
    dataframe_to_xlsx(df, uniquefile(filepath))
    print("\rData written to file.  ")


# Export a table with the given parameters. This functions behaves the same as main(), but without the user prompts
# This is the function to call when importing this library
def table_export(url : str = 'https://en.wikibooks.org/wiki/Vehicle_Identification_Numbers_(VIN_codes)/World_Manufacturer_Identifier_(WMI)', filepath : str = 'data.xlsx', tablenum : int = 0):
    req = connect_to_url(url)
    df = extract_table(req,tablenum)
    dataframe_to_xlsx(df,uniquefile(filepath))



# Connect to a URL and return the request response
def connect_to_url(url : str) -> requests.Response:
    # Try connecting and handle any error
    try:
        request = requests.get(url)
    except requests.exceptions.HTTPError as err:
        print("\rHTTP Error:",err)
    except requests.exceptions.ConnectionError as err:
        print("\rError connecting:",err)
    except requests.exceptions.Timeout as err:
        print("\rTimeout error:",err)
    except requests.exceptions.RequestException as err:
        print("\rUnknown error processing request:",err)
    else: 
        # If no error, return the request
        return request

    # If you got here, then an error happened, so need to exit
    sys.exit(-1)


# Use a request to extract a table into a pandas dataframe
def extract_table(req : requests.Response, tablenum: int) -> pd.DataFrame:

    # Parse the html to find the right table
    soup = BeautifulSoup(req.text,'html.parser')
    tables = soup.find_all('table')

    print(tables[0].prettify())
    return
    # Check for errors in data pulled
    if len(tables) - 1 < tablenum:
        SystemExit("Cannot grab table #{} from {} possible table(s).".format(tablenum,len(tables)))
    table = tables[tablenum]


    # Extract table body, replacing breaks with real new line characters. If there is no table body, just use the table itself
    tablebody = table.find('tbody') if table.find('tbody') else table
    for br in tablebody.find_all('br'):
        br.replace_with('\n')

    # Extract headers from table body
    headers = [i.text.strip() for i in tablebody.find_all('th')]

    # Extract table records from data in tablebody, make sure to remove blank lists and text
    content = [tuple(x) for x in [[td.text.strip() for td in tr.find_all('td')] for tr in tablebody.find_all('tr')] if x]
    
    # Convert list of tuples to dataframe with given column names by headers, and return
    return pd.DataFrame(content,columns=headers)


# Write a dataframe to a given xlsx filepath
def dataframe_to_xlsx(df : pd.DataFrame, filepath : str):
    
    # Check if the file name is for an xlsx spreadsheet
    if os.path.splitext(filepath)[1] == '.xlsx':
        # Write to file
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Data', index=False)
        writer.save()
    else:
        print("Invalid xlsx file name or missing extension.")


# Simple function that makes any filepath unique by adding a number
def uniquefile(path : str) -> str:
    # Pull the file name without the file extension, initialize the number to add
    file, extension = os.path.splitext(path)
    counter = 2

    # While the file already exists in the directory, try again with a new number added
    while os.path.exists(path):
        path = file + str(counter) + extension
        counter += 1
    
    # Return the new (or same) path to file
    return path


# When running as main, run main
if __name__ == "__main__":
    main()
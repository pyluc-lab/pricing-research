import pandas as pd
import pathlib
import win32com.client as win32
import pythoncom
import logging
import os
import time
from modules.constants import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException


def set_logging():
    """
    Sets up logging for the application. Creates a directory for logs if it doesn't exist
    and configures the logging basic settings. Validates the log file name, log level,
    and encoding format.

    Exceptions are caught and logged as errors if any issue arises during the setup.
    """
    try:
        # Set the directory for logs
        log_dir = 'Logs'
        
        # Create the directory if it doesn't exist
        os.makedirs(log_dir, exist_ok=True)


        # Define the log file path
        log_file = os.path.join(log_dir, 'app.log')

        # Check if the log file name is a valid string ending with '.log'
        if not isinstance(log_file, str) or not log_file.endswith('.log'):
            raise ValueError("File name must end with '.log'.")

        # Set the log level
        log_level = logging.INFO
        # Validate the log level
        if log_level not in [logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL]:
            raise ValueError("Invalid log level.")

        # Set the encoding for the log file
        encoding = 'UTF-8'
        # Validate the encoding
        if not isinstance(encoding, str) or not encoding.isascii():
            raise ValueError("Codification must be a valid ASCII string")

        # Configure basic logging settings
        logging.basicConfig(
            filename=log_file,
            level=log_level,
            encoding=encoding
        )

        # Log information if the log file is created successfully
        logging.info(f"Log file created in: {log_file}")
    except Exception as e:
        # Log an error message if any exception occurs
        logging.error(f"Unexpected error while creating log file: {e}")
    

def set_outlook():
    """
    Initializes Microsoft Outlook using COM automation. This function only supports
    Windows systems and will log an error if attempted on other operating systems.

    Returns the Outlook application object if successful, or None if an error occurs.
    """
    
    # Check if the operating system is Windows
    if os.name != 'nt':
        logging.error("This function is only supported by Windows systems")
        return None

    try:
        # Initialize the COM library
        pythoncom.CoInitialize()
        
        # Create a new instance of the Outlook application
        outlook = win32.Dispatch("outlook.application")

        # Log successful creation of Outlook
        logging.info("Outlook created successfully")
        return outlook
    except Exception as e:
        # Log any exceptions during Outlook initialization
        logging.error(f"Unexpected error while creating outlook: {e}")
        return None

def set_browser():
    """
    Initializes a Chrome browser instance using WebDriver with automatic driver management.
    
    Returns the Chrome WebDriver object if successful, or None if an error occurs.
    """
    
    try:
        # Install and set up the ChromeDriver using WebDriver Manager
        service = Service(ChromeDriverManager().install())
        
        # Initialize the Chrome browser
        browser = webdriver.Chrome(service=service)
        
        return browser
    except Exception as e:
        # Log any exceptions during browser initialization
        logging.error(f"Error: Could not initialize browser: {e}")
        return None
        
def load_dataframe(data_base_file_path=None):
    """
    Loads an Excel file into a pandas DataFrame.

    Parameters:
    data_base_file_path (str): The file path to the Excel file. Defaults to "data_base/search.xlsx" if not provided.

    Returns:
    DataFrame or None: Returns a pandas DataFrame if the file is successfully loaded, or None if an error occurs.
    """
    
    # Default file path if no argument is provided
    if data_base_file_path is None:
        data_base_file_path = r"data_base/search.xlsx"

    # Check if the file exists
    if not os.path.exists(data_base_file_path):
        logging.error(f"File not found: {data_base_file_path}")
        return None
    
    try:
        # Load the Excel file into a DataFrame using openpyxl engine
        dataframe = pd.read_excel(data_base_file_path, engine='openpyxl')
        return dataframe
    except Exception as e:
        # Log any exceptions that occur during the file loading
        logging.error(f"Unexpected error while loading dataframe: {e}")
        return None

def format_prices(value):
    """
    Converts a price string formatted in a Brazilian currency style to a float.

    Parameters:
    value (str): The price string (e.g., 'R$ 1.234,56').

    Returns:
    float: The numerical representation of the price.
    """
    # Remove currency symbol and thousands separator, replace decimal comma with dot
    f_value = value.replace('R$', '').replace('.', '').replace(',', '.').strip()
    # Consider only the numerical part before any additional text
    f_value = f_value.split()[0]
    return float(f_value)

def scroll_down(browser, time_pause=0.1):
    """
   Incrementally scrolls down a webpage using the specified browser object.

   Parameters:
   browser: An instance of a Selenium WebDriver used for automating web interaction.
   time_pause (float): Duration to pause between scrolls, in seconds (default is 0.1 seconds).

   Behavior:
   Scrolls the page down by 500 pixels at a time.
   Continues scrolling until reaching 12,500 pixels or detecting the end of the page.
   Includes a final pause to ensure the page has loaded completely.

    """
    scroll = 0
    while scroll < 12500:
        browser.execute_script(f"window.scrollTo(0, {scroll})")  # Execute a JavaScript function on the browser
        scroll += 500
        time.sleep(time_pause)  # Wait 0.2 seconds between scrolls for smoothness

        # Check if the end of the page has been reached
        end_of_page = browser.execute_script("return window.innerHeight + window.scrollY >= document.body.scrollHeight")
        if end_of_page:
            break  # Stop scrolling if the bottom of the page is reached
    time.sleep(1)  # Final wait to allow full page to loa

def wait(browser, element_type, element_name, wait_time=10):
    """
    Waits for a specific element to be present on a webpage.

    Parameters:
    browser: The web browser automation object (e.g., from Selenium WebDriver).
    element_type (str): The type of locating strategy (e.g., 'id', 'name', 'css_selector').
    element_name (str): The name or identifier of the element to locate.
    wait_time (int, optional): Maximum time to wait before timing out.

    Returns:
    None: Logs an error and exits the function if inputs are invalid or an exception occurs.
    """
    # Validate input types
    if not isinstance(element_type, str) or not isinstance(element_name, str):
        logging.error("element_type or element_name must be a str.")
        return

    try:
        # Wait for the presence of the element
        WebDriverWait(browser, wait_time).until(
            EC.presence_of_element_located((getattr(By, element_type.upper()), element_name))
        )
    except Exception as e:
        logging.error(f"An error occurred while waiting for the element: {e}")


def search_on_google_shopping(browser, url, df_search):
    """
    Conducts product searches on Google Shopping using a Selenium browser instance.

    Parameters:
    - browser: Selenium WebDriver instance.
    - url: URL for the Google Shopping page to start the search.
    - df_search: DataFrame containing product details with columns ['product', 'min_price', 'max_price', 'banned_terms'].

    Returns:
    - dict_products: Dictionary with product names as keys and their details (price and link) as values.
    """

    # Validate the URL input
    if not isinstance(url, str):
        logging.error("URL must be a string.")
        return None
    
    # Validate the DataFrame input
    if not isinstance(df_search, pd.DataFrame) or df_search.empty:
        logging.error("Invalid or empty DataFrame.")
        return None
    
    # Initialize dictionary to store searched product details
    dict_products = {}

    # Attempt to load the URL
    try:
        browser.get(url)
    except Exception as e:
        logging.error(f"Unable to load URL: {e}")
        return dict_products

    # Iterate through the products in the DataFrame
    for product_name_on_dataframe in df_search['product']:
        try:
            # Extract price range and banned terms for each product
            min_price = float(df_search.loc[df_search['product'] == product_name_on_dataframe, 'min_price'].values[0])
            max_price = float(df_search.loc[df_search['product'] == product_name_on_dataframe, 'max_price'].values[0])
            banned_terms = str(df_search.loc[df_search['product'] == product_name_on_dataframe, 'banned_terms'].values[0])
            # Split banned terms into list
            banned_terms_list = [term for term in banned_terms.replace(';', ',').split(',')]
        except ValueError as e:
            logging.error(f"Error parsing price for {product_name_on_dataframe}: {e}")
            continue

        # Attempt to perform a search on Google
        try:
            search_bar = browser.find_element(By.ID, 'APjFqb')
            search_bar.send_keys(product_name_on_dataframe + Keys.ENTER)
        except NoSuchElementException:
            logging.error("Search bar not found.")
            continue
        
        # Wait until the shopping tab button is clickable
        wait(browser, 'XPATH', '//*[@id="hdtb-sc"]/div/div/div[1]/div/div[2]')
        
        # Click the Shopping tab button
        try:
            shopping_button = browser.find_element(By.XPATH, '//*[@id="hdtb-sc"]/div/div/div[1]/div/div[2]')
            shopping_button.click()
        except NoSuchElementException:
            logging.error(f"Shopping button not found for {product_name_on_dataframe}.")
            continue
        
        # Wait for the product list to load
        wait(browser, 'CLASS_NAME', 'i0X6df')
        scroll_down(browser)

        # Retrieve and process search results
        try:
            list_result = browser.find_elements(By.CLASS_NAME, 'i0X6df')
        except NoSuchElementException:
            logging.error(f"Results not found for {product_name_on_dataframe}.")
            continue

        # Process each product result
        for element in list_result:
            try:
                # Extract product name and price
                product_name = element.find_element(By.CLASS_NAME, 'EI11Pd').text
                product_price = format_prices(element.find_element(By.CLASS_NAME, 'a8Pemb').text)
                
                # Check conditions and store valid results
                if all(term not in product_name for term in banned_terms_list) and min_price <= product_price <= max_price:
                    product_link = element.find_element(By.TAG_NAME, 'a').get_attribute('href')
                    dict_products[product_name] = {
                        'product_name': product_name,
                        'product_price': f'R${product_price}',
                        'product_link': product_link
                    }
            except (NoSuchElementException, ValueError):
                continue

        # Re-load the initial URL after processing each search result
        try:
            browser.get(url)
        except Exception as e:
            logging.error(f"Error returning to homepage: {e}")

        # Wait until the search bar is ready again
        wait(browser, 'ID', 'APjFqb')

    return dict_products


def search_on_mercado_livre(browser, url, df_search):
    """
    Performs product searches on Mercado Livre using a Selenium browser instance.

    Parameters:
    - browser: Selenium WebDriver instance.
    - url: URL for the Mercado Livre page to start the search.
    - df_search: DataFrame containing product details with columns ['product', 'min_price', 'max_price', 'banned_terms'].

    Returns:
    - dict_products: Dictionary with product names as keys and their details (price and link) as values.
    """

    # Validate the URL input
    if not isinstance(url, str):
        logging.error("URL must be a string.")
        return None
    
    # Validate the DataFrame input
    if not isinstance(df_search, pd.DataFrame) or df_search.empty:
        logging.error("Invalid or empty DataFrame.")
        return None
    
    # Initialize dictionary to store searched product details
    dict_products = {}

    # Attempt to load the URL
    try:
        browser.get(url)
    except Exception as e:
        logging.error(f"Unable to load URL: {e}.")
        return dict_products

    # Iterate through the products in the DataFrame
    for product_name_on_dataframe in df_search['product']:
        try:
            # Extract price range and banned terms for each product
            min_price = float(df_search.loc[df_search['product'] == product_name_on_dataframe, 'min_price'].values[0])
            max_price = float(df_search.loc[df_search['product'] == product_name_on_dataframe, 'max_price'].values[0])
            banned_terms = str(df_search.loc[df_search['product'] == product_name_on_dataframe, 'banned_terms'].values[0])
            # Split banned terms into list
            terms_by_comma = banned_terms.split(', ')
            banned_terms_list = [item for term in terms_by_comma for item in term.split(';')]
        except ValueError as e:
            logging.error(f"Error parsing price for {product_name_on_dataframe}: {e}")
            continue
        
        # Wait for the search bar to be ready
        wait(browser, 'ID', 'cb1-edit')
        
        # Attempt to perform a search on Mercado Livre
        try:
            search_bar = browser.find_element(By.ID, 'cb1-edit')
            search_bar.send_keys(product_name_on_dataframe + Keys.ENTER)
        except NoSuchElementException as e:
            logging.error("Search bar not found.")
            continue
        
        # Wait for the product list to load and scroll down the page
        wait(browser, 'CLASS_NAME', 'poly-card__content')
        scroll_down(browser)

        # Retrieve and process search results
        try:
            list_result = browser.find_elements(By.CLASS_NAME, 'poly-card__content')
        except NoSuchElementException as e:
            logging.error(f"Results not found for {product_name_on_dataframe}.")
            continue
        
        # Process each product result
        try:
            for element in list_result:
                try:
                    # Extract product name and price
                    product_name = element.find_element(By.TAG_NAME, 'h2').text
                    product_price = format_prices(element.find_element(By.CLASS_NAME, 'andes-money-amount__fraction').text)

                    # Check conditions and store valid results
                    for term in banned_terms_list:
                        if term in product_name:
                            continue
                        if min_price <= product_price <= max_price:
                            product_link = element.find_element(By.TAG_NAME, 'a').get_attribute('href')
                            dict_products[product_name] = {
                                'product_name': product_name, 
                                'product_price': f'R${product_price}', 
                                'product_link': product_link
                            }
                except (NoSuchElementException, ValueError):
                    continue
        except Exception as e:
            logging.error(f"Error processing product info for {product_name_on_dataframe}: {e}")

        # Re-load the initial URL after processing each search result
        try:
            browser.get(url)
        except Exception as e:
            logging.error(f"Error returning to homepage: {e}")
        
        # Wait until the search bar is ready again
        wait(browser, 'ID', 'cb1-edit')

    return dict_products


def search_on_amazon(browser, url, df_search):
    """
    Performs product searches on Amazon using a Selenium browser instance.

    Parameters:
    - browser: Selenium WebDriver instance.
    - url: URL for the Amazon page to start the search.
    - df_search: DataFrame containing product details with columns ['product', 'min_price', 'max_price', 'banned_terms'].

    Returns:
    - dict_products: Dictionary with product names as keys and their details (price and link) as values.
    """

    # Validate the URL input
    if not isinstance(url, str):
        logging.error("URL must be a string.")
        return None
    
    # Validate the DataFrame input
    if not isinstance(df_search, pd.DataFrame) or df_search.empty:
        logging.error("Invalid or empty DataFrame.")
        return None
    
    # Initialize dictionary to store searched product details
    dict_products = {}

    # Attempt to load the URL
    try:
        browser.get(url)
    except Exception as e:
        logging.error(f"Unable to load URL: {e}.")
        return dict_products

    # Iterate through the products in the DataFrame
    for product_name_on_dataframe in df_search['product']:
        try:
            # Extract price range and banned terms for each product
            min_price = float(df_search.loc[df_search['product'] == product_name_on_dataframe, 'min_price'].values[0])
            max_price = float(df_search.loc[df_search['product'] == product_name_on_dataframe, 'max_price'].values[0])
            banned_terms = str(df_search.loc[df_search['product'] == product_name_on_dataframe, 'banned_terms'].values[0])
            # Split banned terms into list
            terms_by_comma = banned_terms.split(', ')
            banned_terms_list = [item for term in terms_by_comma for item in term.split(';')]
        except ValueError as e:
            logging.error(f"Error parsing price for {product_name_on_dataframe}: {e}")
            continue
        
        # Wait for the search bar to be ready
        wait(browser, 'ID', 'twotabsearchtextbox')
        
        # Attempt to perform a search on Amazon
        try:
            search_bar = browser.find_element(By.ID, 'twotabsearchtextbox')
            search_bar.send_keys(product_name_on_dataframe + Keys.ENTER)
        except NoSuchElementException as e:
            logging.error("Search bar not found.")
            continue
        
        # Wait for the product list to load and scroll down the page
        wait(browser, 'CLASS_NAME', 's-asin')
        scroll_down(browser)

        # Retrieve and process search results
        try:
            list_result = browser.find_elements(By.CLASS_NAME, 's-asin')
        except NoSuchElementException as e:
            logging.error(f"Results not found for {product_name_on_dataframe}.")
            continue
        
        # Process each product result
        try:
            for element in list_result:
                try:
                    # Extract product name and price
                    product_name = element.find_element(By.TAG_NAME, 'h2').text
                    product_price = format_prices(element.find_element(By.CLASS_NAME, 'a-price').text)

                    # Check conditions and store valid results
                    for term in banned_terms_list:
                        if term in product_name:
                            break
                    else:
                        if min_price <= product_price <= max_price:
                            product_link = element.find_element(By.CLASS_NAME, 'a-link-normal').get_attribute('href')
                            dict_products[product_name] = {
                                'product_name': product_name, 
                                'product_price': f'R${product_price}', 
                                'product_link': product_link
                            }
                except (NoSuchElementException, ValueError):
                    continue
        except Exception as e:
            logging.error(f"Error processing product info for {product_name_on_dataframe}: {e}")

        # Re-load the initial URL after processing each search result
        try:
            browser.get(url)
        except Exception as e:
            logging.error(f"Error returning to homepage: {e}")
        
        # Wait until the search bar is ready again
        wait(browser, 'ID', 'twotabsearchtextbox')

    return dict_products

def creating_Dataframe_with_results(dict_):
    """
    Converts a dictionary containing product details into a Pandas DataFrame.

    Parameters:
    - dict_: Dictionary where keys are product names and values are dictionaries
             with product details (such as price and link).

    Returns:
    - df_: DataFrame containing product details, indexed by default integer index.
    
    Raises:
    - Logs an error if dict_ is not a dictionary or is empty.
    - Logs an error if an exception occurs during DataFrame creation.
    """

    # Check if the input is a dictionary
    if not isinstance(dict_, dict):
        logging.error("Invalid dict format.")
        return None

    # Check if the dictionary is empty
    if len(dict_) < 1:
        logging.error(f"Dict {dict_} is empty.")
        return None

    try:
        # Convert dictionary to DataFrame with the dictionary keys as index
        df_ = pd.DataFrame.from_dict(dict_, orient='index')
        # Reset index to default integer index and drop the existing index
        df_.reset_index(drop=True, inplace=True)
    except Exception as e:
        # Log any exceptions encountered during DataFrame creation
        logging.error(f"Unexpected error while making dict into DataFrame: {e}.")
        return None

    return df_

def get_or_create_folder_for_results_file():
    """
    Ensures there is a directory for storing results files. If the directory
    doesn't exist, it will be created.

    Returns:
    - backup_path: Path object representing the directory for results files.

    Error Handling:
    - Logs an error if there is an exception during directory creation.
    - Logs an informative message if the directory is created successfully.
    """

    # Define the directory path for storing results files
    backup_path = pathlib.Path('results')

    # Check if the directory exists; if not, try to create it
    if not backup_path.is_dir():
        try:
            # Create directory, including parent directories as needed
            backup_path.mkdir(parents=True, exist_ok=True)
            logging.info(f"Folder: '{backup_path}' created successfully.")
        except Exception as e:
            # Log any errors encountered during directory creation
            logging.error(f"Error while creating folder '{backup_path}': {e}")
            return None

    # Return the Path object representing the directory
    return backup_path

def creating_excel_file_with_dataframes(dataframe, backup_path, file_name):
    """
    Saves a given DataFrame to an Excel file within a specified directory.

    Parameters:
    - dataframe: The DataFrame to be written to the Excel file.
    - backup_path: Path object representing the directory where the file will be saved.
    - file_name: Name of the Excel file to be created (without extension).

    Error Handling:
    - Logs an error if the input is not a valid DataFrame.
    - Logs an error if the backup path does not exist.
    - Logs an error if the file name is invalid or missing.
    """
    
    # Check if the provided dataframe is a valid pandas DataFrame
    if not isinstance(dataframe, pd.DataFrame):
        logging.error(f"{dataframe} is not a valid DataFrame.")
        return
    
    # Check if the provided backup path is a valid existing path
    if not backup_path.exists():
        logging.error(f"{backup_path} is not a valid path.")
        return
    
    # Check if the file name provided is valid and non-empty
    if not file_name:
        logging.error("Invalid file name.")
        return
    
    # Save DataFrame to an Excel file at specified location
    dataframe.to_excel(backup_path / f'{file_name}.xlsx', index=False)
    
def send_email(email_address, outlook, backup_path):
    """
    Sends an email with attachments from a specified directory.

    Parameters:
    - email_address: The recipient's email address.
    - outlook: The Outlook application object used to create and send the email.
    - backup_path: Path object representing the directory containing files to be attached.

    Error Handling:
    - Logs an error if any attachment file cannot be found.
    - Handles general exceptions during email creation and sending.
    """
    try:
        # Create a new email item using Outlook
        mail = outlook.CreateItem(0)
        mail.To = email_address
        mail.Subject = "Results of pricing research script."
        mail.Body = 'Pipipi Popopo'

        # List all files in the backup directory
        files_backup_folder = backup_path.iterdir()
        files_names = [file.name for file in files_backup_folder if file.is_file()]

        # Attach each file to the email
        for name in files_names:
            attachment = pathlib.Path.cwd() / backup_path / name
            if not attachment.exists():
                logging.error(f"Error: File {attachment} not found.")
                continue  # Skip to the next file
            mail.Attachments.Add(str(attachment))

        # Send the email
        mail.Send()
        logging.info(f"Email sent to {email_address} successfully.")

    except IndexError as e:
        logging.error(f"Error: Could not find data for email. Details: {e}")
    except FileNotFoundError as e:
        logging.error(f"Error: File not found. Details: {e}")
    except Exception as e:
        logging.error(f"Unexpected error while sending emails: {e}")


def main():
    """
    Main function to execute the price research script.

    - Sets up logging and the Outlook application.
    - Loads a DataFrame for search terms.
    - Initiates a web browser for web scraping.
    - Searches products on Google Shopping, Mercado Livre, and Amazon.
    - Creates DataFrames from search results and stores them in Excel files.
    - Sends an email with the results' file attachments.
    """
    
    # Setup logging and Outlook
    set_logging()
    outlook = set_outlook()

    # Load search terms into DataFrame
    df_search = load_dataframe()

    # Set up the web browser for scraping
    browser = set_browser()

    # Perform searches and store results
    products_results_on_google_shopping = search_on_google_shopping(browser, GOOGLE_URL, df_search)
    products_results_on_mercado_livre = search_on_mercado_livre(browser, MERCADO_LIVRE_URL, df_search)
    products_results_on_amazon = search_on_amazon(browser, AMAZON_URL, df_search)

    # Close the web browser
    browser.quit()

    # Create DataFrames from the search results
    google_shopping_df = creating_Dataframe_with_results(products_results_on_google_shopping)
    mercado_livre_df = creating_Dataframe_with_results(products_results_on_mercado_livre)
    amazon_df = creating_Dataframe_with_results(products_results_on_amazon)

    # Prepare the backup directory for results
    backup_path = get_or_create_folder_for_results_file()

    # Create Excel files with the results
    creating_excel_file_with_dataframes(google_shopping_df, backup_path, 'Google_Shopping')
    creating_excel_file_with_dataframes(mercado_livre_df, backup_path, 'Mercado_Livre')
    creating_excel_file_with_dataframes(amazon_df, backup_path, 'Amazon')

    # Send an email with results as attachments
    send_email(EMAIL_ADDRESS, outlook, backup_path)


try:
    if __name__ == "__main__":
        main()
except Exception as e:
    logging.critical(f"Unexpected error while excetuing script: {e}")
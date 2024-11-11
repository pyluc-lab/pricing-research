# Price Search with Selenium

## Description

This program was developed to work with the Brazilian versions of the websites. Selenium is extremely sensitive to changes made by site developers. Any changes in the websites that alter the XPATH, class names, IDs, etc., will require a minor revision of this code to also update the XPATH, class names, IDs, etc., in the corresponding site functions.

This project automates price searches on online platforms such as Google Shopping, Mercado Livre, and Amazon. It utilizes Selenium, a powerful tool for browser automation, enabling direct navigation on target pages. Although price comparison could be done via APIs, using Selenium was intentional for testing and improving skills with this library, focusing on DOM manipulation, data extraction, and process automation in a real navigation environment.

The system captures information configured in an Excel file (`search.xlsx`), allowing users to specify the product, desired price range, and terms to exclude to refine searches. The results are exported as `.xlsx` files and automatically emailed to a designated address.

## Features

* **Selenium Automation:** Real-time search by accessing browser pages and interacting directly with site elements.
* **Search Configuration via Excel:** Customize products of interest, price limits, and exclusion conditions.
* **Integrated Export:** Results are exported in Excel format.
* **Automatic Email Sending:** Integration with Outlook to send results directly via email.
* **Detailed Logging:** Logging system that tracks errors and provides status information of the execution.

```shell
> data_base
    search.xlsx
> Logs
    app.log
> modules
    __init__.py
    constants.py
> README
    readme eng
    readme port
> results
    (files will be saved here after the search)
main.py
Copy
Copy
```

### Directory and File Description

* **data_base/** : Contains crucial files for search definition, e.g., `search.xlsx`.
* **Logs/** : Maintains an activity log of the application in the file `app.log`.
* **modules/** : Includes essential modules such as `__init__.py` and `constants.py`.
* **constants.py** : The user must insert a valid email to receive the result files.
* **README/** : Includes complete project documentation available in English and Portuguese.
* **results/** : Directory designated to store `.xlsx` result files post-search execution.
* **main.py** : Main file that initiates the project execution.

## Usage Instructions

1. **Email Configuration:**
   * Open `modules/constants.py`.
   * Enter a valid email address to ensure receipt of the generated files.
2. **Search Configuration in `search.xlsx`:**
   * Open the `search.xlsx` file located in `data_base/`.
   * Enter the name of the product you wish to search.
   * Define the desired minimum and maximum purchase price.
   * Add keywords that, if found, will disregard the search to avoid irrelevant results.
3. **Project Execution:**
   * Run `main.py` to start the price search.
   * Check the `results/` directory to access the generated `.xlsx` files.
   * The operation log will be available in `Logs/app.log`.

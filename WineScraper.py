# Program that looks through a given location of Vinmonopolet and lists all
# alcholic drinks by their price per cl alcohol
#
# Outputs to an xlsx file (MS Office Excel)
#
# Reqs.:
#   bs4, selenium, pandas, openpyxl, and mozillas geckodriver.exe
#
# Made by Augustin Winther... Bless this mess of code <3
#

# Import selenium stuff
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as DriverWait
from selenium.webdriver.support import expected_conditions as expct_cond
from selenium.webdriver.firefox.service import Service as firefox_service
from selenium.webdriver.firefox.options import Options as firefox_options

from bs4 import BeautifulSoup
from bs4.dammit import UnicodeDammit
from os import path
from math import ceil as round_up
import re
import sys
import pandas as pd
import os

def absolute_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = path.dirname(__file__)
    return path.join(base_path, relative_path)

def string_to_float(string):
    string = string.replace(",",".")
    string = float(re.sub("[^\d.]", "", string))
    return string

def wait_for_element(driver=None, xpath=None):
    element = DriverWait(driver, 30).until(expct_cond.presence_of_element_located((By.XPATH, xpath)))
    return element

def initiate_web_driver(driver_type=None, driver_path=None):
    if driver_type.lower() == "firefox":
        options = firefox_options()
        options.add_argument('--headless')
        service = firefox_service(log_path=os.devnull, executable_path=driver_path)
        driver = webdriver.Firefox(options=options, service=service)
    else:
        driver = None
    return driver

class Store(object):
    def __init__(self, name, amount, button):
        self.name = name
        self.amount = amount
        self.button = button

# PROGRAM STARTS HERE #
if __name__ == "__main__":
    
    # Initiate web driver
    print("\n Looking for installed web browsers...", end='\r')
    try:
        driver = initiate_web_driver(driver_type='firefox', 
                                     driver_path=absolute_path('./driver/geckodriver.exe'))
        print(" Using Firefox as web browser!              ")
    except Exception:
        input(" Couldn't find any compatible browsers...\n"
              " Please install Firefox: https://www.mozilla.org/en-US/firefox/new/")
        sys.exit()

    # Connect to www.vinmonopolet.no
    print("\n Connecting to www.vinmonopolet.no...", end='\r')
    try:
        driver.get("https://www.vinmonopolet.no/search?q=:relevance&searchType=product&currentPage=0")
        print(" Connected to www.vinmonopolet.no!        ")
    except Exception:
        driver.close()
        input(" Couldn't connect to www.vinmonopolet.no. Close the program and try again...\n")
        sys.exit()

    # Find the "Butikker" (stores) button and click it
    try:
        stores_button = wait_for_element(driver, xpath='//button[@class="expandable__header expandable__icon--right"]')
    finally:
        stores_button.click()

    # Get a list all the html store link items
    store_link_items = driver.find_elements(By.XPATH, value='//ul[@class="facet__list facet__list--scroll-list"]//li')

    # Go through all html store link items. 
    # Extract info. Create Store object with info and add it to store_list
    store_list = []
    for item in store_link_items:
        button = item.find_element(by="tag name", value='button')
        name = item.find_element(by="class name", value='facet-value__name').text
        amount = item.find_element(by="class name", value='facet-value__count').text

        # Force name to Unicode
        name = UnicodeDammit(name).unicode_markup

        # Remove characters from 'amount' and turn it into an int
        amount = int(re.sub("[^\d]","", amount))

        # Create a Store object and add it to store_list
        this_store = Store(name, amount, button)
        store_list.append(this_store)

    # Ask user for which store to index
    input_store_name = input("\n Please type the Vinmonopolet location you'd wish to check out: ")
    while len(input_store_name) < 2:
        input_store_name = input("\n Please type the Vinmonopolet location you'd wish to check out: ")

    # Check which store(s) the input matches
    store_to_check = None
    possible_stores = []
    for store in store_list:
        if input_store_name.lower() == store.name.lower():
            store_to_check = store
            break
        elif input_store_name.lower() in store.name.lower():
            possible_stores.append(store)
            store_to_check = possible_stores

    # If store does not exists => quit.
    # If multiple stores share input name, than ask user which store they mean
    if store_to_check == None:
        driver.close()
        print(f"\n Couldn't find any Vinmonopolet stores at the location: {input_store_name}")
        input(f" Close the program and try again...")
        sys.exit()
    elif store_to_check == possible_stores:
        print(f"\n We found several locations. Which one do you want to use?")
        for store in possible_stores:
            index = possible_stores.index(store)
            print(f" [{index}] {store.name}")
        print("")
        answer = None
        while answer not in range(len(possible_stores)):
            answer = input(f" Please type the number corresponding to the location: ")
            try:
                answer = int(answer)
            except Exception:
                answer = None
        store_to_check = possible_stores[answer]
    
    # Create data frame to store all product info with header
    all_products_df = pd.DataFrame(columns=[ 'id',
                                             'name',
                                             'type',
                                             'price',
                                             'volume',
                                             'alcohol_percent',
                                             'alcohol_volume',
                                             'price_per_alcohol_volume',
                                             'link'])

    # Get new url for the store to check byt clicking the store link
    # and wait for the url in the driver to change
    previous_url = driver.current_url
    store_to_check.button.click()
    while (driver.current_url == previous_url):
        pass

    # Create base_url with varaible page number
    base_url = driver.current_url
    base_url = base_url[0:-1]  # Removes page number at end of url as this needs to be variable

    # Get the last page's page number
    products_per_page = 24
    product_amount = store_to_check.amount
    last_page_num = round_up(product_amount / products_per_page) - 1  # -1 since the first page_num = 0

    # Go thorugh all products in the store
    print(f"\n Going through Vinmonopolet at {store_to_check.name}!")
    page_num = 0
    while page_num <= last_page_num:
        url = base_url + str(page_num)
        
        # Wait until the page has loaded (Using first product name as reference)
        driver.get(url)
        wait_for_element(driver, xpath='//div[@class="product__name"]')

        # Get the html code and turn it intp soup. Then find all products
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        product_list = soup.findAll('li', {'class' : 'product-item'})

        # Go through all products and extract their delish info
        for product in product_list:
            # Print progress
            product_num = page_num * products_per_page + product_list.index(product)+1
            print(f" Progress: {product_num}/{product_amount}", end="\r")

            # Get the url of the product
            product_link = product.find('a', {'class' : 'product-item__image-container'})['href']
            product_link = "https://www.vinmonopolet.no" + product_link
            
            # Wait until the page has loaded (Using h1 product title as load reference)
            driver.get(product_link)
            wait_for_element(driver, xpath='//h1')
            
            # Get the html code and turn it intp soup
            html = driver.page_source
            soup = BeautifulSoup(html, "html.parser")

            # Try to assign all item values. 
            # If it doesn't work, the product is most likely not an alcoholic drink
            try:
                product_name = soup.find('h1', {'class' : 'product__name'}).text
                product_type = soup.find('p', {'class' : 'product__category-name'}).text
                product_price = soup.find('span', {'class' : 'product__price'}).text
                product_volume = soup.find('span', {'class' : 'product__amount'}).text
                product_alcohol_percent = soup.find('span', {'class' : 'product__contents-list__content-percentage'}).text
            except Exception:
                continue

            # Skip if product is alcohol free
            if product_alcohol_percent == "0%":
                continue
            elif "alkoholfritt" in product_type.lower():
                continue

            # Format variables
            product_id = int(product_link.split("/")[-1])
            product_name = UnicodeDammit(product_name).unicode_markup
            product_type = UnicodeDammit(product_type).unicode_markup
            product_price = string_to_float(product_price)
            product_volume = string_to_float(product_volume)
            product_alcohol_percent = string_to_float(product_alcohol_percent)
            product_alcohol_volume = round(product_volume * (product_alcohol_percent/100), 1)
            product_price_per_alcohol_volume = round(product_price / product_alcohol_volume, 2)

            # Create product df and it to the main data frame
            product_df = pd.DataFrame({ 'id'    : [product_id],
                                        'name'  : [product_name],
                                        'type'  : [product_type],
                                        'price' : [product_price],
                                        'volume': [product_volume],
                                        'alcohol_percent': [product_alcohol_percent],
                                        'alcohol_volume' : [product_alcohol_volume],
                                        'price_per_alcohol_volume' : [product_price_per_alcohol_volume],
                                        'link'  : [product_link]})
            all_products_df = pd.concat([all_products_df, product_df], ignore_index=True)
        
        page_num += 1
    
    print(f" Finished: {product_num}/{product_amount}!")
    
    # Sort the data by price_per_alcohol_volume   
    print("\n Sorting alcohol products by NOK per centi liters alcohol...")
    all_products_df.sort_values(["price_per_alcohol_volume"], 
                                axis=0, 
                                ascending=[True], 
                                inplace=True)
                            
    # Export the data to a excel file in same directory as program
    documents_path = path.expanduser('~\\Documents')
    if not path.exists(documents_path):
        documents_path = path.expanduser('~\\Dokumenter')
    xlsx_file = f"Alkohol ({store_to_check.name}).xlsx"
    xlsx_file = path.join(documents_path, xlsx_file)
    dupe_num = 2
    while path.exists(xlsx_file):
        xlsx_file = f"Alkohol ({store_to_check.name}) ({dupe_num}).xlsx"
        xlsx_file = path.join(documents_path, xlsx_file)
        dupe_num += 1
    xlsx_file_path = absolute_path(xlsx_file)
    all_products_df.to_excel(xlsx_file_path, index=False)

    # Close the program
    driver.close()
    input(f" Done! Data saved to: {xlsx_file}\n\n You can now exit the program...")
    sys.exit()
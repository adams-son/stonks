from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import csv
from urllib3.exceptions import NewConnectionError
import json
import re
import itertools

firstTime = True
bad_stocks = 0

def getSymbols(URL):
    
    max_retries = 5
    for i in range(max_retries):
        try:
            driver.get(URL)
            # rest of your code
            break
        except NewConnectionError:
            if i < max_retries - 1:  # i is zero indexed
                print(f"Connection refused. Retrying in 5 seconds... (attempt {i+1})")
                time.sleep(5)
            else:
                print("Failed to establish a new connection after multiple attempts.")
                raise

    
    with open('output.csv', 'a', newline='') as file:
        writer = csv.writer(file)

        # Wait until the tables are loaded
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "(//table[contains(@class, 'u-standardTable')])[position()<=2]")))

        tables = driver.find_elements(By.XPATH, "(//table[contains(@class, 'u-standardTable')])[position()<=2]")


        row_data = []

        for i in range(len(tables[0].find_elements(By.XPATH, ".//tbody/tr"))):

            row0 = tables[0].find_elements(By.XPATH, ".//tbody/tr")[i]
            row1 = tables[1].find_elements(By.XPATH, ".//tbody/tr")[i]

            column0 = row0.find_elements(By.XPATH, "td")[1]
            column1 = row1.find_elements(By.XPATH, "td")[0]
            column2 = row1.find_elements(By.XPATH, "td")[7]
            column3 = row1.find_elements(By.XPATH, "td")[8]
            
            span0 = column0.find_element(By.XPATH, ".//span")
            class0 = span0.get_attribute('class')

            row_data.append([class0, column1.text, column2.text, column3.text])

        # Write the list to the CSV
        for row in row_data:
            writer.writerow(row)

        # driver.quit()

def getValues(symbol):
    global firstTime
    global bad_stocks

    if  firstTime:
        driver.get("https://finance.yahoo.com/")
        # time.sleep(3)

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "scroll-down-btn")))
        scroll_button = driver.find_elements(By.ID, "scroll-down-btn")
        scroll_button[0].click()

        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//button[@type='submit' and @value='reject']")))
        button = driver.find_elements(By.XPATH, "//button[@type='submit' and @value='reject']")
        button[0].click()
        firstTime = False
        time.sleep(3)

        driver.delete_cookie("A1")
        cookie = {
        "name": "A1",
        "value": "d=AQABBHaGLmUCEPjxbhR-Wzz5ThxeSssDTUYFEgABCAHTL2VhZfW6b2UBAiAAAAcIc4YuZWOGcAs&S=AQAAAvIYgxNKFc-6J-N5gtJbiQw",
        "domain": ".yahoo.com",
        "path": "/",
        "secure": True,
        "expiry": int(time.time()) + 2 * 24 * 60 * 60  # Expiry date is 2 days from now
        }

        driver.add_cookie(cookie)
        
    url1 = f"https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/{symbol}?lang=en-US&region=US&padTimeSeries=true&type=quarterlyNetPPE%2CquarterlyCurrentAssets%2CquarterlyCurrentLiabilities%2CquarterlyEBIT%2CquarterlyEnterpriseValue&merge=false&period1=493590046&period2=1698254026&corsDomain=finance.yahoo.com"
    driver.get(url1)
    
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//pre")))
        elements = driver.find_elements(By.XPATH, "//pre")

        data = json.loads(elements[0].text)
        # enterprise_value = data["timeseries"]["result"][0]["quarterlyEnterpriseValue"][-1]["reportedValue"]["raw"]
        # ebit = data["timeseries"]["result"][1]["quarterlyEBIT"][-1]["reportedValue"]["raw"]
        # net_ppe = data["timeseries"]["result"][2]["quarterlyNetPPE"][-1]["reportedValue"]["raw"]
        # current_liabilities = data["timeseries"]["result"][3]["quarterlyCurrentLiabilities"][-1]["reportedValue"]["raw"]
        # current_assets = data["timeseries"]["result"][4]["quarterlyCurrentAssets"][-1]["reportedValue"]["raw"]
        latest_data = {}

        for item in data["timeseries"]["result"]:
            metric_type = item["meta"]["type"][0]  # Get the metric type (e.g., "quarterlyNetPPE")
            metric_data = item[metric_type]
            latest_entry = metric_data[-1]
            value = latest_entry["reportedValue"]["raw"]
            latest_data[metric_type] = value

        sorted_keys = sorted(latest_data.keys())
        latest_data_list = [symbol] + [f"{key}: {latest_data[key]}" for key in sorted_keys]

        output_file = 'newSymbols.csv'
        with open(output_file, mode='a', newline='') as outfile:
            writer = csv.writer(outfile)
            
            sorted_keys = sorted(latest_data.keys())
            latest_data_list = [symbol] + [f"{key}: {latest_data[key]}" for key in sorted_keys]
            writer.writerow(latest_data_list)
        # writer.writerow(latest_data_list)

    except Exception as e:
        print(f"{symbol} + {e}")
        bad_stocks += 1


def formatCSV(row):
    first_column = row[0]
    last_column = row[-1]

    last_word_in_first_column = first_column.split()[-1]

    if row[-2] == "TSX Venture Exchange": #CA
        row[-1] = last_column + ".V" 
    elif row[-2] == "Canadian Securities Exchange":
        row[-1] = last_column + ".CA"
    elif row[-2] == "Toronto Stock Exchange":
        row[-1] = last_column + ".TO" 
    elif last_word_in_first_column == "SE":
        row[-1] = last_column + ".ST"
    elif last_word_in_first_column == "NO": 
        row[-1] = last_column + ".OL"
    elif last_word_in_first_column == "DK": 
        row[-1] = last_column + ".CO"
    elif last_word_in_first_column == "DE": #Maybe not optimal
        row[-1] = last_column + ".DE"
    elif last_word_in_first_column == "IT":
        row[-1] = last_column + ".MI"
    elif last_word_in_first_column == "FI":
        row[-1] = last_column + ".HE"
    elif last_word_in_first_column == "FR":
        row[-1] = last_column + ".PA"
    elif last_word_in_first_column == "BE":
        row[-1] = last_column + ".BR"
    elif last_word_in_first_column == "NL":
        row[-1] = last_column + ".AS"
    elif last_word_in_first_column == "PT":
        row[-1] = last_column + ".LS"

    # Replace spaces and small chars
    row[-1] = row[-1].replace(" ", "-")
    row[-1] = re.sub(r'[a-z]', '', row[-1])

    getValues(row[-1])
    
    return row


def createList(row):
    # Return on Capital = EBIT/(Net Working Capital + Net Fixed Assets)
    # Net Working Capital:
    # Net Working Capital = Current Assets - Current Liabilities

    # Net Fixed Assets:
    # Net Fixed Assets = Gross Fixed Assets - Accumulated Depreciation  (Net PPE)

    # Tangible Capital Employed:
    # Tangible Capital Employed = Net Working Capital + Net Fixed Assets

    # Tangible Capital Employed = 25,000 + 174,533,000 = 174,558,000

    # Earnings Yield = EBIT/Enterprise Value

    # Enterprise Value = Market Capitalization + Total Debt - Cash and Cash Equivalents
    # Market Capitalization = Price of one share * Number of outstanding shares

    #Symbol ,quarterlyCurrentAssets: quarterlyCurrentLiabilities: quarterlyEBIT: quarterlyEnterpriseValue: quarterlyNetPPE: 
    symbol = row[0]
    quarterlyCurrentAssets = float(row[1].split(":")[1].strip())
    quarterlyCurrentLiabilities = float(row[2].split(":")[1].strip())
    quarterlyEBIT = float(row[3].split(":")[1].strip())
    quarterlyEnterpriseValue = float(row[4].split(":")[1].strip())
    quarterlyNetPPE = float(row[5].split(":")[1].strip())

    net_working_capital = quarterlyEBIT / ((quarterlyCurrentAssets - quarterlyCurrentLiabilities) + quarterlyNetPPE)
    earnings_yield = quarterlyEBIT / quarterlyEnterpriseValue

    output_file = 'calculated.csv'
    with open(output_file, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([symbol, net_working_capital, earnings_yield])

def append_row_numbers(sorted_data):
    return [[i + 1] + row for i, row in enumerate(sorted_data)]

def process_csv(file1_path, file2_path, output_path, top_entries=15):
    # Read data from file1
    with open(file1_path, 'r') as file1:
        reader1 = csv.reader(file1)
        data1 = [next(reader1) for _ in range(top_entries)]

    # Read data from file2
    with open(file2_path, 'r') as file2:
        reader2 = csv.reader(file2)
        data2 = list(reader2)

    # Initialize the output data
    output_data = []

    # Compare the values in index 0 and generate output
    for row1 in data1:
        index_0_value = row1[0]
        found_in_file2 = any(index_0_value == row2[0] for row2 in data2)

        if found_in_file2:
            output_data.append([index_0_value, row1[1], 'HOLD'])
        else:
            output_data.append([index_0_value, row1[1], 'BUY'])

    # Check for values only in file2
    for row2 in data2:
        index_0_value = row2[0]
        found_in_file1 = any(index_0_value == row1[0] for row1 in data1)

        if not found_in_file1:
            output_data.append([index_0_value, row2[1], 'SELL'])

    # Write the output data to a new file
    with open(output_path, 'w', newline='') as output_file:
        writer = csv.writer(output_file)
        writer.writerows(output_data)

def copy_except_sell(input_path, output_path):
    with open(input_path, 'r') as input_file:
        reader = csv.reader(input_file)
        data = [row for row in reader if row[2] != 'SELL']

    with open(output_path, 'w', newline='') as output_file:
        writer = csv.writer(output_file)
        writer.writerows(data)

if __name__ == "__main__":

    # clear file
    with open('output.csv', 'w', newline='') as file:
        pass
    with open('newSymbols.csv', 'w', newline='') as file:
        pass
    with open('calculated.csv', 'w', newline='') as file:
        pass


    chrome_options = Options()
    # chrome_options.add_argument("detach")
    # chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    
    offset = 0
    max_res = 100
    while True:
        url = f"https://www.avanza.se/frontend/template.html/marketing/advanced-filter/advanced-filter-template?1697966914150&widgets.marketCapitalInSek.filter.lower=&widgets.marketCapitalInSek.filter.upper=&widgets.marketCapitalInSek.active=true&widgets.numberOfOwners.filter.lower=&widgets.numberOfOwners.filter.upper=&widgets.numberOfOwners.active=true&parameters.startIndex={offset}&parameters.maxResults={max_res}&parameters.selectedFields%5B0%5D=LATEST&parameters.selectedFields%5B1%5D=DEVELOPMENT_TODAY&parameters.selectedFields%5B2%5D=DEVELOPMENT_ONE_YEAR&parameters.selectedFields%5B3%5D=MARKET_CAPITAL_IN_SEK&parameters.selectedFields%5B4%5D=PRICE_PER_EARNINGS&parameters.selectedFields%5B5%5D=DIRECT_YIELD&parameters.selectedFields%5B6%5D=NBR_OF_OWNERS&parameters.selectedFields%5B7%5D=LIST&parameters.selectedFields%5B8%5D=TICKER_SYMBOL"
        
        try:
            getSymbols(url)
            offset += max_res
        except:
            break

    # start_row = 9045

    input_file = 'output.csv'
    output_file = 'newSymbols.csv'
    with open(input_file, mode='r', newline='') as infile:
        reader = csv.reader(infile)

        # for _ in itertools.islice(reader, start_row):
        #     pass

        for row in reader:
            formatCSV(row)

    with open(output_file, mode='r+', newline='') as file:
        reader = csv.reader(file)
        writer = csv.writer(file)

        for row in reader:
            createList(row)

    with open('calculated.csv', mode='r', newline='') as file:
        reader = csv.reader(file)
        data = [row for row in reader]

    sorted_net_working_capital = sorted(data, key=lambda row: float(row[1]), reverse=True)
    sorted_earnings_yield = sorted(data, key=lambda row: float(row[2]), reverse=True)

    sorted_net_working_capital_with_numbers = append_row_numbers(sorted_net_working_capital)
    sorted_earnings_yield_with_numbers = append_row_numbers(sorted_earnings_yield)

    # Write sorted data with row numbers to separate CSV files
    with open('sorted_net_working_capital.csv', mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(sorted_net_working_capital_with_numbers)

    with open('sorted_earnings_yield.csv', mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(sorted_earnings_yield_with_numbers)

    ###################################
    net_working_capital = {}
    with open('sorted_net_working_capital.csv', newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            net_working_capital[row[1]] = int(row[0])

    # Load the data from sorted_earnings_yield.csv and update the dictionary
    with open('sorted_earnings_yield.csv', newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if row[1] in net_working_capital:
                net_working_capital[row[1]] += int(row[0])

    # Create a new CSV file with the combined and sorted data
    with open('combined_data.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        sorted_data = sorted(net_working_capital.items(), key=lambda x: x[1])
        for key, value in sorted_data:
            writer.writerow([key, value])


    process_csv('combined_data.csv', 'memory.csv', 'final.csv')
    copy_except_sell('final.csv', 'memory.csv')

    print(f"Bad Stocks: {bad_stocks}")
    
    driver.quit()  


from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from openpyxl import load_workbook, Workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import datetime
import re
def setup_driver():
    """Set up Selenium WebDriver for Firefox."""
    options = webdriver.FirefoxOptions()
    options.headless = True  # Set to False if you want to see the browser window
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
    return driver

def fetch_technical_indicators(driver):
    tableTA = driver.find_element(By.XPATH, "//a[@data-test='technical-indicators-title']/parent::h2/parent::div//table/tbody")
    rows = tableTA.find_elements(By.XPATH, "./tr")
    
    names = []
    values = []
    actions = []
    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            name = row.find_element(By.XPATH, "./td[1]/span").get_attribute("innerHTML") # Name in first column

            # Extract value (typically in the second column)
            value = row.find_element(By.XPATH, "./td[2]").get_attribute("innerHTML")  # Value in second column

            # Extract action (typically in the third column)
            action = row.find_element(By.XPATH, "./td[3]").get_attribute("innerHTML")  # Action in third column

            # Append the extracted data to the lists
            names.append(name)
            values.append(value)
            actions.append(action)
        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return names, values, actions
def fetch_bankrate(driver):
    tableTA = driver.find_element(By.XPATH, "//table[contains(@class,'centralBankSideBlockTbl')]")
    rows = tableTA.find_elements(By.XPATH, "./tbody/tr")
    
    banks = []
    interestrates = []
    dates = []
    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            bank = row.find_element(By.XPATH, "./td[3]/a").get_attribute("innerHTML") # Name in first column

            # Extract value (typically in the second column)
            interestrate = row.find_element(By.XPATH, "./td[4]").get_attribute("innerHTML")  # Value in second column

            # Extract action (typically in the third column)
            date = row.find_element(By.XPATH, "./td[5]").get_attribute("innerHTML")  # Action in third column

            # Append the extracted data to the lists
            banks.append(bank)
            interestrates.append(interestrate)
            dates.append(date)

        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return  banks, interestrates, dates 
def fetch_indices(driver,index):
    table = driver.find_element(By.XPATH, f"(//div[contains(@class,'text-xs leading-4')])[{index}]")
    
    rows = table.find_elements(By.XPATH, "./div/div[@class='table-row']")
    #print(rows[0].get_attribute("innerHTML"))
    indices = []
    prices = []
    changes = []
    percentchanges = []
    
    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            print(row.get_attribute("innerHTML"))
            indice= row.find_element(By.XPATH, "./div[1]/a").get_attribute("innerHTML")  # Name in first column
            print(indice)
            
            # Extract value (typically in the second column)
            price = row.find_element(By.XPATH, "./div[2]").get_attribute("innerHTML")  # Value in second column
            change = row.find_element(By.XPATH, "./div[3]/span").get_attribute("innerHTML")  # Signal in second column

            # Extract action (typically in the third column)
            percentchange = row.find_element(By.XPATH, "./div[4]/span").get_attribute("innerHTML")

            # Append the extracted data to the lists
            indices.append(indice)
            prices.append(price)
            changes.append(change)
            percentchanges.append(percentchange)
            print(indices, prices, changes, percentchanges)
        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return indices, prices, changes, percentchanges

def fetch_calendar(driver):
    table = driver.find_element(By.XPATH, "//table[@id='economicCalendarData']")
    
    rows = table.find_elements(By.XPATH, "./tbody/tr[position()>1]")
    #print(rows[0].get_attribute("innerHTML"))
    event_times =[]
    curs= [] 
    imps =[]
    events=[]
    actuals = []
    forecasts = []
    previouss = []

    theday = driver.find_element(By.XPATH, "//td[@class='theDay']").get_attribute("innerHTML").replace("&nbsp;", "").strip()
    
    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            event_time = row.find_element(By.XPATH, "./td[1]").get_attribute("innerHTML").replace("&nbsp;", "").strip()  # Name in first column
            cur = row.find_element(By.XPATH, "./td[2]").text.replace("&nbsp;", "").strip()  # Value in second column
            imp = row.find_element(By.XPATH, "./td[3]").text.replace("&nbsp;", "").strip()  # Signal in second column
            event = row.find_element(By.XPATH, "./td[4]/a").get_attribute("innerHTML").replace("&nbsp;", "").strip()  # Action in third column
            actual = row.find_element(By.XPATH, "./td[5]").get_attribute("innerHTML").replace("&nbsp;", "").strip()  # Action in third column
            forecast = row.find_element(By.XPATH, "./td[6]").get_attribute("innerHTML").replace("&nbsp;", "").strip()  # Action in third colum
            previous = row.find_element(By.XPATH, "./td[7]/span").get_attribute("innerHTML").replace("&nbsp;", "").strip()  # Action in third column
            event_times.append(event_time)
            curs.append(cur)
            imps.append(imp)
            events.append(event)
            actuals.append(actual)
            forecasts.append(forecast)
            previouss.append(previous)



            print(event_time, cur, imp, event, actual, forecast, previous)
        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return event_times, curs, imps, events, actuals, forecasts, previouss, theday

def fetch_moving_averages(driver):
    tableMA = driver.find_element(By.XPATH, "//a[@data-test='moving-average-title']/parent::h2/parent::div//table/tbody")
    rows = tableMA.find_elements(By.XPATH, "./tr")
    
    manames = []
    simples = []
    simplesignals = []
    expos = []
    exposignals = []
    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            maname = row.find_element(By.XPATH, "./td[1]/span").get_attribute("innerHTML")  # Name in first column

            # Extract value (typically in the second column)
            simple = row.find_element(By.XPATH, "./td[2]/div/div").get_attribute("innerHTML")  # Value in second column
            simplesignal = row.find_element(By.XPATH, "./td[2]/div/td").get_attribute("innerHTML")  # Signal in second column

            # Extract action (typically in the third column)
            expo = row.find_element(By.XPATH, "./td[3]/div/div").get_attribute("innerHTML") 
            exposignal = row.find_element(By.XPATH, "./td[3]/div/td").get_attribute("innerHTML")  # Signal in third column

            # Append the extracted data to the lists
            manames.append(maname)
            simples.append(simple)
            simplesignals.append(simplesignal)
            expos.append(expo)
            exposignals.append(exposignal)
        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return manames, simples, simplesignals, expos, exposignals



def fetch_pivot_points(driver):
    tablePV = driver.find_element(By.XPATH, "//a[@data-test='pivot-points-title']/parent::h2/parent::div//table/tbody")
    rows = tablePV.find_elements(By.XPATH, "./tr")
    
    pvnames = []
    s3s = []
    s2s = []
    s1s = []
    pps = []
    r1s = []
    r2s = []
    r3s = []
    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            pvname = row.find_element(By.XPATH, "./td[1]/span").get_attribute("innerHTML")

            # Extract pivot point levels (columns 2-8)
            s3 = row.find_element(By.XPATH, "./td[2]/span").get_attribute("innerHTML") 
            s2 = row.find_element(By.XPATH, "./td[3]/span").get_attribute("innerHTML") 
            s1 = row.find_element(By.XPATH, "./td[4]/span").get_attribute("innerHTML") 
            pp = row.find_element(By.XPATH, "./td[5]/span").get_attribute("innerHTML") 
            r1 = row.find_element(By.XPATH, "./td[6]/span").get_attribute("innerHTML") 
            r2 = row.find_element(By.XPATH, "./td[7]/span").get_attribute("innerHTML") 
            r3 = row.find_element(By.XPATH, "./td[8]/span").get_attribute("innerHTML")

            # Append the extracted data to the lists
            pvnames.append(pvname)
            s3s.append(s3)
            s2s.append(s2)
            s1s.append(s1)
            pps.append(pp)
            r1s.append(r1)
            r2s.append(r2)
            r3s.append(r3)
        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return pvnames, s3s, s2s, s1s, pps, r1s, r2s, r3s

def click_daily_button(driver):
    """Click the 'Daily' button to update the data."""
    try:
        daily_button = driver.find_element(By.XPATH, "//button[contains(text(),'Daily')]")
        driver.execute_script("arguments[0].click();", daily_button)
        print("Clicked the 'Daily' button.")
        time.sleep(5)  # Allow time for data update
    except Exception as e:
        print(f"Error clicking the 'Daily' button: {e}")

def extract_icb_data(driver, url):
    driver.get(url)
    time.sleep(5)  # Allow time for page to load

    indices, price, change, percentchange = fetch_indices(driver,1)
    commodity, cprice, cchange, cpercentchange = fetch_indices(driver,2)
    bonds, bprice, bchange, bpercentchange = fetch_indices(driver,3)

    dfindices = pd.DataFrame({
        "Indices": indices,
        "Price": price,
        "Change": change,
        "Percent Change": percentchange
    })

    dfcommodity = pd.DataFrame({
        "Commodity": commodity,
        "Price": cprice,
        "Change": cchange,
        "Percent Change": cpercentchange
    })
    dfbonds = pd.DataFrame({
        "Bonds": bonds,
        "Price": bprice,
        "Change": bchange,
        "Percent Change": bpercentchange
    })
    return dfindices, dfcommodity, dfbonds


def extract_ec_data(driver, url):
    driver.get(url)
    time.sleep(5)  # Allow time for page to load

    event_time, cur, imp, event, actual, forecast, previous, theday = fetch_calendar(driver)
    bank, interestrate, date = fetch_bankrate(driver)

    dfcalendar = pd.DataFrame({
        "Event Time": event_time,
        "Currency": cur,
        "Imp": imp,
        "Event": event,
        "Actual Rate": actual,
        "Forecast Rate": forecast,
        "Previous Rate": previous
    })
    dfbankrate = pd.DataFrame({
        "Bank": bank,
        "Interstrate": interestrate,
        "Date": date
    })

    
    return dfcalendar, dfbankrate, theday


def extract_cur_data(driver, url):
    driver.get(url)
    time.sleep(5)  # Allow time for page to load

    Pair,	Last,	Open,	High,	Low,	Chg,	Chgper,	Time = fetch_curmarket(driver)
    perPair,	per15Minutes,	perHourly,	perDaily, perWeek,	perMonth,	perYTD,	per3Years, pipPair, pip15Minutes, pipHourly, pipDaily, pipWeek, pipMonth, pipYTD, pip3years= fetch_change(driver)
    
    dfcurrencymarket = pd.DataFrame({
        "Pair": Pair,
        "Last": Last,
        "Open": Open,
        "High": High,
        "Low": Low,
        "Change": Chg,
        "Change Percentage": Chgper,
        "Time": Time
    })
    
    dfpercentchange = pd.DataFrame({
        "Pair": perPair,
        "15 Minutes": per15Minutes,
        "Hourly": perHourly,
        "Daily": perDaily,
        "Weekly": perWeek,
        "Monthly": perMonth,
        "Year to Date": perYTD,
        "3 Years": per3Years
    })
    dfpipchange = pd.DataFrame({
        "Pair": pipPair,
        "15 Minutes": pip15Minutes,
        "Hourly": pipHourly,
        "Daily": pipDaily,
        "Weekly": pipWeek,
        "Monthly": pipMonth,
        "Year to Date": pipYTD,
        "3 Years": pip3years
    })
    



    
    return dfcurrencymarket, dfpercentchange, dfpipchange
def fetch_change(driver):
    tablechange = driver.find_element(By.XPATH, "//table[@id='dailyTab']/tbody")
    rows = tablechange.find_elements(By.XPATH, "./tr")
    perpairs = []
    per15minutess = []
    perHourlys = []
    perDailys = []
    perWeeks = []
    perMonths = []
    perYTDs = []
    per3yesarss = []
    pipPairs =[]
    pip15minutess = []
    pipHourlys = []
    pipDailys = []
    pipWeeks = []
    pipMonths = []
    pipYTDs = []
    pip3yesarss = []

    for row in rows:
        try:
            # Extract name (typically in the first column)
            perPair = row.find_element(By.XPATH, "./td[2]/a").get_attribute("innerHTML") # Pair
            per15minutes = row.find_element(By.XPATH, "./td[3]").get_attribute("innerHTML") # 15 Minutes
            perHourly = row.find_element(By.XPATH, "./td[4]").get_attribute("innerHTML") # Hourly
            perDaily = row.find_element(By.XPATH, "./td[5]").get_attribute("innerHTML") # Daily
            perWeek = row.find_element(By.XPATH, "./td[6]").get_attribute("innerHTML") # Weekly
            perMonth = row.find_element(By.XPATH, "./td[7]").get_attribute("innerHTML") # Monthly
            perYTD = row.find_element(By.XPATH, "./td[8]").get_attribute("innerHTML") # Year to Date
            per3years = row.find_element(By.XPATH, "./td[9]").get_attribute("innerHTML") # 3 Years

            perpairs.append(perPair)
            per15minutess.append(per15minutes)
            perHourlys.append(perHourly)
            perDailys.append(perDaily)
            perWeeks.append(perWeek)
            perMonths.append(perMonth)
            perYTDs.append(perYTD)
            per3yesarss.append(per3years)

            
            
            
            
        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    try:
        pip_button = driver.find_element(By.XPATH, "//a[contains(text(),'PIP Change')]")
        driver.execute_script("arguments[0].click();", pip_button)
        print("Clicked the 'PIP' button.")
        time.sleep(5)  # Allow time for data update
    except Exception as e:
        print(f"Error clicking the 'PIP' button: {e}")
   
    time.sleep(5)
    tablepipchange = driver.find_element(By.XPATH, "//table[@id='dailyTab']/tbody")
    piprows = tablepipchange.find_elements(By.XPATH, "./tr")

    for row in piprows:
        try:
            pipPair = row.find_element(By.XPATH, "./td[2]/a").get_attribute("innerHTML") # Pair in first column
            pip15minutes = row.find_element(By.XPATH, "./td[3]").get_attribute("innerHTML") # 15 Minutes in first column
            pipHourly = row.find_element(By.XPATH, "./td[4]").get_attribute("innerHTML") # Hourly in first column
            pipDaily = row.find_element(By.XPATH, "./td[5]").get_attribute("innerHTML") # Daily in first column
            pipWeek = row.find_element(By.XPATH, "./td[6]").get_attribute("innerHTML") # Weekly in first colum
            pipMonth = row.find_element(By.XPATH, "./td[7]").get_attribute("innerHTML") # Monthly in first column
            pipYTD = row.find_element(By.XPATH, "./td[8]").get_attribute("innerHTML") # Year to Date in first column
            pip3years = row.find_element(By.XPATH, "./td[9]").get_attribute("innerHTML") # 3 Years in first column

            pipPairs.append(pipPair)
            pip15minutess.append(pip15minutes)
            pipHourlys.append(pipHourly)
            pipDailys.append(pipDaily)
            pipWeeks.append(pipWeek)
            pipMonths.append(pipMonth)
            pipYTDs.append(pipYTD)
            pip3yesarss.append(pip3years)
        except Exception as e:
            continue


    return perpairs, per15minutess, perHourlys, perDailys, perWeeks, perMonths, perYTDs, per3yesarss, pipPairs, pip15minutess, pipHourlys, pipDailys, pipWeeks, pipMonths, pipYTDs, pip3yesarss

def fetch_curmarket(driver):
    tablecur = driver.find_element(By.XPATH, "//table[@id='cr1']/tbody")
    rows = tablecur.find_elements(By.XPATH, "./tr")


    pairs = []
    lasts = []
    opens = []
    highs = []
    lows = []
    chgs = []
    chgpers = []
    curtimes = []


    
    for row in rows:
        try:
            # Extract name (typically in the first column)
            pair = row.find_element(By.XPATH, "./td[2]/a").get_attribute("innerHTML") # Name in first column
            last = row.find_element(By.XPATH, "./td[3]").get_attribute("innerHTML") # Last price
            open = row.find_element(By.XPATH, "./td[4]").get_attribute("innerHTML") # Open price
            high = row.find_element(By.XPATH, "./td[5]").get_attribute("innerHTML") # High price
            low = row.find_element(By.XPATH, "./td[6]").get_attribute("innerHTML") # Low price
            change = row.find_element(By.XPATH, "./td[7]").get_attribute("innerHTML") # Change
            chgper = row.find_element(By.XPATH, "./td[8]").get_attribute("innerHTML") # Change percentage
            curtime = row.find_element(By.XPATH, "./td[9]").get_attribute("innerHTML") # Time
            # Append the extracted data to the lists
            pairs.append(pair)
            lasts.append(last)
            opens.append(open)
            highs.append(high)
            lows.append(low)
            chgs.append(change)
            chgpers.append(chgper)
            curtimes.append(curtime)

            # Extract value (typically in the second column)

        except Exception as e:
            # Skip rows where extraction fails
            continue
    
    return pairs, lasts, opens, highs, lows, chgs, chgpers, curtimes 

    
def extract_technical_data(driver, url):
    """Extract technical data from the specified URL."""
    driver.get(url)
    time.sleep(5)  # Allow time for page to load

    # Click the 'Daily' button to get the correct data
    click_daily_button(driver)

    # Fetch different tables
    names, values, actions = fetch_technical_indicators(driver)
    manames, simples, simplesignals, expos, exposignals = fetch_moving_averages(driver)
    pvnames, s3s, s2s, s1s, pps, r1s, r2s, r3s = fetch_pivot_points(driver)
    
    
    # Convert fetched data into DataFrames
    dfTA = pd.DataFrame({
        "Name": names,
        "Value": values,
        "Action": actions
    })

    # Moving Averages DataFrame
    dfMA = pd.DataFrame({
        "Name": manames,
        "Simple": simples,
        "Simple Signal": simplesignals,
        "Exponential": expos,
        "Exponential Signal": exposignals
    })

    # Pivot Points DataFrame
    dfPV = pd.DataFrame({
        "Name": pvnames,
        "S3": s3s,
        "S2": s2s,
        "S1": s1s,
        "PP": pps,
        "R1": r1s,
        "R2": r2s,
        "R3": r3s
    })


    return dfTA, dfMA, dfPV

def save_to_excel_multiple_sheets(df_dict, filename="Forex2.xlsx"):
    """Save data from multiple URLs into different sheets in the same Excel file."""
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for sheet_name, dfs in df_dict.items():
            dfTA, dfMA, dfPV = dfs  # Unpack the DataFrames for each URL
            # Save the DataFrames to their respective sheets
            dfTA.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=2)
            dfMA.to_excel(writer, sheet_name=sheet_name, index=False, startrow=20, startcol=2)
            dfPV.to_excel(writer, sheet_name=sheet_name, index=False, startrow=20, startcol=10)
    print(f"Data saved to {filename}")

def shorten_url(url):
    """Extract currency pair (e.g., 'eur-usd') from the URL to use as the sheet name."""
    # Extract the part of the URL that contains the currency pair
    match = re.search(r'/currencies/([a-zA-Z\-]+)-technical', url)
    if match:
        return match.group(1)  # Extracted currency pair (e.g., 'eur-usd')
    return url 



    
def main(urls, filename = f"FOREX_{datetime.datetime.today().strftime("%Y-%m-%d")}.xlsx"):
    """Extract data from multiple URLs and save each sheet immediately."""
    driver = setup_driver()

    try:
        with pd.ExcelWriter(filename, engine='xlsxwriter', mode="w") as writer:  
            
            
            
            url_icb = "https://www.investing.com/"
            dfindices, dfcommodify, dfbond = extract_icb_data(driver, url_icb)
            sheet_name = 'Key Indicators'
            dfindices.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=2)
            dfcommodify.to_excel(writer, sheet_name=sheet_name, index=False, startrow=15, startcol=2)
            dfbond.to_excel(writer, sheet_name=sheet_name, index=False, startrow=30, startcol=2)

            url_ec = "https://www.investing.com/economic-calendar/"
            dfec, dfcbrates, theday = extract_ec_data(driver, url_ec)
            sheet_name = 'Economic Calendar'
            sheet_name2 = 'Central Bank Rates'
            dfec.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3, startcol=2)
            df_theday = pd.DataFrame([[theday]], columns=["Date"])
            df_theday.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, startcol=2)
            dfcbrates.to_excel(writer, sheet_name=sheet_name2, index= False, startrow = 2, startcol = 2)
            print("All data saved to Forex2.xlsx")

            
            url_currencies = "https://www.investing.com/currencies/"
            dfcurrencymarket, dfpercentchange, dfpipchange= extract_cur_data(driver, url_currencies)
            sheet_name = 'Currency Market'
            sheet_name2 = "Percent Change"
            sheet_name3 = "Pip Change"
            dfcurrencymarket.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=2)
            dfpercentchange.to_excel(writer, sheet_name=sheet_name2, index=False, startrow=2, startcol=2)
            dfpipchange.to_excel(writer, sheet_name=sheet_name3, index=False, startrow=2, startcol=2)
            print("Currency Market, Percent Change, Pip Change data saved to Forex2.xlsx")

            for url in urls:
                print(f'Processing: {url}')
                
                # Extract data from the specified URL
                dfTA, dfMA, dfPV = extract_technical_data(driver, url)

                # Get sheet name
                sheet_name = shorten_url(url)

                # Save each DataFrame as soon as it is extracted
                dfTA.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=2)
                dfMA.to_excel(writer, sheet_name=sheet_name, index=False, startrow=20, startcol=2)
                dfPV.to_excel(writer, sheet_name=sheet_name, index=False, startrow=20, startcol=10)
                
                print(f"Saved {sheet_name} to {filename}")
            print(f"All data saved to {filename}")




    except Exception as e:
        print(f"Error in main execution: for currencypairs data {e}")
    
    finally:
        driver.quit()


# Example usage: Call the main function with a list of URLs
urls = [
    "https://www.investing.com/currencies/eur-usd-technical",
    "https://www.investing.com/currencies/gbp-usd-technical",
    "https://www.investing.com/currencies/aud-usd-technical",
    "https://www.investing.com/currencies/usd-jpy-technical",
    "https://www.investing.com/currencies/eur-gbp-technical", 
    "https://www.investing.com/currencies/usd-cad-technical",
    "https://www.investing.com/currencies/usd-chf-technical",
    "https://www.investing.com/currencies/usd-nzd-technical",

      # Add more URLs as needed
]
main(urls)



#https://www.investing.com/economic-calendar/
#https://www.investing.com/currencies/

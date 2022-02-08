import datetime
from selenium import webdriver
import pandas as pd
import xlsxwriter
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


def setup_sheet(workbook, start_date_in):
    # Create a workbook and add a worksheet.
    worksheet = workbook.add_worksheet()
    header_format = workbook.add_format()
    header_format.set_align('center')
    title = 'Weekly Report for Week of ' + start_date_in
    worksheet.merge_range('A1:I1', title, header_format)
    worksheet.write_row('B2', ['Store 1', 'Store 2', 'Store 3', 'Store 4', 'Store 5', 'Store 6', 'Store 7', 'Store 8'],
                        header_format)
    worksheet.write_column('A3', ['INCOME', 'BUSINESS', 'SALES TAX', 'REBATE', 'COMMISSION FEES', 'MO COMMISSION FEES',
                                  'FUEL PROFIT', 'CHECK FEES', 'TOTAL INCOME', '', 'EXPENSES', 'PURCHASE',
                                  'CASH PURCHASE', 'CASH EXPENSE', 'CHECK EXPENSE', 'MAINTENANCE', 'UTILITIES',
                                  'SALES TAX', 'PAYROLL TAX', 'PAYROLL CK', 'PAYROLL CASH', 'INSURANCE', 'RENT', 'LOAN',
                                  'TOTAL EXPENSES', '', '', 'TOTAL PROFIT/LOSS'], header_format)
    category_format = workbook.add_format()
    category_format.set_bg_color('#F4B084')
    category_format.set_bold()
    worksheet.write('A3', 'INCOME', category_format)
    worksheet.write('A13', 'EXPENSES', category_format)
    worksheet.set_column(0, 0, 20)

    sum_format = workbook.add_format()
    sum_format.set_bg_color('#BDD7EE')
    sum_format.set_bold()
    for col_num in range(1, 9):
        col = chr(65 + col_num)
        worksheet.write_formula(10, col_num, '=SUM(%s$4:%s$10)' % (col, col), sum_format)
        worksheet.write_formula(26, col_num, '=SUM(%s$14:%s$26)' % (col, col), sum_format)
        worksheet.write_formula(29, col_num, '=%s$11 - %s$27' % (col, col), sum_format)
    worksheet.write('A11', 'TOTAL INCOME', sum_format)
    worksheet.write('A27', 'TOTAL EXPENSES', sum_format)
    worksheet.write('A30', 'TOTAL PROFIT/LOSS', sum_format)

    return worksheet


# Method to Create and Start the driver
def setup_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('headless')
    chrome_options.add_argument('window-size=1920x1080')
    chrome_options.add_argument("disable-gpu")
    driver = webdriver.Chrome('chromedriver', options=chrome_options)

    return driver


# Method to Login into Crony POS
def login(driver, username, password):
    url = 'https://online.cronypos.com/Login.aspx'
    driver.get(url)  # url of website to login
    driver.find_element(By.ID, 'txtUser').send_keys(username)  # puts the user
    driver.find_element(By.ID, 'txtpassword').send_keys(password)  # puts the password
    driver.find_element(By.TAG_NAME, 'button').click()  # enters the site

    print('Successful Login!')


# Method to Get into a specific store
def go_to_store(driver, store_num):
    store_list = ['6', '7', '4', '3', '11', '29', '30', '99']
    link = 'https://online.cronypos.com/ChangeStore.aspx?ID='
    url = link + store_list[store_num - 1]
    driver.get(url)

    print('At Store ', store_num)


# Method that takes a $xx string and turns it into a float
def get_float(var):
    if type(var) == float:
        return var
    else:
        return float(var[1:].replace(",", ""))


def fill_profit_form(driver, start_date_in):
    # start_date_in = "08/16/2021"
    start_date = datetime.datetime.strptime(start_date_in, "%m/%d/%Y")
    start_text = start_date_in.replace("/", "%2F")
    end_date = start_date + datetime.timedelta(days=6)
    end_date_in = end_date.strftime("%m/%d/%Y")
    end_text = end_date_in.replace("/", "%2F")
    link = "https://online.cronypos.com/index.aspx?comm=REPORTS&dfa=%s&dta=%s" % (start_text, end_text)
    driver.get(link)  # url for reports page


def get_profit(driver, start_date_in):
    fill_profit_form(driver, start_date_in)
    # gets all of the values
    table = driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/div[2]/div[2]/div/table").get_attribute("outerHTML")
    soup = BeautifulSoup(table, 'html.parser')
    soup_table = soup.find_all("table")
    tables = pd.read_html(str(soup_table))[0]
    tables.columns = tables.iloc[0]
    tables = tables[1:]

    # Collects all of the income data PLUS the Payroll Check Expense
    taxable_sales = tables.loc[tables['Sales'] == 'Taxable Sales:', 'Amount'].iloc[0]
    taxable_sales = get_float(taxable_sales)

    nontaxable_sales = tables.loc[tables['Sales'] == 'NONTax Sales:', 'Amount'].iloc[0]
    nontaxable_sales = get_float(nontaxable_sales)

    sales_tax = tables.loc[tables['Sales'] == 'Tax Collected:', 'Amount'].iloc[0]
    sales_tax = get_float(sales_tax)

    rebates = tables.loc[tables['Sales'] == 'Rebates/income:', 'Amount'].iloc[0]
    rebates = get_float(rebates)

    check_fees = tables.loc[tables['Sales'] == 'Fees:', 'Amount'].iloc[0]
    check_fees = get_float(check_fees)

    # This is an Expense
    payroll_checks = tables.loc[tables['Sales'] == 'Payroll Checks:', 'Amount'].iloc[0]
    payroll_checks = get_float(payroll_checks)

    expense_cash = tables.loc[tables['Sales'] == 'Expense Cash:', 'Amount'].iloc[0]
    expense_cash = get_float(expense_cash)

    # combines tax and non-tax sales to get total business
    business = taxable_sales + nontaxable_sales
    print('############################  INCOME  #############################')
    print('Business:' + str(business))  # Cell B4
    print('Sales Tax:' + str(sales_tax))  # Cell B5
    print('Rebate:' + str(rebates))  # Cell B6
    # print('Fuel Profit:' + str(fuel_profit))  # Cell B9
    print('Check Fees:' + str(check_fees))  # Cell B10
    print('Payroll Check:' + str(payroll_checks))  # Cell B22
    print('Expense Cash:' + str(expense_cash))
    return [business, sales_tax, rebates, check_fees, payroll_checks, expense_cash]


def fill_expense_form(driver, start_date_in):
    # start_date_in = "01/01/2021"
    start_date = datetime.datetime.strptime(start_date_in, "%m/%d/%Y")
    end_date = start_date + datetime.timedelta(days=5)
    end_date_in = end_date.strftime("%m/%d/%Y")
    link = "https://online.cronypos.com/indexentry.aspx?comm=checkbook"
    driver.get(link)  # url for reports page
    driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/div[2]/div[1]/div/div/div[1]/div/a").click()
    driver.implicitly_wait(5)
    start = driver.find_element(By.NAME, 'txtFrom')
    start.clear()
    start.send_keys(start_date_in)
    end = driver.find_element(By.NAME, 'txtto')
    end.clear()
    end.send_keys(end_date_in)
    expense_dropdown = Select(driver.find_element(By.ID, "txtStatus"))
    expense_dropdown.select_by_visible_text("EXPENSE")
    driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/div[2]/div[1]/div/div/div[2]/div/div/form/input[2]").click()


def get_expense(driver, start_date_in):
    fill_expense_form(driver, start_date_in)
    sales_tax = ['WEBFILE', 'WEBFILE TAX']
    payroll_tax = ["IRS"]
    purchase_expense = ['ALEXANDER OIL FUEL', 'ZEEE TRADING INC.', 'COZZINI BROS', 'SUBURBAN PROPANE', 'SUBURBAN',
                        'GAMA', 'GROCER SUPPLY', 'ORKIN PEST CONTROL', 'QUEST FUEL', 'THOMAS PETROLUM',]
    rent_expense = ['VASSAR COMMERCIAL PROPERTIES, LLC', 'AUSTIN AFFORDABLE HOUSING COPORATION']
    utility = ['NUCO2', 'TEXAS DISPONSAL', 'TEXAS DISPONSAL SYSTEMS', 'TEXAS DISPONSAL AUTO OAY',
               'TIME WARNER CABLE', 'SPECTRUM', 'WASTE CONNECTION', 'COUNTY LINE SPECIAL UTILITY DISTRICT',
               'GOFORTH SPECIAL UTILITY DISTRICT', 'WASTE MANAGEMENT', 'CREEDMOOR MAHA WATER SUPPLY CORP',
               'TIME WARNEBR CABLE', 'TIME WARNER', 'PEDERNALES ELEC.', 'COUNTRY LINE SPECIAL UTILITY DISTRICT',
               'PEDERNALES ELECTRIC', 'PRO HYGIENIC SERVICES', 'US TEST', 'MAXWELL SPECIAL UTILITY DISTRICT']

    maintenance = ['TEXAS LED', 'TEXAS LED LIGHTING', 'METAL MART',
                   'FELIPE ROMERO', 'FELIPE ROGELIO ROMERO', 'ROGELIO ROMERO', 'Juan Carlos',
                   'JUAN CARLOS HERNANDEZ', 'JUAN HERNANDEZ', 'JUAN C HERNANDEZ', 'JESUS RAMIREZ CASTANEDA',
                   "JESUS RAMIREZ", "HECTOR HERNANDEZ", 'GERMAN SANCHEZ', "PENSKE TRUCK LEASING CO. LP",
                   'HAMILTON ELECTRIC', 'HAMILTON ELECTRIC WORKS, INC', 'HAMILTON ELECTRIC WORKS']

    check_expenses = ['ACCOUNT ANALYSIS SERVICES', "ACCOUNT ALALYSIS", 'ACCOUNT ANALYSIS SERVICES CHARGE',
                      'ACCOUNTANALYSIS SERVICES GHARGE', 'ANALYSIS CHARGE', 'UNITED REFRIGERATION',
                      'CITY OF UHLAND', 'PROFESSIONAL PLOTTER TECH', 'BRECK O`STEEN', 'CUMMINS ALLISON',
                      'LEASE SERVICES', 'TEJAS READY MIX CONCRETE', 'BRINKS INCORPORATED', 'EFREN HERNANDEZ',
                      'COMMERCIAL KITCHEN', 'MAXIMUS UNIVERSAL PROTECTION', 'UNITED REFRIGERATION INC',
                      'RAHIM MOMIN', 'ZAHRA MOMIN', 'RISE BROADBAND', 'A1 PUMP INC', 'BOBBY COOPER', 'I3 POS',
                      'PIRKEY BARBER', 'SIGN EXPO', 'SC GLOBAL', 'SC MAINTENANCE',
                      'ALARM CONNECTION', 'MAS MEDIOS LLC', 'ATT PAYMENT', 'El Show Del Chivo', 'EL SHOW DEL CHIVO',
                      'JOSE DELGADO', 'JRD TRADE COMPANY', 'SERGIO FUENTES', 'SIGN EXPO AUSTIN',
                      'BENJAMIN GONZALEZ TREJO', 'ACE MART', '360 INDUSTRIAL SUPPLY', 'ESCOBEDO GROUP', 'JOSE ZUÑIGA',
                      "MERCHANTBNKCD DISCOUNT", 'MERCHANTBNKCD FEE', 'ONE GAS TEXAS', "ONE GAS TEXAS PR",
                      'TXWORKFORCECOMM', 'TX WORKFORCE COMM', 'TXWORKFORCECOMMI', 'TXWORKFORCE', 'TXWORFORCE',
                      "BANK CARD DISC", 'BANCARD MTOT', 'BANK CARD', 'BANKCARD', 'BANKCARD DISC', 'WILLY NANAYAKKAR',
                      'BANKCARD MTOT', 'UNIVISION', "TEXAS DEPARTMENT OF LICENSING & REGULATION",
                      "MILTON MATEO MALDONADO", 'ALIUM', "ALIUM TECHNOLGY INC.", 'ALIUM TECHNOLOGY',
                      'ALIUM TECHNOLOGY INC.', 'ADM CPA', 'CITY OF AUSTIN', 'TEXAS SDU CHILD SUPPORT',
                      'TEXAS SDUCHILD SUPPORT', 'ADT SECURITY', 'BLUEBONNET', 'CHILD SUPPORT OFFICE',
                      '11TH AGENCY', '11th AGENCY', 'UNITED REFRIGERATION INC', 'MERCHANT BANCK CHARGE BACK',
                      'MERCHANT BANKCD CHARGEBACK', 'ADT SECURITY SERVICES', 'PLATINIUM CHECK SERVICES',
                      'GEIGER COMMUNICATION', 'TEXAS DEPARTMENT OF STATE HEALTH SERVICE', 'KANSAS PAYMENT CENTER',
                      'KEMPER AUTO INSURANCE', 'EMAGINENET', 'EMAGINET TECH.', 'HOME DEPOT', 'MEHAD INSURANCE',
                      'ATM LINK', 'ATMLINK', 'CANON COPIERS', 'SANJEEV SHARMA', 'TEXAS LOTTERY']

    misc_expenses = ['AGA KHAN FOUNDATION', 'AMEX EPAYMENT', 'ASCENTIUM CAPITAL', 'BANCORP SOUTH BANK',
                     'CAPITAL ONE', 'ESTEEM FINANCE LOAN',
                     'CAPITOL ONE', 'DELTA FINANCE', 'FUSION CELLULAR',
                     'HAJJAR PETERS LLP', 'INFINITY INSURANCE', 'MANESIA PARTNERS LTD', 'MIR CONSULTANS',
                     'MIR CONSULTANT', 'AMERICAN HOMES', 'PIONER FINANCIAL',
                     'MIR CONSULTANTS', 'NEW YORK LIFE', 'NILOFAR KAROVADIYA', 'NIZARI PFCU',
                     'PRIMERICA LIFE INSURANCE', 'PRIMERICAA01', 'PROG COUNTY MUT', 'R BANK', 'R BANK ACHTTRANSFER',
                     'ROGERS & WHITLEY LLP.', 'TD AUTO FINANCE', 'VIKAS ANAND']
    ignore_expense = ['ASIF CK SERVICES', 'POCO LOCO 8 LLC']
    num_rent = 0
    num_purchase_expense = 0
    num_sales_tax = 0
    num_payroll_tax = 0
    num_utilities = 0
    num_maintenance = 0
    num_check_expense = 0
    num_misc_expense = 0

    table = driver.find_element(By.ID, "checksTable").get_attribute("outerHTML")
    soup = BeautifulSoup(table, 'html.parser')
    soup_table = soup.find_all("table")
    tables = pd.read_html(str(soup_table))[0]
    not_named = []
    not_used = []
    for row in tables.itertuples(name=None):
        if row[1] != 'Voided':
            if row[4] in sales_tax:
                num_sales_tax += get_float(row[6])
            elif row[4] in payroll_tax:
                # print(row[4])
                num_payroll_tax += get_float(row[6])
            elif row[4] in utility:
                # print(row[4])
                num_utilities += get_float(row[6])
            elif row[4] in maintenance:
                # print(tables['Payeename'][i])
                num_maintenance += get_float(row[6])
            elif row[4] in check_expenses:
                # print(tables['Payeename'][i])
                num_check_expense += get_float(row[6])
            elif row[4] in misc_expenses:
                # print(tables['Payeename'][i])
                num_misc_expense += get_float(row[6])
            elif row[4] in purchase_expense:
                # print(tables['Payeename'][i])
                num_purchase_expense += get_float(row[6])
            elif row[4] in rent_expense:
                # print(tables['Payeename'][i])
                num_rent += get_float(row[6])
            elif row[4] in ignore_expense:
                pass
            elif len(row[4]) > 100:
                not_named.append(row[2])
            else:
                if row[4] not in not_used:
                    not_used.append(row[4])
    print('############################  EXPENSES  #############################')
    print('Misc Expense:' + str(num_misc_expense))
    print('Check Expense:' + str(num_check_expense))
    print('Maintenance:' + str(num_maintenance))
    print('Utilities:' + str(num_utilities))
    print('Sales Tax:' + str(num_sales_tax))
    print('Payroll Tax:' + str(num_payroll_tax))
    print('Vendors that were not processed:')
    not_used.sort()
    print(not_used)
    print('Checks that do not have vendor name:')
    print(not_named)
    return [num_check_expense, num_maintenance, num_utilities, num_sales_tax, num_payroll_tax, num_purchase_expense,
            num_rent]
    # print(type(tables['Amt'][0]))
    # print(tables.columns)
    # print(type(tables))


def get_fuel_profit(driver, store_num, start_date_in):
    """
    Store 1 : Needs Calculations
    Store 2 : Ignore
    Store 3 : 10 Cents x Reg Volume 20 cents x Diesel, Supreme Volume
    Store 4 : Needs Calculation
    Store 5 : Provided
    Store 6 : Ignore
    Store 7 : Ignore
    Store 8 : Provided
    """
    # start_date_in = "08/16/2021"
    if store_num != 2 and store_num != 6 and store_num != 7:
        start_date = datetime.datetime.strptime(start_date_in, "%m/%d/%Y")
        start_text = start_date_in.replace("/", "%2F")
        end_date = start_date + datetime.timedelta(days=6)
        end_date_in = end_date.strftime("%m/%d/%Y")
        end_text = end_date_in.replace("/", "%2F")
        link = "https://online.cronypos.com/index.aspx?comm=gassummary&dfa=%s&dta=%s" % (start_text, end_text)
        driver.get(link)  # url for reports page
        if store_num == 1 or store_num == 4:
            table = driver.find_elements(By.TAG_NAME, "h4")
            reg_cost = get_float(table[0].get_attribute("innerHTML")[-5:])
            sup_cost = get_float(table[1].get_attribute("innerHTML")[-5:])
            plus_cost = (reg_cost * 0.65) + (sup_cost * 0.35)
            diesel_cost = get_float(table[2].get_attribute("innerHTML")[-5:])

            volume_table = driver.find_elements(By.TAG_NAME, "table")[1].get_attribute("outerHTML")
            vol_list = pd.read_html(volume_table)[0].iloc[7]

            reg_profit = float(vol_list['Regular Profit']) - reg_cost * float(vol_list['Regular Vol'])
            plus_profit = float(vol_list['Plus Profit']) - plus_cost * float(vol_list['Plus Vol'])
            sup_profit = float(vol_list['Super Profit']) - sup_cost * float(vol_list['Super Vol'])
            diesel_profit = float(vol_list['Diesel Profit']) - diesel_cost * float(vol_list['Diesel Vol'])
            tot_profit = reg_profit + plus_profit + sup_profit + diesel_profit
            return tot_profit
        elif store_num == 3:
            volume_table = driver.find_elements(By.TAG_NAME, "table")[1].get_attribute("outerHTML")
            vol_list = pd.read_html(volume_table)[0].iloc[7]
            tot_profit = 0.10 * float(vol_list['Regular Vol']) + 0.135 * float(vol_list['Plus Vol']) \
                         + 0.2 * float(vol_list['Plus Vol']) + 0.2 * float(vol_list['Diesel Vol'])
            return tot_profit
        elif store_num == 5:
            table = driver.find_element(By.XPATH,
                                        "/html/body/div[3]/div[1]/div[2]/div/div/div[6]/div/div/div[2]/table/tfoot/tr/td[6]")
            return get_float(table.get_attribute("innerHTML"))
        elif store_num == 8:
            table = driver.find_element(By.XPATH,
                                        "/html/body/div[3]/div[1]/div[2]/div/div/div[5]/div/div/div[2]/table/tfoot/tr/td[6]")
            return get_float(table.get_attribute("innerHTML"))
    return 0.0


def fill_purchase_form(driver, start_date_in):
    # start_date_in = "01/01/2021"
    start_date = datetime.datetime.strptime(start_date_in, "%m/%d/%Y")
    end_date = start_date + datetime.timedelta(days=6)
    end_date_in = end_date.strftime("%m/%d/%Y")
    link = "https://online.cronypos.com/indexentry.aspx?comm=polist"
    driver.get(link)  # url for reports page
    # driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/div[2]/div[1]/div/div/div[1]/div/a").click()
    driver.implicitly_wait(5)
    start = driver.find_element(By.NAME, 'txtFrom')
    start.clear()
    start.send_keys(start_date_in)
    end = driver.find_element(By.NAME, 'txtto')
    end.clear()
    end.send_keys(end_date_in)
    driver.execute_script("arguments[0].click();", WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[3]/div[1]/div[2]/div[1]/div/div/div[2]/div/div/form/input[2]"))))


def get_purchases(driver, start_date_in, expense_amt):
    fill_purchase_form(driver, start_date_in)
    table = driver.find_elements(By.XPATH,
                                 "/html/body/div[3]/div[1]/div[2]/div[3]/div/div/div[2]/div/div/div/div[4]/div[1]/div/table/thead/tr/th[8]")
    return get_float(table[0].get_attribute("innerHTML")[13:]) + expense_amt


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    crony_driver = setup_driver()
    login(crony_driver, 'mike', '1122*')
    price_list = []
    report_start_date = input('Enter the Start Date: ')
    # price_list.append(float(input('Enter the Price of Regular Gas: ')))
    # price_list.append(float(input('Enter the Price of Supreme Gas: ')))
    # price_list.append(float(input('Enter the Price of Diesel Gas: ')))
    file_name_date = datetime.datetime.strptime(report_start_date, "%m/%d/%Y")
    file_name_date = file_name_date.strftime("%m%d%Y")
    book_name = 'Report' + file_name_date + '.xlsx'
    report_workbook = xlsxwriter.Workbook(filename=book_name)
    report_sheet = setup_sheet(report_workbook, report_start_date)
    for store in range(1, 9):
        print()
        go_to_store(crony_driver, store)
        profit_list = get_profit(crony_driver, report_start_date)
        expense_list = get_expense(crony_driver, report_start_date)
        purchase_amt = get_purchases(crony_driver, report_start_date, expense_list[5])
        # Sets the Fuel Profit to 0 and if it has a gas station
        # send it to the fuel profit method to collect data
        fuel_profit = get_fuel_profit(crony_driver, store, report_start_date)
        col = chr(65 + store)
        report_sheet.write_column('%s4' % col,
                                  [profit_list[0], profit_list[1], profit_list[2], 0, 0, fuel_profit, profit_list[3]])
        report_sheet.write_column('%s14' % col,
                                  [purchase_amt, 0, profit_list[5], expense_list[0], expense_list[1], expense_list[2],
                                   expense_list[3], expense_list[4], profit_list[4], 0, 0, expense_list[6]])
    report_workbook.close()
    crony_driver.close()

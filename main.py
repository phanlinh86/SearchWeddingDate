# webdrive downloaded from https://chromedriver.chromium.org/downloads ChromeDriver 99.0.4844.17
# May need to delete TriggerReset key in regedit to prevent Window Defender reset the browser every time open
# https://support.google.com/chrome/forum/AAAAP1KN0B0_pTMBYdpwUE/?hl=en&gpf=%23!topic%2Fchrome%2F_pTMBYdpwUE
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import re

from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import date, timedelta
import pandas as pd
from matplotlib import pyplot as plt

DRIVER_PATH = r"C:\Users\PhanLinh\PycharmProjects\WebScrapping\chromedriver_97.exe"
URL = "https://xemboi.com.vn/xem-ngay-cuoi-hoi"

BRIDE_DATE_LABEL = "cô dâu"
GROOM_DATE_LABEL = "chú rể"
WEDDING_DATE_LABEL = "ngày cưới"


BRIDE_BIRTHDAY = "10-09-1987"
GROOM_BIRTHDAY = "11-02-1986"
WEDDING_DATE_START_MONTH = 4
WEDDING_DATE_END_MONTH = 12
WEDDING_DATE_YEAR = 2022

def getDayMonthYearFromDate(date):
    data = re.split('(\d+)-(\d+)-(\d+)', date)
    day = data[1]
    month = data[2]
    year = data[3]
    return (day,month,year)

def modify_date_on_web(element_dict,date):
    day,month,year = getDayMonthYearFromDate(date)
    element_dict['day'].select_by_value(day)
    element_dict['month'].select_by_value(month)
    element_dict['year'].select_by_value(year)

def get_info_from_ketluan(ket_luan_str):
    data = re.split("(\d+)/100\n(\w+)",ket_luan_str)
    score = data[1]
    result_str = data[2] + data[3]
    return(score,result_str)

def Main():
    web_driver =  webdriver.Chrome(DRIVER_PATH)
    web_driver.get(URL)
    bride_element_dict = {}
    groom_element_dict = {}
    wedding_element_dict = {}
    try :
        form_result = WebDriverWait(web_driver,10).until(EC.presence_of_element_located((By.CLASS_NAME,"form_xemngaycuoi")))
        # Get elements needed to input dates
        list_group_result = form_result.find_elements_by_class_name("list_group")
        # Bride dropdown list elements
        bride_element_dict['day'] = Select(form_result.find_element_by_class_name("ngayv")) # day
        bride_element_dict['month']  = Select(form_result.find_element_by_class_name("thangv")) # month
        bride_element_dict['year']  = Select(form_result.find_element_by_class_name("namv")) # year
        # Groom dropdown list elements
        groom_element_dict['day'] = Select(form_result.find_element_by_class_name("ngayc")) # day
        groom_element_dict['month'] = Select(form_result.find_element_by_class_name("thangc")) # month
        groom_element_dict['year'] = Select(form_result.find_element_by_class_name("namc")) # year
        # Wedding date dropdown list elements
        wedding_element_dict['day'] = Select(form_result.find_element_by_class_name("ngayl")) # day
        wedding_element_dict['month']  = Select(form_result.find_element_by_class_name("thangl")) # month
        wedding_element_dict['year']  = Select(form_result.find_element_by_class_name("naml")) # year
        #list_day.select_by_value("19")
        #list_month.select_by_value("04")
        #list_year.select_by_value("2022")
        for list_group_element in list_group_result:
            """
            if BRIDE_DATE_LABEL not in list_group_element.find_element_by_tag_name("label").text \
                and GROOM_DATE_LABEL not in list_group_element.find_element_by_tag_name("label").text \
                and WEDDING_DATE_LABEL not in list_group_element.find_element_by_tag_name("label").text :
            """
            try:
                list_group_element.find_element_by_tag_name("label")
            except:
                # List button is the one that doesn't have label
                list_button =  list_group_element

        # Input date for bride & groom
        modify_date_on_web(bride_element_dict,BRIDE_BIRTHDAY)
        modify_date_on_web(groom_element_dict, GROOM_BIRTHDAY)
        # Searching
        modify_date_on_web(wedding_element_dict, "01-%02d-%4d" % (WEDDING_DATE_START_MONTH,WEDDING_DATE_YEAR))
        list_button.click()
        # Search for result
        ket_luan_element = WebDriverWait(web_driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ket_luan")))

        # Collect data for all date in the range
        final_result_dict =  {}
        final_result_dict['date'] = []
        final_result_dict['result'] = []
        final_result_dict['score'] = []
        start_date = date(WEDDING_DATE_YEAR, WEDDING_DATE_START_MONTH, 1)
        end_date = date(WEDDING_DATE_YEAR, WEDDING_DATE_END_MONTH, 31)
        delta = timedelta(days=1)
        while start_date <= end_date:
            current_date = start_date.strftime("%d-%m-%Y")
            final_result_dict['date'].append(current_date)
            modify_date_on_web(wedding_element_dict, "%02d-%02d-%4d" % (start_date.day,start_date.month, start_date.year))
            list_button.click()
            time.sleep(3) # wait a while
            ket_luan_element = WebDriverWait(web_driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "ket_luan")))
            #time.sleep(20)  # Wait for new result updated
            start_date += delta
            score, result = get_info_from_ketluan(ket_luan_element.text)
            final_result_dict['result'].append(result)
            final_result_dict['score'].append(score)
            print("Date %s : Score %s/100 - %s " % (current_date,score,result))
        final_result_df = pd.DataFrame(final_result_dict)
        print(final_result_df)
        fig = plt.figure()
        plt.stem(final_result_df['date'],final_result_df['score'].astype(int))
        fig.canvas.set_window_title("Score vs Date")
        final_result_df.to_excel("wedding_score.xlsx")
    finally:
        web_driver.quit()

    #time.sleep(5)
    #web_driver.quit()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    Main()


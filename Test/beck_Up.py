import smtplib
import time
import os
import pandas as pd
import logging

from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, \
    ElementNotVisibleException
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# -----------------------------------------------------
# MANAGE CHROME OPTIONS
# -----------------------------------------------------
chrome_options = Options()
chrome_options.add_argument('disable-notifications')
chrome_options.add_argument("--disable-web-security")
chrome_options.add_argument("--disable-popup-blocking")
# Chrome adds various arguments, if you do not want those arguments added, pass them into excludeSwitches. A common example is to turn the popup blocker back on.
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
# Setting the detach parameter to true will keep the browser open after the driver process has been quit.
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(chrome_options=chrome_options, executable_path="../Drivers/chromedriver.exe")
driver.implicitly_wait(10)
driver.maximize_window()

time.sleep(0.5)

# -----------------------------------------------------
# CUSTOMIZABLE PARAMETERS
# -----------------------------------------------------
url = "https://my.fibank.bg/EBank/public/offices"
ur1 = "https://www.google.com/"
privacy_Chrome_Element = "/html/body/div[2]/div[2]/div[3]/span/div/div/div/div[3]/div[1]/button[2]/div"
fibank_Logo = "/html/body/div[1]/div[1]/nav/div/div[1]/a"
work_Chains = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/ul/li[4]/a/span[1]"
work_Chains1 = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/ul/li[3]/a/span[1]"
img_AllChains = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/button/span[1]"
long_WorkingTime = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/ul/li[3]/a/span[1]"
drop_Down_Menu = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/button"
readText = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/ul/li[1]/a/span[1]"
read_Long_WorkTime = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/button/span[1]"
div_Main_Chain = "/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[5]/div/div/div/ul/li"


# -----------------------------------------------------
# CUSTOMIZABLE FUNCTIONS
# -----------------------------------------------------
# FIND AND VERIFY ELEMENT EXISTENCE
def find_Element(elementXpath):
    log_Files("Start find_Element(): finding element..")
    try:
        time.sleep(0.3)
        element = driver.find_element(By.XPATH, elementXpath)
        time.sleep(0.3)
        if element:
            # print("Find")
            # print('"' + elementXpath + '"')
            log_Files("find_Element(): Element foulnd " + str(elementXpath))
            return (elementXpath)
        else:
            print("not find")
            log_Files("", "find_Element(): Element not foulnd " + str(elementXpath))
            return False
    except NoSuchElementException as e:
        log_Files("startMessage", errorMessage=None)
        log_Files("", "Exception find_Element(): in finding element " + str(elementXpath) + "." + str(e))
        print(e)
        return False
    return True


# CLICK AND VERIFY ELEMENT
def click_Element(elementXpath):
    log_Files("Start click_Element(): clicking element..")
    try:
        driver.find_element(By.XPATH, elementXpath).click()
        time.sleep(1)
        # print("Click")
        # print('"' + elementXpath + '"')
        return (elementXpath)

    except NoSuchElementException as e:
        log_Files("", "Exception click_Element(): in finding element " + str(elementXpath) + "." + str(e))
        print(e)
        return False
    return True


# READ TEXT
def readText(elementXpath):
    log_Files("Start readText(): reading element..")
    try:
        element = driver.find_element(By.XPATH, elementXpath).text
        time.sleep(0.5)
    # print(element)
    except ElementNotVisibleException as e:
        log_Files("", "Exception readText(): in reding element " + str(elementXpath) + "." + str(e))
        print(e)
        return False
    return element


# CREATE EXEL FILE
# create_Xlsx_File("C:\PythonApp","fibank_branches.xlsx")
def create_Xlsx_File(directoryPath, fileName):
    log_Files("Start create_Xlsx_File()..")
    try:
        log_Files("Start create_Xlsx_File(): creating dirPath..")
        if not os.path.isdir(directoryPath):
            os.makedirs(directoryPath)

        filePath = directoryPath + '\\' + fileName

        if not os.path.isfile(filePath):
            log_Files("Start create_Xlsx_File(): creating xlsx file: " + filePath)
            writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
            writer.close()

    except (FileNotFoundError, PermissionError, OSError) as e:
        log_Files("", "Exception create_Xlsx_File(): in creating xlsx file" + ". " + str(e))
        print(e)
        print("Error opening file")

    return filePath


# SEND EMAIL
# send_Email("madimarinova89@gmail.com","ivan951020nikolov@gmail.com","C:\PythonApp\Fibank_branches.xlsx")
def send_Email(from_Sender, to_Receiver, attachment_Path):
    # import smtplib
    #  from email.mime.multipart import MIMEMultipart
    # from email.mime.text import MIMEText
    # from email.mime.base import MIMEBase
    # from email import encoders
    log_Files("Start send_Email(): " + "From: " + str(from_Sender) + ", " + "To: " + str(to_Receiver))
    try:
        # instance of MIMEMultipart
        msg = MIMEMultipart()

        # storing the senders email address
        msg['From'] = from_Sender

        # storing the receivers email address
        msg['To'] = to_Receiver

        # storing the subject
        msg['Subject'] = "???????"

        # string to store the body of the mail
        # body = "??????" Plain
        body = """\
                            <html>
                              <head></head>
                              <body>
                                <p>????????? ?-???/ ?-??,<br><br>
                                ? ????????? ?? ????????? ?? ?? ?????, ???? ????? ???? ?? ??????? ?? ??????? "?????????? RPA ??????????" ,<br>
                                ?? ???????? ???? ??? Github, ????? ? xlsx ??????? ???????? ?? ?????? ???????? ????? ???? ???????<br><br>
                                ??:??? ???? ??? ???? ??? ???? ? ??-????? ???? ?? script-a ? ?????????????. ? ?????? ? ????? ?? ??????, ?????? ????? ?? ?? ?????? ??????????. :)<br><br>
                                ????? ? ??????? ???!<br><br>
                                ????????,<br>
                                ??????? ????????<br>
                                ???: 0893400059
                                </p>
                              </body>
                            </html>
                            """

        # attach the body with the msg instance
        # msg.attach(MIMEText(body, 'plain'))
        msg.attach(MIMEText(body, 'html'))

        # open the file to be sent
        filename = "Fibank_branches.xlsx"
        # attachment = open("C:\PythonApp\Fibank_branches.xlsx", "rb")
        attachment = open(attachment_Path, "rb")

        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

        # To change the payload into encoded form
        p.set_payload((attachment).read())

        # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        # attach the instance 'p' to instance 'msg'
        msg.attach(p)

        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)

        # start TLS for security
        s.starttls()

        # Authentication
        s.login(from_Sender, "ngaj dvpr hnia neqm")

        # Converts the Multipart msg into a string
        text = msg.as_string()

        # sending the mail
        s.sendmail(from_Sender, to_Receiver, text)

        # terminating the session
        s.quit()

    except Exception as e:
        log_Files("", "Exception send_Email(): in sending email" + ". " + str(e))
        if hasattr(e, 'message'):
            print(e.message)
        else:
            print(e)


# LOG FILES
def log_Files(startMessage=None, errorMessage=None):
    try:
        logging.basicConfig(filename="..\Logs\Logs.log", level="DEBUG",
                            format='%(asctime)s : %(levelname)s : %(message)s', datefmt='%d%m%y %I:%M:%S')

        # logger = logging.getLogger()
        # logger.setLevel(logging.DEBUG)

        logging.info(startMessage)
        # logging.debug("This is debug")
        # logging.warning("This is warning")
        # logging.info(endMessage)
        logging.error(errorMessage)
    # logging.info("\n")
    # logging.critical("This is critical\n")

    except Exception as e:
        log_Files("", "Exception log_Files()" + ". " + str(e))
        if hasattr(e, 'message'):
            print(e.message)
        else:
            print(e)


# -----------------------------------------------------
# OPEN URL IN CHROME
# -----------------------------------------------------
# open_URL(url,privacy_Chrome_Element,fibank_Logo)
def open_URL(url, elementXpath, elementXpath1):
    log_Files("*** START STEP 1: open_URL(): url: " + str(url) + " ***")
    try:
        driver.get(url)
        print("url load")

        time.sleep(1)
        if find_Element(elementXpath):
            click_Element(find_Element(elementXpath))

        log_Files("*** START STEP 1: open_URL(): COMPLETED SUCCESSFULLY ***")
    except NoSuchElementException as e:
        log_Files("", "Exception open_URL(): " + str(url) + ". " + str(e))
        print("There is no privacy popup")
        print(e)
        return False
    finally:
        if find_Element(elementXpath1) == False:
            log_Files("", "Exception open_URL(): refresh page. Could not find " + str(elementXpath1))
            print("FALSE")
            driver.refresh()
            time.sleep(0.5)
    return True


# -----------------------------------------------------
# FIND ALL OPEN CHAINS DURING THE WEEKEND
# -----------------------------------------------------
# find_Open_Chains_During_The_Weekend(img_AllChains,drop_Down_Menu,work_Chains,work_Chains1,read_Long_WorkTime)
def find_Open_Chains_During_The_Weekend(elementXpath, elementXpath1, elementXpath2, elementXpath3, elementXpath4):
    log_Files("*** START STEP 2: find_Open_Chains_During_The_Weekend(): find open chains during the weekend ***")
    try:
        find_Element(elementXpath)
        click_Element(elementXpath1)

        if find_Element(elementXpath2) != False or find_Element(elementXpath3) != False:
            try:
                click_Element(elementXpath3)
                print("-----")
            except:
                click_Element(elementXpath2)
                print("-----")

        # click_Element(find_Element(elementXpath1))
        '''
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB * 3)
        time.sleep(0.5)
        #  actions.perform()
        actions.send_keys(Keys.ENTER)
        actions.perform()
        time.sleep(0.5)
        word = readText(elementXpath3)
        time.sleep(1)
        # print(word)
        # print(str(word))
        print("-----")
        my_bytes = word.encode('utf-8')
        print(my_bytes)

        # ?? decode bytes to str
        my_str_again = my_bytes.decode('utf-8')
        print(my_str_again)
        print("-----")
        #
        if word != "? ???????? ???. ?????":
            i = 0
            for i in range(3):
                click_Element(elementXpath2)
                time.sleep(0.5)
                actions.send_keys(Keys.TAB * i)
                actions.send_keys(Keys.ENTER)
                actions.perform()
                time.sleep(0.5)
                if word == "? ???????? ???. ?????":
                    break;
                else:
                    print("Cannot find ? ???????? ???. ?????")
                i += 1
        '''
        log_Files("*** START STEP 2: find_Open_Chains_During_The_Weekend(): COMPLETED SUCCESSFULLY ***")
    except Exception as e:
        log_Files("",
                  "Exception find_Open_Chains_During_The_Weekend(): can not ind open chains during the weekend." + str(
                      e))
        if hasattr(e, 'message'):
            print(e.message)
        else:
            print(e)
    # except ElementNotVisibleException as e:
    #    print(e)


# -----------------------------------------------------
# EXTRACT DESCRIPTION DATA
# -----------------------------------------------------
#  extract_Chain_Data(div_Main_Chain)
def extract_Chain_Data(elementXpath):
    log_Files("*** START STEP 3: extract_Chain_Data(): extract data ***")
    try:
        find_Element(elementXpath)
        allChain_List = []
        for store in (driver.find_elements(By.XPATH, elementXpath)):
            allChain_List.append(store.text)

        # print(allChain_List)
        # print(len(allChain_List))

        newChainList = []
        for i in range(len(allChain_List)):
            li = list(allChain_List[i].split("\n"))
            newChainList.append(li)

        # print(newChainList[0])
        # print(newChainList[8])
        # print('------')
        # print(newChainList[8][3])
        # print(newChainList)

        log_Files("*** START STEP 3: extract_Chain_Data(): COMPLETED SUCCESSFULLY ***")
    except ElementNotVisibleException as e:
        log_Files("", "Exception extract_Chain_Data(): extract data " + str(e))
        print(e)

    return newChainList


# -----------------------------------------------------
# CREATE TABLE AND IMPORT DATA
# -----------------------------------------------------
# import_Data_In_Sheet(div_Main_Chain,"C:\PythonApp","fibank_branches.xlsx")
def import_Data_In_Sheet(elementXpath, directory, filePath):
    log_Files("*** START STEP 4: import_Data_In_Sheet(): import data ***")
    try:
        # dataframe Name and Age columns
        li = extract_Chain_Data(elementXpath)

        df = pd.DataFrame({'??? ?? ????': [],
                           '?????': [],
                           '???????': [],
                           '???.????? ??????': [],
                           '???.????? ??????': []})

        for i in range(len(li)):
            # print(li[i])
            name = li[i][2]
            address = li[i][3]
            if len(li[i]) > 13:
                phone = li[i][11]
                saturday_Work_Time = li[i][8]
                sunday_Work_Time = li[i][9]

            else:
                phone = li[i][9]
                saturday_Work_Time = li[i][7]
                sunday_Work_Time = 'N/A'

            df = df.append(
                {'??? ?? ????': name, '?????': address, '???????': phone, '???.????? ??????': saturday_Work_Time,
                 '???.????? ??????': sunday_Work_Time},
                ignore_index=True)

        fileName = create_Xlsx_File(directory, filePath)
        print(fileName)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
        writer1 = pd.ExcelWriter('..\Reports\Fibank_branches.xlsx', engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        df.to_excel(writer1, sheet_name='Sheet1', index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        writer1.close()

        # Read
        reader = pd.read_excel(fileName)
        print(reader)

        log_Files("*** START STEP 4: import_Data_In_Sheet(): COMPLETED SUCCESSFULLY ***")

    except Exception as e:
        log_Files("", "Exception import_Data_In_Sheet(): can not import data " + str(e))
        if hasattr(e, 'message'):
            print(e.message)
        else:
            print(e)

    return fileName


########################
# MAIN
########################
'''
open_URL(url, privacy_Chrome_Element, fibank_Logo)
find_Open_Chains_During_The_Weekend(img_AllChains, drop_Down_Menu, work_Chains, work_Chains1, read_Long_WorkTime)

extract_Chain_Data(div_Main_Chain)

#import_Data_In_Sheet(div_Main_Chain, "C:\PythonApp", "fibank_branches.xlsx")
#send_Email("madimarinova89@gmail.com","ivan951020nikolov@gmail.com","C:\PythonApp\Fibank_branches.xlsx")
fileName = import_Data_In_Sheet(div_Main_Chain, "C:\PythonApp", "fibank_branches.xlsx")
send_Email("madimarinova89@gmail.com","ivan951020nikolov@gmail.com",fileName)

########################
# TEST
########################

driver.get(url)
print("url load")

'''

body = """\
                    <html>
                      <head></head>
                      <body>
                        <p> ----------- WEEKEND WORKING TIME AUTOMATION DONE-----------<br>
                        </p>
                      </body>
                    </html>
                    """

to = ["madimarinova89@gmail.com"]
def send_Email_EDIT(from_Sender, to_Receiver, body, attachment_Path = None):
    log_Files("Start send_Email(): " + "From: " + str(from_Sender) + ", " + "To: " + str(to_Receiver))
    try:

        msg = MIMEMultipart()
        msg['From'] = from_Sender
        msg['To'] =", ".join(to_Receiver)
        msg['Subject'] = "WEEKEND WORKING TIME AUTOMATION"
        msg.attach(MIMEText(body, 'html'))

        filename = "Fibank_branches.xlsx"
        attachment = open(attachment_Path, "rb")

        p = MIMEBase('application', 'octet-stream')
        p.set_payload((attachment).read())
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(from_Sender, "ngaj dvpr hnia neqm")
        text = msg.as_string()

        s.sendmail(from_Sender, to_Receiver, text)
        s.quit()

    except Exception as e:
        log_Files("", "Exception send_Email_EDIT(): in sending email" + ". " + str(e))
        if hasattr(e, 'message'):
            print(e.message)
        else:
            print(e)

send_Email("madimarinova89@gmail.com", to, body,"C:\PythonApp\Fibank_branches.xlsx")
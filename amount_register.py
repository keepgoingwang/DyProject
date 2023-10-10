import time
import json
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException, ElementNotVisibleException

import requests
import re
from lxml import etree

class Chorme_page():

    def __init__(self,url=None) -> None:
        # url = 'https://www.microsoft.com/zh-cn/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook/'
        self.service = webdriver.ChromeService(executable_path=r'.\Tools\chromedriver-win64\chromedriver.exe')
        self.url = url

    def register_outlook_email(self,email, password):
        self.driver = webdriver.Chrome(service=self.service)
        self.wait = WebDriverWait(self.driver, 10) # 确保页面元素加载完成

        # 打开Outlook注册页面
        self.driver.get(self.url) # https://signup.live.com/signup
        
        # 点击创建邮件
        creat_button = self.driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-outline-primary-white[data-bi-cn="CreateFreeAccount"]')
        creat_button.click()

        # 同意协议
        agree_button = self.driver.find_element(By.ID, 'iSignupAction')
        agree_button.click()

        # 输入邮箱地址

        email_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="新建电子邮件"]')))  # driver.find_element(By.ID, 'MemberName')
        email_element.send_keys(email)

        # 点击下一步按钮
        next_button = self.driver.find_element(By.ID, 'iSignupAction')
        next_button.click()


        # 输入密码
        time.sleep(2)
        password_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="创建密码"]'))) # driver.find_element
        password_element.send_keys(password)

        # 点击下一步按钮
        next_button = self.driver.find_element(By.ID, 'iSignupAction')
        next_button.click()


        # 输入姓,名
        first_name = ''.join(random.sample(list('ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwxyz'),3))
        last_name = ''.join(random.sample(list('ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwxyz'),6))

        first_name_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="姓"]')))
        first_name_element.send_keys(first_name)
        last_name_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="名"]'))) # By.CSS_SELECTOR, 'input[placeholder="名"]
        last_name_element.send_keys(last_name)

        # 点击下一步按钮
        next_button = self.driver.find_element(By.ID, 'iSignupAction')
        next_button.click()

        # 设置日期
        year = random.randint(1978,2000)
        year_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="年"]')))
        year_element.send_keys(year)

        month_element = self.wait.until(EC.element_to_be_clickable((By.ID, 'BirthMonth')))
        month_select = Select(month_element)
        month_select.select_by_index(random.randint(1,12))

        date_element = self.wait.until(EC.element_to_be_clickable((By.ID, 'BirthDay')))
        date_select = Select(date_element)
        date_select.select_by_index(random.randint(1,28))

        
        # 点击下一步按钮
        next_button = self.driver.find_element(By.ID, 'iSignupAction')
        next_button.click()

        print('等待用户进行非人机验证')
        if_done = str(input('是否完成验证(Y or N):'))
        idx = 0
        while if_done.upper() == 'N':
             
            if_done = str(input('请先完成非人机验证(Y or N):'))
            idx += 1
            if idx == 2:
                quit()
        # 隐私政策
        fi = True
        while fi:
            final_button = self.driver.find_element(By.CSS_SELECTOR,"button")
            if final_button:
                final_button.click()
                fi = False
        # 禁止保持登录
        try:
            false_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input#idBtn_Back")))
            false_button.click()
        except NoSuchElementException or TimeoutException:
            print('元素定位失败')

        print('注册成功')
        time.sleep(10)
        # 关闭浏览器
        self.driver.quit()


    def register_heygen_account(self, email, password):
        # self.url = 'https://app.heygen.com/guest/templates?cid=ab45d201'

        self.driver = webdriver.Chrome(service=self.service)
        self.wait = WebDriverWait(self.driver, 30) # 确保页面元素加载完成
        
        # 打开注册页面
        self.driver.get(self.url)

        # 点击使用Heygen
        # THFF_button = self.driver.find_element(By.CSS_SELECTOR, "a.hero-btn.pc-show.w-button[target='_blank'][href*='heygen']")
        # THFF_button.click()

        # 点击跳转注册页面
        
        register_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "span.css-nbgej.btnContent")))
        register_button.click()

        # 输入邮箱,发送验证码
        time.sleep(5)
        account_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Email"]')))
        account_element.send_keys(email)

        time.sleep(1)
        send_code_element = self.driver.find_element(By.CSS_SELECTOR, "button.css-o9bvpj") # [disabled] span.css-7orpxh
        send_code_element.click()

        first_window_handle = self.driver.current_window_handle # 获取第一个窗口句柄

        # 登录 outlook 邮箱获取验证码

        self.driver.execute_script('window.open("https://www.microsoft.com/zh-cn/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook")')
        self.driver.switch_to.window(self.driver.window_handles[1]) # 切换至outlook邮箱

        time.sleep(10)
        signup_button = self.driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-outlook')
        signup_button.click()

        self.driver.switch_to.window(self.driver.window_handles[-1]) # 切换至登录页面
        email_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="电子邮件、电话或 Skype"]')))
        email_element.send_keys(email)

        next_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input#idSIButton9.button_primary")))
        next_button.click()

        password_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="密码"]')))
        password_element.send_keys(password)

        go_next_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input#idSIButton9")))
        go_next_button.click()

        
        time.sleep(2)
        # 确认隐私保护,禁止保持登录
        try:
            con_button = self.driver.find_element(By.CSS_SELECTOR, "span.ms-Button-label.label-117")
            con_button.click()
            false_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input#idBtn_Back")))
            false_button.click()
        except NoSuchElementException or TimeoutException:
            try:
                false_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input#idBtn_Back")))
                false_button.click()
            except NoSuchElementException or TimeoutException:
                print('元素定位失败')
                quit()
        
        # 获取验证码
        print('点击垃圾箱')
        time.sleep(15)
        # rubbish_folder = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.C2IG3.LPIso.oTkSL.iDEcr.wk4Sg[aria-selected="true"]')))
        js = 'document.querySelector(\'div.C2IG3.LPIso.oTkSL.iDEcr[aria-selected="false"] span.gtcPn\').click()'
        cli_conn = True
        while cli_conn:
            try:
                # rubbish_folder.click()
                self.driver.execute_script(js)
                cli_conn = False
                print('点击垃圾箱成功')
            except (ElementClickInterceptedException, ElementNotInteractableException, ElementNotVisibleException):
                print('点击垃圾箱失败，继续点击')
                time.sleep(2)
        # 此处一直未能正常点击垃圾箱（虽然显示点击成功,但并未跳转）一直停留在收件箱，导致后续url出错不能正确获取验证码，需解决。
        # 解决方案：使用js
        
        time.sleep(5)
        print('获取最新的邮件')
        first_element = self.wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="groupHeader今天"]/../child::*[position()=2]'))) # 定位 id=groupHeader今天 的同级下一个div元素
        f_conn = True
        while f_conn:
            try:
                first_element.click() 
                current_url = self.driver.current_url
                # print('current_url',current_url)
                if not current_url.endswith('junkemail'):
                    f_conn = False
                    print('获取最新的邮件成功')
                    print(current_url)
            except (ElementClickInterceptedException, ElementNotInteractableException, ElementNotVisibleException):
                print('获取最新的邮件失败，继续点击')
                time.sleep(2)


        re_code = True
        while re_code:
            try:
                nums = self.driver.find_element(By.XPATH, '//*[@id="UniqueMessageBody"]/div/div/div/div/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr[4]/td/div/p').text
                print(nums)
                if nums:
                    num_code = int(nums[0])
                    re_code = False
                else:
                    if_conn_code = str(input('验证码获取失败,是否重新获取(Y or N):')).upper()
                    if if_conn_code not in ['Y','YES']:
                        print('用户不继续,退出')
                        quit()
            except NoSuchElementException or TimeoutException:
                num_code = int(input('请输入验证码:'))
                re_code = False

        print('验证码:',num_code)

        # 切回第一个页面进行设置
        self.driver.switch_to.window(first_window_handle)
        
        # 输入验证码
        code_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Verification Code"]')))
        code_element.send_keys(num_code)
        code_next_button = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[class="css-7spms1"]')))
        code_next_button.click()

        # 输入密码
        first_password_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="New Password"]')))
        first_password_element.send_keys(password)
        confirm_password_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Confirm New Password"]')))
        confirm_password_element.send_keys(password)

        done_button = self.driver.find_element(By.CSS_SELECTOR, 'button[class="css-8bykl"]')
        done_button.click()
        
        # 确保已经跳转到heygen页面注册成功
        cur_url = self.driver.current_url
        tr = True
        while tr:
            time.sleep(5)
            cur_url_1 = self.driver.current_url
            if cur_url_1 != cur_url:
                tr = False

        print("注册成功")

        self.driver.quit()



def save_emails_to_json(emails):
    with open(r'.\emails.json', 'w') as file:
        json.dump(emails, file)



def email_main():
    url = 'https://www.microsoft.com/zh-cn/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook/'
    # 获取要注册的邮箱数量
    num_emails = int(input("请输入要注册的邮箱数量(最多12个):"))
    while num_emails > 12 or num_emails == '':
        print("数量超过限制或未输入数量，请重新输入")
        num_emails = int(input("请输入要注册的邮箱数量(最多12个):"))

    # 选择是否手动输入账号密码
    
    if_input = str(input("是否自行输入邮箱账号和密码(Yes or No):"))
    
    if_input = if_input.upper()

    while not (if_input == 'YES' or if_input == 'Y' or if_input == 'NO' or if_input == 'N'): 
        print('输入有误,请重新输入')
        if_input = str(input("是否自行输入邮箱账号和密码(Yes or No):")).upper()
        
    
    # 初始化邮箱和密码列表
    emails = {}

    # 逐个注册邮箱
    for i in range(num_emails):
        page = Chorme_page(url)
        if if_input == 'YES' or  if_input == 'Y':
            email = input(f"请输入第{i+1}个邮箱名称：")
            password = input(f"请输入第{i+1}个邮箱密码：")
        else:
            print(f"请输入第{i+1}个邮箱注册")
            account_num = random.randint(5,12)
            email_f = ''.join(random.sample(list('abcdefghijklmnopqrstuvwxyz0123456789'),account_num))
            while email_f[0].isdigit(): # 判断是否第一个元素为数字
                email_f = email_f[1:]
            email_b = str('@outlook.com')
            email = email_f # + email_b
            password_1 = ''.join(random.sample(list('ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwxyz'),8))
            password_2 = ''.join(random.sample(list('0123456789'),2))
            password = str(password_1) + str(password_2)
        
        try:
            page.register_outlook_email(email, password)
            emails[email + email_b] = password
            
            time.sleep(2)

        except Exception as e:
            print(e)
            if_conn = str(input("邮箱注册失败, 是否进行下一个注册(Y or N):")).upper()
            if if_conn not in ['Y', 'YES']:
                print(emails)
                if len(emails) > 0:
                    save_emails_to_json(emails)
                    quit()


        
    # 保存邮箱和密码到JSON文件
    print(emails)
    save_emails_to_json(emails)

    print("邮箱注册完成,并已保存到emails.json文件中")



def heygen_main():

    heygen_url = str(input('请输入HeyGen的注册链接,若不输入则使用默认链接:'))
    if heygen_url == '':
        heygen_url = 'https://app.heygen.com/guest/templates?cid=ab45d201'
    
    if_manual = str(input('是否自定义邮箱账号,密码(Y or N):'))
    if_manual = if_manual.upper()
    while not (if_manual == 'YES' or if_manual == 'Y' or if_manual == 'NO' or if_manual == 'N'): 
        print('输入有误,请重新输入')
        if_manual = str(input("是否自定义邮箱账号,密码(Yes or No):")).upper()
    
    if if_manual == 'Y' or if_manual =='YES':
        con =True
        while con:
            page = Chorme_page(heygen_url)
            account = input(f"请输入账号名称:")
            password = input(f"请输入账号密码:")
            try:
                page.register_heygen_account(account, password)
                print(f'账号注册成功')
            except:
                print(f'账号注册失败')

            if_con = str(input('是否继续进行注册:(Yes or No):')).upper()
            if if_con == 'Y' or if_con == 'YES':
                con = True
            else:
                con = False
        
        print('注册完成, 退出程序')
        quit()


    else:
        con = True
        while con:
            file_path = str(input('请输入批量注册的账号密码文件路径:'))
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    print('预计注册数量:',len(data))
            except:
                print('文件路径有误,需重新输入')
                continue

            for idx, account in enumerate(data.keys()):
                page = Chorme_page(heygen_url)
                password = data[account]
                # try:
                page.register_heygen_account(account, password)
                print(f'第{idx+1}个账号注册成功,进行第下一个')
                time.sleep(2)
                # except:
                #     print(f'第{idx+1}个账号注册失败,进行第下一个')
            print(f'本批次共计{len(data)}注册成功')

            if_con = str(input('是否继续进行批量注册:(Yes or No):')).upper()
            if if_con == 'Y' or if_con == 'YES':
                con = True
            else:
                con = False
            
        
        print('注册完成, 退出程序')
        quit()


    

def main():
    
    choice = str(input('请输入要进行的注册流程(1:outlook邮箱, 2:heygen账号):'))
    if choice == '1':
        email_main()
    elif choice == '2':
        heygen_main()
    else:
        print("输入有误,请重新运行")
        quit()

if __name__ == "__main__":
    
    main()
    
    # {"pxqt3mr@outlook.com": "VUWSzMqH18"}
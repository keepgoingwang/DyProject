import time
import json
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

def register_outlook_email(email, password):
    # 创建Chrome浏览器实例
    service = webdriver.ChromeService(executable_path=r'.\Tools\chromedriver-win64\chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    # 打开Outlook注册页面
    driver.get('https://www.microsoft.com/zh-cn/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook/') # https://signup.live.com/signup
    
    # 点击创建邮件
    creat_button = driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-outline-primary-white[data-bi-cn="CreateFreeAccount"]')
    creat_button.click()

    # 同意协议
    agree_button = driver.find_element(By.ID, 'iSignupAction')
    agree_button.click()
    time.sleep(2)

    # 输入邮箱地址
    email_element = driver.find_element(By.CSS_SELECTOR, 'input[aria-label="新建电子邮件"]')   # driver.find_element(By.ID, 'MemberName')
    email_element.send_keys(email)

    # 点击下一步按钮
    next_button = driver.find_element(By.ID, 'iSignupAction')
    next_button.click()


    # 输入密码
    time.sleep(2)
    password_element = driver.find_element(By.CSS_SELECTOR, 'input[aria-label="创建密码"]')
    password_element.send_keys(password)

    # 点击下一步按钮
    next_button = driver.find_element(By.ID, 'iSignupAction')
    next_button.click()


    # 输入姓,名
    first_name = ''.join(random.sample(list('ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwxyz'),3))
    last_name = ''.join(random.sample(list('ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwxyz'),6))

    first_name_element = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="姓"]')
    first_name_element.send_keys(first_name)
    last_name_element = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="名"]')
    last_name_element.send_keys(last_name)

    # 点击下一步按钮
    next_button = driver.find_element(By.ID, 'iSignupAction')
    next_button.click()

    # 设置日期
    year = random.randint(1978,2000)
    year_element = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="年"]')
    year_element.send_keys(year)

    month_element = driver.find_element(By.ID, 'BirthMonth')
    month_select = Select(month_element)
    month_select.select_by_index(random.randint(1,12))

    date_element = driver.find_element(By.ID, 'BirthDay')
    date_select = Select(date_element)
    date_select.select_by_index(random.randint(1,28))

    print('注册成功')
    # 注册成功后休眠10秒
    time.sleep(10)

    # 关闭浏览器
    driver.quit()

def save_emails_to_json(emails):
    with open(r'.\emails.json', 'w') as file:
        json.dump(emails, file)

def main():
    # 获取要注册的邮箱数量
    num_emails = int(input("请输入要注册的邮箱数量(最多12个):"))
    while num_emails > 12 or num_emails == '':
        print("数量超过限制或未输入数量，请重新输入")
        num_emails = int(input("请输入要注册的邮箱数量(最多12个):"))

    # 选择是否手动输入账号密码
    
    if_input = str(input("是否自行输入邮箱账号和密码(Yes 或者 No):"))
    
    if_input = if_input.upper()

    while not (if_input == 'YES' or if_input == 'Y' or if_input == 'NO' or if_input == 'N'): 
        print('输入有误,请重新输入')
        if_input = str(input("是否自行输入邮箱账号和密码(Yes 或者 No):")).upper()
        
    
    # 初始化邮箱和密码列表
    emails = {}

    # 逐个注册邮箱
    for i in range(num_emails):
        if if_input == 'YES' or  if_input == 'Y':
            email = input(f"请输入第{i+1}个邮箱名称：")
            password = input(f"请输入第{i+1}个邮箱密码：")
        else:
            account_num = random.randint(5,12)
            email_f = ''.join(random.sample(list('abcdefghijklmnopqrstuvwxyz0123456789'),account_num))
            email_b = str('@outloot.com')
            email = email_f # + email_b
            password_1 = ''.join(random.sample(list('ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwxyz'),8))
            password_2 = ''.join(random.sample(list('0123456789'),2))
            password = str(password_1) + str(password_2)
        register_outlook_email(email, password)
        emails[email] = password

    # 保存邮箱和密码到JSON文件
    print(emails)
    save_emails_to_json(emails)
    print("邮箱注册完成，并已保存到emails.json文件中")

if __name__ == "__main__":
    main()
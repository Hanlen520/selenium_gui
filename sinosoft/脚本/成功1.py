from selenium.webdriver.chrome.options import Options
try:
    path = "测试用例/UI测试.xls"
    sheet_name = "test_sheet"
    mobile_emulation = {"deviceName":"iPhone 6"}  
    chrome_options = Options()  
    chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)  
    driver = webdriver.Chrome(chrome_options = chrome_options)
    driver.implicitly_wait(10)
    url = read_by_rowname(path, sheet_name, "url")
    driver.get(url)
    sleep(5)
    js="var q=document.documentElement.scrollTop=300"  
    driver.execute_script(js)
    sleep(2)
    bullets = driver.find_elements_by_class_name("swiper-pagination-bullet")
    bullets[0].click()
    sleep(2)
    driver.find_element_by_class_name("page-footer-right").click()
    sleep(2)
    add1 = driver.find_elements_by_class_name("btn-x-number__icon")
    add1[3].click()
    sleep(1)
    driver.find_element_by_class_name("page-footer-right").click()
    sleep(3)
    driver.find_element_by_class_name("choice-country-search-input").send_keys("澳大利亚")
    sleep(2)
    driver.find_element_by_link_text("澳大利亚").click()
    sleep(2)
    elements = driver.find_elements_by_class_name("mint-header-button")
    elements[5].click()
    sleep(3)
    driver.find_element_by_class_name("page-footer-right").click()
    sleep(3)
    ss=driver.find_elements_by_class_name("mint-field-core")
    ss[0].send_keys("李大嘴")
    sleep(1)
    ss[1].send_keys("411023199610291515")
    sleep(1)
    ss[2].send_keys("18519149575")
    sleep(1)
    ss[3].send_keys("360535694@qq.com")
    sleep(1)
    ss[5].send_keys("lidazui")
    sleep(1)
    driver.find_element_by_class_name("page-footer-right").click()
    sleep(10)
    code = driver.find_element_by_xpath("//*[@id='app']/section/div[2]/div[1]/span[2]").text
    print("code="+code)
    money = driver.find_element_by_class_name("money").text
    print("money="+money)
    if(code!="" and money!=""):
        print_excel("成功","订单编号和金额都不为空",driver)
    sleep(3)
except:
    raise RuntimeError(traceback)
finally:
    driver.quit()
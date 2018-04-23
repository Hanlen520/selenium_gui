try:
    path = "测试用例/UI测试.xls"
    sheet_name = "test_sheet"
    driver = webdriver.Chrome()
    driver.implicitly_wait(10)
    url = readByRowName(path, sheet_name, "url")
    driver.get(url)
    driver.find_element_by_id("kw").send_keys("自动化测试")
    sleep(3)
    driver.find_element_by_id("su").click()
    
except:
    raise RuntimeError(traceback)
finally:
    driver.quit()
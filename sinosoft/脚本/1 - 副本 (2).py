try:
    path = "测试用例/UI测试.xls"
    sheet_name = "test_sheet"
    driver = webdriver.Chrome()
    driver.implicitly_wait(10)
    url = read_by_rowname(path, sheet_name, "url")
    driver.get(url)
    print_excel("555","666",driver)
    
except:
    raise RuntimeError(traceback)
finally:
    driver.quit()
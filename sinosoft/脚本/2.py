try:
    print(2)
    path = "测试用例/UI测试.xls"
    sheet_name = "test_sheet"
    driver = webdriver.Chrome()
    driver.implicitly_wait(10)
    url = read_by_rowname(path, sheet_name, "url")
    driver.get(url)
    sleep(3)
    print_excel("成功","222")
    
except:
    print_excel("失败", str(traceback.format_exc()))
    raise RuntimeError(traceback)
finally:
    driver.quit()
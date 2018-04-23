def go():
    from selenium import webdriver
    from time import sleep
    path = "C:/Users/Administrator/Desktop/python-selenium/epicc.xls"
    sheetName = "abc"
    driver = webdriver.Chrome()
    driver.implicitly_wait(10)
    driver.maximize_window()
    driver.get('https://www.epicc.com.cn/m')
    driver.find_element_by_class_name("immediate-offer").click()
    driver.switch_to.frame("showIframecarIndexNew")
    driver.find_element_by_id("citynameS").click()
    driver.switch_to.frame("overPageIframe")
    sleep(5)
    driver.find_element_by_id("onlynum").click()
    driver.find_element_by_id("onlynum").send_keys(readByRowName(path, sheetName, "city_name"))
    sleep(3)
    driver.find_element_by_xpath("//*[@id='country_0']/a").click()
    driver.find_element_by_class_name("c-field").send_keys(readByRowName(path, sheetName, "tel"))
    driver.find_element_by_class_name("code__button").click()
    sleep(30)
    driver.find_element_by_class_name("home__button--nextstep").click()
    driver.find_element_by_id("placeholder").send_keys(readByRowName(path, sheetName, "car_id"))
    sleep(3)
    c_fields = driver.find_elements_by_class_name("c-field")
    c_fields[2].send_keys(readByRowName(path, sheetName, "engine_number"))
    driver.find_element_by_class_name("u-ellipsis-2").click()
    sleep(3)
    hot_link = driver.find_elements_by_class_name("brand-hot__link")
    hot_link[0].click()
    sleep(3)
    body_link = driver.find_elements_by_class_name("brand-body__link")
    body_link[0].click()
    sleep(3)
    series_field = driver.find_elements_by_class_name("series-field")
    series_field[0].click()
    sleep(3)
    series_field_c = driver.find_elements_by_class_name("series-popup-card__item")
    series_field_c[1].click()
    sleep(3)
    carinfo = driver.find_elements_by_class_name("carinfo-grid__text")
    carinfo[1].click()
    sleep(3)
    car_datetime = driver.find_elements_by_class_name("car-datetime-confirm")
    car_datetime[1].click()
    sleep(3)
    c_fields1 = driver.find_elements_by_class_name("c-field")
    c_fields1[7].send_keys(readByRowName(path, sheetName, "user_name"))
    sleep(3)
    c_fields1[8].send_keys(readByRowName(path, sheetName, "id_card"))
    driver.find_element_by_class_name("carinfo-button--submit").click()
    sleep(15)
    money = driver.find_element_by_class_name("monney").text
    print(money)

if __name__ == "__main__":
    go()







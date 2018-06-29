from selenium import webdriver
from selenium.webdriver.common import action_chains, keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import os
from time import sleep

from sinosoft.test import excel_util

dict_fujiaxian = {}
dict_fujiaxian["003"] = "0";
dict_fujiaxian["004"] = "1";
dict_fujiaxian["005"] = "2";

def login(excel_list):
    driver=webdriver.Ie()
    driver.maximize_window()
    driver.get("http://10.18.121.226:8001/prpall/")
    sleep(1)
    driver.switch_to.frame("fraInterface")
    driver.find_element_by_xpath("//input[@name ='UserCode']").send_keys('1000000000')
    driver.find_element_by_xpath("//input[@name ='Password']").send_keys('0000')
    driver.find_element_by_xpath("//input[@name ='ComCode']").send_keys('0000000000')
    driver.find_element_by_xpath("//input[@name ='RiskCode']").send_keys('1101')
    driver.find_element_by_xpath("//a[@class ='login_but']").click()
    driver.switch_to.default_content()
    #首页左选框
    driver.switch_to.frame("fraMenu")
    driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
    driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
    driver.find_element_by_xpath("//form[@name ='fm']").find_elements_by_tag_name('td')[3].click()
    sleep(0.5)
    driver.find_element_by_xpath("//form[@name ='fm']").find_elements_by_tag_name('td')[4].click()
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    #归属机构
    driver.switch_to.frame("fraInterface")
    ActionChains(driver).double_click(driver.find_element_by_css_selector("[name='ComName']")).perform()
    sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(0)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    sleep(0.5)
    driver.switch_to.default_content()
    #归属业务员
    driver.switch_to.frame("fraInterface")
    ActionChains(driver).double_click(driver.find_element_by_css_selector("[name='Handler1Name']")).perform()
    sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(0)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    sleep(1)
    #境内境外标识
    driver.switch_to.frame("fraInterface")
    AbroadFlag = Select(driver.find_element_by_css_selector("[name='AbroadFlag']"))
    AbroadFlag.select_by_index(2)
    #实际业务员
    ActionChains(driver).double_click(driver.find_element_by_css_selector("[name='HandlerName']")).perform()
    sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(0)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    sleep(1)
    #业务来源
    driver.switch_to.frame("fraInterface")
    driver.find_element_by_xpath("//input[@name ='BusinessCategory' and @value=4]").click()
    sleep(0.5)
    #客户名称
    driver.find_element_by_css_selector('#customername1').send_keys('中国船舶')
    ActionChains(driver).double_click(driver.find_element_by_css_selector('#customername1')).perform()
    driver.switch_to.default_content()
    sleep(1)
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(2)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    sleep(1)
    driver.switch_to.frame("fraInterface")
    driver.find_element_by_xpath("//img[@name ='InsuredImg']").click()
    driver.switch_to.default_content()
    sleep(1)
    js = "var q=document.documentElement.scrollTop=300"
    driver.execute_script(js)
    sleep(2)
    #保险期间
    driver.switch_to.frame("fraInterface")
    for i in range(9):
        driver.find_element_by_xpath("//input[@name ='StartDate']").send_keys(keys.Keys.BACK_SPACE)
    sleep(2)
    driver.find_element_by_xpath("//input[@name ='StartDate']").send_keys(excel_list[2])
    #被保险人信息
    driver.find_element_by_xpath("//img[@name ='AddressImg']").click()
    sleep(0.5)
    #邮政编码
    driver.find_elements_by_xpath("//input[@name ='AddressCode']")[1].send_keys("123321")
    #地址
    driver.find_elements_by_xpath("//input[@name ='AddressName']")[1].send_keys("北京市")

    js = "var q=document.documentElement.scrollTop=300"
    driver.execute_script(js)
    sleep(2)
    #船舶信息(船舶代码)
    now_handle = driver.current_window_handle
    print(now_handle)
    driver.find_element_by_xpath("//img[@name ='ItemShipImg']").click()
    sleep(1)
    driver.find_element_by_css_selector('#ShipCNameId').click()
    sleep(1)
    all_handles = driver.window_handles
    print(all_handles)
    for handle in all_handles:
        if handle != now_handle:
            driver.switch_to.window(handle)
            driver.find_element_by_xpath("//input[@name ='ModelNameSubmit']").click()
            sleep(1)
            driver.find_element_by_css_selector("#commonTable").find_elements_by_tag_name("tbody")[0].find_elements_by_tag_name("tr")[0].find_elements_by_tag_name("td")[0].click()
    driver.switch_to.window(now_handle)
    #船舶类型
    driver.switch_to.frame("fraInterface")
    ActionChains(driver).double_click(driver.find_element_by_xpath("//input[@name ='ShipTypeName']")).perform()
    driver.switch_to.default_content()
    sleep(2)
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(0)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    #船体材料
    sleep(1)
    driver.switch_to.frame("fraInterface")
    ActionChains(driver).double_click(driver.find_element_by_xpath("//input[@name ='shipMeterialName']")).perform()
    driver.switch_to.default_content()
    sleep(2)
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(0)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    #航行范围 / 航程
    sleep(2)
    driver.switch_to.frame("fraInterface")
    driver.find_element_by_xpath("//input[@name ='SailScope']").send_keys('10000')

    js = "var q=document.documentElement.scrollTop=300"
    driver.execute_script(js)
    #主险
    driver.find_element_by_xpath("//img[@name ='ItemKindMainImg']").click()
    driver.find_element_by_xpath("//input[@name ='button_ItemKindMain_Insert']").click()
    sleep(2)
    ActionChains(driver).double_click(driver.find_elements_by_xpath("//input[@name ='KindNameMain']")[1]).perform()
    driver.switch_to.default_content()
    sleep(1)
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(2)
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    sleep(2)
    driver.switch_to.frame("fraInterface")
    driver.find_elements_by_xpath("//input[@name ='AmountMain']")[1].send_keys('10000')
    driver.find_elements_by_xpath("//input[@name ='RateMain']")[1].clear()
    driver.find_elements_by_xpath("//input[@name ='RateMain']")[1].send_keys('4')
    driver.find_elements_by_xpath("//input[@name ='PremiumMain']")[1].click()

    #附加险
    driver.find_element_by_xpath("//img[@name ='ItemKindSubImg']").click()
    sleep(1)
    driver.find_element_by_xpath("//input[@name ='button_ItemKindSub_Insert']").click()
    sleep(2)
    ActionChains(driver).double_click(driver.find_elements_by_xpath("//input[@name ='KindNameSub']")[1]).perform()
    driver.switch_to.default_content()
    sleep(1)
    driver.switch_to.frame("fraCode")
    codeselect=Select(driver.find_element_by_css_selector("[name='codeselect']"))
    codeselect.select_by_index(int(dict_fujiaxian["003"]))
    driver.find_element_by_xpath("//input[@name ='SelectIt']").click()
    driver.switch_to.default_content()
    sleep(2)
    driver.switch_to.frame("fraInterface")
    driver.find_elements_by_xpath("//input[@name ='AmountSub']")[1].send_keys('10000')
    driver.find_elements_by_xpath("//input[@name ='RateSub']")[1].clear()
    driver.find_elements_by_xpath("//input[@name ='RateSub']")[1].send_keys('4')
    driver.find_elements_by_xpath("//input[@name ='PremiumSub']")[1].click()
    sleep(0.5)
    #币别信息
    driver.find_element_by_xpath("//input[@name ='button_Fee_Refresh']").click()

    #保存
    driver.find_element_by_xpath("//input[@name ='buttonSave']").click()
    sleep(2)
    driver.switch_to_alert().accept()

if __name__ == '__main__':
    my_list = excel_util.getExcelList()
    for i in range(len(my_list)):
        print(my_list[i])
        login(my_list[i])
        for j in range(len(my_list[i][7])):
            print(my_list[i][7][j])



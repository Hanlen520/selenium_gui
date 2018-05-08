import traceback
from selenium import webdriver
from time import sleep
import time

def go():
    try:
        driver = webdriver.Chrome()
        driver.implicitly_wait(10)
        driver.get("https://www.epicc.com.cn/m")
        # driver.find_element_by_id("kw").send_keys("自动化测试")
        sleep(3)
        # text = driver.find_element_by_id("setf").text
        # print("text="+text)
        driver.save_screenshot('a.png')
        # driver.find_element_by_id("su").click()
        sleep(9999)
    except:
        raise RuntimeError(traceback)
    finally:
        pass
        # driver.quit()

if __name__ == '__main__':
    go()





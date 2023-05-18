from selenium import webdriver
from PIL import Image
from selenium.webdriver import ActionChains
import os,time,random
import xlrd
from xlwt import *
from xlrd import open_workbook
from xlutils.copy import copy
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
#异常代码
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchFrameException
from selenium.common.exceptions import TimeoutException
import sys
class ExpressInfo:
    def __init__(self, dilivery, receiver,index):
        self.dilivery = dilivery
        self.receiver = receiver
        self.index = index

def get_track(distance):
    track = []
    current = 0
    mid = distance*3/4
    t = random.randint(2, 3)/10
    v=0
    while current < distance:
        if current < mid:
            a = 2
        else:
            a = -3
        v0 = v
        v = v0+a*t
        move = v0*t+1/2*a*t*t
        current += move
        track.append(round(move))
    return track
# 生成拖拽移动轨迹，加3是为了模拟滑过缺口位置后返回缺口的情况
all_num_list = []

def read_excel(workbook,num):
    # 获取所有sheet
    print(workbook.sheet_names())
    # 根据sheet索引或者名称获取sheet内容
    # sheet2 = workbook.sheet_by_index(1)  # sheet索引从0开始
    sheet1 = workbook.sheet_by_name('Sheet1')
    # sheet的名称，行数，列数
    # print(sheet1.name, sheet1.nrows, sheet1.ncols)
    # 获取整行和整列的值（数组）
    # rows = sheet1.row_values(3)  # 获取第四行内容
    # cols = sheet1.col_values(1)  # 获取第2列内容
    # 确保 num 不大于表单的实际行数
    nrows = sheet1.nrows
    if num <= nrows - 1:
        cell = sheet1.cell(num, 0)
        # 获取单元格内容
        value = cell.value
        print(value)
        return value
    else:
        print(f"num 超出表单的实际行数，表单只有 {nrows} 行数据。")
        return None
    # print(sheet1.cell_value(1, 0).encode('utf-8'))
    # print(sheet1.row(1)[0].value.encode('utf-8'))
    # # 获取单元格内容的数据类型
    # print(sheet1.cell(1, 0).ctype)

def get_all_excel(workbook):
    # 获取所有sheet
    print(workbook.sheet_names())
    # 根据sheet索引或者名称获取sheet内容
    # sheet2 = workbook.sheet_by_index(1)  # sheet索引从0开始
    sheet1 = workbook.sheet_by_name('Sheet1')
    # sheet的名称，行数，列数
    # print(sheet1.name, sheet1.nrows, sheet1.ncols)
    # 获取整行和整列的值（数组）
    # rows = sheet1.row_values(3)  # 获取第四行内容
    # cols = sheet1.col_values(1)  # 获取第2列内容
    # 确保 num 不大于表单的实际行数
    nrows = sheet1.nrows
    for num in range(0,nrows):
        all_num_list.append(read_excel(workbook,num))
   



# 打开文件
print('请将待查询的表格放入D盘，第一列为单号，第二列为状态，并将名称修改为‘dan1.xls’\n')
number = 0
n = int(input("请输入每次查询的顺丰快递数量:"))
workbookPath = str(input("请输入xls文件目录"))
# number = int(input("请输入从第几行开始查询:"))
workbook = xlrd.open_workbook(rf'{workbookPath}')
# workbook = xlrd.open_workbook(r'/Users/chenzhe/Desktop/未命名文件夹 3/dan1.xls')
wb = copy(workbook)
ws = wb.get_sheet(0)
get_all_excel(workbook)
danhao = {}
status = ['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
status_xpath = ['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
danhao[read_excel(workbook,0)] = None
options = webdriver.ChromeOptions()
driver_path = "/Users/chenzhe/chromedriver"
# options.add_argument('--headless')  # 使用无头模式，即不打开浏览器窗口
driver = webdriver.Chrome(executable_path=driver_path, options=options)
driver.implicitly_wait(20)
driver.get('http://www.sf-express.com/cn/sc/dynamic_function/waybill/#search/bill-number/')
driver.maximize_window()
driver.implicitly_wait(1)
login_button = driver.find_element(By.XPATH, "//button[contains(text(), '登录/注册')]")
driver.execute_script("arguments[0].click();", login_button)
# login_button.click()
# 等待登录页面消失
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.XPATH, ".//div[@class='svg-wrap cursor-point']")))
time.sleep(1)
agree_button = driver.find_element(By.XPATH, ".//div[@class='svg-wrap cursor-point']")
driver.execute_script("arguments[0].click();", agree_button)
# agree_button.click()
print('请在网页扫码登录');
wait = WebDriverWait(driver, 60)
wait.until_not(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '登录/注册')]")))

def findElement():
            driver.switch_to.default_content()
            #查到当前
            flag1 = 1
            print(f"findElement方法开始调用")
            while flag1:
                time.sleep(0.2)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@class='waybill-detail-wrapper']")))
                waybill_wrapper_1 = driver.find_elements(By.XPATH,'//div[@class="waybill-wrapper no-border"]')
                waybill_wrapper_2 = driver.find_elements(By.XPATH,'//div[@class="waybill-wrapper"]')
                waybill_wrapper_1.extend(waybill_wrapper_2)
                # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//div[@class="waybill-detail-wrapper"]//button[@class="show-details-btn"]')))
                # buttons = driver.find_elements(By.XPATH,'//div[@class="waybill-detail-wrapper"]//button[@class="show-details-btn"]')
                count = 0
                print(f"waybill_wrapper_1迭代开始")
                for index, waybill_wrapper in enumerate(waybill_wrapper_1):
                    numberElement = waybill_wrapper.find_element(By.XPATH,'.//div[@class="bill-num"]//span[@class="number"]')
                    print(f"迭代单号{numberElement.text}")
                    # 使用 ActionChains 模块中的 move_to_element() 方法将页面滚动到该 button 元素所在的位置
                    button = waybill_wrapper.find_element(By.XPATH,'.//button[@class="show-details-btn"]')
                    time.sleep(0.5)
                    # actions = ActionChains(driver)
                    # actions.move_to_element(button)
                    # actions.perform()
                    # time.sleep(1)
                    driver.execute_script("arguments[0].click();", button)
                    print(f"迭代次数{count}")
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@class='pop-wrap']")))
                    pop = driver.find_element(By.XPATH, "//div[@class='pop-wrap']")
                    content = pop.find_element(By.XPATH, ".//div[@class='content']")
                    items = content.find_elements(By.XPATH,".//div[@class='item']")
                    print(f"items:{items}")
                    diliverInfoElement = items[0]
                    receiverElement = items[1]  
                    pElements = diliverInfoElement.find_elements(By.XPATH, ".//div[@class='info']//p")
                    print(f"pElements:{pElements}")
                    diliverNamePhone = pElements[0].text
                    diliverAddress = pElements[1].text
                    diliverInfoStr = diliverNamePhone+","+diliverAddress
                    print(f"diliverInfoStr:{diliverInfoStr}")
                    pRElements = receiverElement.find_elements(By.XPATH, ".//div[@class='info']//p")
                    receiverNamePhone = pRElements[0].text
                    receiverAddress = pRElements[1].text
                    receiverInfoStr = receiverNamePhone+","+receiverAddress
                    print(f"receiverInfoStr:{receiverInfoStr}")
                    time.sleep(0.5)
                    close = pop.find_element(By.XPATH, "//div[@class='svg cursor-point']")
                    driver.execute_script("arguments[0].click();", close)
                    time.sleep(0.5)
                    WebDriverWait(driver, 20).until_not(EC.presence_of_element_located((By.XPATH, "//div[@class='pop-wrap']")))
                    danhao[numberElement.text] = ExpressInfo(diliverInfoStr,receiverInfoStr,index)
                    count += 1
                print(f"danhao:{danhao}")
                for num,info in danhao.items():
                    print(f"fo num:{num}，info:{info}")
                    if((num != None) and (info != None)):
                        ws.write(number-n+info.index, 0, num)
                        ws.write(number-n+info.index, 1, info.dilivery)
                        ws.write(number-n+info.index, 2, info.receiver)
                        print(f"写入{number-n+info.index}{info}")
                        wb.save('/Users/chenzhe/Desktop/未命名文件夹 3/dan1.xls')
                        time.sleep(0.5)
                flag1 = 0
                    # tab = pop.find_element(By.XPATH, "//div[contains(text(), '电子存根')]")
                    # tab.click()
                    
                # for i in range(0, n):
                #         status_xpath[i] = '//*[@id="waybill-'+danhao[i]+'"]/span'
                #         status[i] = driver.find_element_by_xpath(status_xpath[i]).text
                #         if status[i] == '运送中':
                #             status[i] = '待收款'
                #         elif status[i] == '已退回':
                #             status[i] = '拒签退回'
                #         elif status[i] == '已签收':
                #             status[i] = '已签收'
                #         else:
                #             status[i] = '有问题'
                # for i in range(0,n):
                #     ws.write(number-n+i, 2, status[i])
                #     flag1 = 0
                #     wb.save('/Users/chenzhe/Desktop/未命名文件夹 3/dan1.xls')
                #     break

while 1:
    if(number >= len(all_num_list)):
        driver.close()
        sys.exit()
    danhao = {}
    for i in range(0,n):
        danhao[read_excel(workbook,number+i)] = None
    number += n
    flag2=1
    while flag2:
        try:
            delete_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='delete']")))
            driver.execute_script("arguments[0].click();", delete_button)
            # delete_button.click()
            flag2 = 0
        except TimeoutException:
            time.sleep(0.2)
    time.sleep(0.5)
    for key in danhao.keys():
         if key != None:
            input_elem = driver.find_element(By.XPATH, ".//div[@class='input-list']/input")
            input_elem.send_keys(key+' ')  # 输入运单号
            time.sleep(0.1)
       
    time.sleep(0.5)

    search_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@class='search-icon']")))
    driver.execute_script("arguments[0].click();", search_button)
    # search_button.click()
    time.sleep(0)
    flag = 1
    while flag :
        try:
            driver.switch_to.frame('tcaptcha_iframe')
            flag = 0
            time.sleep(0.2)
            # getElementImage(driver.find_element_by_xpath('//*[@id="slideBlock"]'))
            flag = 9
            yundong = [0,260,240,220,260,240,220,260,240,220]
            while flag:
                flag3 = 1
                while flag3:# 直到找到缺块
                    try:
                        thumb = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tcaptcha_drag_thumb"]')))
                        flag3 = 0
                    except TimeoutException:
                        flag3 = 1
                        time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tcaptcha_drag_thumb"]')))
                butten0 = driver.find_element(By.XPATH, '//*[@id="tcaptcha_drag_thumb"]')
                action = ActionChains(driver)
                track_list = get_track(yundong[flag])  # 生成轨迹
                flag -= 1
                time.sleep(0.8)
                action.click_and_hold(butten0)  # 根据轨迹拖拽缺块
                for track in track_list:
                    action.move_by_offset(track, 0)
                action.release(butten0).perform()  # 拖拽缺块
                time.sleep(1.5)#等待缺块归位
                try:
                    thumb = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tcaptcha_drag_thumb"]')))#看是否还能找到缺块
                    time.sleep(0.1)
                    continue
                except TimeoutException:#找不到缺块证明已经查找完
                        print(f"TimeoutException调用方法findElement")
                        findElement()
                        flag = 0
                time.sleep(0.2)
        except NoSuchFrameException:
                print(f"NoSuchFrameException调用方法findElement")
            #查到当前
                findElement()
                flag = 0
                    
   
#务必记得加入quit()或close()结束进程，不断测试电脑只会卡卡西
driver.close()
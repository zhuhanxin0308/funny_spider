import sys
import os
import time
import random
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.switch_to import SwitchTo
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import cv2 as cv
import pandas as pd
import re

''' 
顺丰快递类
'''
class SFOrderProcessor:
    # 属性 输出文件后缀,用于区别源文件
    suffix = "_checked"
    # 输入excel文件的订单号 所在列的列名
    order_name = "顺丰快递单号"
    # 拆分单元大小 最大一次查询二十个
    split_num = 20 
    # 属性浏览器对象
    driver = None
    # 属性，存放查询结果数组
    result = []
    # 存放原始数据的数组
    row_data = []
    # 存放不符合条件的订单号
    bad_orders = []
    # 存放合法的订单号
    good_orders = []
    # 正则表达式，用于设定订单号校验规则
    regx = r"^SF[0-9]{10,20}$"
    # 文件全路径名
    fullpath = None
    # 滑动验证码重试次数
    retry = 3
    
    # 定义构造函数
    def __init__(self, input):
        # 检测输入，如果时路径，读取excel,获取订单列表，如果是列表直接使用
        self.check_input(input)
        print("开始处理，即将弹出浏览器，请不要操作浏览器，请等待浏览器自动关闭")
        # 创建浏览器
        try:
            self.driver = webdriver.Firefox()
        except Exception as e:
            print("启动浏览器实例失败，请确认是否安装对应的库和浏览器")
            exit()
        # 拆分
        arrlist = [self.good_orders[i:i + self.split_num] for i in range(0,len(self.good_orders), self.split_num)]
        # 开始循环
        for item in arrlist:
            #try:
            self.process(item)
            #except:
            #    pass
            # 控制抓取频率，放置被封
            time.sleep(random.randint(5,8))
        # 完成后导出
        pd_data = pd.DataFrame(self.result)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)
        # 新的文件名
        if self.fullpath:
            arr = self.fullpath.split('.')
            filename = arr[0] + self.suffix + '.' + arr[-1]
        else:
            filename = './output' + self.suffix + '.xlsx'
        pd_data.to_excel(filename, sheet_name = "查询结果",index = False, na_rep = 0,inf_rep = 0)
        self.driver.quit()
        print('共发现' + str(len(self.good_orders) + len(self.bad_orders))+ '个订单号，去掉不合法的' + str(len(self.bad_orders)) + '个，共处理订单' + str(len(self.good_orders)) + '条')
    

    # 订单号校验，以防止无效输入,计算合法订单号
    def get_right_orders(self, orders):
        for order in orders:
            if re.match(self.regx, str(order)) is not None:
                self.good_orders.append(str(order))
            else:
                self.bad_orders.append(str(order))
    
    def get_data_from_excel(self, filepath):
        try:
            track_sheet = pd.read_excel(filepath)
            return track_sheet[self.order_name]
        except:
            print("读取excel失败，请检查excel文件")
            exit()

    def check_input(self, input):
        # 如果传的订单号列表直接进行校验
        if type(input) == list:
            self.get_right_orders(input)
        elif type(input) == str:
            # 检测文件是否存在
            self.fullpath = input
            if os.path.exists(input):
                self.get_right_orders(self.get_data_from_excel(input))
        if len(self.good_orders) == 0:
            print("没有获取到合法的订单号")
            exit()

    # 工人类可以计算距离
    def get_distance(self, bg_img_path='./bg.png', slider_img_path='./slider.png'):
        """获取滑块移动距离"""
        # 背景图片处理
        bg_img = cv.imread(bg_img_path, 0)  # 读入灰度图片
        bg_img = cv.GaussianBlur(bg_img, (3, 3), 0)  # 高斯模糊去噪
        bg_img = cv.Canny(bg_img, 50, 150)  # Canny算法进行边缘检测
        # 滑块做同样处理
        slider_img = cv.imread(slider_img_path, 0)
        slider_img = cv.GaussianBlur(slider_img, (3, 3), 0)
        slider_img = cv.Canny(slider_img, 50, 150)
        # 寻找最佳匹配
        res = cv.matchTemplate(bg_img, slider_img, cv.TM_CCOEFF_NORMED)
        # 最小值，最大值，并得到最小值, 最大值的索引
        min_val, max_val, min_loc, max_loc = cv.minMaxLoc(res)
        # 例如：(-0.05772797390818596, 0.30968162417411804, (0, 0), (196, 1))
        top_left = max_loc[0]  # 横坐标
        return top_left/2
    
    ''' 
    工人类抓取并处理内容
    orders 订单号列表
    '''
    def process(self, orders):
        # 将订单号列表转成字符串
        order_str = ','.join([str(i) for i in orders])
        # 导航到包含滑动验证码的网页
        self.driver.get("https://www.sf-express.com/we/ow/chn/sc/waybill/waybill-detail/" + order_str)

        # 等待页面加载完成
        wait = WebDriverWait(self.driver, 15)
        slider = wait.until(EC.presence_of_element_located((By.ID, "tcaptcha_iframe")))
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "tcaptcha_iframe"))
        # 获取滑块和背景图像的元素
        wait.until(EC.presence_of_element_located((By.ID, "slideBlock")))
        wait.until(EC.presence_of_element_located((By.ID, "slideBg")))
        slider_block = self.driver.find_element(By.ID, "slideBlock")
        bg = self.driver.find_element(By.ID, "slideBg")
        slider_knob = self.driver.find_element(By.ID, "tcaptcha_drag_thumb")

        block_img = slider_block.get_attribute('src')
        bg_img = bg.get_attribute('src')
        while block_img is None:
            slider_block = self.driver.find_element(By.ID, "slideBlock")
            block_img = slider_block.get_attribute('src')
            slider_knob = self.driver.find_element(By.ID, "tcaptcha_drag_thumb")

        while bg_img is None:
            bg = self.driver.find_element(By.ID, "slideBg")
            bg_img = bg.get_attribute('src')
            slider_knob = self.driver.find_element(By.ID, "tcaptcha_drag_thumb")
        g = requests.get(bg_img).content
        b = requests.get(block_img).content
        with open('bg.png', "wb") as f1,open('block.png', 'wb') as f2:
            f1.write(g)
            f2.write(b)
        jvli = self.get_distance('bg.png', 'block.png')
        # 计算距离后将图片文件删除
        os.remove('block.png')
        os.remove('bg.png')
        # 创建一个 ActionChains 对象
        actions = ActionChains(self.driver)
        # 将鼠标悬停在滑块上，按下左键，并向右拖动滑块
        actions.drag_and_drop_by_offset(slider_knob, jvli-26, 0).perform()
        actions.release()
        self.driver.switch_to.default_content()
        # 等待验证结果
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "route-list")))
        # 等待地图加载
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "map")))
        # 如果一页超过十条，自动点击更多
        
        if len(orders) > 10:
            fetch_more = self.driver.find_element(By.CLASS_NAME, 'fetch-more')
            # 滚动到元素位置
            self.driver.execute_script("arguments[0].scrollIntoView();", fetch_more)
            # 防止加载动画还没结束就点击
            time.sleep(1)
            fetch_more.click()

        # 查询结果加载完成
        # 查找所有的.route-list元素
        route_list =  self.driver.find_elements(By.CLASS_NAME, "route-list")
        # 查找所有的.number 元素
        num_list = self.driver.find_elements(By.CLASS_NAME, 'number')
        # 订单号
        ems_num = []
        #遍历查询出来的订单号
        for item in num_list:
            ems_num.append(item.get_attribute("textContent"))
        #遍历查询出来的订单的状态信息
        index = 0
        for item in route_list:
           el = item.find_element(By.CLASS_NAME, 'first').get_attribute("textContent").split(" ", maxsplit=4)
           num = ems_num[index]
           status = el[0]
           date_time = el[2] + ' ' + el[3]
           detail = el[4]
           index = index + 1
           # 构造数据
           self.result.append({self.order_name:num,'签收时间':date_time,'订单状态':status,'备注':detail})

        


if __name__ == '__main__':
    try:
      path = sys.argv[1]
    except:
        print('请输入文件路径')
        exit()
    SFOrderProcessor(path)
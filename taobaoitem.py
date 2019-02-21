# -*- coding:UTF-8 -*-

import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from urllib.parse import quote
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from pyquery import PyQuery as pq
import xlwt

CHROME_DRIVER = '/usr/local/bin/chromedriver'
TB_LOGIN_URL = 'https://login.taobao.com/member/login.jhtml'
MAX_PAGE = 10
KEYWORD = '面膜'

class SessionException(Exception):
    """
    会话异常类
    """
    def __init__(self, message):
        super().__init__(self)
        self.message = message

    def __str__(self):
        return self.message

class TaoBaoSearch(object):
    def __init__(self):
        self.browser = None
        self.productlist = list()
        self.page = 1

    def login(self, username, password):
        print("初始化浏览器")
        self.__init_browser()
        time.sleep(2)
        print("切换密码输入框")
        self.__switch_to_password_mode()
        time.sleep(2)
        print('输入账号')
        self.__write_username(username)
        time.sleep(2)
        print('输入密码')
        self.__write_password(password)
        time.sleep(2)
        print('判断是否有验证码')
        if self.__lock_exist():
            self.__unlock()
        self.__submit()
        print('登录成功')
        time.sleep(1)
        print('搜索商品')
        self.__search()
        # 搜索商品
        # self.__searchItem()

    #创建浏览器对象
    def __init_browser(self):
        options = Options()
        options.add_argument('--proxy-server=http://127.0.0.1:9000')
        self.browser = webdriver.Chrome(executable_path=CHROME_DRIVER, options=options)
        self.browser.implicitly_wait(3)
        self.browser.maximize_window()
        self.browser.get(TB_LOGIN_URL)

        self.wait = WebDriverWait(self.browser, 10)

    def __switch_to_password_mode(self):
        """
        切换到密码模式
        :return:
        """
        if self.browser.find_element_by_id('J_QRCodeLogin').is_displayed():
            self.browser.find_element_by_id('J_Quick2Static').click()

    def __write_username(self, username):
        """
        输入账号
        :param username:
        :return:
        """
        username_input_element = self.browser.find_element_by_id('TPL_username_1')
        username_input_element.clear()
        username_input_element.send_keys(username)

    def __write_password(self, password):
        """
        输入密码
        :param password:
        :return:
        """
        password_input_element = self.browser.find_element_by_id("TPL_password_1")
        password_input_element.clear()
        password_input_element.send_keys(password)

    def __lock_exist(self):
        """
        判断是否存在滑动验证
        :return:
        """
        return self.__is_element_exist('#nc_1_wrapper') and self.browser.find_element_by_id(
            'nc_1_wrapper').is_displayed()

    def __unlock(self):
        """
        执行滑动解锁
        :return:
        """
        bar_element = self.browser.find_element_by_id('nc_1_n1z')
        ActionChains(self.browser).drag_and_drop_by_offset(bar_element, 350, 0).perform()
        time.sleep(0.5)
        self.browser.get_screenshot_as_file('error.png')
        if self.__is_element_exist('.errloading > span'):
            error_message_element = self.browser.find_element_by_css_selector('.errloading > span')
            error_message = error_message_element.text
            self.browser.execute_script('noCaptcha.reset(1)')
            raise SessionException('滑动验证失败, message = ' + error_message)

    def __submit(self):
        """
        提交登录
        :return:
        """
        self.browser.find_element_by_id('J_SubmitStatic').click()
        time.sleep(0.5)
        if self.__is_element_exist("#J_Message"):
            error_message_element = self.browser.find_element_by_css_selector('#J_Message > p')
            error_message = error_message_element.text
            raise SessionException('登录出错, message = ' + error_message)


    def index_page(self, page):
        """
        根据页码获取商品列表
        :param page: 页码
        """
        print('正在爬取第', page, '页')
        self.page = page
        try:
            url = 'https://s.taobao.com/search?q=' + quote(KEYWORD)
            self.browser.get(url)
            if page > 1:
                input = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-pager div.form > input')))
                submit = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, '#mainsrp-pager div.form > span.btn.J_Submit')))
                input.clear()
                input.send_keys(page)
                submit.click()
            self.wait.until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, '#mainsrp-pager li.item.active > span'), str(page)))
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.m-itemlist .items .item')))
            self.get_products()
        except NoSuchElementException:
            self.browser.get_screenshot_as_file('error.png')
            self.index_page(page)

    def get_products(self):
        """
        提取商品数据
        """
        print('解析网页数据')
        html = self.browser.page_source
        doc = pq(html)
        items = doc('#mainsrp-itemlist .items .item').items()
        products = []
        for item in items:
            product = list()
            product.append(item.find('.pic .img').attr('data-src'))
            product.append(item.find('.price').text())
            product.append(item.find('.deal-cnt').text())
            product.append(item.find('.title').text())
            product.append(item.find('.shop').text())
            product.append(item.find('.location').text())
            # product = {
            #     'image': item.find('.pic .img').attr('data-src'),
            #     'price': item.find('.price').text(),
            #     'deal': item.find('.deal-cnt').text(),
            #     'title': item.find('.title').text(),
            #     'shop': item.find('.shop').text(),
            #     'location': item.find('.location').text()
            # }
            products.append(product)
            self.productlist.append(products)
            print(self.productlist)
        self.__write_product(products)

    def __write_product(self, products):
        workbook = xlwt.Workbook()
        sheet1 = workbook.add_sheet('sheet1')
        titles = ['图片', '价格', '销量', '标题', '店铺', '归属地']
        for i in range(len(titles)):
            sheet1.write(0, i, titles[i])
        for i in range(len(products)):
            for j in range(6):
                sheet1.write(i + 1 + (len(self.productlist) - len(products)), j, products[i][j])
        workbook.save('Workbook3.xls')
        print('创建execel完成！')

    def __search(self):
        """
        遍历每一页
        """
        for i in range(1, MAX_PAGE + 1):
            self.index_page(i)
        self.browser.close()

    def __is_element_exist(self, selector):
        """
        检查是否存在指定元素
        :param selector:
        :return:
        """
        try:
            self.browser.find_element_by_css_selector(selector)
            return True
        except NoSuchElementException:
            return False

if __name__ == '__main__':
    tb = TaoBaoSearch()
    tb.login('username', 'password')
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from mylogclass import MyLogClass
import openpyxl


class FuPin():
    def __init__(self, user, password, path):
        try:
            self.user = user
            self.password = password
            self.path = path
            self.driver = webdriver.Chrome()
            self.wait = WebDriverWait(self.driver, 10)
            self.log = MyLogClass()
        except Exception as e:
            print(e)
            time.sleep(10)

    def read_excel(self):
        try:
            wb = openpyxl.load_workbook(self.path)
            ws = wb.active
            for row in list(ws.rows)[1:]:
                id = row[5].value
                data = {
                    '工资性收入': str(row[8].value),
                    '生产经营性收入': str(row[9].value),
                    '计划生育金': str(row[10].value),
                    '低保金': str(row[11].value),
                    '特困供养金': str(row[12].value),
                    '养老保险金': str(row[13].value),
                    '生态补偿金': str(row[14].value),
                    '其他转移性收入': str(row[15].value),
                    '资产收益扶贫分红收入': str(row[16].value),
                    '其他财产性收入': str(row[17].value),
                    '生产性支出': str(row[18].value),
                }
                yield id, data
        except Exception as e:
            print(e)
            time.sleep(10)

    def run(self):
        try:
            # 登陆
            self.login()
            # 等待登陆成功
            self.login_status()
            # 进入系统
            self.into_system()
            # 打开搜索窗口
            self.search_win()

            for id, data in self.read_excel():
                try:
                    # 查询
                    self.search(id)
                    # 设置二号子项
                    self.set_item2()
                    # 设置三号子项
                    self.set_item3(data)
                    # 设置五号子项
                    self.set_item5()
                    self.log.logger.info('完成:  ' + id)
                except Exception as e:
                    self.log.logger.warning('失败:  ' + id)
        except Exception as e:
            print(e)
            time.sleep(10)

    def login(self):
        self.driver.get("http://106.38.235.201:7080/portal/#")
        self.driver.maximize_window()
        username = self.wait.until(EC.element_to_be_clickable((By.ID, "username")))
        username.send_keys(self.user)
        password = self.wait.until(EC.element_to_be_clickable((By.ID, "password")))
        password.send_keys(self.password)

    def login_status(self):
        while True:
            if '全国扶贫开发信息系统' == self.driver.title:
                return True
            time.sleep(2)

    def into_system(self):
        menuIcon1 = self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "menuIcon1")))
        menuIcon1.click()
        time.sleep(1)
        dialog = self.wait.until(
            EC.presence_of_element_located((By.XPATH, '//div[@id="dialog"]//ul[@class="link"]//li[1]/a')))
        dialog.click()
        time.sleep(1)

    def search_win(self):
        try:
            扶贫对象 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@title,"扶贫对象")]')))
            扶贫对象.click()
            基础信息维护 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@title,"基础信息维护")]')))
            基础信息维护.click()
            年度2020 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@title,"2020年度")]')))
            年度2020.click()
            贫困户 = self.wait.until(EC.element_to_be_clickable((By.XPATH,
                                                              '//a[contains(@title,"2020年度")]/following-sibling::nui-main-menu-sub//a[contains(@title,"贫困户")]')))
            贫困户.click()
            time.sleep(1)
        except Exception as e:
            self.search_win()

    def search(self, id):
        try:
            input_id = self.wait.until(EC.presence_of_element_located((By.ID, 'aab004')))
            input_id.clear()
            input_id.send_keys(id)
            查询 = self.wait.until(EC.element_to_be_clickable((By.ID, 'on_query')))
            查询.click()
            time.sleep(1)
            名字 = self.wait.until(EC.element_to_be_clickable((By.XPATH,
                                                             '//tbody[contains(@class,"ui-datatable-data ui-widget-content ui-datatable-hoverable-rows")]/tr[1]/td[4]//a')))
            名字.click()
            time.sleep(1)
        except Exception as e:
            self.search(id)

    def set_item2(self):
        二_生产生活条件 = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"二、生产生活条件")]/..')))
        二_生产生活条件.click()
        time.sleep(0.5)
        xiala = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aac316"]//div[contains(@class,"ui-dropdown-trigger ui-state-default ui-corner-right")]')))
        xiala.click()
        time.sleep(0.5)
        xiala = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "硬化路")]')))
        xiala.click()
        time.sleep(0.5)

    def set_item3(self, data):
        三_上年度收入和患病信息 = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"三、上年度收入和患病信息")]/..')))
        三_上年度收入和患病信息.click()
        time.sleep(0.5)
        for k, v in data.items():
            _input = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//label[contains(text(),"' + k + '")]/../following-sibling::div/input')))
            _input.clear()
            time.sleep(0.1)
            _input.send_keys(v)
        time.sleep(0.5)

    def set_item5(self):
        五_帮扶责任人结对信息 = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"五、帮扶责任人结对信息")]/..')))
        五_帮扶责任人结对信息.click()
        time.sleep(0.5)
        r_count = len(self.driver.find_elements_by_xpath(
            '//p-tabpanel[contains(@header,"五、帮扶责任人结对信息")]//tbody/tr/td[1]/p-dtradiobutton/div'))

        for n in range(r_count):
            r = self.driver.find_elements_by_xpath(
                '//p-tabpanel[contains(@header,"五、帮扶责任人结对信息")]//tbody/tr/td[1]/p-dtradiobutton/div')[n]
            r.click()
            updateTime = self.wait.until(EC.element_to_be_clickable((By.ID, 'updateTime')))
            updateTime.click()
            time.sleep(0.1)
            input_time = self.driver.find_elements_by_xpath('//p-dialog[contains(@header,"修改帮扶时间")]//input')[1]
            input_time.send_keys('xxxxxxxxxxxxxxxx')
            保存时间 = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//p-dialog[contains(@header,"修改帮扶时间")]//button[@id="saveTime"]')))
            保存时间.click()
            time.sleep(0.1)
            确定 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"确定")]')))
            确定.click()
            time.sleep(0.1)
            r = self.driver.find_elements_by_xpath(
                '//p-tabpanel[contains(@header,"五、帮扶责任人结对信息")]//tbody/tr/td[1]/p-dtradiobutton/div')[n]
            r.click()
            time.sleep(0.1)
            updateTime = self.wait.until(EC.element_to_be_clickable((By.ID, 'updateTime')))
            updateTime.click()
            time.sleep(0.1)
            input_time = self.driver.find_elements_by_xpath('//p-dialog[contains(@header,"修改帮扶时间")]//input')[1]
            input_time.send_keys('2025年12月31日')
            保存时间 = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//p-dialog[contains(@header,"修改帮扶时间")]//button[@id="saveTime"]')))
            保存时间.click()
            time.sleep(0.1)
            确定 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"确定")]')))
            确定.click()
            time.sleep(0.1)

        保存 = self.wait.until(EC.element_to_be_clickable((By.ID, 'on_save')))
        保存.click()
        time.sleep(0.5)
        确定 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"确定")]')))
        确定.click()
        time.sleep(0.5)
        关闭 = self.wait.until(EC.element_to_be_clickable((By.ID, 'on_cancel')))
        关闭.click()
        time.sleep(0.5)


if __name__ == '__main__':
    user = '37132300901'
    password = '123abc@A'
    path = '1.xlsx'

    OBJ = FuPin(user, password, path)
    OBJ.run()
    time.sleep(10)
    OBJ.driver.quit()

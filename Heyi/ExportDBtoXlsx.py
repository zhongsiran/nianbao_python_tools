import pymysql
from openpyxl import load_workbook
from openpyxl import Workbook
from string import Template
from datetime import date
import sys

class ExportMysqlToXlsx:
    def __init__(self):

        self.host = 'localhost'
        self.host_2 = '10.2.94.28'  # 合益所
        # self.host_2 = '10.21.201.175' # 花城所
        self.port = 3306
        self.user = 'homestead'
        self.passwd = 'secret'
        self.db = 'nianbao'

        self.conn = self.get_conn()
        self.cur = self.conn.cursor()
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet_headrow = ['注册号','名称','地址','登记电话','法定代表人','联络员','联络员电话']
        self.worksheet_headrow.extend(['历史电话记录','本年度电话记录','跟进情况'])
        self.worksheet_headrow.extend(['年报状态','电话情况','联系员电话情况','企业类型','片区'])
        self.db_table_name = '2019_nianbao_corp'
        self.results = tuple()
        self.row = 1

    def get_conn(self):
        print('连接中......')
        try:
            conn = pymysql.connect(host=self.host, port=self.port, user=self.user, passwd=self.passwd, db=self.db, charset='utf8')
            print('连接本机数据库成功')
        except pymysql.err.OperationalError as e:
            print('没有在本机发现数据库，尝试连接%s' % self.host_2)
            try:
                conn = pymysql.connect(host=self.host_2, port=self.port, user=self.user, passwd=self.passwd, db=self.db, charset='utf8')
                print('连接成功')
            except pymysql.err.OperationalError as e:
                print('无法连接到服务器，请检查设置，程序即将退出')
                print('错误信息：')
                print(e)
                sys.exit()
        
        return conn

    def query_all(self, cur, sql, args):
        print('查询中......')
        cur.execute(sql, args)
        return cur.fetchall()

    def get_db_data(self):
        sql = '''SELECT 
        Regnum, CorpName, Addr, Phone, RepPerson, ContactPerson, ContactPhone,
        PhoneCallHistoryRecord, PhoneCallRecord, Status,
        nian_bao_status, phone_status, cphone_status, type, division
        FROM %s
        ''' % self.db_table_name
        results = self.query_all(self.cur, sql, None)
        if results:
            self.results = results
        else:
            print('查询不到记录')
            self.results = ()

    def save_to_xlsx(self):
        if self.worksheet_headrow:
            for i, j in enumerate(self.worksheet_headrow):
                self.worksheet.cell(row=1, column=i+1, value=j)
        result_row = 2
        for result in self.results:
            for col_idx, col_content in enumerate(result):
                self.worksheet.cell(row=result_row, column=col_idx+1, value=col_content)
            result_row = result_row + 1
        self.save_workbook()

    def save_workbook(self):
        now = date.today()
        today = "%d-%d-%d" %(now.year, now.month, now.day)
        save_name = today + '导出记录表.xlsx'
        try:
            self.workbook.save(save_name)
        except PermissionError as e:
            print(e)
            print('出错了：无法保存文件“%s”，请检查是否已经打开同名文件。'% save_name)
        
    def close_cur(self):
        self.cur.close()

    def close_conn(self):
        self.conn.close()

if __name__ == "__main__":
    E = ExportMysqlToXlsx()
    E.get_db_data()
    E.save_to_xlsx()
    E.close_cur()
    E.close_conn()
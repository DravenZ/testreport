# encoding: utf-8

"""
__author__ = "Zhang Pengfei"
__date__ = 2018/11/7
"""

from MySqlDB import MysqlDb
import time
import datetime


class QueryData(object):

    user = 'chadwick'
    password = 'Al2ab7U$UM7Ux2FJJj3ztsuMWErND'
    host = "203.6.234.220"
    port = 3306
    db_name = 'zentao'
    sql = ''
    sqldb = MysqlDb()
    conn = sqldb.conn(host, port, db_name, user, password)

    date = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    now_date = datetime.datetime.now().date()

    start_time = "'" + '2018-11-2' + "'"

    def __init__(self, product=1):
        if product == 1:
            self.duan = {
                'WEB': (10, -1),
                'SERVER': (17, 8)
            }
            self.world_farm_product_id = (14, -1)
        elif product == 2:
            self.duan = {
                'ANDROID': (16, 12, 13),
                'IOS': (11, 14, 15)
            }
            self.world_farm_product_id = (15, -1)
        else:
            self.duan = {
                'WEB': (4, -1),
                'ANDROID': (5, -1),
                'IOS': (6, -1),
                'SERVER': (7, 8)
            }
            self.world_farm_product_id = (2, 10)

    def query_create_all(self):
        bug_list = []
        for i in self.duan.values():
            sql_create_all = "select severity,count(*) from zt_bug WHERE DATE_FORMAT(openedDate,'%%y-%%m-%%d')  >= DATE_FORMAT(%s,'%%y-%%m-%%d')\
                                        and product in %s and project in %s GROUP BY severity;" % (self.start_time, self.world_farm_product_id, i)
            # print(sql_create_all)
            bug_list.append(self.sqldb.sqlsearch(sql_create_all, self.conn))
        return bug_list

    def query_personnel_list(self):
        sql = "SELECT zu.realname FROM zt_bug zb LEFT JOIN zentao.zt_user zu ON zu.account=" \
              "(CASE zb.resolvedBy WHEN '' THEN zb.assignedTo ELSE zb.resolvedBy END) " \
              "WHERE zb.product IN %s AND " \
              "DATE_FORMAT(zb.openedDate,'%%y-%%m-%%d')>=DATE_FORMAT(%s,'%%y-%%m-%%d') " \
              "GROUP BY zu.realname;" % (self.world_farm_product_id, self.start_time)
        print(sql)
        return self.sqldb.sqlsearch(sql, self.conn)

    def query_bug_list_by_person(self):
        bug_list = []
        for i in range(1, 5):
            sql = "SELECT zu.realname, count(*) FROM zt_bug zb LEFT JOIN zentao.zt_user zu ON zu.account=" \
              "(CASE zb.resolvedBy WHEN '' THEN zb.assignedTo ELSE zb.resolvedBy END) " \
              "WHERE zb.product IN %s AND " \
              "DATE_FORMAT(zb.openedDate,'%%y-%%m-%%d')>=DATE_FORMAT(%s,'%%y-%%m-%%d') " \
              "AND severity = %s GROUP BY zu.realname;" % (self.world_farm_product_id, self.start_time, str(i))
            print(sql)
            bug_list.append(self.sqldb.sqlsearch(sql, self.conn))
        return bug_list

    def query_bug_by_date(self):
        bug_list = []
        for j in range(4):
            yes_time = self.now_date + datetime.timedelta(days=-(3-j))
            print(yes_time)
            str_date = "'" + str(yes_time) + "'"
            bug_list_day = []
            for i in self.duan.values():
                sql_create_all = "select count(*) from zt_bug WHERE DATE_FORMAT(openedDate,'%%y-%%m-%%d')  = " \
                                 "DATE_FORMAT(%s,'%%y-%%m-%%d') and product in %s and project in %s" % \
                                 (str_date, self.world_farm_product_id, i)
                # print(sql_create_all)
                bug_list_day.append(self.sqldb.sqlsearch(sql_create_all, self.conn))
            bug_list.append(bug_list_day)
        return bug_list

    def query_bug_list(self):
        sql_bug_list = "SELECT zb.product,zb.title,zb.severity,zu.realname,zb.id FROM zt_bug zb " \
                       "LEFT JOIN zentao.zt_user zu ON zu.account=(CASE zb.resolvedBy " \
                       "WHEN '' THEN zb.assignedTo ELSE zb.resolvedBy END) " \
                       "WHERE zb.product IN %s AND zb.STATUS='active' " \
                       "AND DATE_FORMAT(zb.openedDate,'%%y-%%m-%%d')>=DATE_FORMAT(%s,'%%y-%%m-%%d') " \
                       "ORDER BY zb.product,zu.realname;" % (self.world_farm_product_id, self.start_time)
        print(sql_bug_list)
        return self.sqldb.sqlsearch(sql_bug_list, self.conn)


if __name__ == '__main__':
    Q = QueryData()
    # print(Q.query_create_all(Q.start_time))
    # print(Q.query_bug_list(Q.start_time))
    print(Q.query_personnel_list())

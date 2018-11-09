# encoding: utf-8

"""
__author__ = "Zhang Pengfei"
__date__ = 2018/11/7
"""

import pymysql


class MysqlDb(object):

    # 连接数据库
    def conn(self, host, port, dbname, user, password):
        try:
            conn = pymysql.connect(host, user, password, dbname, port, charset='utf8')
            return conn
        except Exception as e:
            print("connect failed", e)

    # 查询
    def sqlsearch(self, sql, db):
        try:
            cur = db.cursor()
            x = cur.execute(sql)                   # 使用cursor进行各种操作
            results = cur.fetchmany(x)
            cur.close()
            return results
        except Exception as e:
            print("search failed", e)

    # 直接增删改
    def sqlDML(self, sql, db):
        # include: insert,update,delete
        try:
            cr = db.cursor()
            cr.execute(sql)
            cr.close()
            db.commit()
        except Exception as e:
            print("sqlDML failed",e)

    # 有参数增删改
    def sqlDML2(self, sql, params, db):
        # execute dml with parameters
        try:
            cr = db.cursor()
            cr.executemany(sql, params)
            cr.close()
            db.commit()
        except Exception as e:
            print("sqlDML2 failed", e)

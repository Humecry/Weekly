#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-07-17 14:05:40
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description：统计上周异常数据
import pyodbc
import pandas as pd
import datetime
from datetime import timedelta
# 引入配置文件
from conf import *
# 引入公共函数
from common import *

# 当前时间
now = datetime.datetime.now()

'''
上周会员卡消费异常
'''
def memberCard(process_bar):
	# 上上周周日
	last_week_start = now - timedelta(days=7+now.isoweekday())
	last_week_sunday = last_week_start.strftime("%Y%m%d")
	last_week_start = last_week_start.strftime("%Y-%m-%d")
	# 上周周日
	last_week_end = now - timedelta(days=now.isoweekday())
	last_week_end = last_week_end.strftime("%Y-%m-%d")
	# 上周周六
	last_week_saturday = now - timedelta(days=1+now.isoweekday())
	last_week_saturday = last_week_saturday.strftime("%Y%m%d")
	# 连接数据库
	try:
		conn = pyodbc.connect(S003)
	except pyodbc.OperationalError as err:
		print("会员卡消费3号店数据库报错:")
		print(err.args)
		return False
	except pyodbc.Error as err:
		print("会员消费3号店数据库报错:")
		print(err.args)
		return False
	except:
		print("会员卡消费报错:3号店数据库连接失败,请确认配置是否正确!")

	process_bar.show_process()

	sql = '''
	SELECT
		a.k# [卡号],
		c.hr [持卡人],
		c.z# [身份证],
		sum( a.je ) [消费额],
		sum( a.oe ) [原价额],
		c.ten [手机号],
		sum( a.ie ) [成本],
		sum( a.je - a.ie ) [毛利],
		count( DISTINCT a.ls ) [消费次数],
		sum( a.je / b.tc ) [积分],
		a.ykd [分店号]
	FROM
		pos.dbo.jhjltab AS a
		LEFT JOIN pos.dbo.jmgtab AS b ON b.gz = a.bg % 10000
		LEFT JOIN pos.dbo.jhdjtab AS c ON a.k# = c.h#
	WHERE
		a.k# > 0
		AND b.tc > 0
		AND a.rt BETWEEN '{startDate}'
		AND '{endDate}'
	GROUP BY
		a.k#,
		c.hr,
		c.z#,
		c.ten,
		a.ykd 
	HAVING
		count( DISTINCT a.ls ) >= 21
	'''.format(startDate=last_week_start, endDate=last_week_end)
	df = pd.read_sql(sql, conn)

	process_bar.show_process()

	# 写入Excel
	filePath = PATH + last_week_sunday + "-" + last_week_saturday + "上周会员卡消费异常.xlsx"
	writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
	df.to_excel(writer,  sheet_name='Sheet1', index=False)
	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']
	# 设置表格样式
	col = []
	for value in df.columns.tolist():
		col.append({'header': value})
	worksheet.add_table('A1:K' + str(len(df) + 1), {'columns': col})
	# 设置字体
	cell_format = workbook.add_format({'font_name': '微软雅黑'})
	worksheet.set_column('A:K', 11, cell_format)
	# 设置列宽
	worksheet.set_column('C:C', 25)
	worksheet.set_column('F:F', 15)
	# 设置数字格式
	cell_format = workbook.add_format({'font_name': '微软雅黑', 'num_format': '0.00'})
	worksheet.set_column('D:E', None, cell_format)
	worksheet.set_column('G:H', None, cell_format)
	worksheet.set_column('J:J', None, cell_format)
	# 冻结窗口
	writer.save()
	return filePath

'''
上周批发毛利低于8个点数据
'''
def wholesaleProfit(last_week_start, last_week_end, process_bar):
	# 连接数据库
	try:
		conn = pyodbc.connect(S000)
	except:
		print("批发毛利报错:0号数据库连接失败,请确认配置是否正确!")
		return False
	sql = '''
	DECLARE @rq smalldatetime 
	SET @rq = CONVERT ( VARCHAR, DATEADD ( d, 2- DATEPART ( w, GETDATE ( ) ), GETDATE ( ) ), 112 ) 
	SELECT
		lqm [分店],批发单号,客户号,客户名,备注,批发日,货号,品名,数量,批发价,进价,[毛利%] 
	FROM
		pos.dbo.低毛利批发商品
		LEFT JOIN pos.dbo.jlqtab ON ( lb = 2 AND fd = c ) 
	WHERE
		批发日 BETWEEN @rq - 7 
		AND @rq 
		AND 档期号 = 0 
	ORDER BY
		c,批发日
	'''
	df = pd.read_sql(sql, conn)

	process_bar.show_process()

	# 写入Excel
	filePath = PATH + last_week_start + "-" + last_week_end + "上周批发毛利低于8点.xlsx"
	writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
	df.to_excel(writer,  sheet_name='Sheet1', index=False)
	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']
	# 设置表格样式
	col = []
	for value in df.columns.tolist():
		col.append({'header': value})
	worksheet.add_table('A1:L' + str(len(df) + 1), {'columns': col})
	# 设置字体
	cell_format = workbook.add_format({'font_name': '微软雅黑'})
	worksheet.set_column('A:L', 14, cell_format)
	# 设置列宽
	worksheet.set_column('D:D', 56)
	worksheet.set_column('E:E', 30)
	worksheet.set_column('F:F', 21)
	worksheet.set_column('H:H', 39)
	# 设置数字格式
	cell_format = workbook.add_format({'font_name': '微软雅黑', 'num_format': '0.00'})
	worksheet.set_column('J:K', None, cell_format)
	cell_format = workbook.add_format({'font_name': '微软雅黑', 'num_format': '0.0%'})
	worksheet.set_column('L:L', None, cell_format)
	# 设置条件格式
	red_format = workbook.add_format({'font_name': '微软雅黑', 'bg_color': 'FF0040'})
	worksheet.conditional_format('L2:L' + str(len(df) + 1), {'type': 'cell',
	                                    'criteria': '<',
	                                    'value':     0,
	                                    'format':    red_format})
	# 冻结窗口
	worksheet.freeze_panes(0, 1)
	writer.save()
	return filePath

'''
上周猪肉负毛利销售数据
'''
def porkProfit(last_week_start, last_week_end, process_bar):
	# 连接数据库
	try:
		conn = pyodbc.connect(S008)
	except:
		print("猪肉负毛利报错:8号店数据库连接失败,请确认配置是否正确!")
		return False
	sql = '''
	DECLARE @rq SMALLDATETIME 
	SET @rq = CONVERT ( VARCHAR, DATEADD( d, 1-DATEPART ( w, GETDATE()), GETDATE()), 112 ) 
	SELECT
		lqm [分店],
		rq [日期],
		sx# [货号],
		na [品名],
		qtl [销量],
		qi [成本],
		qo [销售额],
		( qo - qi ) / qi [毛利率] 
	FROM
		pos.dbo.jsptab,
		pos.dbo.bzsrbak
		JOIN pos.dbo.jlqtab ON ( jlqtab.lb= 2 AND c = sn ) 
	WHERE
		s = sx# 
		AND rq BETWEEN @rq - 6 
		AND @rq 
		AND hs = 1201 
		AND QI > QC 
		AND qtl >0
	'''
	df = pd.read_sql(sql, conn)

	process_bar.show_process()

	# 写入Excel
	filePath = PATH + last_week_start + "-" + last_week_end + "猪肉负毛利销售明细.xlsx"
	writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
	df.to_excel(writer,  sheet_name='Sheet1', index=False)
	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']
	# 设置表格样式
	col = []
	for value in df.columns.tolist():
		col.append({'header': value})
	worksheet.add_table('A1:H' + str(len(df) + 1), {'columns': col})
	# 设置字体
	cell_format = workbook.add_format({'font_name': '微软雅黑'})
	worksheet.set_column('A:H', 11, cell_format)
	# 设置列宽
	worksheet.set_column('B:B', 21)
	# 数字格式
	cell_format = workbook.add_format({'font_name': '微软雅黑', 'num_format': '0.00'})
	worksheet.set_column('E:G', None, cell_format)
	red_format = workbook.add_format({'font_name': '微软雅黑', 'bg_color': 'FF0040'})
	cell_format = workbook.add_format({'font_name': '微软雅黑', 'num_format': '0.0%'})
	worksheet.set_column('H:H', None, cell_format)
	# 设置条件格式
	worksheet.conditional_format('H2:H' + str(len(df) + 1), {'type': 'cell',
	                                    'criteria': '<',
	                                    'value':     0,
	                                    'format':    red_format})
	writer.save()
	return filePath

def main(process_bar):
	# 上周会员卡消费异常
	path1 = memberCard(process_bar)
	# 上周的第一天周一，最后一天周日
	last_week_start = now - timedelta(days=6+now.isoweekday())
	last_week_start = last_week_start.strftime("%Y%m%d")
	last_week_end = now - timedelta(days=now.isoweekday())
	last_week_end = last_week_end.strftime("%Y%m%d")
	# 上周批发毛利低于8个点数据
	path2 = wholesaleProfit(last_week_start, last_week_end, process_bar)
	# 上周猪肉负毛利销售数据
	path3 = porkProfit(last_week_start, last_week_end, process_bar)
	return path1, path2, path3

if __name__ == '__main__':
	# 进度条
	max_steps = 4
	process_bar = ShowProcess(max_steps, 'OK')
	main(process_bar)
	time.sleep(20)
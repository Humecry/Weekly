#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-06-24 11:21:03
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description: 统计上月客流数据,日客流趋势与任意时间通道分布统计的都是进客流量，并且都乘以1.5倍

import pyodbc
import xlsxwriter
import time
import datetime
from datetime import date 
from dateutil.relativedelta import relativedelta
# 引入配置文件
from conf import *
# 引入公共函数
from common import *

def flow(firstDate, lastDate, process_bar, type='InSum'):
	# 连接客流数据库
	try:
		conn = pyodbc.connect(PASSENGER_FLOW_SERVER)
		cursor = conn.cursor()
		cursor2 = conn.cursor()
	except:
		print('客流月数据报错:客流数据库连接失败,请确认配置是否正确!')
		return False
	# 日期字典
	week = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
	attr = ['平时', '平时', '平时', '平时', '平时', '周末', '周末']
	last_month = {
		'int': [], # 存放上月的日期，整型，格式：20180622
		'str': [], # 存放上月的日期，字符串，格式：2018-06-22
		'week': [],
		'attr': []
	}
	l = (lastDate-firstDate).days+1
	for i in range(l):
		day = firstDate  + datetime.timedelta(days=i)
		last_month['int'].append(int((day).strftime("%Y%m%d")))
		last_month['str'].append(day.strftime("%Y-%m-%d"))
		last_month['week'].append(week[day.weekday()])
		last_month['attr'].append(attr[day.weekday()])
	# 新建工作簿
	filePath = PATH + str(last_month['int'][0]) + '-' + str(last_month['int'][l-1]) + '乐海客流系统数据.xlsx'
	workbook = xlsxwriter.Workbook(filePath)
	'''
	1.日客流趋势
	'''
	# 新建工作表
	worksheet = workbook.add_worksheet('日客流趋势')
	# 冻结窗格
	worksheet.freeze_panes(2, 3)
	# 设置统一表格格式
	format_comm = {
		'font_name': '微软雅黑',
		'font_size': 9,
		'bold': True,
		'text_wrap': True, # 自动换行
		'border': 1,
		'align': 'center',
		'valign': 'vcenter',
	}
	cell_format = workbook.add_format(format_comm)
	# 设置客流数据格式
	intFormat1 = workbook.add_format(dict(format_comm, **{'bold': False, 'num_format': '0'}))
	intFormat2 = workbook.add_format(dict(format_comm, **{'bg_color':'gray', 'num_format': '0'}))
	percentFormat = workbook.add_format(dict(format_comm, **{'num_format': '0.0%'}))
	percentFormat2 = workbook.add_format(dict(format_comm, **{'num_format': '0.00%'}))
	# 设置标题
	yellow = workbook.add_format(dict(format_comm, **{'bg_color':'yellow', 'font_size':22}))
	worksheet.merge_range(0,0,0,18,'乐海城市广场日客流趋势', yellow) 
	# 设置标题行高
	worksheet.set_row(0, 40) 
	# 设置列宽
	worksheet.set_column('A:A', 10)
	worksheet.set_column('R:R', 10)
	worksheet.set_column('D:Q', 7)
	# 字段名
	row = ['日期',	'星期',	'属性',	'8:00',	'09:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00',	'17:00', '18:00', '19:00', '20:00',	'21:00', '22:00~22:30',	'合计']
	worksheet.write_row('A2', row, cell_format)
	worksheet.write_column(2, 1, last_month['week'], cell_format)
	worksheet.write_column(2, 2, last_month['attr'], cell_format)
	# 设置星期六与星期日为红字体
	red = workbook.add_format({'font_color': 'red'})
	worksheet.conditional_format('B3:B'+str(l+2), {
			'type': 'text',
		    'criteria': 'containing',
		    'value': '星期六',
		    'format': red
	    })
	worksheet.conditional_format('B3:B'+str(l+2), {
			'type': 'text',
		    'criteria': 'containing',
		    'value': '星期日',
		    'format': red
	    })
	# 创建图表
	chart = workbook.add_chart({'type': 'line'})
	# 查询及写入月初到月底，8点到22点进客流量
	i = 0
	for day in last_month['int']:
		k = 0
		# 写入日期
		worksheet.write(i+2, 0, last_month['str'][i], cell_format)
		# 查询客流量
		cursor.execute("SELECT * FROM [dbo].[Summary_Sixty] WHERE DateKey = '{0}' AND SiteKey = '{1}' AND CONVERT(VARCHAR(10), CountDate, 108) >= '{2}' AND CONVERT(VARCHAR(10), CountDate, 108) <= '{3}'".format(day,  "P00001", "08:00:00", "22:00:00"))
		rows = cursor.fetchall()
		for row in rows:
			# 写入各时间段的进客流量，在实际值上乘以1.5倍
			worksheet.write(i + 2, k + 3, row.InSum*1.5, intFormat1)		
			k += 1
		# 画图
		chart.add_series({
				'name': last_month['str'][i]+last_month['week'][i]+last_month['attr'][i],
				'categories': '=日客流趋势!D2:R2',
				'values': '=日客流趋势!D{0}:R{0}'.format(i+3),
			})
		# 按天求和
		worksheet.write_formula(i+2, 18, '=SUM(D{0}:R{0})'.format(i+3), intFormat1)
		i += 1

		process_bar.show_process()

	# 设置图表大小
	chart.set_size({'width':900,'height':400})
	# 插入图表的位置
	worksheet.insert_chart('D' + str(l + 6), chart)
	# 按时间段求和
	for i in range(16):
		# 合计
		worksheet.write_formula(l+2, i+3, '=SUM({0}3:{0}{1})'.format(chr(i+68), l+2), intFormat2)
		# 合计百分比
		worksheet.write_formula(l+3, i+3, '={0}{1}/S{1}'.format(chr(i+68), l+3), percentFormat)
	worksheet.merge_range(l+2,0,l+2,2,'合计', cell_format) 
	worksheet.merge_range(l+3,0,l+3,2,'合计百分比', cell_format)


	'''
	2.任意时间通道分布
	'''
	# 新建工作表
	worksheet = workbook.add_worksheet('任意时间通道分布')
	# 冻结窗格
	worksheet.freeze_panes(2, 3)
	# 设置统一表格格式
	worksheet.set_column('A:C', None, cell_format)
	# 设置列宽
	worksheet.set_column('A:A', 10)
	worksheet.set_column('B:P', 7)
	# 排除的场所编码
	exceptArr = ('P00001', 'P00001S00010', 'P00001S00011', 'P00001S00013')
	row2 = ['日期', '星期', '属性']
	# 获取各区域上月合计排名,从高到低
	cursor.execute("SELECT SiteKey, sum(InSum) AS InSum FROM [dbo].[Summary_Day] WHERE DateKey >= '{0}' AND DateKey <= '{1}' AND SiteKey NOT IN {2} GROUP BY SiteKey ORDER BY InSum DESC".format(last_month['int'][0], last_month['int'][l-1], exceptArr))
	rows = cursor.fetchall()
	countArea = 0 # 场所计数
	for item in rows:
		for i in range(l):
			row = None
			# 查询指定场所的一天进客流量，在实际值上乘以1.5倍
			cursor2.execute("SELECT * FROM [dbo].[Summary_Day] WHERE DateKey = '{0}' AND SiteKey = '{1}'".format(last_month['int'][i], item.SiteKey))
			rows2 = cursor2.fetchall()
			for row in rows2:
				worksheet.write(i+2, countArea+3, row.InSum*1.5, intFormat1)
			# 客流量查询为空时，数据写为0
			if row is None:
				worksheet.write(i+2, countArea+3, 0, intFormat1)
		# 未知场所
		row2.append(DIC.get(item.SiteKey, '未知'))
		countArea += 1

		process_bar.show_process()

	row2.extend(['合计',	str(int(last_month['int'][0]/10000)-1)+'年同期合计', '增长率'])
	# 设置标题
	worksheet.merge_range(0,0,0,len(row2)-1, '乐海城市广场任意时间通道分布(进)', yellow) 
	# 设置标题行高
	worksheet.set_row(0, 40) 
	worksheet.set_row(1, 40) 
	# 写入表头
	worksheet.write_row('A2', row2, cell_format)
	worksheet.write_column(2, 1, last_month['week'])
	worksheet.write_column(2, 2, last_month['attr'])
	# 设置星期六与星期日为红字体
	red = workbook.add_format({'font_color': 'red'})
	worksheet.conditional_format('B3:B'+str(l+2), {
			'type': 'text',
		    'criteria': 'containing',
		    'value': '星期六',
		    'format': red
	    })
	worksheet.conditional_format('B3:B'+str(l+2), {
			'type': 'text',
		    'criteria': 'containing',
		    'value': '星期日',
		    'format': red
	    })
	# 创建图表
	chart2 = workbook.add_chart({'type': 'line'})
	i = 0
	for day in last_month['str']:
		# 写入日期
		worksheet.write(i+2, 0, day)
		# 按天写入上月客流合计
		worksheet.write_formula(i+2, countArea+3, '=SUM(D{0}:{1}{0})'.format(i+3, chr(68+countArea-1)), intFormat1)
		# 查询去年同期客流合计
		try:
			cursor.execute("SELECT SUM(InSum) AS InSum FROM [dbo].[Summary_Day] WHERE CountDate = '{0}' AND SiteKey NOT IN {1}".format(str(int(day[0:4])-1) + day[4:], exceptArr))
			rows = cursor.fetchall()
		except:
			print("客流月数据注意:上月2月份有29号，而去年同期无29号，所以29号无同期数据!")
		for row in rows:
			# 写入去年年合计进客流量
			worksheet.write(i+2, countArea+4, row.InSum*1.5, intFormat1)
			# 同期增长率
			worksheet.write_formula(i+2, countArea+5, '=({1}{0}-{2}{0})/{2}{0}'.format(i+3, chr(68+countArea), chr(68+countArea+1)), percentFormat2)
		# 画图
		chart2.add_series({
				'name': last_month['str'][i]+last_month['week'][i]+last_month['attr'][i],
				'categories': '=任意时间通道分布!D2:{0}2'.format(chr(68+countArea-1)),
				'values': '=任意时间通道分布!D{0}:{1}{0}'.format(i+3, chr(68+countArea-1)),
			})
		i += 1

		process_bar.show_process()

	# 设置图表大小
	chart2.set_size({'width':900,'height':400})
	# 插入图表的位置
	worksheet.insert_chart('D' + str(l+7), chart2)
	# 按区域求和求平均值
	for i in range(countArea+2):
		# 平均
		worksheet.write_formula(l+2, i+3, '=SUM({0}3:{0}{1})/{2}'.format(chr(i+68), l+2, l), intFormat2)
		# 合计
		worksheet.write_formula(l+3, i+3, '=SUM({0}3:{0}{1})'.format(chr(i+68), l+2), intFormat2)
		# 合计百分比
		if i >= countArea+1:
			break
		worksheet.write_formula(l+4, i+3, '={0}{2}/{1}{2}'.format(chr(i+68), chr(countArea+68), l+4), percentFormat)
	# 合并单元格
	worksheet.merge_range(l+2,0,l+2,2,'平均', cell_format) 
	worksheet.merge_range(l+3,0,l+3,2,'合计', cell_format) 
	worksheet.merge_range(l+4,0,l+4,2,'合计百分比', cell_format) 
	# 平均值增长率
	worksheet.write_formula(l+2, countArea+5, '=({0}{2}-{1}{2})/{1}{2}'.format(chr(68+countArea), chr(68+countArea+1), l+3), percentFormat2)
	# 合计增长率
	worksheet.write_formula(l+3, countArea+5, '=({0}{2}-{1}{2})/{1}{2}'.format(chr(68+countArea), chr(68+countArea+1), l+4), percentFormat2)

	workbook.close()
	return filePath
def main(type, process_bar):
	today = date.today()
	if type == 'lastweek':
		firstDate = today - datetime.timedelta(today.weekday()+7)
		lastDate = today - datetime.timedelta(today.weekday()+1)
	elif type == 'lastmonth':
		d = today - relativedelta(months=1) # 这个1指上月
		firstDate = date(d.year, d.month,1)
		lastDate = date(today.year, today.month, 1) - relativedelta(days=1) # 这里获取上月最后一天
	else:
		print(type + '参数错误,客流数据生成失败!')
		return False
	filePath = flow(firstDate, lastDate, process_bar)
	return filePath
if __name__ == '__main__':
	# 设置统计时间段
	firstDate = date(2017, 10, 1)
	lastDate = date(2017, 10, 31)
	# 进度条
	max_steps = lastDate.day * 2 + len(DIC)
	process_bar = ShowProcess(max_steps, 'OK')
	# 任意时间
	flow(firstDate, lastDate, process_bar)
	# 上周
	# main('lastweek', process_bar)
	# 上月
	# main('lastmonth', process_bar)
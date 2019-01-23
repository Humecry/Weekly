#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-07-20 17:08:55
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description: 生成每周日报, 发送到企业微信, 附加定时发送功能

import requests
import urllib
import time
import json
# from apscheduler.schedulers.blocking import BlockingScheduler
# 引入配置文件
from conf import *
# 引入公共函数
from common import *
# 引入自建模块
import jd
import passengerFlow
import unusual

# 企业微信接口
class Wxwork:
	# 初始化
	def __init__(self):
		self.token = self.get_token() #获取令牌
		self.chat = "12379832426587121255" #测试群Id
	# 获取token
	def get_token(self):
		token_url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken"
		token_params = {
			"corpid": "wxcee54f67c8e413c9",
			"corpsecret": "HJ3PW1Yu-GTRQF6zyTSg_j4q0Ga6bQWlVq0LtUiDXkQ"
		}
		# 读取token缓存
		f = open("token.txt", "r")
		token_cache = f.read()
		f.close()
		token_cache = eval(token_cache)
		# 判断token缓存是否过期
		if time.time() < token_cache["expires"]:
			self.token = token_cache["access_token"]
		else:
			# 获取token
			self.token = requests.get(token_url, token_params).json()
			self.token["expires"] = time.time() + self.token["expires_in"]
			f = open("token.txt", "w")
			f.write(str(self.token))
			f.close()
			self.token = self.token["access_token"]
		return self.token
	# 获取应用信息
	def get_app_info(self):
		url = "https://qyapi.weixin.qq.com/cgi-bin/agent/get"
		params = {
			"access_token": self.token,
			"agentid": "1000003" #应用ID
		}
		self.app = requests.get(url, params).json()
		return self.app
	# 获取标签ID
	def get_tags(self):
		url = "https://qyapi.weixin.qq.com/cgi-bin/tag/list"
		params = {
			"access_token": self.token,
		}
		self.tags = requests.get(url, params).json()
		return self.tags
	# 获取成员信息
	def get_user_info(self, id):
		url = "https://qyapi.weixin.qq.com/cgi-bin/user/get"
		params = {
			"access_token": self.token,
			"userid": id
		}
		self.users = requests.get(url, params).json()
		return self.users
	# 发送文本消息
	def send_text(self, message, tagId=4, userId=None):
		url = "https://qyapi.weixin.qq.com/cgi-bin/message/send"
		params = {
			"access_token": self.token,
		}
		jsonData = {
		   "touser" : userId, #用户ID
		   "toparty" : "", #部门ID
		   "totag" : tagId,
		   "msgtype" : "text",
		   "agentid" : 1000003, #应用ID
		   "text" : {
			   "content" : message
		   },
		   "safe":0
		}
		response = requests.post(url, params=params, json=jsonData).json()
		return response
	# 上传临时素材，素材上传得到media_id，该media_id仅三天内有效
	def upload_file(self, fileName):
		try:
			url = "https://qyapi.weixin.qq.com/cgi-bin/media/upload"
			params = {
				"access_token": self.token,
				"type": "file"
			}
			# requests库不支持上传以中文文件名的文件
			files = {'file': open(fileName.encode('utf-8'), 'rb')}
			response = requests.post(url, params=params, files=files).json()
			self.media = response["media_id"]
		except:
			print("requests不支持中文名文件上传，需要对requests原库进行更改。具体参考：https://www.zhihu.com/question/49583910")
		finally:
			return response
	# 发送附件
	def send_file(self, fileName, userId=None, partyId=None, tagId="4"):
		# 上传临时素材
		responseUpload = self.upload_file(fileName)
		url = "https://qyapi.weixin.qq.com/cgi-bin/message/send"
		params = {
			"access_token": self.token,
		}
		jsonData = {
		   "touser" : userId,
		   "toparty" : partyId,
		   "totag" : tagId,
		   "msgtype" : "file",
		   "agentid" : 1000003, #应用ID
		   "file" : {
				"media_id" : self.media
		   },
		   "safe":0
		}
		responseSend = requests.post(url, params=params, json=jsonData).json()
		return responseUpload, responseSend
	# 创建群聊会话
	def creat_group(self, userList, owner="2913", name=None, chatId=None):
		url = "https://qyapi.weixin.qq.com/cgi-bin/appchat/create"
		params = {
			"access_token": self.token,
		}
		jsonData = {
			"name" : name,
			"owner" : owner,
			"userlist" : userList,
			"chatid" : chatId
		}
		response = requests.post(url, params=params, json=jsonData).json()
		self.chat = response["chatid"]
		return response
	# 获取群聊信息
	def get_group_info(self):
		url = "https://qyapi.weixin.qq.com/cgi-bin/appchat/get"
		params = {
			"access_token": self.token,
			"chatid": self.chat
		}
		response = requests.get(url, params=params).json()
		return response
	# 向群聊发送文本消息
	def send_text2chat(self, message):
		url = "https://qyapi.weixin.qq.com/cgi-bin/appchat/send"
		params = {
			"access_token": self.token,
		}
		jsonData = {
			"chatid": self.chat,
			"msgtype":"text",
			"text":{
				"content" : message
			},
			"safe":0
		}
		response = requests.post(url, params=params, json=jsonData).json()
		return response
	# 向群聊发送附件
	def send_file2chat(self, fileName):
		try:
			# 上传临时素材
			responseUpload = self.upload_file(fileName)
			url = "https://qyapi.weixin.qq.com/cgi-bin/appchat/send"
			params = {
				"access_token": self.token,
			}
			jsonData = {
			   "chatid" : self.chat,
			   "msgtype" : "file",
			   "file" : {
					"media_id" : self.media
			   },
			   "safe":0
			}
			responseSend = requests.post(url, params=params, json=jsonData).json()
			if responseSend['errcode'] == 0:
				print(fileName + "发送成功OK!✔️")
		except:
			print("发送附件失败!")
		finally:	
			return responseUpload, responseSend
	def send_excel2chat(self, type):
		# 上周京东美团
		if type == 'jd':
			self.send_file2chat(jd.main())
		# 上周客流
		elif type == 'passengerFlowLastWeek':
			self.send_file2chat(passengerFlow.main('lastweek'))
		# 上周异常
		elif type == 'unusual':
			for fileName in unusual.main():
				self.send_file2chat(fileName)
		elif type == 'passengerFlowLastMonth':
			self.send_file2chat(passengerFlow.main('lastmonth'))
		else:
			print('send_excel2chat报错:非法参数!')
		print('----------------------------------------------------------------')

# 将字典格式化为json字符串
def echo(dic):
	jsonString = json.dumps(dic, ensure_ascii=False, indent=4)
	print(jsonString)
	return jsonString

# 在本地生成所有统计文件
def createFile():
	arw = arrow.now()
	# 进度条
	max_steps = len(JDshops) * 2  + arw.shift(months=-1).ceil('month').day * 2 + len(DIC) + 7*2 + len(DIC) + 4
	process_bar = ShowProcess(max_steps, '恭喜, 成功导出数据!')

	# 上周京东美团
	jd.main(process_bar)
	# 上周客流
	passengerFlow.main('lastweek', process_bar)
	# 上周异常
	unusual.main(process_bar)
	# 上月客流
	passengerFlow.main('lastmonth', process_bar)

# 设置时间自动发送文件到群里
def setTime2Do():
	# 设定时间
	monday = {
		'day_of_week': '2', # 0-6为周一到周日
		'hour': 10,
		'minute': 29
	}
	firstDay = {
		'day': '25', # 几号
		'hour': 10,
		'minute': 29
	}
	scheduler = BlockingScheduler()
	
	wx = Wxwork()
	# 上周京东美团
	scheduler.add_job(wx.send_excel2chat, 'cron', ['jd'], **monday)
	# 上周客流
	scheduler.add_job(wx.send_excel2chat, 'cron', ['passengerFlowLastWeek'], **monday)
	# 上周异常
	scheduler.add_job(wx.send_excel2chat, 'cron', ['unusual'], **monday)
	# 上月客流
	scheduler.add_job(wx.send_excel2chat, 'cron', ['passengerFlowLastMonth'], **firstDay)
	scheduler.start()

if __name__ == '__main__':
	# 定时执行
	# setTime2Do()
	
	# 手动执行
	# 仅生成文件
	createFile()
	# 生成并发送文件
	# wx = Wxwork()
	# wx.send_excel2chat('jd')
	# wx.send_excel2chat('passengerFlowLastWeek')
	# wx.send_excel2chat('unusual')
	# wx.send_excel2chat('passengerFlowLastMonth')
	time.sleep(30)
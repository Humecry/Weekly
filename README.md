# Weekly
## 配置文件
新建conf.py文件, 保存在同级目录里
```python
# 配置文件

# Excel保存路径
PATH = 'excel/'

# 京东到家与美团外卖cookie
netSales = {
	'JDcookie': 'store.o2o.jd.com1=',
	'MTcookie': 'token=; acctId=31330952',
}

# 京东到家的商铺与编号
JDshops = {'塔埔店': '11728789', '绿苑店': '11728788', '华森店': '11800856'}

# 客流数据库配置
PASSENGER_FLOW_SERVER = (
				'DRIVER={SQL Server};'
			    'SERVER=192.168.105.1,1433;'
			    'DATABASE=;'
			    'UID=;'
			    'PWD=;')
# 客流通道编码与场所名称字典
DIC = {
	'P00001S00002': '3-4号客梯',
	'P00001S00004': 'B2F超市扶梯入口',
	'P00001S00005': '1-2号客梯',
	'P00001S00006': '麦当劳',
	'P00001S00007': '屈臣氏',
	'P00001S00008': 'C大门',
	'P00001S00009': 'D大门侧门',
	'P00001S00015': 'B大门',
	'P00001S00016': 'A大门',
	'P00001S00017': '1-2号客梯1F侧门',
	'P00001S00018': 'D大门',
	'P00001S00019': '85度C',
	'P00001S00020': '尊宝披萨'
}

# 会员卡,批发,猪肉异常数据库配置
S000 = (
		'DRIVER={SQL Server};'
	    'SERVER=192.168.105.2;'
	    'UID=sa_;'
	    'PWD=;')

S003 = (
		'DRIVER={SQL Server};'
	    'SERVER=192.168.2.1;'
	    'UID=sa_;'
	    'PWD=;')

S008 = (
		'DRIVER={SQL Server};'
	    'SERVER=192.168.103.1;'
	    'UID=;'
	    'PWD=;')
```
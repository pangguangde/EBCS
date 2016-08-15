# _*_ coding: utf-8
__author__ = 'guangde'

import csv
import traceback
import xlsxwriter
import datetime


'''
注意：

1. 文件名有命名规范
	中通和申通外部的文件名应为 ‘外部中通/申通.csv’
	系统内部的文件名应为 ‘系统混合.csv’

2. csv 文件的列也有要求
	‘外部中通/申通.csv’ 文件从左到右依次为 运单号、省份、重量、价格
	‘系统混合.csv’     文件从左到右依次为 运单号、省份、重量、快递公司名称

'''


'''
申通计价规则:
	1.江浙沪皖地区:	1kg 以内按首重算, 超出部分每 100g 收取1元,		不足 100g 按 100g 算
	2.外围地区:		1kg 以内按首重算, 超出部分每 100g 收取0.63元,	不足 100g 按 100g 算

中通计价规则:
	1.江浙沪皖地区:	1kg 以内按首重算, 超出部分每 1kg 收取1元,				不足 1kg 按 1kg 算
	2.外围地区:		重量 * 续重单价
'''


def parse_csv(file_path, is_company, company_name):
	if file_path.find('/') >= 0:
		filename = file_path.split('/')[-1].encode('utf8')
	else:
		filename = file_path.split('\\')[-1]
	logger('加载' + filename + '数据')
	csvfile = file(file_path, 'rb')
	reader = csv.reader(csvfile)

	dict = {}
	orders = []
	count = 0

	cpname_col = None
	order_col = None
	province_col = None
	weight_col = None
	for column in reader:
		if is_company:
			if not weight_col:
				print column
				cpname_col = column.index('快递公司名称')
				order_col = column.index('快递单号')
				province_col = column.index('省')
				city_col = column.index('市')
				weight_col = column.index('订单毛重')
				continue
			else:
				price = 0
				cpname = column[cpname_col]
				order = column[order_col]
				province = column[province_col]
				city = column[city_col]
				weight = column[weight_col]
		else:
			if company_name == '韵达':
				order = column[0]
				province = column[1]
				weight = column[2]
				price = column[3]
				cpname = company_name
				
			else:
				order = column[0]
				province = column[1]
				weight = column[2]
				price = column[3]
				cpname = company_name
		if dict.has_key(order):
			orders.append(order)
			count += 1
		else:
			if is_company:
				dict.setdefault(order, {'province': province, 'price': price, 'weight': weight, 'cpname': cpname, 'city': city})
			else:
				dict.setdefault(order, {'province': province, 'price': price, 'weight': weight, 'cpname': cpname})
	csvfile = file('%s 重复运单号(%s).csv' % (filename, count), 'wb')
	writer = csv.writer(csvfile)
	writer.writerow(['运单号'])
	writer.writerows([[o] for o in orders])
	csvfile.close()
	print '%s 数据已经加载完毕,重复数：%s' % (filename, count)
	return dict


def shentong_price(province, weight):
	price_dict = {
		'浙江': (4, 1),
		'上海': (4, 1),
		'江苏': (4, 1),
		'安徽': (4, 1),
		'广东': (6.3, 0.63),
		'山东': (6.3, 0.63),
		'北京': (6.3, 0.63),
		'天津': (6.3, 0.63),
		'福建': (6.3, 0.63),
		'江西': (6.3, 0.63),
		'湖北': (6.3, 0.63),
		'湖南': (6.3, 0.63),
		'河北': (6.3, 0.63),
		'河南': (6.3, 0.63),
		'陕西': (6.3, 0.63),
		'辽宁': (6.3, 0.63),
		'云南': (6.3, 0.63),
		'四川': (6.3, 0.63),
		'重庆': (6.3, 0.63),
		'山西': (6.3, 0.63),
		'广西': (6.3, 0.63),
		'吉林': (6.3, 0.63),
		'贵州': (6.3, 0.63),
		'甘肃': (6.3, 0.63),
		'海南': (6.3, 0.63),
		'青海': (6.3, 0.63),
		'宁夏': (6.3, 0.63),
		'新疆': (6.3, 0.63),
		'西藏': (6.3, 0.63),
		'黑龙': (6.3, 0.63),
		'内蒙': (6.3, 0.63)
	}
	try:
		prov_key = province[0: 6]
		price_info = price_dict[prov_key]
		price_line = price_info[0]
		extra_price = price_info[1]
		if weight <= 1:
			return float('%.4f' % price_line)
		else:
			extra_weight = weight - 1
			extra_weight = hectogram(extra_weight)
			price = price_line + extra_weight * extra_price
			return float('%.4f' % price)
	except Exception, e:
		exstr = traceback.format_exc()
		print exstr


def zhongtong_price(province, weight):
	price_dict = {
		'浙江': (3.7, 1),  
		'上海': (3.7, 1),  
		'江苏': (3.7, 1),  
		'安徽': (3.7, 1),  
		'广东': (5.2, 0.4),
		'山东': (5.2, 0.4),
		'北京': (5.2, 0.4),
		'天津': (5.2, 0.4),
		'福建': (5.2, 0.4),
		'江西': (5.2, 0.4),
		'湖北': (5.2, 0.4),
		'湖南': (5.2, 0.4),
		'河北': (5.2, 0.4),
		'河南': (5.2, 0.4),
		'陕西': (5.6, 0.6),
		'辽宁': (5.6, 0.6),
		'云南': (5.6, 0.6),
		'四川': (5.6, 0.6),
		'重庆': (5.6, 0.6),
		'山西': (5.6, 0.6),
		'广西': (5.6, 0.6),
		'吉林': (5.6, 0.6),
		'贵州': (5.6, 0.6),
		'甘肃': (8.2, 0.8),
		'海南': (8.2, 0.8),
		'青海': (8.2, 0.8),
		'宁夏': (8.2, 0.8),
		'新疆': (8.2, 0.8),
		'西藏': (8.2, 0.8),
		'黑龙': (5.6, 0.6),
		'内蒙': (8.2, 0.8)
	}
	try:
		prov_key = province[0: 6]
		price_info = price_dict[prov_key]
		price_line = price_info[0]
		extra_price = price_info[1]
		if weight < 1:
			return float('%.4f' % price_line)
		else:
			extra_weight = weight - 1
			if prov_key in ['浙江', '江苏', '上海', '安徽']:
				extra_weight = kg(extra_weight)
				price = price_line + extra_weight
			else:
				price = hectogram(weight) * extra_price
			return float('%.4f' % price)
	except Exception, e:
		exstr = traceback.format_exc()
		print exstr


def yunda_price(province, city, weight):
	price_dict = {
		'上海': (3.5,	3.5,	4.0,	1.0),
		'浙江': (4.0,	4.0,	4.5,	1.0),
		'江苏': (4.0,	4.0,	4.5,	1.0),
		'安徽': (4.0,	4.0,	4.5,	1.0),
		'广东': (4.5,	5.5,	7.0,	3.5),
		'山东': (4.5,	5.5,	7.0,	3.5),
		'北京': (4.5,	5.5,	7.0,	3.5),
		'福建': (4.5,	5.5,	7.0,	3.5),
		'天津': (4.5,	5.5,	7.0,	3.5),
		'湖南': (4.5,	5.5,	7.0,	3.5),
		'河南': (4.5,	5.5,	7.0,	3.5),
		'江西': (4.5,	5.5,	7.0,	3.5),
		'湖北': (4.5,	5.5,	7.0,	3.5),
		'河北': (4.5,	5.5,	7.0,	3.5),
		'四川': (5.0,	5.5,	8.0,	5.0),
		'重庆': (5.0,	5.5,	8.0,	5.0),
		'山西': (5.0,	5.5,	8.0,	5.0),
		'广西': (5.0,	5.5,	8.0,	5.0),
		'陕西': (5.0,	5.5,	8.0,	5.0),
		'辽宁': (5.0,	6.0,	8.5,	6.0),
		'云南': (5.0,	6.0,	8.5,	6.0),
		'吉林': (5.0,	6.0,	8.5,	6.0),
		'海南': (5.0,	6.0,	8.5,	6.0),
		'贵州': (5.0,	6.0,	8.5,	6.0),
		'黑龙': (5.0,	6.0,	8.5,	6.0),
		'甘肃': (7.5,	10.0,	13.5,	7.0),
		'青海': (7.5,	10.0,	13.5,	7.0),
		'宁夏': (7.5,	10.0,	13.5,	7.0),
		'内蒙': (8.5,	13.0,	17.0,	8.0),
		'新疆': (9.5,	16.0,	23.0,	14.0),
		'西藏': (9.5,	16.0,	23.0,	14.0)
	}
	try:
		extra_weight = weight - 1
		extra_weight = kg(extra_weight)
		if city in ['唐山市', '张家口市', '承德市']:
			if weight <= 0.5:
				return 5
			elif weight <= 1:
				return 5.5
			elif weight <= 1.5:
				return 8
			else:
				price = 8 + extra_weight * 5.0
				return float('%.4f' % price)
		prov_key = province[0: 6]
		price_info = price_dict[prov_key]
		price_line1 = price_info[0]
		price_line2 = price_info[1]
		price_line3 = price_info[3]
		extra_price = price_info[3]
		if weight <= 0.5:
			return float('%.4f' % price_line1)
		elif weight <= 1:
			return float('%.4f' % price_line2)
		elif weight <= 1.5:
			return float('%.4f' % price_line3)
		else:
			price = price_line3 + extra_weight * extra_price
			return float('%.4f' % price)
	except Exception, e:
		exstr = traceback.format_exc()
		print exstr


def calculate_price(weight, province, city, company_name, order_code):
	if company_name == '中通':
		return zhongtong_price(province, weight)
	elif company_name == '申通':
		return shentong_price(province, weight)
	elif company_name == '韵达':
		return yunda_price(province, city, weight)
	else:
		print '未识别的快递公司'
	return None


def compute(company_name, waibu_file, company_file):
	logger('计算%s数据' % company_name)
	waibu_dict = parse_csv(waibu_file, is_company=False, company_name=company_name)
	transform_file_decode(company_file)
	company_dict = parse_csv('系统混合.csv', is_company=True, company_name=company_name)
	data = []
	data_1 = []

	workbook = xlsxwriter.Workbook('比对结果(%s).xlsx' % company_name)
	worksheet = workbook.add_worksheet()

	format_1 = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'right'})
	format_2 = workbook.add_format({'bold': False, 'font_color': 'green', 'align': 'right'})
	format_3 = workbook.add_format({'bold': False, 'font_color': 'gray', 'align': 'right'})
	format_4 = workbook.add_format({'align': 'right'})
	column_num = 1
	worksheet.write(0, 0, u'运单号')
	worksheet.write(0, 1, u'省份')
	worksheet.write(0, 2, u'城市')
	worksheet.write(0, 3, u'毛重(外)')
	worksheet.write(0, 4, u'毛重')
	worksheet.write(0, 5, u'毛重差(外-内)')
	worksheet.write(0, 6, u'价格(外)')
	worksheet.write(0, 7, u'价格')
	worksheet.write(0, 8, u'价格差(外-内)')
	cp_count = 0
	for item in company_dict.items():
		if waibu_dict.has_key(item[0]):
			cp_count += 1
			waibu_weight = float(waibu_dict[item[0]]['weight'])
			company_weight = float(item[1]['weight'])

			waibu_price = float(waibu_dict[item[0]]['price'])
			company_price = calculate_price(company_weight, waibu_dict[item[0]]['province'], item[1]['city'], company_name, item[0])

			weight_diff = waibu_weight - company_weight
			price_diff = float('%.4f' % (waibu_price - company_price))
			sub_str = ''
			if price_diff > 0:
				data_1.append([price_diff, (
					item[0],
					item[1]['province'],
					'%.3f' % waibu_weight,
					'%.3f' % company_weight,
					'%.3f' % weight_diff,
					'%.2f' % waibu_price,
					'%.2f' % company_price,
					'%.2f' % price_diff,
					item[1]['city']
				)])
	print 'compare count=%s' % cp_count
	data_1.sort(reverse=True)
	for i in data_1:
		data.append(i[1])
		worksheet.write(column_num, 0, i[1][0].decode('utf8'))
		worksheet.write(column_num, 1, i[1][1].decode('utf8'))
		worksheet.write(column_num, 2, i[1][8].decode('utf8'))
		worksheet.write(column_num, 3, i[1][2].decode('utf8'), format_4)
		worksheet.write(column_num, 4, i[1][3].decode('utf8'), format_4)
		if float(i[1][4]) > 0:
			worksheet.write(column_num, 5, i[1][4].decode('utf8'), format_1)
		elif float(i[1][4]) < 0:
			worksheet.write(column_num, 5, i[1][4].decode('utf8'), format_2)
		else:
			worksheet.write(column_num, 5, i[1][4].decode('utf8'), format_3)
		worksheet.write(column_num, 6, i[1][5].decode('utf8'), format_4)
		worksheet.write(column_num, 7, i[1][6].decode('utf8'), format_4)
		worksheet.write(column_num, 8, i[1][7].decode('utf8'), format_1)

		column_num += 1
	workbook.close()

	waibu_list = [item for item in waibu_dict.keys()]
	company_list = []

	for item in company_dict.items():
		if item[1]['cpname'] == company_name:
			company_list.append(item[0])
	waibu_set = set(waibu_list)
	company_set = set(company_list)

	extra_wai = waibu_set - company_set
	extra_nei = company_set - waibu_set

	logger('检查不互有的运单')

	print '外部有而公司没有的运单数: %s' % len(extra_wai)
	csvfile = file('外部有而公司没有的运单号(%s-%s).csv' % (company_name, len(extra_wai)), 'wb')
	writer = csv.writer(csvfile)
	writer.writerow(['运单号'])
	writer.writerows([[o] for o in extra_wai])
	print '公司有而外部没有的运单数: %s' % len(extra_nei)
	csvfile = file('公司有而外部没有的运单号(%s-%s).csv' % (company_name, len(extra_nei)), 'wb')
	writer = csv.writer(csvfile)
	writer.writerow(['运单号'])
	writer.writerows([[o] for o in extra_nei])
	print '-' * 60
	csvfile.close()


def transform_file_decode(filename):
	try:
		fin = open(filename, 'r')
		fout = open('系统混合.csv', 'w')
		out_str = ''
		lines = fin.readlines()
		row_num = len(lines)
		count = 1
		for line in lines:
			if count == row_num:  # 已自动去除最后一行
				break
			out_str += line.decode('gbk').encode('utf8')
			count += 1

		fout.write(out_str)
		fin.close()
		fout.close()
	except Exception, e:
		print e


def get_width(o):
	"""Return the screen column width for unicode ordinal o."""
	widths = [
		(126, 1), (159, 0), (687, 1), (710, 0), (711, 1),
		(727, 0), (733, 1), (879, 0), (1154, 1), (1161, 0),
		(4347, 1), (4447, 2), (7467, 1), (7521, 0), (8369, 1),
		(8426, 0), (9000, 1), (9002, 2), (11021, 1), (12350, 2),
		(12351, 1), (12438, 2), (12442, 0), (19893, 2), (19967, 1),
		(55203, 2), (63743, 1), (64106, 2), (65039, 1), (65059, 0),
		(65131, 2), (65279, 1), (65376, 2), (65500, 1), (65510, 2),
		(120831, 1), (262141, 2), (1114109, 1)
	]
	if o == 0xe or o == 0xf:
		return 0
	for num, wid in widths:
		if o <= num:
			return wid
	return 1


def logger(info):
	str_len = 0
	for i in info.decode('utf8'):
		str_len += get_width(ord(i))
	left_len = 20
	print '%s%s%s' % (('-' * left_len), info, ('-' * (60 - left_len - str_len)))


def kg(weight):
	weight = float('%.5f' % weight)
	return int(weight) if(int(weight) == weight) else (int(weight) + 1)


def hectogram(weight):
	weight = float('%.5f' % weight)
	a = weight * 10
	return int(a) if(int(a) == a) else (int(a) + 1)

compute('中通', '外部中通.csv', '杭州洛迈混合快递记录表格1.csv')
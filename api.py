# coding=utf-8
import csv
import json
import os
import traceback
from json import JSONEncoder

import sys
import xlsxwriter
from flask import Flask, request
from flask.ext.restful import Resource, Api

project_dir = os.path.split(os.path.realpath(sys.argv[0]))[0]
print project_dir


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
	if is_company:
		for column in reader:
			if not weight_col:
				cpname_col = column.index('快递公司名称')
				order_col = column.index('快递单号')
				province_col = column.index('省')
				weight_col = column.index('订单毛重')
				continue
			else:
				price = 0
				cpname = column[cpname_col]
				order = column[order_col]
				province = column[province_col]
				weight = column[weight_col]
			if dict.has_key(order):
				orders.append(order)
				count += 1
			else:
				dict.setdefault(order, {'province': province, 'price': price, 'weight': weight, 'cpname': cpname})
	else:
		for column in reader:
			order = column[0]
			province = column[1]
			weight = column[2]
			price = column[3]
			cpname = company_name
			if dict.has_key(order):
				orders.append(order)
				count += 1
			else:
				dict.setdefault(order, {'province': province, 'price': price, 'weight': weight, 'cpname': cpname})
	csvfile = file('%s/result/%s 重复运单号(%s).csv' % (project_dir, filename, count), 'wb')
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
		'浙江': (3.5, 1),
		'上海': (3.5, 1),
		'江苏': (3.5, 1),
		'安徽': (3.5, 1),
		'广东': (5.0, 4),
		'山东': (5.0, 4),
		'北京': (5.0, 4),
		'天津': (5.0, 4),
		'福建': (5.0, 4),
		'江西': (5.0, 4),
		'湖北': (5.0, 4),
		'湖南': (5.0, 4),
		'河北': (5.0, 4),
		'河南': (5.0, 4),
		'黑龙': (5.0, 6),
		'陕西': (5.0, 6),
		'辽宁': (5.0, 6),
		'云南': (5.0, 6),
		'四川': (5.0, 6),
		'重庆': (5.0, 6),
		'山西': (5.0, 6),
		'吉林': (5.0, 6),
		'广西': (5.0, 6),
		'贵州': (5.0, 6),
		'甘肃': (5.0, 8),
		'海南': (5.0, 8),
		'青海': (5.0, 8),
		'宁夏': (5.0, 8),
		'内蒙': (5.0, 8),
		'西藏': (8.2, 8),
		'新疆': (8.2, 8)
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
				price = price_line + extra_weight * extra_price
			return float('%.4f' % price)
	except Exception, e:
		exstr = traceback.format_exc()

		print exstr


def calculate_price(weight, province, company_name, order_code):
	if company_name == '中通':
		return zhongtong_price(province, weight)
	elif company_name == '申通':
		return shentong_price(province, weight)
	else:

		print '未识别的快递公司'
	return None


def compute(company_name, waibu_file, company_file):
	logger('计算%s数据' % company_name)
	waibu_dict = parse_csv(waibu_file, is_company=False, company_name=company_name)
	transform_file_decode(company_file)
	filename = u'系统混合'
	company_dict = parse_csv(('%s/tmp/%s.csv' % (project_dir, filename.encode('utf8'))).decode('utf8'), is_company=True, company_name=company_name)
	data = []
	data_1 = []

	workbook = xlsxwriter.Workbook('%s/result/比对结果(%s).xlsx' % (project_dir, company_name))
	worksheet = workbook.add_worksheet()

	format_1 = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'right'})
	format_2 = workbook.add_format({'bold': False, 'font_color': 'green', 'align': 'right'})
	format_3 = workbook.add_format({'bold': False, 'font_color': 'gray', 'align': 'right'})
	format_4 = workbook.add_format({'align': 'right'})
	column_num = 1
	worksheet.write(0, 0, u'运单号')
	worksheet.write(0, 1, u'省份')
	worksheet.write(0, 2, u'毛重(外)')
	worksheet.write(0, 3, u'毛重')
	worksheet.write(0, 4, u'毛重差(外-内)')
	worksheet.write(0, 5, u'价格(外)')
	worksheet.write(0, 6, u'价格')
	worksheet.write(0, 7, u'价格差(外-内)')
	cp_count = 0
	for item in company_dict.items():
		if waibu_dict.has_key(item[0]):
			cp_count += 1
			waibu_weight = float(waibu_dict[item[0]]['weight'])
			company_weight = float(item[1]['weight'])

			waibu_price = float(waibu_dict[item[0]]['price'])
			company_price = calculate_price(company_weight, waibu_dict[item[0]]['province'], company_name, item[0])

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
					'%.2f' % price_diff
				)])

	print 'compare count=%s' % cp_count
	data_1.sort(reverse=True)
	for i in data_1:
		data.append(i[0])
		worksheet.write(column_num, 0, i[1][0].decode('utf8'))
		worksheet.write(column_num, 1, i[1][1].decode('utf8'))
		worksheet.write(column_num, 2, i[1][2].decode('utf8'), format_4)
		worksheet.write(column_num, 3, i[1][3].decode('utf8'), format_4)
		if float(i[1][4]) > 0:
			worksheet.write(column_num, 4, i[1][4].decode('utf8'), format_1)
		elif float(i[1][4]) < 0:
			worksheet.write(column_num, 4, i[1][4].decode('utf8'), format_2)
		else:
			worksheet.write(column_num, 4, i[1][4].decode('utf8'), format_3)
		worksheet.write(column_num, 5, i[1][5].decode('utf8'), format_4)
		worksheet.write(column_num, 6, i[1][6].decode('utf8'), format_4)
		worksheet.write(column_num, 7, i[1][7].decode('utf8'), format_1)
		column_num += 1
	worksheet.write(column_num, 0, u'总计')
	worksheet.write(column_num, 1, '')
	worksheet.write(column_num, 2, '')
	worksheet.write(column_num, 3, '')
	worksheet.write(column_num, 4, '')
	worksheet.write(column_num, 5, '')
	worksheet.write(column_num, 6, '')
	worksheet.write(column_num, 7, sum(data))
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
	csvfile = file('%s/result/外部有而公司没有的运单号(%s-%s).csv' % (project_dir, company_name, len(extra_wai)), 'wb')
	writer = csv.writer(csvfile)
	writer.writerow(['运单号'])
	writer.writerows([[o] for o in extra_wai])
	print '公司有而外部没有的运单数: %s' % len(extra_nei)
	csvfile = file('%s/result/公司有而外部没有的运单号(%s-%s).csv' % (project_dir, company_name, len(extra_nei)), 'wb')
	writer = csv.writer(csvfile)
	writer.writerow(['运单号'])
	writer.writerows([[o] for o in extra_nei])
	print '-' * 60
	csvfile.close()


def transform_file_decode(filename):
	try:
		fin = open(filename, 'r')
		fout = open('%s/tmp/系统混合.csv' % project_dir, 'w')
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
	return int(weight) if (int(weight) == weight) else (int(weight) + 1)


def hectogram(weight):
	weight = float('%.5f' % weight)
	a = weight * 10
	return int(a) if (int(a) == a) else (int(a) + 1)


app = Flask(__name__)
api = Api(app)


@app.route('/zhongtong_file')
def recv_zhongtong_path():
	data = json.load(open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'r'))
	data['zhongtong_filepath'] = request.args.get('file_path')
	json.dump(data, open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'w'))
	return JSONEncoder().encode({'succ': 200, 'msg': 'zhongtong file path received'})


@app.route('/shentong_file')
def recv_shentong_path():
	data = json.load(open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'r'))
	data['shentong_filepath'] = request.args.get('file_path')
	json.dump(data, open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'w'))
	return JSONEncoder().encode({'succ': 200, 'msg': 'shentong file path received'})


@app.route('/neibu_file')
def recv_neibu_path():
	data = json.load(open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'r'))
	data['neibu_filepath'] = request.args.get('file_path')
	json.dump(data, open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'w'))
	return JSONEncoder().encode({'succ': 200, 'msg': 'neibu file path received'})


@app.route('/check')
def check():
	data = json.load(open('/Users/pangguangde/Documents/APIServer/res/filepath.json', 'r'))
	zhongtong_path = data.get('zhongtong_filepath')
	shentong_path = data.get('shentong_filepath')
	neibu_path = data.get('neibu_filepath')
	mode = request.args.get('mode')
	if neibu_path is None:
		return JSONEncoder().encode({'succ': -1, 'msg': 'neibu file path is require!'})
	else:
		if mode == 'zhongtong':
			compute('中通', zhongtong_path, neibu_path)
		elif mode == 'shentong':
			compute('申通', shentong_path, neibu_path)
		else:
			compute('中通', zhongtong_path, neibu_path)
			compute('申通', shentong_path, neibu_path)
		return JSONEncoder().encode({'succ': 200, 'msg': 'check finish!'})


if __name__ == '__main__':
	app.run(debug=True)

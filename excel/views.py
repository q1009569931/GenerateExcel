from django.shortcuts import render
from django.views.generic.base import View
from django.http import HttpResponse
import json
from django.conf import settings
import os
import xlwt
import random
import time
from django.core.cache import cache #引入缓存模块
from decimal import Decimal
# Create your views here.

# 公差数据
# 尺寸公差
size_diff = [0, 3, 6, 30, 120, 400, 1000, 2000]
size_diff_high = [0.05, 0.05, 0.1, 0.15, 0.2, 0.3, 0.5, 0]
size_diff_middle = [0.1, 0.2, 0.3, 0.5, 0.8, 1.2, 2.0, 3.0, 4.0]
size_diff_low = [0.2, 0.3, 0.5, 0.8, 1.2, 2.0, 3.0, 4.0]
size_diff_lowest = [0, 0.5, 1.0, 1.5, 2.5, 4.0, 6.0, 8.0]
# 倒角公差
r_diff = [0, 3, 6, 30]
r_diff_high = [0.2, 0.5, 1.0, 2.0]
r_diff_low = [0.4, 1.0, 2.0, 4.0]

def ListToMinNum(L):
	# 把列表的公差细分成0.01间隔的数
	# 为避免浮点误差，使用Decimal
	for i in range(0, len(L)):
		low = Decimal(str(0 - L[i]))
		high = Decimal(str(L[i]))
		NumList = []
		while low <= high:
			NumList.append(low)
			low += Decimal(str(0.01))
		L[i] = NumList[:]

ListToMinNum(size_diff_high)
ListToMinNum(size_diff_middle)
ListToMinNum(size_diff_low)
ListToMinNum(size_diff_lowest)
ListToMinNum(r_diff_high)
ListToMinNum(r_diff_low)

def get_data_diff_real(diff, diff_rank, num_require):
	for i in range(0, len(diff)-1):
			if diff[i] < num_require <= diff[i+1]:
				diff = diff_rank[i]
				diff = random.choice(diff)
				num_real = num_require + diff

				diff = str(diff)
				num_real = str(num_real)
				num_require = str(num_require.quantize(Decimal("0.0")))

				result = {
					"num_require": num_require,
					"diff": diff,
					"num_real": num_real
				}

				return result

# 3个函数处理不同的数据类型
def hole(data):
	hole_diff = data["hole_diff"]
	num_input = data["num_input"]

	num_require = Decimal(num_input).quantize(Decimal("0.00"))
	hole_diff = Decimal(hole_diff).quantize(Decimal("0.00"))
	num_real = num_require + hole_diff

	diff = str(hole_diff)
	num_real = str(num_real)
	num_require = str(num_require.quantize(Decimal("0.0")))

	result = {
		"num_require": num_require,
		"diff": diff,
		"num_real": num_real
	}

	return result


def r(data):
	num_require = data["num_input"]
	rank = data["rank"]

	num_require = Decimal(num_require).quantize(Decimal("0.00"))
	if rank == "high" or rank == "middle":
		return get_data_diff_real(r_diff, r_diff_high, num_require)
	else:
		return get_data_diff_real(r_diff, r_diff_low, num_require)

def line(data):
	num_require = data["num_input"]
	rank = data["rank"]

	num_require = Decimal(num_require).quantize(Decimal("0.00"))
	if rank == "high":
		return get_data_diff_real(size_diff, size_diff_high, num_require)
	elif rank == "middle":
		return get_data_diff_real(size_diff, size_diff_middle, num_require)
	elif rank == "low":
		return get_data_diff_real(size_diff, size_diff_low, num_require)
	else:
		return get_data_diff_real(size_diff, size_diff_lowest, num_require)

# 把函数放在字典里，减少if else语句的使用
dict_name = {
	"line": line,
	"r": r,
	"hole": hole
}

# 根据时间戳生成excel的路径+文件名
def excelpath():
	timestamp = str(time.time()).split(".")[-1]
	num = random.randint(1,10)
	filename = timestamp + str(num) + ".xls"
	return os.path.join(settings.MEDIA_ROOT, filename)


# 字体格式
def FontStyle(name, height, bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style

# 写入excel
def write_excel(datas):
	f = xlwt.Workbook()
	sheet = f.add_sheet('仪器', cell_overwrite_ok=True)
	row0 = ['要求值', '测试值', '差值']

	# 写入第一行
	for i in range(0, len(row0)):
		sheet.write(0, i, row0[i], FontStyle('Times New Roman',220,True))

	line = 1

	# 遍历列表，写入数据
	for data in datas:
		num_require = data["num_require"]
		num_real = data["num_real"]
		diff = data["diff"]
		sheet.write(line, 0, num_require)
		sheet.write(line, 1, num_real)
		sheet.write(line, 2, diff)
		line += 1

	EXCELPATH = excelpath()
	# 保存文件
	if os.path.exists(EXCELPATH):
		os.remove(EXCELPATH)
	f.save(EXCELPATH)
	return EXCELPATH

class GDate(View):
	def get(self, request):
		return render(request, 'index.html')

	def post(self, request):
		data = json.loads(request.body.decode())
		name = data["name"]
		func_for_data = dict_name.get(name)
		result = func_for_data(data)

		return HttpResponse(json.dumps(result), content_type="application/json")

class Excel(View):
	def post(self, request):
		# 第一次ajax传送数据过来生成excel表，第二次请求是form表单提交用于下载
		if request.META['CONTENT_TYPE'].find('text') != -1:
			data = json.loads(request.body.decode()) # 接受数据经转换后，是列表里面嵌套字典。
			EXCELPATH = write_excel(data)
			# 路径问题，使用redis，cookie为键，路径为值
			# 使用redis的字符串类型存储
			cache.set(request.META["HTTP_COOKIE"], EXCELPATH)
		EXCELPATH = cache.get(request.META["HTTP_COOKIE"])
		file = open(EXCELPATH, 'rb')
		response = HttpResponse(file)
		response['Content-Type'] = 'application/octet-stream' #设置头信息，告诉浏览器这是个文件
		response['Content-Disposition'] = 'attachment;filename="result.xls"'
		return response

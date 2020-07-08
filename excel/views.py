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
# 尺寸公差表
size_require = [0, 3, 6, 30, 120, 400, 1000, 2000]
size_diff_high = [0.05, 0.05, 0.1, 0.15, 0.2, 0.3, 0.5, 0]
size_diff_middle = [0.1, 0.2, 0.3, 0.5, 0.8, 1.2, 2.0, 3.0, 4.0]
size_diff_low = [0.2, 0.3, 0.5, 0.8, 1.2, 2.0, 3.0, 4.0]
size_diff_lowest = [0, 0.5, 1.0, 1.5, 2.5, 4.0, 6.0, 8.0]
# 倒角公差表
r_require = [0, 3, 6, 30]
r_diff_high = [0.2, 0.5, 1.0, 2.0]
r_diff_low = [0.4, 1.0, 2.0, 4.0]

# 实际差值
real_diff = [-0.02, -0.01, 0.00, 0.01, 0.02]

def get_diff(diff_list, num_list, num_require):
	'''
	从公差表获得公差
	'''
	for i in range(0, len(num_list)-1):
		if num_list[i] < num_require <= num_list[i+1]:
			return diff_list[i]

# 线性公差
def line(num_require, rank, hole_diff):
	if rank == "high":
		diff = get_diff(size_diff_high, size_require, num_require)
	elif rank == "middle":
		diff = get_diff(size_diff_middle, size_require, num_require)
	elif rank == "low":
		diff = get_diff(size_diff_low, size_require, num_require)
	else:
		diff = get_diff(size_diff_lowest, size_require, num_require)
	return "±" + str(diff)

# 倒角公差
def r(num_require, rank, hole_diff):
	if rank == "high" or rank == "middle":
		diff = get_diff(r_diff_high, r_require, num_require)
	else:
		diff = get_diff(r_diff_low, r_require, num_require)
	return "±" + str(diff)

# 孔位公差
def hole(num_require, rank, hole_diff):
	return "±" + hole_diff

# 把函数放到字典里，减少if else判断
func_dict = {
	"line": line,
	"r": r,
	"hole": hole
}
# 字母对应的name
name_dict = {
	"l": "line",
	"r": "r",
	"h": "hole"
}

def get_result(data):
	num_require = float(data["num_input"])
	the_real_diff = random.choice(real_diff)
	name = name_dict[data["name"]]
	rank = data["rank"]
	hole_diff = data["hole_diff"]

	func_get_data = func_dict[name]
	diff = func_get_data(num_require, rank, hole_diff)
	num_real = num_require + the_real_diff

	# 全部变成字符串
	num_require = str(Decimal(num_require).quantize(Decimal("0.0")))
	num_real = str(Decimal(num_real).quantize(Decimal("0.00")))
	the_real_diff = str(Decimal(the_real_diff).quantize(Decimal("0.00")))

	result = {
		"num_require": num_require,
		"diff": diff,
		"num_real": num_real,
		"real_diff": the_real_diff
	}

	return result



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
		print(data)
		result = get_result(data)

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

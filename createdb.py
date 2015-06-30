#---------------------------------------------------------------
#    程序：CIQ2k查询数据库
#    版本：0.0.3
#    作者：懒秀才
#    日期：2014-05-28
#    语言：Python 3
#    操作：CIQ2k查询数据库新建、导入数据等操作
#    函数：
#          createCiq2kMdb()		创建空数据库[ok]2014-05-28
#          createExit()			创建出境表结构[ok]2014-05-28
#          createEntry()		创建入境表结构[ok]2014-05-28
#          createExitDisq()		创建出境不合格表结构[ok]2014-05-28
#          createEntryDisq()	创建入境检验不合格表结构[ok]2014-05-28
#          createCert()			创建出境签证表结构[ok]2014-05-28
#          createEntrySit()		创建疫情表结构[ok]2014-05-28
#          createCon()			创建集装箱表结构[ok]2014-05-28
#          createExitProd()		创建出境生产企业表结构[ok]2014-05-30
#          connectCiq2KMdb()	连接数据库[ok]2014-05-28
#          insertExit()			导入出境表数据(追加)[ok]2014-05-28
#          insertEntry()		导入入境表数据(追加)[ok]2014-05-28
#          insertExitDisq()		导入出境检验不合格表数据(追加)[ok]2014-05-28
#          insertEntryDisq()	导入入境检验不合格表数据(追加)[ok]2014-05-28
#          insertCert()			导入签证表数据(追加)[ok]2014-05-28
#          insertEntrySit()		导入入境疫情表数据(追加)[ok]2014-05-30
#          insertCon()			导入集装箱表数据(追加)[ok]2014-05-30
#          insertExitProd()		导入出境生产企业表数据(追加)[ok]2014-05-30
#          insertBasicDb()		创建基础数据表
#---------------------------------------------------------------

# 0.0.2 2014-05-30	增加创建基础数据表
# 0.0.2 2014-05-30	追加表数据引入更新计数器
# 0.0.3 2015-06-30	修改追加数据导入文本文件名



#! /usr/bin/env python
#coding=utf-8

import os
import datetime,time
import pypyodbc

def createCiq2kMdb():
	pypyodbc.win_create_mdb('D:\\db.mdb')

def connectCiq2KMdb():
	
	conn = pypyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};DBQ=D:\\db.mdb')
	cur = conn.cursor()

def temp():
#--------------------------------------------------------------------
#    功能：在当前目录新建数据库db.mdb，建立表结构，导入对应基础数据表
#
#          出境表结构[ok]2014-05-28
#          入境表结构[ok]2014-05-28
#          出境检验不合格表结构[ok]2014-05-28
#          入境检验不合格表结构[ok]2014-05-28
#          出境签证表结构
#          集装箱表结构
#          入境疫情表结构
#
#          出境生产企业表结构
#          出境报检企业表结构
#          入境报检企业表结构
#          入境受货人表结构
#
#          施检科室表结构及数据
#          国家表结构及数据
#          欧盟表结构及数据
#          协议国家表结构及数据
#--------------------------------------------------------------------

# 创建基础数据表
def createBasicDb():
	
	connectCiq2KMdb()
	
	# 创建科室表
	cur.execute('''CREATE TABLE 科室 (
	ID COUNTER PRIMARY KEY,
	科室代码 VARCHAR(12),
	科室名称 VARCHAR(50);''')
	cur.commit()
	
	# 创建国家地区表
	cur.execute('''CREATE TABLE 国家 (
	ID COUNTER PRIMARY KEY,
	国家代码 VARCHAR(4),
	国家地区名称 VARCHAR(50);''')
	cur.commit()
	
	# 创建抽检模式表
	cur.execute('''CREATE TABLE 抽检模式 (
	ID COUNTER PRIMARY KEY,
	抽检代码 VARCHAR(1),
	抽检模式名称 VARCHAR(50);''')
	cur.commit()
	
	# 创建检验监管模式表
	cur.execute('''CREATE TABLE 检验监管模式 (
	ID COUNTER PRIMARY KEY,
	检验监管模式代码 VARCHAR(4),
	检验监管模式名称 VARCHAR(50);''')
	cur.commit()
		
	
	conn.close()
	
# 创建出境表
def createExit():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 出境 (
	ID COUNTER PRIMARY KEY,
	报检号 VARCHAR(20),
	报检类别代码 VARCHAR(4),
	报检单位注册号 VARCHAR(10),
	发货人代码 VARCHAR(10),
	发货人企业性质代码 VARCHAR(4),
	发货人行政区划代码 VARCHAR(8),
	发货人海关注册代码 VARCHAR(10),
	运输方式代码 VARCHAR(4),
	贸易方式代码 VARCHAR(4),
	启运口岸代码 VARCHAR(8),
	到达口岸代码 VARCHAR(8),
	贸易国家地区代码 VARCHAR(4),
	检验检疫方式代码 VARCHAR(4),
	检验依据类别代码 VARCHAR(4),
	是否抽检 VARCHAR(1),
	是否二次抽检 VARCHAR(1),
	是否隔离检疫 VARCHAR(1),
	货物用途代码 VARCHAR(4),
	检验检疫结果代码 VARCHAR(4),
	查验结果代码 VARCHAR(4),
	报检日期 datetime,
	收费日期 datetime,
	检验检疫完成日期 datetime,
	统计日期 datetime,
	上报日期 datetime,
	接收日期 datetime,
	检验检疫机构代码 VARCHAR(8),
	HS编码 VARCHAR(12),
	是否法检 VARCHAR(1),
	HS数重量 float,
	HS计量单位 VARCHAR(4),
	商品统计分类代码 VARCHAR(12),
	商品统计数重量 float,
	商品统计单位代码 VARCHAR(4),
	产地代码 VARCHAR(8),
	包装种类代码 VARCHAR(4),
	包装件数 float,
	货值美元 float,
	检验检疫项目代码 VARCHAR(50),
	检验检疫不合格内容代码 VARCHAR(50),
	不合格数重量 float,
	不合格金额美元 float,
	检验不合格原因代码 VARCHAR(4),
	检验不合格处理代码 VARCHAR(4),
	检疫不合格原因代码 VARCHAR(4),
	检疫处理方法代码 VARCHAR(4),
	检疫具体处理方法代码 VARCHAR(4),
	检疫处理机构代码 VARCHAR(8),
	检疫处理部门代码 VARCHAR(4),
	货物名称 VARCHAR(50),
	是否危险品 VARCHAR(1),
	是否出境重点商品 VARCHAR(1),
	是否出境大宗商品 VARCHAR(1),
	是否需要出境许可证 VARCHAR(1),
	民用品标志 VARCHAR(1),
	监管条件 VARCHAR(20),
	area_sta_bunc VARCHAR(50),
	施检科室代码 VARCHAR(12),
	生产企业代码 VARCHAR(10),
	检验监管模式代码 VARCHAR(4),
	检毕日期 datetime;''')

	cur.commit()
	conn.close()

# 创建入境表
def createEntry():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 入境 (
	ID COUNTER PRIMARY KEY,
	报检号 VARCHAR(20),
	报检类别代码 VARCHAR(4),
	报检单位注册号 VARCHAR(10),
	受货人代码 VARCHAR(10),
	发货人企业性质代码 VARCHAR(4),
	发货人行政区划代码 VARCHAR(8),
	发货人海关注册代码 VARCHAR(10),
	运输方式代码 VARCHAR(4),
	贸易方式代码 VARCHAR(4),
	启运口岸代码 VARCHAR(8),
	经停口岸代码 VARCHAR(8),
	贸易国家地区代码 VARCHAR(4),
	检验检疫方式代码 VARCHAR(4),
	检验依据类别代码 VARCHAR(4),
	是否抽检 VARCHAR(1),
	是否二次抽检 VARCHAR(1),
	是否隔离检疫 VARCHAR(1),
	货物用途代码 VARCHAR(4),
	检验检疫结果代码 VARCHAR(4),
	查验结果代码 VARCHAR(4),
	报检日期 datetime,
	收费日期 datetime,
	检验检疫完成日期 datetime,
	统计日期 datetime,
	上报日期 datetime,
	接收日期 datetime,
	检验检疫机构代码 VARCHAR(8),
	HS编码 VARCHAR(12),
	是否法检 VARCHAR(1),
	HS数重量 float,
	HS计量单位 VARCHAR(4),
	商品统计分类代码 VARCHAR(12),
	商品统计数重量 float,
	商品统计单位代码 VARCHAR(4),
	原产国代码 VARCHAR(4),
	包装种类代码 VARCHAR(4),
	包装件数 float,
	货值美元 float,
	检验检疫项目代码 VARCHAR(50),
	检验检疫不合格内容代码 VARCHAR(50),
	不合格数重量 float,
	不合格金额美元 float,
	检验不合格原因代码 VARCHAR(4),
	检验不合格处理代码 VARCHAR(4),
	检疫不合格原因代码 VARCHAR(4),
	检疫处理方法代码 VARCHAR(4),
	检疫具体处理方法代码 VARCHAR(4),
	检疫处理机构代码 VARCHAR(8),
	检疫处理部门代码 VARCHAR(4),
	货物名称 VARCHAR(50),
	是否危险品 VARCHAR(1),
	是否入境重点商品 VARCHAR(1),
	是否入境大宗商品 VARCHAR(1),
	是否需要入境许可证 VARCHAR(1),
	民用品标志 VARCHAR(1),
	监管条件 VARCHAR(20),
	area_sta_bunc VARCHAR(50),
	施检科室代码 VARCHAR(12),
	入境口岸代码 VARCHAR(8),
	目的地代码 VARCHAR(8),
	启运国家地区 VARCHAR(4),
	疫情名称代码 VARCHAR(8),
	疫情级别代码 VARCHAR(1),
	是否废旧品 VARCHAR(1),
	是否外商投资财产 VARCHAR(1),
	检验监管模式代码 VARCHAR(4),
	索赔金额美元 float,
	检毕日期 datetime,
	制造商 VARCHAR(50);''')

	cur.commit()
	conn.close()

# 创建出境检验不合格表
def createExitDisq():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 出境检验不合格 (
	ID COUNTER PRIMARY KEY,
	报检号 VARCHAR(20),
	报检类别代码 VARCHAR(4),
	报检单位注册号 VARCHAR(10),
	发货人代码 VARCHAR(10),
	发货人企业性质代码 VARCHAR(4),
	发货人行政区划代码 VARCHAR(8),
	发货人海关注册代码 VARCHAR(10),
	运输方式代码 VARCHAR(4),
	贸易方式代码 VARCHAR(4),
	启运口岸代码 VARCHAR(8),
	到达口岸代码 VARCHAR(8),
	贸易国家地区代码 VARCHAR(4),
	检验检疫方式代码 VARCHAR(4),
	检验依据类别代码 VARCHAR(4),
	是否抽检 VARCHAR(1),
	是否二次抽检 VARCHAR(1),
	是否隔离检疫 VARCHAR(1),
	货物用途代码 VARCHAR(4),
	检验检疫结果代码 VARCHAR(4),
	查验结果代码 VARCHAR(4),
	报检日期 datetime,
	收费日期 datetime,
	检验检疫完成日期 datetime,
	统计日期 datetime,
	上报日期 datetime,
	接收日期 datetime,
	检验检疫机构代码 VARCHAR(8),
	HS编码 VARCHAR(12),
	是否法检 VARCHAR(1),
	HS数重量 float,
	HS计量单位 VARCHAR(4),
	商品统计分类代码 VARCHAR(12),
	商品统计数重量 float,
	商品统计单位代码 VARCHAR(4),
	产地代码 VARCHAR(8),
	包装种类代码 VARCHAR(4),
	包装件数 float,
	货值美元 float,
	检验检疫项目代码 VARCHAR(50),
	检验检疫不合格内容代码 VARCHAR(50),
	不合格数重量 float,
	不合格金额美元 float,
	检验不合格原因代码 VARCHAR(4),
	检验不合格处理代码 VARCHAR(4),
	检疫不合格原因代码 VARCHAR(4),
	检疫处理方法代码 VARCHAR(4),
	检疫具体处理方法代码 VARCHAR(4),
	检疫处理机构代码 VARCHAR(8),
	检疫处理部门代码 VARCHAR(4),
	货物名称 VARCHAR(50),
	是否危险品 VARCHAR(1),
	是否出境重点商品 VARCHAR(1),
	是否出境大宗商品 VARCHAR(1),
	是否需要出境许可证 VARCHAR(1),
	民用品标志 VARCHAR(1),
	监管条件 VARCHAR(20),
	area_sta_bunc VARCHAR(50),
	施检科室代码 VARCHAR(12),
	生产企业代码 VARCHAR(10),
	检验监管模式代码 VARCHAR(4),
	检验具体项目代码串 VARCHAR(50),
	检验检出数量 float,
	结果描述 VARCHAR(50);''')

	cur.commit()
	conn.close()

# 创建入境检验不合格表
def createEntryDisq():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 入境检验不合格 (
	ID COUNTER PRIMARY KEY,
	报检号 VARCHAR(20),
	报检类别代码 VARCHAR(4),
	报检单位注册号 VARCHAR(10),
	受货人代码 VARCHAR(10),
	发货人企业性质代码 VARCHAR(4),
	发货人行政区划代码 VARCHAR(8),
	发货人海关注册代码 VARCHAR(10),
	运输方式代码 VARCHAR(4),
	贸易方式代码 VARCHAR(4),
	启运口岸代码 VARCHAR(8),
	经停口岸代码 VARCHAR(8),
	贸易国家地区代码 VARCHAR(4),
	检验检疫方式代码 VARCHAR(4),
	检验依据类别代码 VARCHAR(4),
	是否抽检 VARCHAR(1),
	是否二次抽检 VARCHAR(1),
	是否隔离检疫 VARCHAR(1),
	货物用途代码 VARCHAR(4),
	检验检疫结果代码 VARCHAR(4),
	查验结果代码 VARCHAR(4),
	报检日期 datetime,
	收费日期 datetime,
	检验检疫完成日期 datetime,
	统计日期 datetime,
	上报日期 datetime,
	接收日期 datetime,
	检验检疫机构代码 VARCHAR(8),
	HS编码 VARCHAR(12),
	是否法检 VARCHAR(1),
	HS数重量 float,
	HS计量单位 VARCHAR(4),
	商品统计分类代码 VARCHAR(12),
	商品统计数重量 float,
	商品统计单位代码 VARCHAR(4),
	原产国代码 VARCHAR(4),
	包装种类代码 VARCHAR(4),
	包装件数 float,
	货值美元 float,
	检验检疫项目代码 VARCHAR(50),
	检验检疫不合格内容代码 VARCHAR(50),
	不合格数重量 float,
	不合格金额美元 float,
	检验不合格原因代码 VARCHAR(4),
	检验不合格处理代码 VARCHAR(4),
	检疫不合格原因代码 VARCHAR(4),
	检疫处理方法代码 VARCHAR(4),
	检疫具体处理方法代码 VARCHAR(4),
	检疫处理机构代码 VARCHAR(8),
	检疫处理部门代码 VARCHAR(4),
	货物名称 VARCHAR(50),
	是否危险品 VARCHAR(1),
	是否入境重点商品 VARCHAR(1),
	是否入境大宗商品 VARCHAR(1),
	是否需要入境许可证 VARCHAR(1),
	民用品标志 VARCHAR(1),
	监管条件 VARCHAR(20),
	area_sta_bunc VARCHAR(50),
	施检科室代码 VARCHAR(12),
	入境口岸代码 VARCHAR(8),
	目的地代码 VARCHAR(8),
	启运国家地区 VARCHAR(4),
	疫情名称代码 VARCHAR(8),
	疫情级别代码 VARCHAR(1),
	是否废旧品 VARCHAR(1),
	是否外商投资财产 VARCHAR(1),
	检验监管模式代码 VARCHAR(4),
	检验具体项目代码串 VARCHAR(50),
	检验检出数量 float,
	结果描述 VARCHAR(50),
	制造商 VARCHAR(50);''')

	cur.commit()
	conn.close()

# 创建出境签证表
def createCert():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 出境签证 (
	ID COUNTER PRIMARY KEY,
	证单种类编号 VARCHAR(4),
	报检号 VARCHAR(20),
	扩展号 VARCHAR(2),
	语种 VARCHAR(4),
	报检类别代码 VARCHAR(4),
	检验检疫机构代码 VARCHAR(8),
	施检部门代码 VARCHAR(12),
	签证内容代码 VARCHAR(4),
	报检日期 datetime,
	收费日期 datetime,
	证单日期 datetime,
	检验检疫完成日期 datetime,
	统计日期 datetime,
	上报日期 datetime,
	接收日期 datetime,
	area_sta_bunc VARCHAR(50);''')

	cur.commit()
	conn.close()

# 创建入境疫情表
def createEntrySit():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 入境疫情 (
	ID COUNTER PRIMARY KEY,
	报检号 VARCHAR(20),
	报检类别代码 VARCHAR(4),
	报检单位注册号 VARCHAR(10),
	受货人代码 VARCHAR(10),
	发货人企业性质代码 VARCHAR(4),
	发货人行政区划代码 VARCHAR(8),
	发货人海关注册代码 VARCHAR(10),
	运输方式代码 VARCHAR(4),
	贸易方式代码 VARCHAR(4),
	启运口岸代码 VARCHAR(8),
	经停口岸代码 VARCHAR(8),
	贸易国家地区代码 VARCHAR(4),
	检验检疫方式代码 VARCHAR(4),
	检验依据类别代码 VARCHAR(4),
	是否抽检 VARCHAR(1),
	是否二次抽检 VARCHAR(1),
	是否隔离检疫 VARCHAR(1),
	货物用途代码 VARCHAR(4),
	检验检疫结果代码 VARCHAR(4),
	查验结果代码 VARCHAR(4),
	报检日期 datetime,
	收费日期 datetime,
	检验检疫完成日期 datetime,
	统计日期 datetime,
	上报日期 datetime,
	接收日期 datetime,
	检验检疫机构代码 VARCHAR(8),
	HS编码 VARCHAR(12),
	是否法检 VARCHAR(1),
	HS数重量 float,
	HS计量单位 VARCHAR(4),
	商品统计分类代码 VARCHAR(12),
	商品统计数重量 float,
	商品统计单位代码 VARCHAR(4),
	原产国代码 VARCHAR(4),
	包装种类代码 VARCHAR(4),
	包装件数 float,
	货值美元 float,
	检验检疫项目代码 VARCHAR(50),
	检疫不合格原因代码 VARCHAR(4),
	检疫处理方法代码 VARCHAR(4),
	检疫具体处理方法代码 VARCHAR(4),
	检疫处理机构代码 VARCHAR(8),
	检疫处理部门代码 VARCHAR(4),
	HS中文名称 VARCHAR(50),
	是否危险品 VARCHAR(1),
	是否入境重点商品 VARCHAR(1),
	是否入境大宗商品 VARCHAR(1),
	是否需要入境许可证 VARCHAR(1),
	民用品标志 VARCHAR(1),
	监管条件 VARCHAR(20),
	area_sta_bunc VARCHAR(50),
	施检科室代码 VARCHAR(12),
	入境口岸代码 VARCHAR(8),
	目的地代码 VARCHAR(8),
	启运国家地区 VARCHAR(4),
	疫情名称代码 VARCHAR(8),
	疫情级别代码 VARCHAR(1),
	是否废旧品 VARCHAR(1),
	是否外商投资财产 VARCHAR(1),
	dis VARCHAR(50),
	货物序号 float,
	检疫不合格数量 float,
	检疫不合格金额 float;''')

	cur.commit()
	conn.close()
	
# 创建集装箱表
def createCon():
	connectCiq2KMdb()
	cur.execute('''CREATE TABLE 集装箱 (
	ID COUNTER PRIMARY KEY,
	报检号 VARCHAR(20),
	集装箱规格代码 VARCHAR(4),
	集装箱数量 float,
	处理集装箱数量 float,
	集装箱结果评定代码 VARCHAR(1),
	集装箱不合格原因代码 VARCHAR(2),
	是否预防性处理 VARCHAR(1),
	集装箱号码串 VARCHAR(200),
	具体处理办法代码 VARCHAR(2),
	卫生处理机构代码 VARCHAR(8),
	卫生处理部门代码 VARCHAR(12),
	适载鉴定结果代码 VARCHAR(50),
	适载鉴定项目代码 VARCHAR(50),
	适载不合格项目代码 VARCHAR(50),
	适载不合格处理方式代码 VARCHAR(50),
	鉴定机构代码 VARCHAR(8),
	鉴定部门代码 VARCHAR(12),
	检疫不合格数量 float,
	鉴定不合格数量 float,
	综合不合格数量 float,
	是否来自疫区 VARCHAR(1),
	出入境标志 VARCHAR(1),
	检验检疫完成日期 datetime,
	检验检疫机构代码 VARCHAR(8),
	检验检疫部门代码 VARCHAR(12),
	area_sta_bunc VARCHAR(50);''')

	cur.commit()
	conn.close()
	
# 创建出境生产企业表结构
def createExitProd():
	connectCiq2KMdb()
	
	cur.execute('''CREATE TABLE 出境生产企业 (
	ID COUNTER PRIMARY KEY,
	出境生产企业代码 VARCHAR(10),
	出境生产企业名称 VARCHAR(50);''')
	
	cur.commit()
	conn.close()

	
#-----------------------------------
#    功能：导入出境表数据(追加)
#-----------------------------------
def insertExit():
	connectCiq2KMdb()
	data = open('出境.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list:  
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化数值，转换数值内容字符串为数值
		for i_n in [29,32,36,37,40,41]:
			decls[i_n] = float(decls[i_n]

		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [20,21,22,23,24,25,60]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 出境 (报检号,报检类别代码,报检单位注册号,发货人代码,发货人企业性质代码,发货人行政区划代码,发货人海关注册代码,运输方式代码,贸易方式代码,启运口岸代码,到达口岸代码,贸易国家地区代码,检验检疫方式代码,检验依据类别代码,是否抽检,是否二次抽检,是否隔离检疫,货物用途代码,检验检疫结果代码,查验结果代码,报检日期,收费日期,检验检疫完成日期,统计日期,上报日期,接收日期,检验检疫机构代码,HS编码,是否法检,HS数重量,HS计量单位,商品统计分类代码,商品统计数重量,商品统计单位代码,产地代码,包装种类代码,包装件数,货值美元,检验检疫项目代码,检验检疫不合格内容代码,不合格数重量,不合格金额美元,检验不合格原因代码,检验不合格处理代码,检疫不合格原因代码,检疫处理方法代码,检疫具体处理方法代码,检疫处理机构代码,检疫处理部门代码,货物名称,是否危险品,是否出境重点商品,是否出境大宗商品,是否需要出境许可证,民用品标志,监管条件,area_sta_bunc,施检科室代码,生产企业代码,检验监管模式代码,检毕日期) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15],decls[16],decls[17],decls[18],decls[19],decls[20],decls[21],decls[22],decls[23],decls[24],decls[25],decls[26],decls[27],decls[28],decls[29],decls[30],decls[31],decls[32],decls[33],decls[34],decls[35],decls[36],decls[37],decls[38],decls[39],decls[40],decls[41],decls[42],decls[43],decls[44],decls[45],decls[46],decls[47],decls[48],decls[49],decls[50],decls[51],decls[52],decls[53],decls[54],decls[55],decls[56],decls[57],decls[58],decls[59],decls[60]))
		if commitCount == 2500 :
			cur.commit()
			commitCount = 1
	cur.commit()
	conn.close()
	
#-----------------------------------
#    功能：导入入境表数据(追加)
#-----------------------------------
def insertEntry():
	connectCiq2KMdb()
	data = open('入境.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list:  
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化数值，转换数值内容字符串为数值
		for i_n in [29,32,36,37,40,41,66]:
			decls[i_n] = float(decls[i_n]

		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [20,21,22,23,24,25,67]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 入境 (报检号,报检类别代码,报检单位注册号,受货人代码,发货人企业性质代码,发货人行政区划代码,发货人海关注册代码,运输方式代码,贸易方式代码,启运口岸代码,经停口岸代码,贸易国家地区代码,检验检疫方式代码,检验依据类别代码,是否抽检,是否二次抽检,是否隔离检疫,货物用途代码,检验检疫结果代码,查验结果代码,报检日期,收费日期,检验检疫完成日期,统计日期,上报日期,接收日期,检验检疫机构代码,HS编码,是否法检,HS数重量,HS计量单位,商品统计分类代码,商品统计数重量,商品统计单位代码,原产国代码,包装种类代码,包装件数,货值美元,检验检疫项目代码,检验检疫不合格内容代码,不合格数重量,不合格金额美元,检验不合格原因代码,检验不合格处理代码,检疫不合格原因代码,检疫处理方法代码,检疫具体处理方法代码,检疫处理机构代码,检疫处理部门代码,货物名称,是否危险品,是否入境重点商品,是否入境大宗商品,是否需要入境许可证,民用品标志,监管条件,area_sta_bunc,施检科室代码,入境口岸代码,目的地代码,启运国家地区,疫情名称代码,疫情级别代码,是否废旧品,是否外商投资财产,检验监管模式代码,索赔金额美元,检毕日期,制造商) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15],decls[16],decls[17],decls[18],decls[19],decls[20],decls[21],decls[22],decls[23],decls[24],decls[25],decls[26],decls[27],decls[28],decls[29],decls[30],decls[31],decls[32],decls[33],decls[34],decls[35],decls[36],decls[37],decls[38],decls[39],decls[40],decls[41],decls[42],decls[43],decls[44],decls[45],decls[46],decls[47],decls[48],decls[49],decls[50],decls[51],decls[52],decls[53],decls[54],decls[55],decls[56],decls[57],decls[58],decls[59],decls[60],decls[61],decls[62],decls[63],decls[64],decls[65],decls[66],decls[67],decls[68]))
		if commitCount == 2500 :
			cur.commit()
			commitCount = 1
	cur.commit()
	conn.close()
	

#-----------------------------------
#    功能：导入出境检验不合格表数据(追加)
#-----------------------------------
def insertExitDisq():
	connectCiq2KMdb()
	data = open('出境不合格.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list:  
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化数值，转换数值内容字符串为数值
		for i_n in [29,32,36,37,40,41,61]:
			decls[i_n] = float(decls[i_n]

		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [20,21,22,23,24,25]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 出境检验不合格 (报检号,报检类别代码,报检单位注册号,发货人代码,发货人企业性质代码,发货人行政区划代码,发货人海关注册代码,运输方式代码,贸易方式代码,启运口岸代码,到达口岸代码,贸易国家地区代码,检验检疫方式代码,检验依据类别代码,是否抽检,是否二次抽检,是否隔离检疫,货物用途代码,检验检疫结果代码,查验结果代码,报检日期,收费日期,检验检疫完成日期,统计日期,上报日期,接收日期,检验检疫机构代码,HS编码,是否法检,HS数重量,HS计量单位,商品统计分类代码,商品统计数重量,商品统计单位代码,产地代码,包装种类代码,包装件数,货值美元,检验检疫项目代码,检验检疫不合格内容代码,不合格数重量,不合格金额美元,检验不合格原因代码,检验不合格处理代码,检疫不合格原因代码,检疫处理方法代码,检疫具体处理方法代码,检疫处理机构代码,检疫处理部门代码,货物名称,是否危险品,是否出境重点商品,是否出境大宗商品,是否需要出境许可证,民用品标志,监管条件,area_sta_bunc,施检科室代码,生产企业代码,检验监管模式代码,检验具体项目代码串,检验检出数量,结果描述) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15],decls[16],decls[17],decls[18],decls[19],decls[20],decls[21],decls[22],decls[23],decls[24],decls[25],decls[26],decls[27],decls[28],decls[29],decls[30],decls[31],decls[32],decls[33],decls[34],decls[35],decls[36],decls[37],decls[38],decls[39],decls[40],decls[41],decls[42],decls[43],decls[44],decls[45],decls[46],decls[47],decls[48],decls[49],decls[50],decls[51],decls[52],decls[53],decls[54],decls[55],decls[56],decls[57],decls[58],decls[59],decls[60],decls[61],decls[62]))
		if commitCount == 2500 :
			cur.commit()
			commitCount = 1
	cur.commit()
	conn.close()
	
	
#-----------------------------------
#    功能：导入入境检验不合格表数据(追加)
#-----------------------------------
def insertEntryDisq():
	connectCiq2KMdb()
	data = open('入境不合格.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list:  
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化数值，转换数值内容字符串为数值
		for i_n in [29,32,36,37,40,41,67]:
			decls[i_n] = float(decls[i_n]

		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [20,21,22,23,24,25]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 入境检验不合格 (报检号,报检类别代码,报检单位注册号,受货人代码,发货人企业性质代码,发货人行政区划代码,发货人海关注册代码,运输方式代码,贸易方式代码,启运口岸代码,经停口岸代码,贸易国家地区代码,检验检疫方式代码,检验依据类别代码,是否抽检,是否二次抽检,是否隔离检疫,货物用途代码,检验检疫结果代码,查验结果代码,报检日期,收费日期,检验检疫完成日期,统计日期,上报日期,接收日期,检验检疫机构代码,HS编码,是否法检,HS数重量,HS计量单位,商品统计分类代码,商品统计数重量,商品统计单位代码,原产国代码,包装种类代码,包装件数,货值美元,检验检疫项目代码,检验检疫不合格内容代码,不合格数重量,不合格金额美元,检验不合格原因代码,检验不合格处理代码,检疫不合格原因代码,检疫处理方法代码,检疫具体处理方法代码,检疫处理机构代码,检疫处理部门代码,货物名称,是否危险品,是否入境重点商品,是否入境大宗商品,是否需要入境许可证,民用品标志,监管条件,area_sta_bunc,施检科室代码,入境口岸代码,目的地代码,启运国家地区,疫情名称代码,疫情级别代码,是否废旧品,是否外商投资财产,检验监管模式代码,检验具体项目代码串,检验检出数量,结果描述,制造商) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15],decls[16],decls[17],decls[18],decls[19],decls[20],decls[21],decls[22],decls[23],decls[24],decls[25],decls[26],decls[27],decls[28],decls[29],decls[30],decls[31],decls[32],decls[33],decls[34],decls[35],decls[36],decls[37],decls[38],decls[39],decls[40],decls[41],decls[42],decls[43],decls[44],decls[45],decls[46],decls[47],decls[48],decls[49],decls[50],decls[51],decls[52],decls[53],decls[54],decls[55],decls[56],decls[57],decls[58],decls[59],decls[60],decls[61],decls[62],decls[63],decls[64],decls[65],decls[66],decls[67],decls[68],decls[69]))

		cur.commit()
	conn.close()
	
	
#-----------------------------------
#    功能：导入签证表数据(追加)
#-----------------------------------
def insertCert():
	connectCiq2KMdb()
	data = open('签证.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list:  
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [8,9,10,11,12,13,14]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 出境签证 (证单种类编号,报检号,扩展号,语种,报检类别代码,检验检疫机构代码,施检部门代码,签证内容代码,报检日期,收费日期,证单日期,检验检疫完成日期,统计日期,上报日期,接收日期,area_sta_bunc) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15]))
		if commitCount == 2500 :
			cur.commit()
			commitCount = 1
	cur.commit()
	conn.close()
	
	
#-----------------------------------
#    功能：导入入境疫情表数据(追加)
#-----------------------------------
def insertEntrySit():
	connectCiq2KMdb()
	data = open('入境疫情.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list:  
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化数值，转换数值内容字符串为数值
		for i_n in [29,32,36,37,61,62,63]:
			decls[i_n] = float(decls[i_n]
		
		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [20,21,22,23,24,25]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 入境疫情 (报检号,报检类别代码,报检单位注册号,受货人代码,发货人企业性质代码,发货人行政区划代码,发货人海关注册代码,运输方式代码,贸易方式代码,启运口岸代码,经停口岸代码,贸易国家地区代码,检验检疫方式代码,检验依据类别代码,是否抽检,是否二次抽检,是否隔离检疫,货物用途代码,检验检疫结果代码,查验结果代码,报检日期,收费日期,检验检疫完成日期,统计日期,上报日期,接收日期,检验检疫机构代码,HS编码,是否法检,HS数重量,HS计量单位,商品统计分类代码,商品统计数重量,商品统计单位代码,原产国代码,包装种类代码,包装件数,货值美元,检验检疫项目代码,检疫不合格原因代码,检疫处理方法代码,检疫具体处理方法代码,检疫处理机构代码,检疫处理部门代码,HS中文名称,是否危险品,是否入境重点商品,是否入境大宗商品,是否需要入境许可证,民用品标志,监管条件,area_sta_bunc,施检科室代码,入境口岸代码,目的地代码,启运国家地区,疫情名称代码,疫情级别代码,是否废旧品,是否外商投资财产,dis,货物序号,检疫不合格数量,检疫不合格金额) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15],decls[16],decls[17],decls[18],decls[19],decls[20],decls[21],decls[22],decls[23],decls[24],decls[25],decls[26],decls[27],decls[28],decls[29],decls[30],decls[31],decls[32],decls[33],decls[34],decls[35],decls[36],decls[37],decls[38],decls[39],decls[40],decls[41],decls[42],decls[43],decls[44],decls[45],decls[46],decls[47],decls[48],decls[49],decls[50],decls[51],decls[52],decls[53],decls[54],decls[55],decls[56],decls[57],decls[58],decls[59],decls[60],decls[61],decls[62],decls[63]))
		if commitCount == 2500 :
			cur.commit()
			commitCount = 1
	cur.commit()
	conn.close()
	
#-----------------------------------
#    功能：导入集装箱表数据(追加)
#-----------------------------------
def insertCon():
	connectCiq2KMdb()
	data = open('集装箱.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	commitCount = 1 # 写入计数
	for l in l_list: 
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		commitCount = commitCount+1
		# 格式化数值，转换数值内容字符串为数值
		for i_n in [2,3,4,17,18,19]:
			decls[i_n] = float(decls[i_n]
		
		# 格式化日期字符串，先转换字符串为日期，再格式转换为字符串
		from datetime import datetime
		for i_d in [22]:
			if len(decls[i_d]) > 5:
				decls[i_d] = datetime.strptime(decls[i_d],"%Y-%m-%d %H:%M:%S")
				decls[i_d] = decls[i_d].strftime("%Y-%m-%d")
			else:
				decls[i_d] = None

		print (decls)
		
		cur.execute('''INSERT INTO 集装箱 (报检号,集装箱规格代码,集装箱数量,处理集装箱数量,集装箱结果评定代码,集装箱不合格原因代码,是否预防性处理,集装箱号码串,具体处理办法代码,卫生处理机构代码,卫生处理部门代码,适载鉴定结果代码,适载鉴定项目代码,适载不合格项目代码,适载不合格处理方式代码,鉴定机构代码,鉴定部门代码,检疫不合格数量,鉴定不合格数量,综合不合格数量,是否来自疫区,出入境标志,检验检疫完成日期,检验检疫机构代码,检验检疫部门代码,area_sta_bunc) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(decls[0],decls[1],decls[2],decls[3],decls[4],decls[5],decls[6],decls[7],decls[8],decls[9],decls[10],decls[11],decls[12],decls[13],decls[14],decls[15],decls[16],decls[17],decls[18],decls[19],decls[20],decls[21],decls[22],decls[23],decls[24],decls[25]))
		if commitCount == 2500 :
			cur.commit()
			commitCount = 1
	cur.commit()
	conn.close()

#--------------------------------------
#    功能：导入出境生产企业表数据(追加)
#--------------------------------------
def insertExitProd():
	connectCiq2KMdb()
	data = open('出境生产企业.txt')
	lines = data.readlines()
	data.close()
	l_list = lines[1:] # 从第二行开始
	for l in l_list:
		decl = l.strip('\n') # 去除行末的换行
		decls = decl.split('\t')
		cur.execute('''INSERT INTO 出境生产企业 (出境生产企业代码,出境生产企业代码) VALUES(?,?)''',(decls[0],decls[1]))
	cur.commit()
	conn.close()
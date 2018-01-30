#!/usr/bin/env python3
# coding=utf-8
from PIL import Image as pimage
import wx
import os
import time
import xlwt
import re
import json
import getpass
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

class Hellowx(wx.Frame):
	def __init__(self, *args, **kw):
		super(Hellowx, self).__init__(None, -1, size=(500, 400), title="评论获取")
		self.username = 'username'
		self.password = 'username'
		self.cookies = {}
		self.token = ''
		self.login_img = 'D:/login.png'

		#Chrome启动参数
		chrome_options = Options()
		chrome_options.add_argument('--headless')
		chrome_options.add_argument('--disable-gpu')
		chrome_options.add_argument('--window-size=1920,1080')
		chrome_options.add_argument('--disable-extensions')
		#指定Chrome路径
		chrome_options.binary_location='D:/Google/Chrome/Application/chrome.exe'

		self.driver = webdriver.Chrome(chrome_options = chrome_options)
		self.login()
		self.getvalue()
		self.Show(True)

	def login(self):
		self.driver.get('https://mp.weixin.qq.com')
		time.sleep(3)

		#输入账号密码
		self.driver.find_element_by_xpath('//input[@name="account"]').send_keys(self.username)
		self.driver.find_element_by_xpath('//input[@name="password"]').send_keys(self.password)
		self.driver.find_element_by_xpath('//a[@title="点击登录"]').click()
		time.sleep(3)
		baidu = self.driver.find_element_by_xpath('//img[@class]')

		#储存二维码
		self.driver.save_screenshot(self.login_img)
		a = pimage.open(self.login_img)
		box = (800,300,1100,600)
		b = a.crop(box)
		b.save(self.login_img)

	def per_article(self, comment_id,total):
		workbook = xlwt.Workbook(encoding = 'utf-8')
		worksheet = workbook.add_sheet('sheet1')
		worksheet.write(0,0,'用户\t时间\t评论\t编号')
		row = 1
		if total%40 == 0:
			num = total//40
		else:
			num = total//40 + 1
		for j in range(num):
			article_url = 'https://mp.weixin.qq.com/misc/appmsgcomment?action=list_comment&begin=' + str(j*40) + '&count=40&comment_id=' + comment_id + '&filtertype=0&day=0&type=2&token=' + self.token + '&lang=zh_CN&f=json&ajax=1'
			res = requests.get(article_url,cookies=self.cookies)
			comments_list = json.loads(res.text)
			comments = json.loads(comments_list["comment_list"])
			for i in comments["comment"]:
				user_id = re.findall('\d{6}',i["content"])
				worksheet.write(row,0,i["nick_name"] + '\t' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(i["post_time"])) + '\t' + i["content"] + '\t' + repr(user_id))
				row += 1
		title = comments["title"]
		user = getpass.getuser()
		path = 'C:/Users/' + user + '/Desktop/' + title.replace('|','') + '.xls'
		workbook.save(path)
		dlg = wx.MessageDialog( self, "获取完毕!文件保存至桌面.", "About Sample Editor", wx.OK)
		dlg.ShowModal()
		dlg.Destroy()

	def per_page(self):
		page = (int(self.page.GetValue())-1)*10
		page_url = 'https://mp.weixin.qq.com/misc/appmsgcomment?action=get_unread_appmsg_comment&begin=' + str(page) +'&count=10&token=' + self.token +'&lang=zh_CN&f=json&ajax=1'
		articles = requests.get(page_url, cookies=self.cookies)
		articles_list = json.loads(articles.text)

		#获取文章ID和评论总数
		i = articles_list["item"][int(self.num.GetValue())-1]
		comment_id = i["comment_id"]
		total = i["total_count"]
		self.per_article(comment_id,total)

	def OnAbout(self,e):
		dlg = wx.MessageDialog( self, "1.先扫码登录\n2.填入文章对应的位置\n3.点击确定\n4.等待提示成功后即可退出", "About Sample Editor", wx.OK)
		dlg.ShowModal()
		dlg.Destroy()

	def getvalue(self):
		#定义菜单栏
		filemenu= wx.Menu()
		menuAbout = filemenu.Append(wx.ID_ABOUT, "&About"," Information about this program")
		menuExit = filemenu.Append(wx.ID_EXIT,"&Exit"," Terminate the program")
		menuBar = wx.MenuBar()
		menuBar.Append(filemenu,"&File")
		self.SetMenuBar(menuBar)
		self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
		self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
		#定义输入框及按钮位置
		self.text_page = wx.StaticText(self, -1, pos=(0,15), label='第几页:', style=wx.LEFT, size=(50,100))
		self.page = wx.TextCtrl(self, -1, pos=(50,10), style=wx.TE_CENTRE)
		self.text_num = wx.StaticText(self, -1, pos=(200,15), label='第几篇:', style=wx.ALIGN_CENTRE_HORIZONTAL, size=(50,100))
		self.num = wx.TextCtrl(self, -1, pos=(250,10), style=wx.TE_CENTRE)
		commit = wx.Button(self, -1, "确定", pos=(400,10))
		self.Bind(wx.EVT_BUTTON, self.showvalue, commit)
		self.Bind(wx.EVT_CLOSE, self.OnExit)
		#显示登录二维码
		png = wx.Image(self.login_img, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
		wx.StaticBitmap(self, -1, png, (100, 50), (png.GetWidth(), png.GetHeight()))
		os.remove(self.login_img)

	def showvalue(self, e):
		time.sleep(3)
		#获取Cookies
		cookie_list = self.driver.get_cookies()
		for i in cookie_list:
			self.cookies[i["name"]] = i["value"]
		url = self.driver.current_url
		self.token = re.findall('&token=(.*?)$',url)[0]
		self.driver.quit()	#退出Chrome
		#检查输入
		if self.page.GetValue() =="":
			dlg = wx.MessageDialog(self, "未输入页数", "", wx.OK)
			dlg.ShowModal()
		elif self.num.GetValue() =="":
			dlg = wx.MessageDialog(self, "未输入篇数", "", wx.OK)
			dlg.ShowModal()
		else:
			self.per_page()

	def OnExit(self, e):
		if self.driver:
			self.driver.quit()
		self.Destroy()

if __name__=="__main__":
	app = wx.App()
	a = Hellowx()
	app.MainLoop()

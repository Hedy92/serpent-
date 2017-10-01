# serpent-
codes that improve efficiency 

Python3
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import csv,re,time
from urllib import request,parse


	with open('/Users/apple/Downloads/keyword.csv','r',encoding='gbk') as csvfile: 
	
	#将搜索关键词按行输入入Excel，并保存为CSV文件；尝试了下encoding须用gbk才能识别
	
		fk=csv.reader(csvfile) #读取关键词
		for line in fk:
			keyword=str(*line) #须用 * 去除括号引号，提取纯字符

			urlsearch='http://www.cpppc.org:8082/efmisweb/ppp/projectLibrary/getPPPList.do?tokenid=null' #不能遗漏 ‘tokenid=null‘。。。多次失败的原因
			reqp=request.Request(urlsearch) #构造request
			reqp.add_header('User-Agent','Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36') 模拟浏览器
			for i in [1,2]: #打开以关键词搜索的页面，使之按项目阶段倒序排列，取前两页，因此显示的大部分项目处于执行阶段，资料最全
				dict={'projName':keyword,'queryPage':i,'sortby':'proj_state','orderby':'desc'}
				urlpage=parse.urlencode(dict) #采用parse来组装参数
				with request.urlopen(reqp,data=urlpage.encode('utf-8')) as fpage: #采用post的方式传递参数  
					regexp=re.compile('"PROJ_RID":"(\w+)"') #用compile来构造正则表达式，寻找project的ID
					htmlp=str(fpage.read(),'utf-8') #使被搜寻页面文本化
					time.sleep(3) #延迟三秒，尽量使文件读取完整
					listp=regexp.findall(htmlp) 

					for projid in listp: #目标:找寻项目里的文件地址
						urlpj='http://www.cpppc.org:8082/efmisweb/ppp/projectLibrary/getProjInfoNational.do?projId='+projid
						reqpj=request.Request(urlpj) 
						reqpj.add_header('User-Agent','Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36')
						regexpdf=re.compile('/ppp/projManage/perview\.do\?fileName=\w+\.pdf&ftpFileName=\d+\.pdf')
						htmlpdf=str(reqpj.read(),'utf-8')
						time.sleep(3)
						listpdf=regexpdf.findall(htmlpdf)			
					
						for urld in listpdf: #目标:下载每个项目里的文件，因无下载选项，采用打印时另存为PDF
							url='http://www.cpppc.org:8082/efmisweb'+urld+'&content=efmisweb&xsg=:8083/'
							browser=webdriver.Chrome() #启动chrome
							browser.get(url)
							actions=ActionChains(browser)
							iframe=browser.find_element_by_xpath('/html/body/iframe') #先定位到iframe，否则find无法识别
							browser.switch_to_frame(iframe)
							browser.find_element_by_id("zoomIn").click() #！！！放大或者缩小一下，才能使文件缓存完整，否则只能打印保存第一页！！！
							browser.switch_to_default_content()
							actions=ActionChains(browser)
							actions.send_keys(Keys.COMMAND,'p').perform() 
							#尝试了wkhtmltopdf，无效；转用webdriver，但是不能控制打印windows,所以设置成跳出打印框，可根据需要决定下载
    		fk.close()
              
              

import win32com.client, MSO, MSPPT
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)
import pandas as pd
import numpy as np 

import os 
import urllib
import urllib.request
from bs4 import BeautifulSoup as bs
import time
import csv
# Application = win32com.client.Dispatch("PowerPoint.Application")
# Application.Visible = True
# Presentation = Application.Presentations.Add()
# Base  = Presentation.Slides.Add(1, ppLayoutBlank)

# button_alpha = Base.Shapes.AddShape(msoShapeRectangle, 400, 10, 150, 40)
# button_alpha.TextFrame.TextRange.Text = 'Alphabetical'

# button_rating = Base.Shapes.AddShape(msoShapeRectangle, 560, 10, 150, 40)
# button_rating.TextFrame.TextRange.Text = 'By rating'



def make_soup(url):
	thepage = urllib.request.urlopen(url)
	soup = bs(thepage,"html.parser")
	
	title =soup.find_all("td",{"class":"titleColumn"})
	string_list = []
	for y in title:
	 	string_list.append(str(y.find_all("a")).split('">')[1].split("</a>")[0].replace(" ","_"))
	
	# images = soup.find_all("td",{"class":"posterColumn"})
	# image_list = []
	# for image in images:
	# 	image_list.append(str(image.find_all("a")).split('src="')[1].split('"')[0])


	# i = 0 
	# for img in image_list:
	# 	i +=1
	# 	imageFile = open("D:\\treemap\\imdb\\"+str(i)+".jpeg",'wb')
	# 	imageFile.write(urllib.request.urlopen(img).read())
	# 	imageFile.close()
	
		
	# movie = soup.find("table",{"class":'lister-item-header'})
	
	# print(soupdata)
	# return soupdata
	return string_list

movies_name = make_soup("http://www.imdb.com/chart/top")

# df = pd.DataFrame(movies_name)
# df.to_csv("name.csv",encoding='utf-8')



width ,height = 28,41

for i,movies in enumerate(movies_name):
	print(i)
	# x = 10 + width * (i % 25)
	# y = 100 + height *(i // 25)
	# r = Base.Shapes.AddShape(msoShapeRectangle, x, y, width, height)




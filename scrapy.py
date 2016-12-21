import requests
import urllib.request
from bs4 import BeautifulSoup
from html.parser import HTMLParser

def crawlBase(link):
	with urllib.request.urlopen(link) as urlToCrawl:
			htmltext = urlToCrawl.read()
	soup = BeautifulSoup(htmltext) 
	# page  = requests.get(link)
	# soup = BeautifulSoup(page.text) 
	return soup

soup = crawlBase('http://www.imdb.com/chart/top')
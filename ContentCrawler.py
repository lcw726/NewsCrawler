import logging
import requests
import configparser
from bs4 import BeautifulSoup
from newsplease import NewsPlease
from Article import Article

config = configparser.ConfigParser()
config.read('app.ini')

#proxies = config._sections['Proxy']

def basicContentCrawler(article, contentType, contentClass, contentToken):
  try:
    text = ''
    dom = requests.get(article.url,  verify = False).text
    soup = BeautifulSoup(dom, 'lxml')
    article_body = soup.find(contentType, contentClass)
    #logging.warning(soup)

    for ele in article_body.find_all(contentToken):
      text += ele.text + '\n'

    article.text = text.strip()
    if article_body.find('img'):
      article.img_url = article_body.find('img').get('src')

      if article.img_url.startswith('//'):
        article.img_url = 'https:' + article.img_url
      elif article.img_url.startswith('http') == False: 
        article.img_url = article.url[:article.url.find('/', 8)] + article.img_url
      #logging.warning(article.img_url)
  except Exception as e:
    logging.warning(article.url + '  Basic Content Error:' + str(e))

def newsPleaseContentCrawler(article):
  dom = requests.get(article.url,  verify = False).text
  nparticle = NewsPlease.from_html(dom)
  article.text = nparticle.text.strip()
  article.img_url = nparticle.image_url

def publishInfoCrawler(article, contentType, contentClass, contentToken):
  try:
    dom = requests.get(article.url,  verify = False).text
    soup = BeautifulSoup(dom, 'lxml')
    #logging.warning(article.url)
    info = soup.find_all(contentType, contentClass)

    #for ele in soup.find_all(contentType, contentClass):
      #ele.find(contentToken)

    if(len(info)>1):
      article.author = info[0].text.strip()
      article.date = info[1].text.strip()
    else:
      article.date = info[0].text.strip()
  except Exception as e:
    logging.warning(e)

def directInfoCrawler(article, dateType, dateClass, authorType, authorClass):
  try:
    dom = requests.get(article.url,  verify = False).text
    soup = BeautifulSoup(dom, 'lxml')
    #logging.warning(article.url)
    if dateType != '':
      if soup.find(dateType,dateClass):
        article.date = soup.find(dateType,dateClass).text.strip()
        #logging.warning(article.date)

    if authorType != '':
      if soup.find(authorType,authorClass):
        article.author = soup.find(authorType,authorClass).text.strip()
        #logging.warning(article.author)
  except Exception as e:
    logging.warning('directInfoCrawler Error:' + str(e))

def listInfoCrawler(article, infoType, infoClass, childType, childClass):
  try:
    dom = requests.get(article.url,  verify = False).text
    soup = BeautifulSoup(dom, 'lxml')

    if childClass == '':
      info = soup.find(infoType, infoClass).find_all(childType)
      article.author = info[0].text.strip()
      article.date = info[1].text.strip()
      #logging.warning(article.author)
      #logging.warning(article.date)
    else:
      info = soup.find(infoType, infoClass).find_all(childType, childClass)
      article.author = info[0].text.strip()
      article.date = info[1].text.strip()
      #logging.warning(article.author)
      #logging.warning(article.date)
  except Exception as e:
    logging.warning(e)
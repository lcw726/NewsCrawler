class Article():
  def __init__(self, title , url, source, text, date, author, img_url, img_name = ''):
    self.title = title.strip()
    self.text = text.strip()
    if url.startswith('//'):
      url = 'https:' + url
    self.url = url
    self.source = source
    self.date = date.strip()
    self.author = author.strip()
    self.img_url = img_url
    self.img_name = img_name

  @classmethod
  def from_list(cls, title, url, source) -> 'Article':
    return cls(title=title.strip(), url = url.strip(), source = source.strip(), text='', date='', author='', img_url='')

  @classmethod
  def from_digiTimes(cls, title, url, source, text, date, img_url) -> 'Article':
    return cls(title=title.strip(), url = url.strip(), source = source.strip(), text=text.strip(), date=date.strip(), author='', img_url=img_url.strip())
  
  @classmethod
  def from_external(cls, url, title, source, author, text) -> 'Article':
    return cls(title=title.strip(), url = url.strip(), source = source.strip(), text=text.strip(), date='', author=author.strip(), img_url='')

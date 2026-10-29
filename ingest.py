"""Загрузка расписания ПКПС: выполняется в GitHub Actions, без токена Telegram."""
import os, re, json, hashlib, urllib.request, time
from datetime import datetime, timezone
from html import unescape
from html.parser import HTMLParser
import xlrd
PAGE='https://pkps-perm.ru/students/raspisanie/fakultet_dizayna_i_servisa_ul_chernyshevskogo_11/'
DAYS=['понедельник','вторник','среда','четверг','пятница','суббота','воскресенье']
GROUP=re.compile(r'^[А-ЯЁA-Z]+-\d{2}-\d{1,2}[КK]?(?:-\d+)?$',re.I)
def key(v):return re.sub(r'\s+','',str(v).upper().replace('–','-').replace('—','-'))
def parse(data):
 w=xlrd.open_workbook(file_contents=data,formatting_info=True);out={}
 for s in w.sheets():
  def cell(r,c):
   v=str(s.cell_value(r,c)).strip()
   if v:return v
   for a,b,x,y in s.merged_cells:
    if a<=r<b and x<=c<y:return str(s.cell_value(a,x)).strip()
   return ''
  def groups(r,c):return [key(v) for v in cell(r,c).splitlines() if GROUP.fullmatch(key(v))]
  for r in range(min(s.nrows,26)):
   for c in range(2,s.ncols):
    if not str(s.cell_value(r,c)).strip():continue
    gs=groups(r,c)
    if not gs:continue
    course=0
    for up in range(r-1,max(-1,r-5),-1):
     m=re.search(r'\b(IV|III|II|I|[1-4])\s*курс',cell(up,c),re.I)
     if m:
      course={'I':1,'II':2,'III':3,'IV':4}.get(m[1].upper(),int(m[1]) if m[1].isdigit() else 0);break
    end=next((y-1 for a,b,x,y in s.merged_cells if a==r and x==c),c)
    lessons={};day=''
    for row in range(r+1,s.nrows):
     if groups(row,c):break
     dv=cell(row,0).lower()
     if dv in DAYS:day=dv
     sl=re.match(r'^\s*(VIII|VII|VI|Vl|IV|III|II|V|I|[1-8])\b',cell(row,1),re.I)
     if not sl or not day:continue
     raw=sl[1].upper();num={'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VL':6,'VII':7,'VIII':8}.get(raw,int(raw) if raw.isdigit() else 0)
     t=re.search(r'(\d{1,2})[.:](\d{2})',cell(row,1));clock=f'{int(t[1]):02}:{t[2]}' if t else ''
     entries=[]
     for col in range(c,end+1):
      text=cell(row,col)
      if text and not any(x['text']==text for x in entries):entries.append({'day':day,'number':num,'time':clock,'text':text,'subgroup':col-c+1})
     for entry in entries:
      if len(entries)==1:entry['subgroup']=None
      lessons.setdefault(day,[]).append(entry)
    for g in gs:out[g]={'group':g,'course':course,'days':lessons,'present':True}
 if not out:raise ValueError('Группы не распознаны')
 return out
class Links(HTMLParser):
 def __init__(self):super().__init__();self.current=None;self.links=[]
 def handle_starttag(self,tag,attrs):
  if tag=='a':self.current=[dict(attrs).get('href',''),'']
 def handle_data(self,data):
  if self.current is not None:self.current[1]+=data
 def handle_endtag(self,tag):
  if tag=='a' and self.current is not None:self.links.append(self.current);self.current=None

def request(url,body=None,secret=None):
 headers={'User-Agent':'PKPS-Schedule/2.0'}
 if body is not None:headers['Content-Type']='application/json'
 if secret:headers['X-Import-Secret']=secret
 req=urllib.request.Request(url,data=json.dumps(body,ensure_ascii=False).encode() if body is not None else None,headers=headers)
 for attempt in range(3):
  try:
   with urllib.request.urlopen(req,timeout=45) as r:return r.read()
  except Exception:
   if attempt==2:raise
   time.sleep(2*(attempt+1))
def main():
 from urllib.parse import urljoin,urlsplit
 worker=os.environ['WORKER_URL'].rstrip('/');secret=os.environ['IMPORT_SECRET']
 if not worker.startswith('https://'):raise ValueError('WORKER_URL должен начинаться с https://')
 manifest=json.loads(request(worker+'/manifest',{},secret));p=Links();p.feed(request(PAGE).decode('utf-8'))
 docs={}
 for href,title in p.links:
  url=urljoin(PAGE,unescape(href));title=' '.join(title.split());date=re.search(r'(\d{2})\.(\d{2})\.(\d{4})',title)
  if urlsplit(url).netloc!=urlsplit(PAGE).netloc or not re.search(r'\.xlsx?$',urlsplit(url).path,re.I):continue
  if re.search(r'полугод|учебных занятий',title,re.I):doc_key='base'
  elif date and re.search(r'изменени|расписан',title,re.I):doc_key=f'{date[3]}-{date[2]}-{date[1]}'
  else:continue
  docs.setdefault(doc_key,{'key':doc_key,'url':url,'title':title})
 if 'base' not in docs:raise ValueError('Основное расписание не найдено')
 selected=['base']+sorted([k for k in docs if k!='base'],reverse=True)[:10]
 for k in selected:
  doc=docs[k];raw=request(doc['url']);digest=hashlib.sha256(raw).hexdigest()
  if manifest.get(k)==digest:print(k,'без изменений');continue
  groups=parse(raw)
  # Сначала сохраняем данные; уведомления выполняются отдельно после импорта.
  response=json.loads(request(worker+'/import',{'key':k,'hash':digest,'data':{'groups':groups,'url':doc['url'],'title':doc['title']}},secret))
  if not response.get('ok'):raise ValueError('Ошибка импорта')
  print(k,'импортировано групп:',len(groups))
 request(worker+'/complete',{'dates':selected[1:]},secret)
 print('Обновление завершено')
if __name__=='__main__':main()

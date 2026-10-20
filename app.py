from http.server import BaseHTTPRequestHandler,HTTPServer
from urllib.parse import parse_qs
from docx import Document
from pathlib import Path
from datetime import date
import json,re
ROOT=Path(__file__).resolve().parent
HTML=Path(__file__).with_name('index.html').read_text(encoding='utf-8')
def put(p,v):
    if p.runs:
        p.runs[0].text=v
        for r in p.runs[1:]: r.text=''
    else: p.add_run(v)
def main():
 class H(BaseHTTPRequestHandler):
  def reply(self,status,data,typ='text/html; charset=utf-8',extra=None):
   if isinstance(data,str): data=data.encode()
   self.send_response(status);self.send_header('Content-Type',typ);self.send_header('Content-Length',str(len(data)))
   for k,v in (extra or {}).items(): self.send_header(k,v)
   self.end_headers();self.wfile.write(data)
  def do_GET(self): self.reply(200,HTML) if self.path=='/' else self.reply(404,'Not found')
  def do_POST(self):
   try:
    form=parse_qs(self.rfile.read(int(self.headers.get('Content-Length',0))).decode());g=lambda k,d='':form.get(k,[d])[0].strip()
    sessions=json.loads(g('sessions','[]'));name=g('name');price=float(g('price','46').replace(',','.'))
    if not name or not sessions: return self.reply(400,'Name und Termine erforderlich.','text/plain')
    dates=', '.join(date.fromisoformat(x).strftime('%d.%m.%Y') for x in sessions)
    doc=Document(ROOT/'Vorlage.docx')
    paras=list(doc.paragraphs)
    for table in doc.tables:
     for row in table.rows:
      for cell in row.cells: paras.extend(cell.paragraphs)
    # Placeholder replacement keeps the supplied document structure and styles.
    mapping={'(Anrede)':g('salutation'),'      (Name)':name,'(Name)':name,'(Adresse)':g('street'),'(Postleitzahl, Ort)':g('zipcity'),'(Honorarnote)':g('invoice_no'),'(Rechnung Datum)':date.fromisoformat(g('invoice_date')).strftime('%d.%m.%Y') if g('invoice_date') else date.today().strftime('%d.%m.%Y'),'(Termin Datums)':dates,'(Versicherungsnummer)':g('insurance')}
    for p in paras:
     old=''.join(r.text for r in p.runs);new=old
     for a,b in mapping.items():new=new.replace(a,b)
     if new!=old:put(p,new)
    t=doc.tables[0]; vals=[('10x' if g('plan')=='10' else '6x'),'Physiotherapie á 45min',f'{price:.2f}'.replace('.',','),f'{price*len(sessions):.2f}'.replace('.',',')]
    for i,v in enumerate(vals):put(t.rows[1].cells[i].paragraphs[0],v)
    put(t.rows[2].cells[-1].paragraphs[0],f'{price*len(sessions):.2f}'.replace('.',','))
    # Fill the patient and insurance text in the first header row, preserving its table cell.
    for p in t.rows[0].cells[0].paragraphs:
     text=''.join(r.text for r in p.runs)
     if '(Name)' in text:
      put(p,'                Patient\n     ('+name+')')
    for p in t.rows[0].cells[-1].paragraphs:
     text=''.join(r.text for r in p.runs)
     if '(Versicherungsnummer)' in text: put(p,'Versicherungsnummer\n('+g('insurance')+')')
    dest=ROOT/'Rechnung_aus_Vorlage.docx';doc.save(dest)
    self.reply(200,dest.read_bytes(),'application/vnd.openxmlformats-officedocument.wordprocessingml.document',{'Content-Disposition':'attachment; filename="Honorarnote.docx"'})
   except Exception as e:self.reply(500,str(e),'text/plain; charset=utf-8')
  def log_message(self,*a):pass
 print('Öffnen: http://127.0.0.1:8765');HTTPServer(('127.0.0.1',8765),H).serve_forever()
if __name__=='__main__':main()

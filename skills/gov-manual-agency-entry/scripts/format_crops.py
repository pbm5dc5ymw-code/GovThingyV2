"""Format review: crop the best printed occurrence of each new/changed row so its TYPOGRAPHY can be checked.

Usage: python format_crops.py <year> <pdf in ~/Downloads> <first body pdf page> <chart-list pdf page> [extra rows]
Needs <scratch>/changed.json ({year: [rows]}) and <scratch>/alias.json ({year: {row: [printed name]}}).
Writes contact sheets (8 crops each). Read every sheet: a unit counts only if it is SET OFF on the page -
its own centered/bold heading, a small-caps run-in ("OFFICE OF X.—This office ..."), or an indented entry in
a printed unit list. Plain sentences ("The X Staff discharges ..."), officials lines and chart boxes do not.
"""
import pymupdf,os,re,json,sys,openpyxl
from PIL import Image, ImageDraw
S='/private/tmp/claude-502/-Users-leo-Documents-GitHub-GovThingyV2/5958fe8f-7909-457c-b84d-a17154b8bd5d/scratchpad'
year=int(sys.argv[1]); pdf=sys.argv[2]; first,body_end=int(sys.argv[3]),int(sys.argv[4])
rows=json.load(open(S+'/changed.json'))[str(year)]+[int(x) for x in sys.argv[5].split(',') if x]
ALIAS=json.load(open(S+'/alias.json')).get(str(year),{})
ws=openpyxl.load_workbook('data/Spreadsheet 1935-1950.xlsx').worksheets[0]
d=pymupdf.open(os.path.expanduser('~/Downloads/'+pdf))
pages=list(range(first-1,body_end-1))
low={i:re.sub(r'\s+',' ',re.sub(r'-\n','',d[i].get_text())).lower() for i in pages}
OFF=re.compile(r'(director|chief|commissioner|administrator|chairman|president|secretary|head|officer|member|manager|superintendent|in charge|archivist|librarian)[ ,.]*(of\s+(the\s+)?)?$')
def norm(s): return re.sub(r'[^a-z ]','',s.lower())
crops=[]
for r in rows:
    name=str(ws.cell(r,1).value).strip()
    terms=ALIAS.get(str(r)) or [re.sub(r'\s*\(.*?\)','',name)]
    cands=[]
    for term in terms:
        t=term.lower(); words=norm(term).split()
        tail=' '.join(words[-2:]) if len(words)>=2 else words[0]
        for i in pages:
            L=low[i]
            for m in re.finditer(re.escape(t),L):
                before=L[max(0,m.start()-40):m.start()]; after=L[m.end():m.end()+4]
                if OFF.search(before.strip()) or re.search(r'_{2,}|\.{4,}|-{3,}',L[m.end():m.end()+12]): sc=0
                elif re.match(r'\s?\.?\s?[—-]',after): sc=3        # run-in heading
                elif re.match(r'\s?(the|this) ',L[m.end():m.end()+6]) : sc=2   # heading line followed by text
                else: sc=1
                cands.append((sc,i,t))
            # OCR-garbled small caps run-ins: look for "<tail>.—"
            for m in re.finditer(re.escape(tail)+r'\s?\.?\s?—',L):
                cands.append((3,i,tail))
    if not cands: crops.append((r,name,None,None,-1)); continue
    sc,i,t=max(cands,key=lambda c:(c[0],-c[1]))
    pg=d[i]; hs=pg.search_for(t) or pg.search_for(' '.join(t.split()[-2:]))
    if hs:
        h=hs[0]; clip=pymupdf.Rect(pg.rect.x0+15,max(0,h.y0-30),pg.rect.x1-15,min(pg.rect.y1,h.y1+42))
    else:
        clip=pg.rect
    fn=f'{S}/fmt/{year}_{r}.png'; pg.get_pixmap(dpi=100,clip=clip).save(fn)
    crops.append((r,name,i+1,fn,sc))
json.dump(crops,open(f'{S}/crops{year}.json','w'))
items=[c for c in crops if c[3]]; per=8; n=0
for k in range(0,len(items),per):
    ims=[Image.open(c[3]) for c in items[k:k+per]]
    w=max(im.width for im in ims); h=sum(min(im.height,260)+22 for im in ims)
    sh=Image.new('RGB',(w,h),'white'); y=0; dr=ImageDraw.Draw(sh)
    for c,im in zip(items[k:k+per],ims):
        im=im.crop((0,0,im.width,min(im.height,260)))
        dr.rectangle([0,y,w,y+20],fill=(255,230,150)); dr.text((4,y+4),f'ROW {c[0]}  {c[1][:70]}  (pdf p{c[2]}) score {c[4]}',fill='black')
        sh.paste(im,(0,y+22)); y+=im.height+22
    sh.save(f'{S}/fmt/sheet{year}_{k//per:02d}.png'); n+=1
print(len(items),'crops',n,'sheets; no hit:',[(c[0],c[1]) for c in crops if not c[3]])

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้าง PowerPoint นำเสนอ "Source Code / Kiro Project" — LIKA Nexus
อธิบาย: แนวทางการพัฒนา · โครงสร้างผลงาน · การประยุกต์ใช้ Kiro
ภาพประกอบ vector · ภาษาทางการ · ไม่ใส่คะแนน
"""
import zipfile, html, math, random

OUT="LIKA_Source_Code.pptx"
EMU_W,EMU_H=12192000,6858000
NAVY="002366"; NAVY2="060a16"; NAVY3="0b1a3a"; GOLD="f5c542"; CYAN="22d3ee"
WHITE="e5edff"; GREY="9fb3d1"; EMER="34d399"; VIOLET="a78bfa"; ORANGE="fb923c"; PINK="f472b6"

def esc(s): return html.escape(str(s),quote=True)
_id=[10]
def nid(): _id[0]+=1; return _id[0]

def _txbody(runs,anchor,align,wrap=True):
    paras=""
    for grp in runs:
        r=""
        for (t,sz,c,b,i) in grp:
            r+=(f'<a:r><a:rPr lang="th-TH" sz="{int(sz*100)}" b="{1 if b else 0}" i="{1 if i else 0}" dirty="0">'
                f'<a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
                f'<a:latin typeface="Tahoma"/><a:cs typeface="Tahoma"/></a:rPr><a:t>{esc(t)}</a:t></a:r>')
        if not grp: r='<a:endParaRPr lang="th-TH"/>'
        paras+=f'<a:p><a:pPr algn="{align}"/>{r}</a:p>'
    return (f'<p:txBody><a:bodyPr wrap="{"square" if wrap else "none"}" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{paras}</p:txBody>')

def tbox(x,y,w,h,runs,anchor="t",align="l",wrap=True):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>{_txbody(runs,anchor,align,wrap)}</p:sp>')

def shape(prst,x,y,w,h,fill,line=None,lw=12700,adj="",grad=None,shadow=False):
    if grad:
        fillxml=(f'<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="{grad[0]}"/></a:gs>'
                 f'<a:gs pos="100000"><a:srgbClr val="{grad[1]}"/></a:gs></a:gsLst>'
                 f'<a:lin ang="5400000" scaled="1"/></a:gradFill>')
    elif fill is None: fillxml='<a:noFill/>'
    else: fillxml=f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    ln=(f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else '<a:ln><a:noFill/></a:ln>')
    sh=('<a:effectLst><a:outerShdw blurRad="90000" dist="40000" dir="5400000" rotWithShape="0">'
        '<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>') if shadow else ''
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="{prst}">{adj or "<a:avLst/>"}</a:prstGeom>{fillxml}{ln}{sh}</p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')

def rrect(x,y,w,h,fill,line=None,lw=10000,r=9000,grad=None,shadow=False):
    return shape("roundRect",x,y,w,h,fill,line,lw,f'<a:avLst><a:gd name="adj" fmla="val {r}"/></a:avLst>',grad,shadow)
def oval(x,y,w,h,fill,line=None,lw=10000,grad=None,shadow=False):
    return shape("ellipse",x,y,w,h,fill,line,lw,"",grad,shadow)
def lineshape(x1,y1,x2,y2,color,w=9525):
    xx,yy=min(x1,x2),min(y1,y2); cx,cy=abs(x2-x1),abs(y2-y1)
    fH='1' if x2<x1 else '0'; fV='1' if y2<y1 else '0'
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{nid()}" name="l"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f'<p:spPr><a:xfrm flipH="{fH}" flipV="{fV}"><a:off x="{xx}" y="{yy}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{w}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:cxnSp>')

def icon_badge(x,y,size,emoji,bg=NAVY,ring=CYAN):
    return (oval(x,y,size,size,bg,line=ring,lw=19050,shadow=True)+
            tbox(x,y,size,size,[[(emoji,int(size/914400*20),GOLD,False,False)]],anchor="ctr",align="ctr"))

def nexus_logo(x,y,s):
    cx,cy=x+s//2,y+s//2; r=int(s*0.34)
    out=[rrect(x,y,s,s,None,grad=(NAVY3,NAVY),r=24000,line=CYAN,lw=12700)]
    pts=[]
    for a in range(0,360,60):
        px=cx+int(r*math.cos(math.radians(a))); py=cy+int(r*math.sin(math.radians(a)))
        pts.append((px,py)); out.append(lineshape(cx,cy,px,py,CYAN,7000))
    for (px,py) in pts: out.append(oval(px-38000,py-38000,76000,76000,"7dd3fc"))
    out.append(oval(cx-70000,cy-70000,140000,140000,GOLD,shadow=True))
    return "".join(out)

def chip(x,y,w,text,color):
    return (rrect(x,y,w,230000,None,line=color,lw=9525,r=50000)+
            tbox(x,y-6000,w,250000,[[(text,9,color,True,False)]],anchor="ctr",align="ctr"))

def slide(shapes,bg=NAVY2):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            f'<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
            f'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'{"".join(shapes)}</p:spTree></p:cSld>'
            f'<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
            f'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
            f'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>')

def starfield(n=36,seed=7):
    random.seed(seed); s=[]
    for _ in range(n):
        x=random.randint(0,EMU_W); y=random.randint(0,EMU_H); r=random.randint(12000,40000)
        s.append(oval(x,y,r,r,GREY))
    return s

def header(kicker,title,page,total=7,accent=GOLD):
    return [rrect(457200,470000,150000,920000,accent,r=50000),
            tbox(760000,450000,10000000,380000,[[(kicker,13,CYAN,True,False)]]),
            tbox(760000,790000,10600000,900000,[[(title,26,WHITE,True,False)]]),
            tbox(11000000,6360000,1000000,330000,[[(f"{page}/{total}",10,GREY,False,False)]],align="r")]

def bullets(x,y,w,items,size=13,gap=540000,color=WHITE,mark="▸",mc=GOLD):
    return [tbox(x,y+i*gap,w,gap,[[(f"{mark} ",12,mc,True,False),(it,size,color,False,False)]]) for i,it in enumerate(items)]

slides=[]

# ====== 1 · TITLE ======
sp=starfield(44)
sp+=[rrect(0,2350000,EMU_W,2000000,None,grad=(NAVY,NAVY3))]
sp+=[nexus_logo(820000,2620000,1420000)]
sp+=[tbox(2550000,2560000,9000000,560000,[[("Source Code / Kiro Project",15,GOLD,True,False)]]),
     tbox(2550000,2980000,9000000,760000,[[("LIKA Nexus",40,WHITE,True,False)]]),
     tbox(2550000,3780000,9000000,480000,[[("แนวทางการพัฒนา · โครงสร้างผลงาน · การประยุกต์ใช้ Kiro",13,CYAN,False,False)]])]
cx=820000
for txt,col in [("Single-file Web App",GOLD),("Spec-driven",CYAN),("Open Source",EMER),("21 Commits",VIOLET)]:
    w=len(txt)*82000+280000; sp.append(chip(cx,4650000,w,txt,col)); cx+=w+150000
sp.append(tbox(820000,5800000,10600000,340000,[[("GitHub: ",10,GOLD,True,False),("github.com/kanchanathsao-crypto/lika-lit-review-helper-",10,WHITE,False,False)]]))
slides.append(slide(sp))

# ====== 2 · โครงสร้างผลงาน (file tree) ======
sp=header("PROJECT STRUCTURE","โครงสร้างของผลงาน",2,7)
sp.append(rrect(760000,1950000,5600000,3900000,"0a1428",line=CYAN,lw=9525,r=5000,shadow=True))
tree=[("📁 LIKA.kiro/",GOLD,0,""),
      ("⭐ lika-nexus.html",WHITE,1,"แอปหลัก (single-file)"),
      ("📁 .kiro/",CYAN,1,"Kiro Project"),
      ("📐 steering/rules.md",EMER,2,"กฎการพัฒนา"),
      ("📐 specs/lit-review-spec.md",EMER,2,"ข้อกำหนด"),
      ("📖 README.md",WHITE,1,"ภาพรวมโครงการ"),
      ("📖 USER_GUIDE.md",GREY,1,"คู่มือผู้ใช้"),
      ("📄 CONCEPT_PAPER.md",GREY,1,"เอกสารแนวคิด"),
      ("📊 LIKA_*.pptx",GOLD,1,"สไลด์นำเสนอ"),
      ("🛠 build_*.py",GREY,1,"สคริปต์สร้างเอกสาร")]
for i,(t,c,ind,note) in enumerate(tree):
    sp.append(tbox(1000000+ind*330000,2140000+i*350000,4400000,340000,[[(t,11.5,c,ind==0,False)]]))
# คำอธิบายฝั่งขวา
sp.append(tbox(6560000,2000000,4900000,360000,[[("จัดระเบียบชัดเจน",13,GOLD,True,False)]]))
sp+=bullets(6560000,2440000,4900000,[
    "โค้ดระบบรวมในไฟล์เดียว ติดตั้งง่าย",
    "โฟลเดอร์ .kiro แยกหลักฐานการพัฒนา",
    "เอกสารครบ: README · คู่มือ · แนวคิด",
    "สคริปต์สร้างสไลด์ตรวจสอบย้อนกลับได้",
], size=12, gap=620000)
slides.append(slide(sp))

# ====== 3 · แนวทางการพัฒนา + tech stack ======
sp=header("DEVELOPMENT APPROACH","แนวทางการพัฒนา",3,7)
sp.append(rrect(760000,1950000,5250000,3850000,NAVY3,line=CYAN,lw=12700,shadow=True))
sp.append(icon_badge(760000+2150000,2130000,900000,"🧩",bg=NAVY,ring=CYAN))
sp.append(tbox(960000,3080000,4900000,400000,[[("หลักการออกแบบ",14,CYAN,True,False)]],align="ctr"))
sp+=bullets(1000000,3520000,4800000,[
    "Single-file Web App — ไม่ต้องมีเซิร์ฟเวอร์",
    "ประมวลผลฝั่งผู้ใช้ (client-side) ทั้งหมด",
    "Privacy-first — ไม่ส่งไฟล์ออกภายนอก",
    "เผยแพร่ผ่าน GitHub Pages",
], size=11.5, gap=560000)
sp.append(rrect(6260000,1950000,5200000,3850000,NAVY3,line=GOLD,lw=12700,shadow=True))
sp.append(tbox(6460000,2120000,4800000,400000,[[("เทคโนโลยีที่ใช้ (Tech Stack)",14,GOLD,True,False)]]))
tech=[("HTML5 + Tailwind CSS","โครงสร้าง + UI (Glassmorphism)",CYAN),
      ("JavaScript (ES6)","กลไกสกัดข้อมูล + ค้นหา",GOLD),
      ("Three.js r128 + OrbitControls","กราฟความรู้ 3 มิติ",EMER),
      ("PDF.js","อ่านไฟล์ PDF ในเบราว์เซอร์",VIOLET),
      ("GitHub Pages","เผยแพร่ใช้งานจริง",ORANGE)]
for i,(t,d,c) in enumerate(tech):
    y=2560000+i*620000
    sp.append(rrect(6460000,y,4800000,540000,"0a1428",line=c,lw=7000,r=10000))
    sp.append(tbox(6640000,y+70000,4500000,300000,[[(t,11.5,c,True,False)]]))
    sp.append(tbox(6640000,y+340000,4500000,220000,[[(d,9.5,WHITE,False,False)]]))
slides.append(slide(sp))

# ====== 4 · กลไก AI engine ======
sp=header("AI ENGINE","กลไกการสกัดและวิเคราะห์ข้อมูล",4)
steps=[("📥","รับเอกสาร","อ่าน PDF/TXT ด้วย PDF.js",GOLD),
       ("🔍","สกัด IMRaD","ดึง 5 ประเด็นสำคัญ",CYAN),
       ("🏷️","จำแนก","22 สาขา + ระเบียบวิธี\n(weighted scoring)",EMER),
       ("🧾","สังเคราะห์","ร่าง Lit-Review + APA",VIOLET)]
x0=760000
for i,(ic,t,d,c) in enumerate(steps):
    cx=x0+i*2720000
    sp.append(rrect(cx,2100000,2550000,2550000,NAVY3,line=c,lw=9525,r=14000,shadow=True))
    sp.append(icon_badge(cx+2550000//2-430000,2300000,860000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+100000,3250000,2350000,360000,[[(t,13,c,True,False)]],align="ctr"))
    for j,ln in enumerate(d.split("\n")):
        sp.append(tbox(cx+120000,3620000+j*360000,2310000,350000,[[(ln,10,WHITE,False,False)]],align="ctr"))
    if i<len(steps)-1:
        sp.append(tbox(cx+2550000-30000,3100000,250000,400000,[[("→",18,GREY,True,False)]],align="ctr"))
sp.append(rrect(760000,4900000,10700000,780000,None,line=CYAN,lw=9525,r=14000))
sp.append(tbox(1000000,5000000,10300000,580000,[[("รองรับทั้งภาษาไทยและอังกฤษ · แปลงปี พ.ศ. เป็น ค.ศ. · ดึงชื่อผู้แต่งแบบเต็มโดยหลีกเลี่ยงชื่อสถาบัน",12,WHITE,False,False)]]))
slides.append(slide(sp))

# ====== 5 · การใช้ Kiro (Spec-driven) ======
sp=header("KIRO PROJECT","การประยุกต์ใช้ Kiro แบบ Spec-driven",5)
steps=[("📐","Steering Rules",".kiro/steering/\nrules.md",GOLD),
       ("📋","Specs",".kiro/specs/\nlit-review-spec.md",CYAN),
       ("🤝","Agent พัฒนา","Kiro เขียนโค้ด\n+ ทดสอบ",EMER),
       ("🚀","Deploy","GitHub Pages\nใช้งานได้จริง",VIOLET)]
x0=760000
for i,(ic,t,d,c) in enumerate(steps):
    cx=x0+i*2720000
    sp.append(rrect(cx,2050000,2550000,2450000,NAVY3,line=c,lw=9525,r=14000,shadow=True))
    sp.append(icon_badge(cx+2550000//2-430000,2230000,860000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+100000,3160000,2350000,360000,[[(t,13,c,True,False)]],align="ctr"))
    for j,ln in enumerate(d.split("\n")):
        sp.append(tbox(cx+120000,3520000+j*350000,2310000,340000,[[(ln,9.5,WHITE,False,False)]],align="ctr"))
    if i<len(steps)-1:
        sp.append(tbox(cx+2550000-30000,3050000,250000,400000,[[("→",18,GREY,True,False)]],align="ctr"))
sp.append(rrect(760000,4750000,10700000,920000,None,line=GOLD,lw=9525,r=14000))
sp.append(tbox(1000000,4850000,10300000,360000,[[("Kiro เป็นเครื่องมือหลักตลอดวงจรพัฒนา",12,GOLD,True,False)]]))
sp.append(tbox(1000000,5200000,10300000,420000,[[("กำหนดกฎ → ออกแบบข้อกำหนด → เขียนและทดสอบโค้ด → เผยแพร่ ตรวจสอบย้อนกลับได้จากประวัติการพัฒนา (commit history) บน GitHub",11,WHITE,False,False)]]))
slides.append(slide(sp))

# ====== 6 · การตรวจสอบ / ช่องทางเข้าถึง ======
sp=header("ACCESS & VERIFICATION","ช่องทางตรวจสอบผลงาน",6)
cards=[("💻","Source Code (GitHub)","โค้ดทั้งหมดเปิดเผย (public)\nพร้อม README อธิบายครบ",CYAN),
       ("📐","Kiro Project",".kiro/ (steering + specs)\nหลักฐาน Spec-driven",GOLD),
       ("📜","Commit History","21 commits แสดงลำดับ\nและแนวทางการพัฒนา",EMER),
       ("🌐","Live Demo","เปิดทดสอบระบบจริง\nได้ทันทีในเบราว์เซอร์",VIOLET)]
for i,(ic,t,d,c) in enumerate(cards):
    cx=760000+(i%2)*5550000; cy=2000000+(i//2)*1700000
    sp.append(rrect(cx,cy,5300000,1520000,NAVY3,line=c,lw=9525,r=14000,shadow=True))
    sp.append(icon_badge(cx+180000,cy+280000,940000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+1300000,cy+260000,3900000,420000,[[(t,13.5,c,True,False)]]))
    for j,ln in enumerate(d.split("\n")):
        sp.append(tbox(cx+1300000,cy+700000+j*360000,3900000,350000,[[(ln,10.5,WHITE,False,False)]]))
slides.append(slide(sp))

# ====== 7 · ลิงก์ส่งมอบ ======
sp=starfield(34,seed=11)
sp+=[rrect(0,1550000,EMU_W,1500000,None,grad=(NAVY,NAVY3))]
sp+=[tbox(760000,1680000,10800000,760000,[[("ช่องทางการส่งมอบ",30,WHITE,True,False)]])]
sp+=[tbox(760000,2480000,10800000,420000,[[("Source Code และ Kiro Project — ตรวจสอบได้ครบทุกด้าน",13,CYAN,False,False)]])]
links=[("💻","Source Code / Kiro Project","github.com/kanchanathsao-crypto/lika-lit-review-helper-",CYAN),
       ("📦","ไฟล์ ZIP (โค้ด + .kiro + เอกสาร)",".../main/LIKA-Nexus-Source.zip",GOLD),
       ("🌐","Live Demo","tinyurl.com/lika-nexus",EMER)]
for i,(ic,t,u,c) in enumerate(links):
    y=3350000+i*820000
    sp.append(rrect(760000,y,10700000,700000,NAVY3,line=c,lw=9525,r=14000,shadow=True))
    sp.append(tbox(1000000,y+120000,900000,460000,[[(ic,20,c,False,False)]],anchor="ctr",align="ctr"))
    sp.append(tbox(1900000,y+130000,9400000,330000,[[(t,13,c,True,False)]]))
    sp.append(tbox(1900000,y+440000,9400000,260000,[[(u,11,WHITE,False,False)]]))
slides.append(slide(sp))

# ======================================================================
N=len(slides)
def rels_xml(items):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+items+'</Relationships>')
content_types=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
'<Default Extension="xml" ContentType="application/xml"/>'
'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
+''.join(f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(N))+'</Types>')
root_rels=rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>')
sldid=''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(N))
presentation=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
f'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
f'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
f'<p:sldIdLst>{sldid}</p:sldIdLst><p:sldSz cx="{EMU_W}" cy="{EMU_H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
pres_rels=rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
'<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
+''.join(f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>' for i in range(N)))
theme=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="LIKA">'
'<a:themeElements><a:clrScheme name="LIKA">'
'<a:dk1><a:srgbClr val="060a16"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
'<a:dk2><a:srgbClr val="002366"/></a:dk2><a:lt2><a:srgbClr val="E5EDFF"/></a:lt2>'
'<a:accent1><a:srgbClr val="002366"/></a:accent1><a:accent2><a:srgbClr val="F5C542"/></a:accent2>'
'<a:accent3><a:srgbClr val="22D3EE"/></a:accent3><a:accent4><a:srgbClr val="34D399"/></a:accent4>'
'<a:accent5><a:srgbClr val="A78BFA"/></a:accent5><a:accent6><a:srgbClr val="FB923C"/></a:accent6>'
'<a:hlink><a:srgbClr val="22D3EE"/></a:hlink><a:folHlink><a:srgbClr val="7DD3FC"/></a:folHlink></a:clrScheme>'
'<a:fontScheme name="LIKA"><a:majorFont><a:latin typeface="Tahoma"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
'<a:minorFont><a:latin typeface="Tahoma"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
'<a:fmtScheme name="LIKA"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
'<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
'<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
'<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
'<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
'<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
'<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>')
slide_master=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
'<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="060a16"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
'</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
'<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')
master_rels=rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
'<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>')
slide_layout=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
'<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
'</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>')
layout_rels=rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>')
def slide_rels(): return rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>')

with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml",content_types)
    z.writestr("_rels/.rels",root_rels)
    z.writestr("ppt/presentation.xml",presentation)
    z.writestr("ppt/_rels/presentation.xml.rels",pres_rels)
    z.writestr("ppt/theme/theme1.xml",theme)
    z.writestr("ppt/slideMasters/slideMaster1.xml",slide_master)
    z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",master_rels)
    z.writestr("ppt/slideLayouts/slideLayout1.xml",slide_layout)
    z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",layout_rels)
    for i,sx in enumerate(slides):
        z.writestr(f"ppt/slides/slide{i+1}.xml",sx)
        z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels",slide_rels())
print(f"สร้างสำเร็จ: {OUT} ({N} สไลด์)")

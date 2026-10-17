#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้าง PowerPoint (.pptx) สวยงามพร้อมภาพประกอบ vector (DrawingML)
สำหรับ LIKA Nexus — Mahidol AI Community 2026
ครอบคลุม: Concept Paper · Source Code/Kiro · Presentation/Demo
ไม่ใช้ไลบรารีภายนอก (เขียน OOXML ตรง + แพ็ก ZIP)
"""
import zipfile, html

OUT = "LIKA_Pitch_Deck.pptx"
EMU_W, EMU_H = 12192000, 6858000
EMU = 914400  # 1 inch

NAVY="002366"; NAVY2="060a16"; NAVY3="0b1a3a"; GOLD="f5c542"; CYAN="22d3ee"
WHITE="e5edff"; GREY="9fb3d1"; EMER="34d399"; VIOLET="a78bfa"; ORANGE="fb923c"; PINK="f472b6"

def esc(s): return html.escape(str(s), quote=True)
_id = [10]
def nid():
    _id[0]+=1; return _id[0]

# ---------------- primitives ----------------
def _txbody(runs, anchor, align, wrap=True):
    paras=""
    for grp in runs:
        r=""
        for (t,sz,c,b,i) in grp:
            r+=(f'<a:r><a:rPr lang="th-TH" sz="{int(sz*100)}" b="{1 if b else 0}" i="{1 if i else 0}" dirty="0">'
                f'<a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
                f'<a:latin typeface="Tahoma"/><a:cs typeface="Tahoma"/></a:rPr><a:t>{esc(t)}</a:t></a:r>')
        if not grp:
            r='<a:endParaRPr lang="th-TH"/>'
        paras+=f'<a:p><a:pPr algn="{align}"/>{r}</a:p>'
    return (f'<p:txBody><a:bodyPr wrap="{"square" if wrap else "none"}" anchor="{anchor}">'
            f'<a:normAutofit/></a:bodyPr><a:lstStyle/>{paras}</p:txBody>')

def tbox(x,y,w,h,runs,anchor="t",align="l",wrap=True):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'{_txbody(runs,anchor,align,wrap)}</p:sp>')

def shape(prst,x,y,w,h,fill,line=None,lw=12700,adj="",grad=None,shadow=False):
    if grad:
        fillxml=(f'<a:gradFill><a:gsLst>'
                 f'<a:gs pos="0"><a:srgbClr val="{grad[0]}"/></a:gs>'
                 f'<a:gs pos="100000"><a:srgbClr val="{grad[1]}"/></a:gs></a:gsLst>'
                 f'<a:lin ang="5400000" scaled="1"/></a:gradFill>')
    elif fill is None:
        fillxml='<a:noFill/>'
    else:
        fillxml=f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    ln=(f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
        else '<a:ln><a:noFill/></a:ln>')
    sh=('<a:effectLst><a:outerShdw blurRad="90000" dist="40000" dir="5400000" rotWithShape="0">'
        '<a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst>') if shadow else ''
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="{prst}">{adj or "<a:avLst/>"}</a:prstGeom>{fillxml}{ln}{sh}</p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')

def rrect(x,y,w,h,fill,line=None,lw=10000,r=9000,grad=None,shadow=False):
    return shape("roundRect",x,y,w,h,fill,line,lw,f'<a:avLst><a:gd name="adj" fmla="val {r}"/></a:avLst>',grad,shadow)
def oval(x,y,w,h,fill,line=None,lw=10000,grad=None,shadow=False):
    return shape("ellipse",x,y,w,h,fill,line,lw,"",grad,shadow)
def line(x1,y1,x2,y2,color,w=9525):
    xx,yy=min(x1,x2),min(y1,y2); cx,cy=abs(x2-x1),abs(y2-y1)
    flipH='1' if x2<x1 else '0'; flipV='1' if y2<y1 else '0'
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{nid()}" name="l"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f'<p:spPr><a:xfrm flipH="{flipH}" flipV="{flipV}"><a:off x="{xx}" y="{yy}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{w}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:cxnSp>')

def icon_badge(x,y,size,emoji,bg=NAVY,ring=CYAN):
    """วงกลมเรืองแสง + อิโมจิกลาง = ภาพประกอบไอคอน"""
    s=[oval(x,y,size,size,bg,line=ring,lw=19050,shadow=True),
       tbox(x,y,size,size,[[(emoji,int(size/EMU*20),GOLD,False,False)]],anchor="ctr",align="ctr")]
    return "".join(s)

def chip(x,y,w,text,color):
    return (rrect(x,y,w,230000,None,line=color,lw=9525,r=50000)+
            tbox(x,y-6000,w,250000,[[(text,9,color,True,False)]],anchor="ctr",align="ctr"))

# ---------------- slide frame ----------------
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

def starfield(n=36):
    import random; random.seed(7); s=[]
    for _ in range(n):
        x=random.randint(0,EMU_W); y=random.randint(0,EMU_H); r=random.randint(12000,42000)
        s.append(oval(x,y,r,r,GREY))
    return s

def header(kicker,title,page,total=10,accent=GOLD):
    return [rrect(457200,470000,150000,920000,accent,r=50000),
            tbox(760000,450000,10000000,380000,[[(kicker,13,CYAN,True,False)]]),
            tbox(760000,790000,10800000,900000,[[(title,28,WHITE,True,False)]]),
            tbox(11000000,6360000,1000000,330000,[[(f"{page}/{total}",10,GREY,False,False)]],align="r")]

def bullets(x,y,w,items,size=14,gap=560000,color=WHITE,mark="▸",mc=GOLD):
    s=[]
    for i,it in enumerate(items):
        s.append(tbox(x,y+i*gap,w,gap,[[(f"{mark} ",13,mc,True,False),(it,size,color,False,False)]]))
    return s

slides=[]

# ============ Slide 1: TITLE ============
sp=starfield(44)
sp+=[rrect(0,2450000,EMU_W,1950000,None,grad=(NAVY,NAVY3))]
sp+=[icon_badge(900000,2720000,1300000,"🛰️",bg=NAVY,ring=CYAN)]
sp+=[tbox(2450000,2650000,9200000,900000,[[("LIKA Nexus",38,WHITE,True,False)]]),
     tbox(2450000,3530000,9200000,500000,[[("ผู้ช่วยทบทวนวรรณกรรม & จัดการความรู้งานวิจัย · ห้องสมุดวิทยาเขตกาญจนบุรี",14,CYAN,False,False)]]),
     tbox(2450000,3980000,9200000,420000,[[("Mahidol AI Community 2026  ·  Research & Innovation",11,GREY,False,False)]])]
# chips แถวล่าง
cx=900000
for txt,col in [("3D Knowledge Graph",GOLD),("AI Auto-Cataloger",CYAN),("Usage Dashboard",EMER),("Privacy-first",VIOLET),("Powered by Kiro",ORANGE)]:
    w=len(txt)*78000+260000
    sp.append(chip(cx,4800000,w,txt,col)); cx+=w+150000
sp.append(tbox(900000,6250000,10800000,360000,[[("Demo: ",10,GOLD,True,False),("https://tinyurl.com/lika-nexus",10,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ Slide 2: ภาพรวม 3 รายการส่งผลงาน ============
sp=header("OVERVIEW — องค์ประกอบการส่งผลงาน","ผลงานประกอบด้วย 3 ส่วนหลัก",2)
cards=[("📄","1 · Concept Paper","อธิบายปัญหา แนวคิด วัตถุประสงค์\nกลุ่มผู้ใช้ และการใช้ AI/Kiro",GOLD),
       ("💻","2 · Source Code / Kiro","โค้ดและ Kiro Project\nตรวจสอบโครงสร้าง + การใช้ Kiro",CYAN),
       ("🎬","3 · Presentation / Demo","สื่อนำเสนอ (≤5 นาที)\nProblem→Solution→Impact",EMER)]
x0=760000
for i,(ic,h,d,c) in enumerate(cards):
    cx=x0+i*3760000
    sp.append(rrect(cx,2050000,3520000,3300000,NAVY3,line=c,lw=12700,shadow=True))
    sp.append(icon_badge(cx+1310000,2300000,900000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+200000,3330000,3120000,500000,[[(h,16,c,True,False)]],align="ctr"))
    for j,ln in enumerate(d.split("\n")):
        sp.append(tbox(cx+200000,3800000+j*430000,3120000,430000,[[(ln,12,WHITE,False,False)]],align="ctr"))
slides.append(slide(sp))

# ============ Slide 3: PROBLEM ============
sp=header("① CONCEPT PAPER — ปัญหา","การทำ Literature Review ใช้เวลานานมาก",3)
sp.append(icon_badge(9900000,2050000,1500000,"⏳",bg=NAVY3,ring=ORANGE))
sp+=bullets(760000,2150000,8600000,[
    "การทบทวนวรรณกรรมใช้เวลาหลายวัน–หลายสัปดาห์ ต้องอ่านทีละฉบับ",
    "สายการแพทย์/วิทยาศาสตร์มีจำนวนเปเปอร์มหาศาลและเพิ่มขึ้นทุกวัน",
    "ต้องสกัดประเด็น เปรียบเทียบระเบียบวิธี และร่างบทความเองทั้งหมด",
    "บรรณารักษ์รับคำขอช่วยทบทวนวรรณกรรมมาก แต่เวลาจำกัด",
    "เครื่องมือที่มีมักเสียเงิน ซับซ้อน หรือส่งไฟล์ขึ้น cloud (เสี่ยงข้อมูล)",
], size=14, gap=620000)
slides.append(slide(sp))

# ============ Slide 4: SOLUTION (มี mockup หน้าจอ) ============
sp=header("① CONCEPT PAPER — Solution & วัตถุประสงค์","เว็บผู้ช่วยทบทวนวรรณกรรมอัตโนมัติ",4)
# mockup หน้าต่างแอป (ซ้าย)
mx,my,mw,mh=760000,2000000,5100000,3400000
sp.append(rrect(mx,my,mw,mh,NAVY3,line=CYAN,lw=9525,r=6000,shadow=True))
sp.append(rrect(mx,my,mw,420000,NAVY,r=6000))
for k,cc in enumerate([PINK,GOLD,EMER]):
    sp.append(oval(mx+160000+k*230000,my+140000,130000,130000,cc))
sp.append(tbox(mx+900000,my+90000,3000000,300000,[[("LIKA Nexus",10,WHITE,True,False)]]))
# กราฟโหนดจำลองใน mockup
cxn=mx+mw*0.6; cyn=my+mh*0.52
sp.append(oval(int(cxn-180000),int(cyn-180000),360000,360000,GOLD,shadow=True))
import math
for a,col in [(0,CYAN),(72,EMER),(144,VIOLET),(216,ORANGE),(288,PINK)]:
    rad=1050000; ex=int(cxn+rad*math.cos(math.radians(a))); ey=int(cyn+rad*math.sin(math.radians(a)))
    sp.append(line(int(cxn),int(cyn),ex,ey,"3b82f6"))
    sp.append(oval(ex-150000,ey-150000,300000,300000,col,shadow=True))
sp.append(tbox(mx,my+mh-360000,mw,300000,[[("3D Knowledge Graph (ตัวอย่างหน้าจอจริง)",9,GREY,False,True)]],align="ctr"))
# ฟีเจอร์ (ขวา)
sp.append(tbox(6150000,2000000,5300000,360000,[[("ระบบทำให้อัตโนมัติ:",13,GOLD,True,False)]]))
sp+=bullets(6150000,2420000,5300000,[
    "รับไฟล์ PDF/TXT หลายฉบับ → สกัดประเด็นสำคัญ",
    "ตารางเปรียบเทียบระเบียบวิธีวิจัย 5 ประเด็น",
    "กราฟความรู้ 3 มิติ + ค้นหาเชิงแนวคิด",
    "ร่าง Lit-Review พร้อมอ้างอิง APA แก้ไขได้",
    "ลดเวลาจากหลายวัน → ไม่กี่นาที",
], size=12.5, gap=560000)
slides.append(slide(sp))

# ============ Slide 5: USERS ============
sp=header("① CONCEPT PAPER — กลุ่มผู้ใช้ / ผู้ได้รับประโยชน์","ออกแบบเพื่อประชาคม MUKA",5)
users=[("👩‍🏫","อาจารย์ / นักวิจัย","ร่าง Lit-Review ตั้งต้น + เทียบระเบียบวิธีทันที",GOLD),
       ("🧑‍💼","บรรณารักษ์ / เจ้าหน้าที่","ให้บริการทบทวนวรรณกรรมเร็วขึ้น + เก็บสถิติผู้ใช้",CYAN),
       ("🧑‍🎓","นักศึกษา","เรียนรู้โครงสร้างงานวิจัยและเริ่มเขียนได้ง่าย",EMER),
       ("🌐","ครอบคลุม 22 สาขา","วิทยาศาสตร์ · สังคมศาสตร์ · มนุษยศาสตร์",VIOLET)]
for i,(ic,t,d,c) in enumerate(users):
    cx=760000+(i%2)*5550000; cy=2050000+(i//2)*1550000
    sp.append(rrect(cx,cy,5300000,1380000,NAVY3,line=c,lw=9525,shadow=True))
    sp.append(icon_badge(cx+180000,cy+220000,940000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+1300000,cy+230000,3900000,480000,[[(t,15,c,True,False)]]))
    sp.append(tbox(cx+1300000,cy+720000,3900000,560000,[[(d,12,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ Slide 6: AI & KIRO ============
sp=header("① CONCEPT PAPER — แนวทางการใช้ AI & Kiro","พัฒนาแบบ Spec-driven ด้วย Kiro",6)
sp.append(rrect(760000,2000000,5250000,3350000,NAVY3,line=CYAN,lw=12700,shadow=True))
sp.append(icon_badge(760000+2150000,2180000,860000,"🤖",bg=NAVY,ring=GOLD))
sp.append(tbox(960000,3080000,4900000,420000,[[("AI / Automated Reasoning",15,GOLD,True,False)]],align="ctr"))
sp+=bullets(1000000,3520000,4800000,[
    "สกัดข้อมูลตามโครงสร้าง IMRaD",
    "จำแนก 22 สาขา + ระเบียบวิธี (weighted scoring)",
    "รองรับไทย–อังกฤษ · ดึงชื่อผู้แต่งเต็ม",
], size=12, gap=560000)
sp.append(rrect(6260000,2000000,5200000,3350000,NAVY3,line=GOLD,lw=12700,shadow=True))
sp.append(icon_badge(6260000+2170000,2180000,860000,"⚙️",bg=NAVY,ring=CYAN))
sp.append(tbox(6460000,3080000,4800000,420000,[[("Kiro = เครื่องมือหลัก",15,CYAN,True,False)]],align="ctr"))
sp+=bullets(6500000,3520000,4800000,[
    ".kiro/steering — มาตรฐานการพัฒนา",
    ".kiro/specs — Acceptance Criteria",
    "Kiro ช่วยเขียนโค้ด ทดสอบ (Node.js) & deploy",
], size=12, gap=560000)
slides.append(slide(sp))

# ============ Slide 7: SOURCE CODE / KIRO PROJECT ============
sp=header("② SOURCE CODE / KIRO PROJECT","โครงสร้างผลงาน & การส่งมอบ",7)
# tree ไฟล์ (ซ้าย)
sp.append(rrect(760000,2000000,5300000,3350000,"0a1428",line=CYAN,lw=9525,r=5000,shadow=True))
tree=[("📁 LIKA.kiro/",GOLD,0),("⭐ lika-nexus.html",WHITE,1),
      ("📄 lit-review-helper.html",GREY,1),("🗂 search.html",GREY,1),
      ("📁 .kiro/",CYAN,1),("📐 steering/rules.md",WHITE,2),
      ("📐 specs/lit-review-spec.md",WHITE,2),("📊 LIKA_Pitch_Deck.pptx",GOLD,1)]
for i,(t,c,ind) in enumerate(tree):
    sp.append(tbox(1000000+ind*360000,2200000+i*380000,4900000,370000,[[(t,12,c,ind==0,False)]]))
# การส่งมอบ (ขวา)
sp.append(tbox(6260000,2000000,5200000,380000,[[("ช่องทางการส่งมอบ & ตรวจสอบ",13,GOLD,True,False)]]))
links=[("💻 Source (GitHub)","github.com/kanchanathsao-crypto/","lika-lit-review-helper-",CYAN),
       ("🌐 Demo ใช้งานจริง","tinyurl.com/lika-nexus","",EMER),
       ("🛠 เทคโนโลยี","HTML5 · Tailwind · Three.js r128","PDF.js · ประมวลผลในเบราว์เซอร์",VIOLET),
       ("✅ พัฒนาด้วย Kiro","Spec-driven · ตรวจ commit history","ได้จาก GitHub",GOLD)]
for i,(t,d1,d2,c) in enumerate(links):
    y=2450000+i*720000
    sp.append(rrect(6260000,y,5200000,640000,NAVY3,line=c,lw=7000,r=10000))
    sp.append(tbox(6440000,y+70000,4900000,330000,[[(t,12,c,True,False)]]))
    sp.append(tbox(6440000,y+360000,4900000,260000,[[((d1+" "+d2).strip(),10,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ Slide 8: FEATURES + DASHBOARD (ภาพกราฟ) ============
sp=header("② PROTOTYPE — ฟีเจอร์เด่น + สถิติผู้ใช้","ระบบ Login & Usage Dashboard",8)
# โดนัทจำลอง (ซ้าย)
cxn,cyn=2300000,3550000
sp.append(oval(cxn-950000,cyn-950000,1900000,1900000,None,line="131c2e",lw=180000))
# ส่วนโค้งสี (จำลองด้วย arcs ใช้ block ซ้อน — ใช้ pie)
def pie(cx,cy,r,start,sweep,color):
    adj=f'<a:avLst><a:gd name="adj1" fmla="val {int(start*60000)}"/><a:gd name="adj2" fmla="val {int((start+sweep)*60000)}"/></a:avLst>'
    return shape("pie",cx-r,cy-r,2*r,2*r,color,adj=adj)
sp.append(pie(cxn,cyn,900000,-90,130,GOLD))
sp.append(pie(cxn,cyn,900000,40,120,CYAN))
sp.append(pie(cxn,cyn,900000,160,110,EMER))
sp.append(oval(cxn-430000,cyn-430000,860000,860000,NAVY2))
sp.append(tbox(cxn-430000,cyn-250000,860000,500000,[[("3 กลุ่ม",12,WHITE,True,False)]],anchor="ctr",align="ctr"))
sp.append(tbox(900000,4650000,2800000,300000,[[("สัดส่วนผู้ใช้ตามประเภท",10,GREY,False,True)]],align="ctr"))
# legend
for i,(t,c) in enumerate([("อาจารย์",GOLD),("เจ้าหน้าที่",CYAN),("นักศึกษา",EMER)]):
    sp.append(oval(760000,2200000+i*330000,200000,200000,c))
    sp.append(tbox(1040000,2170000+i*330000,2300000,280000,[[(t,11,WHITE,False,False)]]))
# กราฟแท่งจำลอง (กลาง-ขวา)
bx=4500000; bh=1500000; heights=[0.4,0.7,0.5,0.9,0.6,1.0,0.8]
sp.append(tbox(bx,2050000,5000000,320000,[[("การเข้าใช้ย้อนหลัง 7 วัน",12,GOLD,True,False)]]))
for i,hh in enumerate(heights):
    h=int(bh*hh); x=bx+i*560000
    sp.append(rrect(x,2500000+(bh-h),360000,h,None,grad=(GOLD,CYAN),r=15000))
# ฟีเจอร์ list
sp+=bullets(4500000,4350000,7000000,[
    "Login แยกประเภท: อาจารย์ · เจ้าหน้าที่ · นักศึกษา (เด้งต้อนรับหน้าแรก)",
    "Dashboard: การ์ดสถิติ · โดนัท · กราฟ 7 วัน · บันทึกล่าสุด",
    "ประมวลผล + เก็บสถิติในเบราว์เซอร์ (privacy-first)",
], size=11.5, gap=470000, mc=CYAN)
slides.append(slide(sp))

# ============ Slide 9: DEMO VIDEO FLOW ============
sp=header("③ PRESENTATION / DEMO VIDEO (≤ 5 นาที)","เส้นเรื่องการนำเสนอ 5 ช่วง",9)
flow=[("Problem","0:00–0:45","Lit Review ใช้เวลานาน",ORANGE,"⏳"),
      ("Solution","0:45–1:30","LIKA ช่วยอัตโนมัติ",GOLD,"💡"),
      ("AI / Kiro","1:30–2:15","Spec-driven + สกัด IMRaD",CYAN,"🤖"),
      ("Prototype","2:15–4:00","สาธิตสด: อัปโหลด→กราฟ→ร่าง",EMER,"🖥️"),
      ("Impact","4:00–4:45","ลดเวลา · 22 สาขา · ต่อยอด",VIOLET,"🚀")]
x0=760000; bw=2000000
for i,(t,tm,d,c,ic) in enumerate(flow):
    cx=x0+i*2170000
    sp.append(rrect(cx,2300000,bw,2500000,NAVY3,line=c,lw=9525,shadow=True))
    sp.append(icon_badge(cx+bw//2-430000,2500000,860000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+100000,3450000,bw-200000,360000,[[(t,13,c,True,False)]],align="ctr"))
    sp.append(tbox(cx+100000,3800000,bw-200000,300000,[[(tm,10,GOLD,False,False)]],align="ctr"))
    sp.append(tbox(cx+100000,4100000,bw-200000,650000,[[(d,10.5,WHITE,False,False)]],align="ctr"))
    if i<len(flow)-1:
        sp.append(tbox(cx+bw-30000,3350000,250000,400000,[[("→",18,GREY,True,False)]],align="ctr"))
slides.append(slide(sp))

# ============ Slide 10: IMPACT / THANK YOU ============
sp=starfield(40)
sp+=[rrect(0,1550000,EMU_W,1600000,None,grad=(NAVY,NAVY3))]
sp+=[tbox(760000,1700000,10800000,800000,[[("Expected Impact & ขอบคุณ",30,WHITE,True,False)]])]
sp+=[tbox(760000,2550000,10800000,450000,[[("ลดเวลา Lit Review จากหลายวัน → ไม่กี่นาที · รองรับ 22 สาขา · Privacy-first",13,CYAN,False,False)]])]
stats=[("หลายวัน→นาที","ลดเวลาทบทวนวรรณกรรม",GOLD),
       ("22 สาขา","วิทย์·สังคม·มนุษยศาสตร์",CYAN),
       ("3 ประเภทผู้ใช้","เก็บสถิติบริการได้จริง",EMER),
       ("ต่อยอดได้","สืบค้นออนไลน์·เชื่อมระบบ",VIOLET)]
for i,(n,d,c) in enumerate(stats):
    cx=760000+i*2720000
    sp.append(rrect(cx,3450000,2550000,1400000,NAVY3,line=c,lw=9525,shadow=True))
    sp.append(tbox(cx+150000,3600000,2250000,600000,[[(n,19,c,True,False)]],align="ctr"))
    sp.append(tbox(cx+150000,4250000,2250000,500000,[[(d,10.5,WHITE,False,False)]],align="ctr"))
sp.append(tbox(760000,5300000,10800000,380000,[[("🌐 Demo: ",11,GOLD,True,False),("https://tinyurl.com/lika-nexus",11,WHITE,False,False)]]))
sp.append(tbox(760000,5720000,10800000,380000,[[("💻 Source: ",11,GOLD,True,False),("github.com/kanchanathsao-crypto/lika-lit-review-helper-",11,WHITE,False,False)]]))
slides.append(slide(sp))

# ======================================================================
# ประกอบไฟล์ .pptx
# ======================================================================
N=len(slides)
content_types=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
'<Default Extension="xml" ContentType="application/xml"/>'
'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
+ ''.join(f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(N))
+ '</Types>')

root_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
'</Relationships>')

sldid=''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(N))
presentation=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
f'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
f'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
f'<p:sldIdLst>{sldid}</p:sldIdLst><p:sldSz cx="{EMU_W}" cy="{EMU_H}"/>'
f'<p:notesSz cx="6858000" cy="9144000"/></p:presentation>')

pres_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
'<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
+ ''.join(f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>' for i in range(N))
+ '</Relationships>')

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
'<a:fmtScheme name="LIKA"><a:fillStyleLst>'
'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
'<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
'<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
'<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
'<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
'<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
'<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
'</a:fmtScheme></a:themeElements></a:theme>')

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

master_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
'<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
'</Relationships>')

slide_layout=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
'<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
'</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>')

layout_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
'</Relationships>')

def slide_rels():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '</Relationships>')

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

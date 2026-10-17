#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้าง PowerPoint "Concept Paper" (.pptx) — LIKA Nexus
ละเอียด · ภาพประกอบ vector · ภาษาทางการ · จัดเนื้อหาให้ตอบเกณฑ์ Technical Review (รอบ 1)
ไม่ใช้ไลบรารีภายนอก (เขียน OOXML ตรง + แพ็ก ZIP)
"""
import zipfile, html, math, random

OUT="LIKA_Concept_Paper.pptx"
EMU_W,EMU_H=12192000,6858000
NAVY="002366"; NAVY2="060a16"; NAVY3="0b1a3a"; GOLD="f5c542"; CYAN="22d3ee"
WHITE="e5edff"; GREY="9fb3d1"; EMER="34d399"; VIOLET="a78bfa"; ORANGE="fb923c"; PINK="f472b6"; RED="f87171"

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
    """โลโก้ nexus node ย่อ"""
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

def header(kicker,title,page,total,accent=GOLD,score=None):
    s=[rrect(457200,470000,150000,920000,accent,r=50000),
       tbox(760000,450000,9200000,380000,[[(kicker,13,CYAN,True,False)]]),
       tbox(760000,790000,9600000,900000,[[(title,26,WHITE,True,False)]]),
       tbox(11000000,6360000,1000000,330000,[[(f"{page}/{total}",10,GREY,False,False)]],align="r")]
    if score:
        s.append(rrect(9850000,470000,1850000,560000,None,line=GOLD,lw=12700,r=16000))
        s.append(tbox(9850000,520000,1850000,250000,[[("เกณฑ์ที่เกี่ยวข้อง",8,CYAN,True,False)]],align="ctr"))
        s.append(tbox(9850000,740000,1850000,260000,[[(score,12,GOLD,True,False)]],align="ctr"))
    return s

def bullets(x,y,w,items,size=13.5,gap=540000,color=WHITE,mark="▸",mc=GOLD):
    return [tbox(x,y+i*gap,w,gap,[[(f"{mark} ",12,mc,True,False),(it,size,color,False,False)]]) for i,it in enumerate(items)]

slides=[]

# ============ 1 · TITLE ============
sp=starfield(46)
sp+=[rrect(0,2350000,EMU_W,2000000,None,grad=(NAVY,NAVY3))]
sp+=[nexus_logo(820000,2620000,1420000)]
sp+=[tbox(2550000,2560000,9000000,620000,[[("Concept Paper",15,GOLD,True,False)]]),
     tbox(2550000,2960000,9000000,760000,[[("LIKA Nexus",40,WHITE,True,False)]]),
     tbox(2550000,3760000,9000000,480000,[[("ระบบผู้ช่วยทบทวนวรรณกรรมและจัดการความรู้งานวิจัย · ห้องสมุดวิทยาเขตกาญจนบุรี",13,CYAN,False,False)]])]
cx=820000
for txt,col in [("Research & Innovation",GOLD),("Powered by Kiro",CYAN),("Privacy-first",EMER),("22 สาขาวิชา",VIOLET)]:
    w=len(txt)*80000+280000; sp.append(chip(cx,4600000,w,txt,col)); cx+=w+150000
sp+=[tbox(820000,5750000,10600000,340000,[[("เสนอต่อ: ",10,GREY,True,False),("Mahidol AI Community 2026 — การพิจารณาผลงานรอบที่ 1 (Technical Review)",10,WHITE,False,False)]]),
     tbox(820000,6120000,10600000,340000,[[("Demo: ",10,GOLD,True,False),("https://tinyurl.com/2u452bd4",10,WHITE,False,False)]])]
slides.append(slide(sp))

# ============ 2 · บทสรุปผู้บริหาร + สารบัญ ============
sp=header("EXECUTIVE SUMMARY","บทสรุปผู้บริหาร",2,10)
sp.append(rrect(760000,1950000,6400000,3750000,NAVY3,line=CYAN,lw=12700,shadow=True))
sp.append(tbox(1000000,2120000,6000000,1500000,[[("LIKA Nexus เป็นเว็บแอปพลิเคชันผู้ช่วยทบทวนวรรณกรรม",13,WHITE,False,False)],
    [("ที่นำไฟล์งานวิจัยหลายฉบับเข้ามาสกัดประเด็นสำคัญโดยอัตโนมัติ",13,WHITE,False,False)],
    [("จัดทำตารางเปรียบเทียบระเบียบวิธีวิจัย กราฟความรู้ 3 มิติ",13,WHITE,False,False)],
    [("และร่างบททบทวนวรรณกรรมพร้อมการอ้างอิง เพื่อลดระยะเวลา",13,WHITE,False,False)],
    [("การทำ Literature Review จากหลายวันเหลือเพียงไม่กี่นาที",13,WHITE,False,False)]]))
sp.append(tbox(1000000,4150000,6000000,400000,[[("ประมวลผลภายในเบราว์เซอร์ทั้งหมด (Privacy-first)",12,EMER,True,False)]]))
sp.append(tbox(1000000,4520000,6000000,400000,[[("พัฒนาด้วย Kiro แบบ Spec-driven · ใช้งานได้จริงแล้ว",12,GOLD,True,False)]]))
# สารบัญ
sp.append(tbox(7450000,1980000,4000000,360000,[[("หัวข้อในเอกสาร",13,GOLD,True,False)]]))
toc=["1  ปัญหา (Problem)","2  แนวคิดและวัตถุประสงค์","3  กลุ่มผู้ใช้ / ผู้ได้รับประโยชน์",
     "4  การประยุกต์ใช้ AI","5  การประยุกต์ใช้ Kiro","6  ต้นแบบ (Prototype)","7  คุณค่า & โอกาสใช้งานจริง"]
for i,t in enumerate(toc):
    sp.append(rrect(7450000,2420000+i*440000,4050000,380000,NAVY3,line=None,r=12000))
    sp.append(tbox(7620000,2440000+i*440000,3800000,340000,[[(t,11,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 3 · PROBLEM ============
sp=header("① ปัญหา (PROBLEM)","การทบทวนวรรณกรรมใช้เวลาและทรัพยากรสูง",3,10,accent=RED)
sp.append(icon_badge(10050000,2000000,1300000,"⏳",bg=NAVY3,ring=RED))
sp+=bullets(760000,2050000,8900000,[
    "การทบทวนวรรณกรรมต้องอ่านเอกสารจำนวนมากทีละฉบับ ใช้เวลาหลายวันถึงหลายสัปดาห์",
    "สาขาวิทยาศาสตร์และการแพทย์มีปริมาณงานวิจัยเพิ่มขึ้นอย่างรวดเร็วและต่อเนื่อง",
    "ผู้วิจัยต้องสกัดประเด็น เปรียบเทียบระเบียบวิธี และเรียบเรียงบทความด้วยตนเองทั้งหมด",
    "บุคลากรห้องสมุดได้รับคำขอสนับสนุนการสืบค้นจำนวนมากภายใต้ข้อจำกัดด้านเวลา",
    "เครื่องมือที่มีอยู่มักมีค่าใช้จ่ายสูง ใช้งานซับซ้อน หรือต้องอัปโหลดเอกสารขึ้นระบบภายนอก",
], size=13.5, gap=620000)
sp.append(rrect(760000,5500000,10700000,760000,None,line=ORANGE,lw=9525,r=14000))
sp.append(tbox(1000000,5600000,10300000,560000,[[("ผลกระทบ: ",12,ORANGE,True,False),
    ("ประสิทธิภาพงานวิจัยและงานบริการวิชาการลดลง และมีความเสี่ยงด้านความลับของงานวิจัยที่ยังไม่เผยแพร่",12,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 4 · แนวคิด & วัตถุประสงค์ ============
sp=header("② แนวคิดและวัตถุประสงค์ของ SOLUTION","แปลงงานทบทวนวรรณกรรมให้เป็นระบบอัตโนมัติ",4,10)
sp.append(tbox(760000,1950000,11000000,360000,[[("แนวคิดหลัก (Concept)",13,GOLD,True,False)]]))
sp.append(rrect(760000,2320000,10700000,760000,NAVY3,line=CYAN,lw=9525,r=12000))
sp.append(tbox(1000000,2420000,10300000,600000,[[("รวบรวม–วิเคราะห์–สังเคราะห์ งานวิจัยหลายฉบับไว้ในระบบเดียว โดยใช้ AI สกัดสาระสำคัญ แล้วนำเสนอผ่านกราฟความรู้ ตารางเปรียบเทียบ และร่างบททบทวนวรรณกรรมที่ปรับแก้ได้",12.5,WHITE,False,False)]]))
sp.append(tbox(760000,3300000,11000000,360000,[[("วัตถุประสงค์ (Objectives)",13,GOLD,True,False)]]))
objs=[("⏱️","ลดระยะเวลา","ย่นเวลาทบทวนวรรณกรรมจากหลายวันเหลือไม่กี่นาที",GOLD),
      ("🧩","เพิ่มคุณภาพ","เปรียบเทียบระเบียบวิธีอย่างเป็นระบบ 5 ประเด็น",CYAN),
      ("🔐","รักษาความลับ","ประมวลผลในเบราว์เซอร์ ไม่ส่งไฟล์ออกภายนอก",EMER),
      ("📚","ครอบคลุม","รองรับงานวิจัย 22 สาขา ทั้งวิทย์-สังคม-มนุษยศาสตร์",VIOLET)]
for i,(ic,t,d,c) in enumerate(objs):
    cx=760000+(i%2)*5550000; cy=3700000+(i//2)*1030000
    sp.append(rrect(cx,cy,5300000,920000,NAVY3,line=c,lw=8000,r=12000))
    sp.append(icon_badge(cx+150000,cy+150000,620000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+900000,cy+140000,4300000,360000,[[(t,13,c,True,False)]]))
    sp.append(tbox(cx+900000,cy+500000,4300000,380000,[[(d,10.5,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 5 · กลุ่มผู้ใช้ ============
sp=header("③ กลุ่มผู้ใช้ / ผู้ได้รับประโยชน์","ออกแบบเพื่อประชาคมวิจัย MUKA",5,10)
users=[("👩‍🏫","อาจารย์ / นักวิจัย","จัดทำร่างบททบทวนวรรณกรรมตั้งต้นและเปรียบเทียบระเบียบวิธีได้รวดเร็ว",GOLD),
       ("🧑‍💼","บุคลากร / บรรณารักษ์","ให้บริการสนับสนุนงานวิจัยได้มีประสิทธิภาพ พร้อมจัดเก็บสถิติการใช้บริการ",CYAN),
       ("🧑‍🎓","นักศึกษา","เรียนรู้โครงสร้างงานวิจัยและเริ่มเขียนบททบทวนวรรณกรรมได้โดยง่าย",EMER),
       ("🏛️","มหาวิทยาลัย","ยกระดับคุณภาพงานวิจัยและบริการวิชาการของวิทยาเขต",VIOLET)]
for i,(ic,t,d,c) in enumerate(users):
    cx=760000+(i%2)*5550000; cy=2000000+(i//2)*1620000
    sp.append(rrect(cx,cy,5300000,1450000,NAVY3,line=c,lw=9525,r=14000,shadow=True))
    sp.append(icon_badge(cx+180000,cy+250000,940000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+1300000,cy+230000,3900000,430000,[[(t,14,c,True,False)]]))
    sp.append(tbox(cx+1300000,cy+660000,3900000,700000,[[(d,11,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 6 · AI (ความคิดสร้างสรรค์/นวัตกรรม 20) ============
sp=header("④ การประยุกต์ใช้ AI","นวัตกรรมการสกัดและสังเคราะห์องค์ความรู้",6,10,accent=CYAN,score="20 คะแนน")
sp.append(rrect(760000,1950000,5250000,3750000,NAVY3,line=CYAN,lw=12700,shadow=True))
sp.append(icon_badge(760000+2150000,2130000,900000,"🧠",bg=NAVY,ring=CYAN))
sp.append(tbox(960000,3080000,4900000,400000,[[("กลไก AI ของระบบ",14,CYAN,True,False)]],align="ctr"))
sp+=bullets(1000000,3520000,4800000,[
    "สกัดสาระตามโครงสร้างงานวิจัย (IMRaD)",
    "จำแนก 22 สาขา + ระเบียบวิธีด้วยการให้น้ำหนักคำสำคัญ",
    "รองรับเอกสารไทย–อังกฤษ และดึงชื่อผู้แต่งเต็ม",
    "สรุปเชิงผู้บริหาร (Executive Summary) ต่อเอกสาร",
], size=11.5, gap=540000)
sp.append(rrect(6260000,1950000,5200000,3750000,NAVY3,line=GOLD,lw=12700,shadow=True))
sp.append(tbox(6460000,2120000,4800000,400000,[[("จุดเด่นเชิงนวัตกรรม",14,GOLD,True,False)]]))
sp+=bullets(6500000,2560000,4800000,[
    "รวม 3 มุมมองในระบบเดียว: กราฟ 3 มิติ · ตารางเทียบ · ร่างบทความ",
    "ค้นหาเชิงแนวคิด (Semantic Search) ภายในคลังเอกสาร",
    "ร่างบททบทวนวรรณกรรมอัตโนมัติพร้อมอ้างอิง APA",
    "ทำงานแบบ Privacy-first โดยไม่พึ่งเซิร์ฟเวอร์ภายนอก",
], size=11.5, gap=640000, mc=GOLD)
slides.append(slide(sp))

# ============ 7 · KIRO (สอดคล้องโจทย์+ใช้ Kiro 20) ============
sp=header("⑤ การประยุกต์ใช้ Kiro Platform","พัฒนาแบบ Spec-driven ด้วย Kiro",7,10,accent=GOLD,score="20 คะแนน")
# ขั้นตอน Kiro
steps=[("📐","Steering Rules",".kiro/steering/rules.md\nกำหนดมาตรฐานการพัฒนา",GOLD),
       ("📋","Specs",".kiro/specs/*.md\nกำหนดเกณฑ์ยอมรับผลงาน",CYAN),
       ("🤝","Agent พัฒนา","Kiro เขียนโค้ด + ทดสอบ\nตามข้อกำหนด",EMER),
       ("🚀","Deploy","เผยแพร่ GitHub Pages\nตรวจสอบได้จริง",VIOLET)]
x0=760000
for i,(ic,t,d,c) in enumerate(steps):
    cx=x0+i*2720000
    sp.append(rrect(cx,2050000,2550000,2550000,NAVY3,line=c,lw=9525,r=14000,shadow=True))
    sp.append(icon_badge(cx+2550000//2-430000,2250000,860000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+100000,3200000,2350000,360000,[[(t,13,c,True,False)]],align="ctr"))
    for j,ln in enumerate(d.split("\n")):
        sp.append(tbox(cx+120000,3560000+j*380000,2310000,370000,[[(ln,10,WHITE,False,False)]],align="ctr"))
    if i<len(steps)-1:
        sp.append(tbox(cx+2550000-30000,3050000,250000,400000,[[("→",18,GREY,True,False)]],align="ctr"))
sp.append(rrect(760000,4850000,10700000,820000,None,line=GOLD,lw=9525,r=14000))
sp.append(tbox(1000000,4950000,10300000,620000,[[("ความสอดคล้องกับโจทย์: ",12,GOLD,True,False),
    ("Kiro เป็นเครื่องมือหลักตลอดวงจรพัฒนา — กำหนดกฎ ออกแบบข้อกำหนด เขียนและทดสอบโค้ด จนถึงการเผยแพร่ ตรวจสอบย้อนกลับได้จาก commit history",11.5,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 8 · PROTOTYPE (คุณภาพต้นแบบ 20) ============
sp=header("⑥ ต้นแบบ (PROTOTYPE)","ระบบที่พัฒนาเสร็จและใช้งานได้จริง",8,10,accent=EMER,score="20 คะแนน")
# mockup
mx,my,mw,mh=760000,2000000,5100000,3350000
sp.append(rrect(mx,my,mw,mh,NAVY3,line=CYAN,lw=9525,r=6000,shadow=True))
sp.append(rrect(mx,my,mw,400000,NAVY,r=6000))
for k,cc in enumerate([PINK,GOLD,EMER]): sp.append(oval(mx+160000+k*230000,my+130000,130000,130000,cc))
sp.append(tbox(mx+900000,my+80000,3000000,300000,[[("LIKA Nexus",10,WHITE,True,False)]]))
cxn=mx+mw*0.58; cyn=my+mh*0.54
sp.append(oval(int(cxn-170000),int(cyn-170000),340000,340000,GOLD,shadow=True))
for a,col in [(0,CYAN),(72,EMER),(144,VIOLET),(216,ORANGE),(288,PINK)]:
    ex=int(cxn+1000000*math.cos(math.radians(a))); ey=int(cyn+1000000*math.sin(math.radians(a)))
    sp.append(lineshape(int(cxn),int(cyn),ex,ey,"3b82f6")); sp.append(oval(ex-140000,ey-140000,280000,280000,col,shadow=True))
sp.append(tbox(mx,my+mh-340000,mw,300000,[[("หน้าจอจริง: 3D Knowledge Graph",9,GREY,False,True)]],align="ctr"))
# คุณสมบัติที่พร้อมใช้
sp.append(tbox(6150000,1980000,5300000,360000,[[("ฟังก์ชันที่พัฒนาเสร็จแล้ว",13,EMER,True,False)]]))
feats=["นำเข้าไฟล์ PDF/TXT และสกัดข้อมูลอัตโนมัติ","ตารางเปรียบเทียบระเบียบวิธีวิจัย 5 ประเด็น",
       "กราฟความรู้ 3 มิติ + ค้นหาเชิงแนวคิด","ร่าง Lit-Review + ส่งออกอ้างอิง APA",
       "ระบบเข้าใช้แยกประเภทผู้ใช้ + แดชบอร์ดสถิติ","ส่งออกบทสรุปรายงานเป็น PDF"]
sp+=bullets(6150000,2420000,5300000,feats,size=11.5,gap=470000,mc=EMER)
sp.append(rrect(6150000,5250000,5300000,420000,None,line=EMER,lw=9525,r=12000))
sp.append(tbox(6330000,5300000,5000000,330000,[[("สถานะ: ",11,EMER,True,False),("เผยแพร่ออนไลน์แล้ว · tinyurl.com/2u452bd4",11,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 9 · FEASIBILITY + คุณค่า (30) ============
sp=header("⑦ คุณค่าและความเป็นไปได้ในการใช้งานจริง","Feasibility & Value",9,10,accent=VIOLET,score="30 คะแนน")
vals=[("💰","ต้นทุนต่ำ","เป็นไฟล์เว็บเดียว ไม่ต้องติดตั้งเซิร์ฟเวอร์ ใช้บริการฟรี",GOLD),
      ("🏫","พร้อมใช้ที่ MUKA","เหมาะกับเครื่องบริการในห้องสมุด ใช้ได้แม้ออฟไลน์",CYAN),
      ("🔐","ปลอดภัย","ไม่ส่งข้อมูลออกนอกเครื่อง เหมาะกับงานยังไม่เผยแพร่",EMER),
      ("📈","ขยายผลได้","ต่อยอดสู่การสืบค้นออนไลน์และเชื่อมระบบห้องสมุด",VIOLET)]
for i,(ic,t,d,c) in enumerate(vals):
    cx=760000+(i%2)*5550000; cy=2000000+(i//2)*1250000
    sp.append(rrect(cx,cy,5300000,1120000,NAVY3,line=c,lw=8000,r=12000))
    sp.append(icon_badge(cx+160000,cy+230000,660000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+940000,cy+200000,4250000,360000,[[(t,13,c,True,False)]]))
    sp.append(tbox(cx+940000,cy+560000,4250000,500000,[[(d,10.5,WHITE,False,False)]]))
sp.append(rrect(760000,4650000,10700000,1000000,None,line=GOLD,lw=9525,r=14000))
sp.append(tbox(1000000,4740000,10300000,400000,[[("โอกาสในการนำไปใช้จริง",12,GOLD,True,False)]]))
sp.append(tbox(1000000,5080000,10300000,520000,[[("นำไปใช้สนับสนุนการทำวิทยานิพนธ์ บริการตอบคำถามวิจัย และจัดการคลังความรู้ของวิทยาเขต โดยสามารถขยายสู่ระบบเก็บสถิติรวมศูนย์และการสืบค้นฐานข้อมูลสากลในอนาคต",11.5,WHITE,False,False)]]))
slides.append(slide(sp))

# ============ 10 · สรุป + checklist เกณฑ์ ============
sp=starfield(36,seed=11)
sp+=[rrect(0,470000,EMU_W,1050000,None,grad=(NAVY,NAVY3))]
sp+=[tbox(760000,560000,11000000,760000,[[("สรุปความสอดคล้องกับเกณฑ์พิจารณา",26,WHITE,True,False)]])]
crit=[("ความสอดคล้องกับโจทย์ & การใช้ Kiro","พัฒนาด้วย Kiro ตลอดวงจร (Steering + Specs)","20",GOLD),
      ("ความคิดสร้างสรรค์และนวัตกรรม","รวมกราฟ 3 มิติ + สกัด AI + ร่างอัตโนมัติ","20",CYAN),
      ("ความเป็นไปได้ในการใช้จริง (Feasibility)","ต้นทุนต่ำ ปลอดภัย พร้อมใช้ที่ MUKA","30",VIOLET),
      ("คุณภาพของต้นแบบ (Prototype)","พัฒนาเสร็จ เผยแพร่ออนไลน์ ใช้งานได้จริง","20",EMER),
      ("ความครบถ้วนของเอกสาร & การนำเสนอ","เอกสารครบทุกหัวข้อ อธิบายชัดเจน","10",ORANGE)]
y0=1780000
for i,(t,d,sc,c) in enumerate(crit):
    cy=y0+i*870000
    sp.append(rrect(760000,cy,10700000,780000,NAVY3,line=c,lw=8000,r=12000))
    sp.append(tbox(950000,cy+90000,650000,600000,[[("✓",20,c,True,False)]],anchor="ctr",align="ctr"))
    sp.append(tbox(1700000,cy+130000,7600000,360000,[[(t,13,WHITE,True,False)]]))
    sp.append(tbox(1700000,cy+470000,7600000,280000,[[(d,10.5,GREY,False,False)]]))
    sp.append(rrect(10350000,cy+180000,950000,420000,None,line=c,lw=9525,r=16000))
    sp.append(tbox(10350000,cy+230000,950000,330000,[[(sc,16,c,True,False)]],align="ctr"))
slides.append(slide(sp))

# ======================================================================
# PACKAGE
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

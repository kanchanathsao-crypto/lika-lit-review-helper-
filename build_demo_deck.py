#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้าง PowerPoint สำหรับ Demo Video (<=5 นาที) — LIKA Nexus
เส้นเรื่อง: Problem -> Solution -> AI/Kiro -> Prototype -> Expected Impact
มี timecode กำกับ + บทพากย์ใน Speaker Notes · ภาพประกอบ vector · ไม่ใส่คะแนน
"""
import zipfile, html, math, random

OUT="LIKA_Demo_Video.pptx"
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
    cx=max(cx,1); cy=max(cy,1)
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

def timebadge(tc,accent):
    # ป้าย timecode มุมขวาบน
    w=1700000
    return (rrect(9950000,470000,w,540000,None,line=accent,lw=12700,r=16000)+
            tbox(9950000,520000,w,250000,[[("⏱ TIMECODE",8,CYAN,True,False)]],align="ctr")+
            tbox(9950000,740000,w,260000,[[(tc,13,accent,True,False)]],align="ctr"))

# transition ระหว่างสไลด์ (สลับ fade / push ตามลำดับ)
def transition_xml(kind="fade"):
    if kind=="push":
        return '<p:transition spd="med" p14:dur="700" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:push dir="l"/></p:transition>'
    if kind=="wipe":
        return '<p:transition spd="med" p14:dur="700" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:wipe dir="l"/></p:transition>'
    return '<p:transition spd="med" p14:dur="600" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:fade/></p:transition>'

# timing tree: ชิ้นงาน (anim_ids) ทยอยปรากฏแบบ fade-in อัตโนมัติ (After Previous)
# ใช้ animEffect fade เข้า — เป็น animation มาตรฐานที่ PowerPoint รองรับและไฟล์ไม่เสีย
def timing_xml(anim_ids):
    if not anim_ids: return ""
    effects=""
    for i,sid in enumerate(anim_ids):
        delay = 0 if i==0 else 300  # ชิ้นแรกเริ่มเมื่อเข้าสไลด์ ชิ้นถัดไปหน่วง 0.3s หลังชิ้นก่อน
        effects+=(
            '<p:par><p:cTn presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="afterEffect">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:to><p:strVal val="visible"/></p:to></p:set>'
            '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn dur="500"/>'
            f'<p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
            '</p:childTnLst></p:cTn></p:par>'
        )
    bld=''.join(f'<p:bldP spid="{sid}" grpId="0"/>' for sid in anim_ids)
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst>{effects}</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'<p:bldLst>{bld}</p:bldLst></p:timing>'
    )

def slide(shapes,bg=NAVY2,trans="fade",anim_ids=None):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            f'<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
            f'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'{"".join(shapes)}</p:spTree></p:cSld>'
            f'{transition_xml(trans)}'
            f'{timing_xml(anim_ids or [])}'
            f'<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
            f'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
            f'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>')

# ---- Speaker Notes slide ----
def notes_xml(text):
    paras=""
    for line in text.split("\n"):
        paras+=(f'<a:p><a:r><a:rPr lang="th-TH" sz="1400" dirty="0"><a:latin typeface="Tahoma"/><a:cs typeface="Tahoma"/></a:rPr>'
                f'<a:t>{esc(line)}</a:t></a:r></a:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            f'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="5943600" cy="8229600"/></a:xfrm></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>'
            f'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>')

def starfield(n=36,seed=7):
    random.seed(seed); s=[]
    for _ in range(n):
        x=random.randint(0,EMU_W); y=random.randint(0,EMU_H); r=random.randint(12000,40000)
        s.append(oval(x,y,r,r,GREY))
    return s

def header(kicker,title,page,total,accent=GOLD):
    return [rrect(457200,470000,150000,920000,accent,r=50000),
            tbox(760000,450000,8800000,380000,[[(kicker,13,CYAN,True,False)]]),
            tbox(760000,790000,8900000,900000,[[(title,25,WHITE,True,False)]]),
            tbox(11000000,6360000,1000000,330000,[[(f"{page}/{total}",10,GREY,False,False)]],align="r")]

def bullets(x,y,w,items,size=13,gap=540000,color=WHITE,mark="▸",mc=GOLD):
    return [tbox(x,y+i*gap,w,gap,[[(f"{mark} ",12,mc,True,False),(it,size,color,False,False)]]) for i,it in enumerate(items)]

import re as _re
def first_id(shape_str):
    m=_re.search(r'<p:cNvPr id="(\d+)"', shape_str)
    return m.group(1) if m else None

slides=[]; notes=[]
_trans_cycle=["fade","push","fade","wipe","fade","push","fade"]
def add(sp, note="", anim_ids=None):
    idx=len(slides)
    tr=_trans_cycle[idx % len(_trans_cycle)]
    slides.append(slide(sp, trans=tr, anim_ids=anim_ids)); notes.append(note)

TOTAL=7

# ====== 1 · INTRO (0:00-0:30) ======
sp=starfield(44)
sp+=[rrect(0,2350000,EMU_W,2050000,None,grad=(NAVY,NAVY3))]
sp+=[nexus_logo(820000,2620000,1450000)]
sp+=[tbox(2600000,2540000,8800000,560000,[[("Demo Video · ไม่เกิน 5 นาที",14,GOLD,True,False)]]),
     tbox(2600000,2960000,8800000,780000,[[("LIKA Nexus",40,WHITE,True,False)]]),
     tbox(2600000,3780000,8800000,480000,[[("ผู้ช่วยทบทวนวรรณกรรมและจัดการความรู้งานวิจัย · ห้องสมุดวิทยาเขตกาญจนบุรี",13,CYAN,False,False)]])]
# เส้นเรื่อง 5 ช่วง
flow=[("Problem",RED),("Solution",GOLD),("AI/Kiro",CYAN),("Prototype",EMER),("Impact",VIOLET)]
cx=1500000
for i,(t,c) in enumerate(flow):
    w=len(t)*95000+360000
    sp.append(chip(cx,4700000,w,t,c)); 
    if i<len(flow)-1: sp.append(tbox(cx+w,4660000,300000,300000,[[("→",14,GREY,True,False)]],align="ctr"))
    cx+=w+330000
sp.append(tbox(820000,5900000,10600000,340000,[[("Demo: ",10,GOLD,True,False),("https://tinyurl.com/lika-nexus",10,WHITE,False,False)]]))
add(sp, "ช่วงที่ 1 · เปิดตัว (0:00-0:30)\n\nบทพากย์:\n\"สวัสดีค่ะ ขอแนะนำ LIKA Nexus ระบบผู้ช่วยทบทวนวรรณกรรมและจัดการความรู้งานวิจัย สำหรับห้องสมุดวิทยาเขตกาญจนบุรี มหาวิทยาลัยมหิดล ในโครงการ Mahidol AI Community 2026 ค่ะ\"\n\nภาพ: สไลด์ปก")

# ====== 2 · PROBLEM (0:30-1:15) ======
sp=header("① PROBLEM — ปัญหา","การทบทวนวรรณกรรมใช้เวลานานมาก",2,TOTAL,accent=RED)
sp.append(timebadge("0:30–1:15",RED))
ic2=icon_badge(9850000,2300000,1300000,"⏳",bg=NAVY3,ring=RED); sp.append(ic2)
blist=bullets(760000,2300000,8700000,[
    "ต้องอ่านงานวิจัยจำนวนมากทีละฉบับ ใช้เวลาหลายวัน–หลายสัปดาห์",
    "สายวิทยาศาสตร์/การแพทย์มีเปเปอร์มหาศาลและเพิ่มขึ้นต่อเนื่อง",
    "ต้องสกัดประเด็น เปรียบเทียบระเบียบวิธี และเขียนเองทั้งหมด",
    "เครื่องมือที่มีมักเสียเงิน ซับซ้อน หรือเสี่ยงต่อความลับข้อมูล",
], size=13.5, gap=720000)
sp+=blist
anim2=[first_id(ic2)]+[first_id(b) for b in blist]
add(sp, anim_ids=anim2, note="ช่วงที่ 2 · Problem (0:30-1:15)\n\nบทพากย์:\n\"ปัญหาที่พบคือ การทำ Literature Review ใช้เวลานานมาก โดยเฉพาะสายวิทยาศาสตร์และการแพทย์ที่มีงานวิจัยจำนวนมาก ผู้วิจัยต้องอ่านทีละฉบับ สกัดประเด็น เปรียบเทียบระเบียบวิธี และเขียนเองทั้งหมด ขณะที่เครื่องมือที่มีอยู่มักเสียเงิน ซับซ้อน หรือต้องอัปโหลดเอกสารขึ้นระบบภายนอก เสี่ยงต่อความลับของงานวิจัยที่ยังไม่เผยแพร่ค่ะ\"")

# ====== 3 · SOLUTION (1:15-2:00) ======
sp=header("② SOLUTION — แนวคิด","เว็บผู้ช่วยทบทวนวรรณกรรมอัตโนมัติ",3,TOTAL,accent=GOLD)
sp.append(timebadge("1:15–2:00",GOLD))
feats=[("📄","รับไฟล์หลายฉบับ","อัปโหลด PDF/TXT แล้วสกัดอัตโนมัติ",GOLD),
       ("📊","เปรียบเทียบ","ตารางระเบียบวิธี 5 ประเด็น",CYAN),
       ("🕸️","กราฟ 3 มิติ","เชื่อมโยงงานวิจัย + ค้นหา",EMER),
       ("✍️","ร่าง + APA","บททบทวนพร้อมอ้างอิง แก้ไขได้",VIOLET)]
anim3=[]
for i,(ic,t,d,c) in enumerate(feats):
    cx=760000+(i%2)*5550000; cy=2150000+(i//2)*1500000
    card=rrect(cx,cy,5300000,1320000,NAVY3,line=c,lw=9525,r=14000,shadow=True); sp.append(card)
    anim3.append(first_id(card))
    sp.append(icon_badge(cx+180000,cy+230000,860000,ic,bg=NAVY,ring=c))
    sp.append(tbox(cx+1200000,cy+230000,3950000,420000,[[(t,14,c,True,False)]]))
    sp.append(tbox(cx+1200000,cy+680000,3950000,560000,[[(d,11.5,WHITE,False,False)]]))
add(sp, anim_ids=anim3, note="ช่วงที่ 3 · Solution (1:15-2:00)\n\nบทพากย์:\n\"LIKA Nexus แก้ปัญหานี้ด้วยการรวมงานวิจัยหลายฉบับไว้ในระบบเดียว เพียงอัปโหลดไฟล์ ระบบจะสรุปให้ 5 ประเด็น คือ วัตถุประสงค์ กลุ่มตัวอย่าง ระเบียบวิธี ผลการศึกษา และข้อจำกัด แล้วแสดงผลเป็นกราฟความรู้ 3 มิติ ตารางเปรียบเทียบ และร่างบททบทวนวรรณกรรมพร้อมอ้างอิง จุดเด่นคือประมวลผลในเบราว์เซอร์ ไม่ส่งไฟล์ออกนอกเครื่องค่ะ\"")

# ====== 4 · AI/KIRO (2:00-2:40) ======
sp=header("③ AI / KIRO","เทคโนโลยีเบื้องหลัง",4,TOTAL,accent=CYAN)
sp.append(timebadge("2:00–2:40",CYAN))
card_ai=rrect(760000,2050000,5250000,3300000,NAVY3,line=CYAN,lw=12700,shadow=True); sp.append(card_ai)
sp.append(icon_badge(760000+2150000,2220000,860000,"🤖",bg=NAVY,ring=CYAN))
sp.append(tbox(960000,3130000,4900000,400000,[[("AI",14,CYAN,True,False)]],align="ctr"))
sp+=bullets(1000000,3560000,4800000,[
    "วิเคราะห์ตามโครงสร้าง IMRaD",
    "จำแนก 22 สาขา + ระเบียบวิธี",
    "รองรับไทย–อังกฤษ",
], size=12, gap=540000)
card_kiro=rrect(6260000,2050000,5200000,3300000,NAVY3,line=GOLD,lw=12700,shadow=True); sp.append(card_kiro)
sp.append(icon_badge(6260000+2170000,2220000,860000,"⚙️",bg=NAVY,ring=GOLD))
sp.append(tbox(6460000,3130000,4800000,400000,[[("Kiro (Spec-driven)",14,GOLD,True,False)]],align="ctr"))
sp+=bullets(6500000,3560000,4800000,[
    "Steering Rules — มาตรฐานการพัฒนา",
    "Specs — ข้อกำหนด",
    "Kiro เขียนโค้ด ทดสอบ และ deploy",
], size=12, gap=540000)
add(sp, anim_ids=[first_id(card_ai),first_id(card_kiro)], note="ช่วงที่ 4 · AI / Kiro (2:00-2:40)\n\nบทพากย์:\n\"ด้าน AI ระบบวิเคราะห์เอกสารตามโครงสร้างงานวิจัย จำแนกได้ 22 สาขาและระเบียบวิธี รองรับทั้งภาษาไทยและอังกฤษ ส่วนการพัฒนาเราใช้ Kiro เป็นเครื่องมือหลักแบบ Spec-driven เริ่มจากกำหนดกฎใน Steering ออกแบบข้อกำหนดใน Specs แล้วให้ Kiro ช่วยเขียนโค้ด ทดสอบ และนำขึ้นใช้งานจริง ตรวจสอบย้อนกลับได้ค่ะ\"")

# ====== 5 · PROTOTYPE (2:40-4:15) ======
sp=header("④ PROTOTYPE — สาธิตสด","อัดหน้าจอเว็บจริง (ช่วงสำคัญที่สุด)",5,TOTAL,accent=EMER)
sp.append(timebadge("2:40–4:15",EMER))
# mockup
mx,my,mw,mh=760000,2050000,4700000,3250000
sp.append(rrect(mx,my,mw,mh,NAVY3,line=CYAN,lw=9525,r=6000,shadow=True))
sp.append(rrect(mx,my,mw,380000,NAVY,r=6000))
for k,cc in enumerate([PINK,GOLD,EMER]): sp.append(oval(mx+150000+k*220000,my+120000,120000,120000,cc))
sp.append(tbox(mx+850000,my+70000,3000000,300000,[[("LIKA Nexus",10,WHITE,True,False)]]))
cxn=mx+mw*0.56; cyn=my+mh*0.56
sp.append(oval(int(cxn-160000),int(cyn-160000),320000,320000,GOLD,shadow=True))
for a,col in [(0,CYAN),(72,EMER),(144,VIOLET),(216,ORANGE),(288,PINK)]:
    ex=int(cxn+900000*math.cos(math.radians(a))); ey=int(cyn+900000*math.sin(math.radians(a)))
    sp.append(lineshape(int(cxn),int(cyn),ex,ey,"3b82f6")); sp.append(oval(ex-130000,ey-130000,260000,260000,col,shadow=True))
# ลำดับสาธิต
sp.append(tbox(5750000,2000000,5700000,360000,[[("ลำดับการสาธิต (อัดหน้าจอ):",13,EMER,True,False)]]))
steps=["1. เข้าสู่ระบบ — เลือกประเภทผู้ใช้","2. อัปโหลดไฟล์ / กดโหลดตัวอย่าง 6 สาขา",
       "3. กราฟ 3 มิติ — หมุน + คลิกโหนด","4. ตารางเปรียบเทียบ 5 ประเด็น",
       "5. ร่าง Lit-Review + ส่งออก APA","6. Usage Dashboard — สถิติผู้ใช้"]
sp+=bullets(5750000,2450000,5700000,steps,size=12,gap=460000,mc=EMER)
add(sp, "ช่วงที่ 5 · Prototype สาธิตสด (2:40-4:15) — ช่วงสำคัญที่สุด ให้เวลามากที่สุด\n\nอัดหน้าจอเว็บจริงตามลำดับ:\n1) เปิดเว็บ เลือกประเภทผู้ใช้ แล้วเข้าสู่ระบบ\n2) ไปแท็บ AI Auto-Cataloger กดโหลดตัวอย่าง 6 สาขา (หรือลากไฟล์ PDF จริง)\n3) ดูกราฟความรู้ 3 มิติ หมุนและคลิกโหนดเพื่อดูรายละเอียดทางขวา\n4) เลื่อนดูตารางเปรียบเทียบระเบียบวิธี 5 ประเด็น\n5) ดูร่างบททบทวนวรรณกรรม แล้วกดส่งออก APA\n6) เปิดแท็บ Usage Dashboard โชว์สถิติผู้ใช้\n\nบทพากย์ (ย่อ): \"มาดูการทำงานจริงกันค่ะ เลือกประเภทผู้ใช้ อัปโหลดงานวิจัย ระบบสร้างกราฟความรู้ ตารางเปรียบเทียบ และร่างบททบทวนให้อัตโนมัติ พร้อมส่งออก APA และดูสถิติการใช้บริการได้ค่ะ\"")

# ====== 6 · EXPECTED IMPACT (4:15-4:50) ======
sp=header("⑤ EXPECTED IMPACT","ผลลัพธ์ที่คาดหวัง",6,TOTAL,accent=VIOLET)
sp.append(timebadge("4:15–4:50",VIOLET))
stats=[("หลายวัน→นาที","ลดเวลาทบทวนวรรณกรรม",GOLD),
       ("22 สาขา","วิทย์ · สังคม · มนุษยศาสตร์",CYAN),
       ("Privacy-first","ประมวลผลในเครื่อง ปลอดภัย",EMER),
       ("ต่อยอดได้","สืบค้นออนไลน์ · เชื่อมระบบห้องสมุด",VIOLET)]
anim6=[]
for i,(n,d,c) in enumerate(stats):
    cx=760000+(i%2)*5550000; cy=2150000+(i//2)*1450000
    card=rrect(cx,cy,5300000,1280000,NAVY3,line=c,lw=9525,r=14000,shadow=True); sp.append(card)
    anim6.append(first_id(card))
    sp.append(tbox(cx+250000,cy+220000,4800000,560000,[[(n,19,c,True,False)]]))
    sp.append(tbox(cx+250000,cy+820000,4800000,420000,[[(d,12,WHITE,False,False)]]))
add(sp, anim_ids=anim6, note="ช่วงที่ 6 · Expected Impact (4:15-4:50)\n\nบทพากย์:\n\"ผลลัพธ์ที่คาดหวังคือ ลดเวลาการทบทวนวรรณกรรมจากหลายวันเหลือไม่กี่นาที ช่วยงานวิจัยของอาจารย์ นักศึกษา และงานบริการของบรรณารักษ์ ระบบใช้ฟรี ไม่ต้องติดตั้ง ปลอดภัยด้วยการประมวลผลในเครื่อง รองรับ 22 สาขา และต่อยอดได้ เช่น การสืบค้นออนไลน์และเชื่อมระบบห้องสมุดในอนาคตค่ะ\"")

# ====== 7 · CLOSING (4:50-5:00) ======
sp=starfield(34,seed=11)
sp+=[rrect(0,2250000,EMU_W,1550000,None,grad=(NAVY,NAVY3))]
sp+=[nexus_logo(4650000,1250000,900000)]
sp+=[tbox(760000,2420000,10800000,760000,[[("ขอบคุณค่ะ / Thank You",30,WHITE,True,False)]],align="ctr")]
sp+=[tbox(760000,3250000,10800000,440000,[[("LIKA Nexus · ห้องสมุดวิทยาเขตกาญจนบุรี มหาวิทยาลัยมหิดล",13,CYAN,False,False)]],align="ctr")]
sp.append(tbox(760000,4600000,10800000,360000,[[("🌐 ",12,GOLD,False,False),("https://tinyurl.com/lika-nexus",12,WHITE,False,False)]],align="ctr"))
sp.append(tbox(760000,5050000,10800000,360000,[[("💻 ",12,GOLD,False,False),("github.com/kanchanathsao-crypto/lika-lit-review-helper-",11,WHITE,False,False)]],align="ctr"))
add(sp, "ช่วงที่ 7 · ปิดท้าย (4:50-5:00)\n\nบทพากย์:\n\"ทั้งหมดนี้คือ LIKA Nexus ผู้ช่วยทบทวนวรรณกรรมสำหรับห้องสมุดวิทยาเขตกาญจนบุรี ขอบคุณค่ะ\"")

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
'<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
'<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
+''.join(f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(N))
+''.join(f'<Override PartName="/ppt/notesSlides/notesSlide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' for i in range(N))
+'</Types>')
root_rels=rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>')
sldid=''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(N))
presentation=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
f'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
f'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
f'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
f'<p:notesMasterIdLst><p:notesMasterId r:id="rIdNM"/></p:notesMasterIdLst>'
f'<p:sldIdLst>{sldid}</p:sldIdLst><p:sldSz cx="{EMU_W}" cy="{EMU_H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
pres_rels=rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
'<Relationship Id="rIdNM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>'
'<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
+''.join(f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>' for i in range(N)))

def theme(name):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{name}">'
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

# notes master (จำเป็นเมื่อมี notesSlide)
notes_master=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
'<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
'<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
'<p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="5943600" cy="8229600"/></a:xfrm>'
'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="th-TH"/></a:p></p:txBody></p:sp>'
'</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:notesMaster>')
notes_master_rels=rels_xml('<Relationship Id="rIdT2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml"/>')

def slide_rels(i):
    return rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{i+1}.xml"/>')
def notes_rels(i):
    return rels_xml('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>'
    f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{i+1}.xml"/>')

with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml",content_types)
    z.writestr("_rels/.rels",root_rels)
    z.writestr("ppt/presentation.xml",presentation)
    z.writestr("ppt/_rels/presentation.xml.rels",pres_rels)
    z.writestr("ppt/theme/theme1.xml",theme("LIKA"))
    z.writestr("ppt/theme/theme2.xml",theme("LIKA-Notes"))
    z.writestr("ppt/slideMasters/slideMaster1.xml",slide_master)
    z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",master_rels)
    z.writestr("ppt/slideLayouts/slideLayout1.xml",slide_layout)
    z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",layout_rels)
    z.writestr("ppt/notesMasters/notesMaster1.xml",notes_master)
    z.writestr("ppt/notesMasters/_rels/notesMaster1.xml.rels",notes_master_rels)
    for i,sx in enumerate(slides):
        z.writestr(f"ppt/slides/slide{i+1}.xml",sx)
        z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels",slide_rels(i))
        z.writestr(f"ppt/notesSlides/notesSlide{i+1}.xml",notes_xml(notes[i]))
        z.writestr(f"ppt/notesSlides/_rels/notesSlide{i+1}.xml.rels",notes_rels(i))
print(f"สร้างสำเร็จ: {OUT} ({N} สไลด์ + speaker notes)")

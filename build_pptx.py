#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้างไฟล์ PowerPoint (.pptx) สำหรับ LIKA Research & Lit-Review Helper
โดยเขียน Office Open XML ตรง ๆ แล้วแพ็กเป็น ZIP (ไม่ต้องใช้ไลบรารีภายนอก)
ธีม: Mahidol Royal Blue #002366 + Gold #F5C542 + Dark Slate #0B1120
"""
import zipfile, html

OUT = "LIKA_Pitch_Deck.pptx"
EMU_W, EMU_H = 12192000, 6858000  # 16:9 (13.333" x 7.5")

# ---------- สี ----------
NAVY   = "002366"
NAVY2  = "0B1120"
GOLD   = "F5C542"
CYAN   = "22D3EE"
WHITE  = "E5EDFF"
GREY   = "9FB3D1"

def esc(s): return html.escape(str(s), quote=True)

# ---------- ตัวช่วยสร้าง shape ข้อความ ----------
def tbox(sid, name, x, y, w, h, runs, anchor="t", align="l"):
    """runs = list ของ (text, size_pt, color, bold, italic)"""
    paras = ""
    for run_group in runs:
        rtxt = ""
        for (text, size, color, bold, italic) in run_group:
            rtxt += (
                f'<a:r><a:rPr lang="th-TH" sz="{int(size*100)}" b="{1 if bold else 0}" '
                f'i="{1 if italic else 0}" dirty="0">'
                f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'<a:latin typeface="Tahoma"/><a:cs typeface="Tahoma"/></a:rPr>'
                f'<a:t>{esc(text)}</a:t></a:r>'
            )
        paras += f'<a:p><a:pPr algn="{align}"/>{rtxt}</a:p>'
    return f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="{anchor}"><a:normAutofit/></a:bodyPr>
<a:lstStyle/>{paras}</p:txBody></p:sp>'''

def rect(sid, name, x, y, w, h, fill, line=None, lw=12700):
    ln = (f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
          if line else '<a:ln><a:noFill/></a:ln>')
    return f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'''

def slide_xml(shapes, bg=NAVY2):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>
<a:effectLst/></p:bgPr></p:bg>
<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{''.join(shapes)}
</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping
 bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
 accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
 hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>'''

# ---------- helper: แถบหัวเรื่อง + เลขหน้า ----------
def header(sid_start, kicker, title, page_no, total=8):
    s = sid_start
    shapes = []
    # แถบทองซ้าย
    shapes.append(rect(s, "accentbar", 457200, 500000, 150000, 900000, GOLD)); s+=1
    shapes.append(tbox(s, "kicker", 750000, 470000, 9000000, 400000,
        [[(kicker, 14, CYAN, True, False)]])); s+=1
    shapes.append(tbox(s, "title", 750000, 820000, 10800000, 1000000,
        [[(title, 30, WHITE, True, False)]])); s+=1
    shapes.append(tbox(s, "pageno", 11200000, 6350000, 800000, 350000,
        [[(f"{page_no}/{total}", 11, GREY, False, False)]], align="r")); s+=1
    return shapes, s

def bullets(sid_start, x, y, w, items, size=15, gap=560000, color=WHITE, bullet="●"):
    s = sid_start; shapes=[]
    for i, it in enumerate(items):
        runs = [[(f"{bullet}  ", 13, GOLD, True, False), (it, size, color, False, False)]]
        shapes.append(tbox(s, f"b{i}", x, y+i*gap, w, gap, runs)); s+=1
    return shapes, s

# ======================================================================
# SLIDES
# ======================================================================
slides = []

# ---- Slide 1: Title ----
sp=[]; s=100
sp.append(rect(s,"band",0,2550000,EMU_W,1750000,NAVY)); s+=1
sp.append(rect(s,"logo",900000,2820000,1150000,1150000,NAVY,line=CYAN,lw=28000)); s+=1
sp.append(tbox(s,"logotx",900000,3150000,1150000,600000,[[("LK",40,GOLD,True,False)]],align="l")); s+=1
sp.append(tbox(s,"t1",2250000,2720000,9200000,900000,
    [[("LIKA Research & Lit-Review Helper",33,WHITE,True,False)]])); s+=1
sp.append(tbox(s,"t2",2250000,3600000,9200000,500000,
    [[("ผู้ช่วยทบทวนวรรณกรรมและงานวิจัย · ห้องสมุดวิทยาเขตกาญจนบุรี (MUKA)",15,CYAN,False,False)]])); s+=1
sp.append(tbox(s,"t3",900000,4650000,10800000,500000,
    [[("Mahidol University Kanchanaburi Campus",12,GREY,False,False)]])); s+=1
sp.append(tbox(s,"t4",900000,6180000,10800000,450000,
    [[("Demo: ",11,GOLD,True,False),("https://tinyurl.com/2u452bd4",11,WHITE,False,False)]])); s+=1
slides.append(slide_xml(sp))

# ---- Slide 2: Problem ----
sp, s = header(100, "① PROBLEM — ปัญหา", "ทำ Literature Review ใช้เวลานานมาก", 2)
sp.append(rect(s,"card",750000,1950000,10700000,3900000,NAVY,line=CYAN,lw=12000)); s+=1
b, s = bullets(s, 1050000, 2250000, 10100000, [
    "การทบทวนวรรณกรรมใช้เวลาหลายวัน–หลายสัปดาห์ ต้องอ่านทีละฉบับด้วยมือ",
    "สายการแพทย์/วิทยาศาสตร์มีจำนวนเปเปอร์มหาศาลและเพิ่มขึ้นทุกวัน",
    "ต้องสกัดประเด็น เปรียบเทียบระเบียบวิธี และร่างบทความเองทั้งหมด",
    "บรรณารักษ์รับคำขอช่วยสืบค้น/ทบทวนวรรณกรรมจำนวนมาก แต่เวลาจำกัด",
    "เครื่องมือที่มีมักเสียเงิน ซับซ้อน หรือส่งไฟล์ขึ้น cloud (เสี่ยงข้อมูลงานยังไม่ตีพิมพ์)",
], gap=640000)
sp += b
slides.append(slide_xml(sp))

# ---- Slide 3: Solution ----
sp, s = header(100, "② SOLUTION — แนวคิดและวัตถุประสงค์", "เว็บผู้ช่วยทบทวนวรรณกรรมอัตโนมัติ", 3)
feats = [
    ("📄 รับไฟล์หลายฉบับ", "อัปโหลด PDF/TXT ≥ 3 ฉบับ อ่านในเบราว์เซอร์"),
    ("🧠 สกัดประเด็นอัตโนมัติ", "วัตถุประสงค์ · กลุ่มตัวอย่าง · วิธี · ผล · ข้อจำกัด"),
    ("📊 ตารางเปรียบเทียบ", "Comparative Methodology Matrix 5 ประเด็น"),
    ("🕸️ กราฟความรู้ 3 มิติ", "เชื่อมโยงงานวิจัยกับระเบียบวิธี (Three.js)"),
    ("✍️ ร่าง Lit-Review", "ย่อหน้าพร้อม [ผู้แต่ง et al., ปี] แก้ไขได้"),
    ("⬇️ ส่งออก APA", "คัดลอก/ส่งออกรายการอ้างอิง APA 7th"),
]
col_w=5050000; x0=750000; y0=2000000
for i,(h_,d_) in enumerate(feats):
    cx = x0 + (i%2)*(col_w+560000)
    cy = y0 + (i//2)*1250000
    sp.append(rect(s,f"fc{i}",cx,cy,col_w,1100000,NAVY,line=CYAN,lw=10000)); s+=1
    sp.append(tbox(s,f"fh{i}",cx+200000,cy+130000,col_w-350000,450000,
        [[(h_,15,GOLD,True,False)]])); s+=1
    sp.append(tbox(s,f"fd{i}",cx+200000,cy+560000,col_w-350000,480000,
        [[(d_,12,WHITE,False,False)]])); s+=1
slides.append(slide_xml(sp))

# ---- Slide 4: Users ----
sp, s = header(100, "กลุ่มผู้ใช้ / ผู้ได้รับประโยชน์", "Target Users & Beneficiaries", 4)
users=[
    ("บรรณารักษ์ MUKA","ให้บริการทบทวนวรรณกรรมได้เร็วขึ้น"),
    ("อาจารย์ / นักวิจัย","ร่าง Lit-Review ตั้งต้น + เทียบระเบียบวิธีทันที"),
    ("นักศึกษา ป.ตรี/บัณฑิต","เรียนรู้โครงสร้างงานวิจัยและเริ่มเขียนได้ง่าย"),
    ("ครอบคลุม 22 สาขา","วิทยาศาสตร์ · สังคมศาสตร์ · มนุษยศาสตร์"),
]
y0=2050000
for i,(u,d) in enumerate(users):
    sp.append(rect(s,f"u{i}",750000,y0+i*920000,10700000,820000,NAVY,line=GOLD,lw=8000)); s+=1
    sp.append(tbox(s,f"ut{i}",1000000,y0+i*920000+120000,3800000,560000,[[(u,16,GOLD,True,False)]])); s+=1
    sp.append(tbox(s,f"ud{i}",5000000,y0+i*920000+150000,6300000,520000,[[(d,14,WHITE,False,False)]])); s+=1
slides.append(slide_xml(sp))

# ---- Slide 5: AI & Kiro ----
sp, s = header(100, "③ AI / KIRO — แนวทางการประยุกต์ใช้", "พัฒนาแบบ Spec-driven ด้วย Kiro", 5)
sp.append(rect(s,"kc1",750000,1980000,5200000,3700000,NAVY,line=CYAN,lw=12000)); s+=1
sp.append(tbox(s,"kh1",1000000,2120000,4800000,450000,[[("🤖 AI / Automated Reasoning",15,GOLD,True,False)]])); s+=1
b,s = bullets(s,1000000,2680000,4700000,[
    "Engine สกัดข้อมูลตามโครงสร้าง IMRaD",
    "จำแนกสาขา/ระเบียบวิธีด้วย weighted keyword scoring",
    "รองรับ ไทย–อังกฤษ แปลง พ.ศ.→ค.ศ. อัตโนมัติ",
], size=13, gap=720000); sp+=b
sp.append(rect(s,"kc2",6250000,1980000,5200000,3700000,NAVY,line=CYAN,lw=12000)); s+=1
sp.append(tbox(s,"kh2",6500000,2120000,4800000,450000,[[("⚙️ Kiro = เครื่องมือหลัก",15,GOLD,True,False)]])); s+=1
b,s = bullets(s,6500000,2680000,4700000,[
    ".kiro/steering/rules.md — มาตรฐานการพัฒนา",
    ".kiro/specs/ — Acceptance Criteria ชัดเจน",
    "Kiro ช่วยเขียนโค้ด ทดสอบ (Node.js) และ deploy",
], size=13, gap=720000); sp+=b
slides.append(slide_xml(sp))

# ---- Slide 6: Prototype / Tech ----
sp, s = header(100, "④ PROTOTYPE — ผลงานต้นแบบ (ใช้งานจริงแล้ว)", "Single-file Web App · Deploy บน GitHub Pages", 6)
sp.append(rect(s,"pc",750000,1980000,5200000,3700000,NAVY,line=GOLD,lw=12000)); s+=1
sp.append(tbox(s,"ph",1000000,2120000,4800000,450000,[[("🛠️ เทคโนโลยี",15,GOLD,True,False)]])); s+=1
b,s=bullets(s,1000000,2680000,4700000,[
    "HTML5 + Tailwind CSS (Glassmorphism)",
    "Three.js r128 + OrbitControls",
    "PDF.js — อ่าน PDF ในเบราว์เซอร์",
    "ธีม Mahidol Royal Blue #002366",
], size=13, gap=640000); sp+=b
sp.append(rect(s,"pc2",6250000,1980000,5200000,3700000,NAVY,line=GOLD,lw=12000)); s+=1
sp.append(tbox(s,"ph2",6500000,2120000,4800000,450000,[[("🔒 จุดเด่น",15,GOLD,True,False)]])); s+=1
b,s=bullets(s,6500000,2680000,4700000,[
    "Privacy-first: ไฟล์ไม่ออกจากเบราว์เซอร์",
    "ฟรี ไม่ต้องติดตั้ง · ใช้ได้แม้ออฟไลน์",
    "Spec-driven ตรวจสอบได้จาก GitHub",
    "Live: tinyurl.com/2u452bd4",
], size=13, gap=640000); sp+=b
slides.append(slide_xml(sp))

# ---- Slide 7: Expected Impact ----
sp, s = header(100, "⑤ EXPECTED IMPACT — ผลลัพธ์ที่คาดหวัง", "คุณค่าและโอกาสในการนำไปใช้จริง", 7)
stats=[
    ("หลายวัน → นาที","ลดเวลาทำ Lit-Review อย่างมาก"),
    ("≥ 3 ฉบับ","ได้ตารางเทียบ + ร่างทันที (ตาม Success Criteria)"),
    ("22 สาขา","ครอบคลุมทุกศาสตร์ รวมมนุษยศาสตร์"),
    ("ต่อยอดได้","สืบค้นออนไลน์ · OCR · เชื่อมระบบห้องสมุด"),
]
y0=2050000
for i,(n,d) in enumerate(stats):
    cx=750000+(i%2)*5550000; cy=y0+(i//2)*1650000
    sp.append(rect(s,f"s{i}",cx,cy,5150000,1480000,NAVY,line=CYAN,lw=10000)); s+=1
    sp.append(tbox(s,f"sn{i}",cx+250000,cy+180000,4650000,650000,[[(n,22,GOLD,True,False)]])); s+=1
    sp.append(tbox(s,f"sd{i}",cx+250000,cy+880000,4650000,500000,[[(d,13,WHITE,False,False)]])); s+=1
slides.append(slide_xml(sp))

# ---- Slide 8: Thank you ----
sp=[]; s=100
sp.append(rect(s,"band",0,2550000,EMU_W,1750000,NAVY)); s+=1
sp.append(tbox(s,"ty",900000,2720000,10800000,900000,[[("ขอบคุณครับ / Thank You",32,WHITE,True,False)]])); s+=1
sp.append(tbox(s,"ty2",900000,3600000,10800000,500000,[[("LIKA Research & Lit-Review Helper — MUKA Library",15,CYAN,False,False)]])); s+=1
sp.append(tbox(s,"ty3",900000,4700000,10800000,450000,[[("🌐 ลองใช้งานจริง: ",12,GOLD,True,False),("https://tinyurl.com/2u452bd4",12,WHITE,False,False)]])); s+=1
sp.append(tbox(s,"ty4",900000,5250000,10800000,450000,[[("💻 Source: ",12,GOLD,True,False),("github.com/kanchanathsao-crypto/lika-lit-review-helper-",12,WHITE,False,False)]])); s+=1
slides.append(slide_xml(sp))

# ======================================================================
# ประกอบไฟล์ .pptx (ZIP)
# ======================================================================
N = len(slides)

content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
'<Default Extension="xml" ContentType="application/xml"/>' \
'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' \
'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' \
'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' \
'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' \
+ ''.join(f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(N)) \
+ '</Types>'

root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' \
'</Relationships>'

sldid = ''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(N))
presentation = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>{sldid}</p:sldIdLst>
<p:sldSz cx="{EMU_W}" cy="{EMU_H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'''

pres_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' \
'<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' \
+ ''.join(f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>' for i in range(N)) \
+ '</Relationships>'

theme = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="LIKA">
<a:themeElements>
<a:clrScheme name="LIKA"><a:dk1><a:srgbClr val="0B1120"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="002366"/></a:dk2><a:lt2><a:srgbClr val="E5EDFF"/></a:lt2>
<a:accent1><a:srgbClr val="002366"/></a:accent1><a:accent2><a:srgbClr val="F5C542"/></a:accent2>
<a:accent3><a:srgbClr val="22D3EE"/></a:accent3><a:accent4><a:srgbClr val="1E3A8A"/></a:accent4>
<a:accent5><a:srgbClr val="7DD3FC"/></a:accent5><a:accent6><a:srgbClr val="9FB3D1"/></a:accent6>
<a:hlink><a:srgbClr val="22D3EE"/></a:hlink><a:folHlink><a:srgbClr val="7DD3FC"/></a:folHlink></a:clrScheme>
<a:fontScheme name="LIKA"><a:majorFont><a:latin typeface="Tahoma"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Tahoma"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="LIKA"><a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>'''

slide_master = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0B1120"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"
 accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
 hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'''

master_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' \
'<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' \
'</Relationships>'

slide_layout = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
 accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"
 accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>'''

layout_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' \
'</Relationships>'

def slide_rels():
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' \
    '</Relationships>'

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("ppt/presentation.xml", presentation)
    z.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
    z.writestr("ppt/theme/theme1.xml", theme)
    z.writestr("ppt/slideMasters/slideMaster1.xml", slide_master)
    z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
    z.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout)
    z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels)
    for i, sx in enumerate(slides):
        z.writestr(f"ppt/slides/slide{i+1}.xml", sx)
        z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", slide_rels())

print(f"สร้างสำเร็จ: {OUT} ({N} สไลด์)")

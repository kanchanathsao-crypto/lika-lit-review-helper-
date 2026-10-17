<div align="center">

# 🛰️ LIKA Nexus
### Research Knowledge Nexus — ระบบผู้ช่วยทบทวนวรรณกรรมและจัดการความรู้งานวิจัย

**ห้องสมุดวิทยาเขตกาญจนบุรี · มหาวิทยาลัยมหิดล (MUKA Library)**

เสนอต่อ: **Mahidol AI Community 2026** · หมวด **Research & Innovation**

[![Demo](https://img.shields.io/badge/🌐_Live_Demo-LIKA_Nexus-f5c542?style=for-the-badge)](https://kanchanathsao-crypto.github.io/lika-lit-review-helper-/lika-nexus.html)
[![Short link](https://img.shields.io/badge/🔗_Short-tinyurl.com%2F2u452bd4-22d3ee?style=for-the-badge)](https://tinyurl.com/2u452bd4)
[![Kiro](https://img.shields.io/badge/Built_with-Kiro_(Spec--driven)-002366?style=for-the-badge)](#-การประยุกต์ใช้-kiro)

</div>

---

## 📌 ภาพรวมโดยย่อ (Overview)

**LIKA Nexus** เป็นเว็บแอปพลิเคชันไฟล์เดียว (single-file) ที่นำไฟล์งานวิจัยหลายฉบับเข้ามา **สกัดประเด็นสำคัญโดยอัตโนมัติ** แล้วนำเสนอผ่าน **กราฟความรู้ 3 มิติ · ตารางเปรียบเทียบระเบียบวิธีวิจัย · ร่างบททบทวนวรรณกรรมพร้อมการอ้างอิง APA** เพื่อลดระยะเวลาการทำ Literature Review จากหลายวันเหลือเพียงไม่กี่นาที

> 🔐 **Privacy-first** — ประมวลผลไฟล์ทั้งหมดภายในเบราว์เซอร์ ไม่ส่งข้อมูลออกนอกเครื่อง เหมาะกับงานวิจัยที่ยังไม่เผยแพร่

---

## 🔗 ลิงก์สำหรับตรวจสอบ (สำหรับคณะกรรมการ)

| รายการ | ลิงก์ |
|---|---|
| 🌐 **ระบบใช้งานจริง (Live Demo)** | https://kanchanathsao-crypto.github.io/lika-lit-review-helper-/lika-nexus.html |
| 🔗 **ลิงก์สั้น** | https://tinyurl.com/2u452bd4 |
| 💻 **Source Code (GitHub)** | https://github.com/kanchanathsao-crypto/lika-lit-review-helper- |
| 📐 **Kiro Project** | โฟลเดอร์ [`.kiro/`](./.kiro) (steering + specs) |

---

## 🗂️ โครงสร้างของผลงาน (Project Structure)

```
LIKA.kiro/
├── lika-nexus.html              ⭐ แอปพลิเคชันหลัก (single-file: HTML + CSS + JS)
├── index.html                   หน้าเริ่มต้น — redirect เข้า lika-nexus.html
│
├── .kiro/                       📐 Kiro Project (หลักฐานการพัฒนาแบบ Spec-driven)
│   ├── steering/
│   │   └── rules.md             กฎ/มาตรฐานการพัฒนา (เทคโนโลยี ธีม การอ้างอิง)
│   └── specs/
│       └── lit-review-spec.md   ข้อกำหนดคุณลักษณะ + เกณฑ์การยอมรับผลงาน
│
├── LIKA_Concept_Paper.pptx      เอกสาร Concept Paper (นำเสนอ)
├── LIKA_Pitch_Deck.pptx         สไลด์นำเสนอ/วิดีโอสาธิต
├── build_concept_paper.py       สคริปต์สร้าง Concept Paper deck
├── build_pptx.py                สคริปต์สร้าง Pitch deck
│
└── _archive-*.html              เวอร์ชันพัฒนาเดิม (เก็บไว้เพื่อการตรวจสอบย้อนหลัง)
```

---

## 🛠️ แนวทางการพัฒนา (Development Approach)

ระบบพัฒนาเป็น **Single-file Web Application** เพื่อให้ติดตั้งและใช้งานง่าย ไม่ต้องมีเซิร์ฟเวอร์:

| ด้าน | รายละเอียด |
|---|---|
| **โครงสร้าง** | HTML5 + CSS + JavaScript ในไฟล์เดียว (`lika-nexus.html`) |
| **UI Framework** | Tailwind CSS (CDN) · ธีม Mahidol Royal Blue `#002366` + Glassmorphism |
| **กราฟ 3 มิติ** | Three.js r128 + OrbitControls (CDN) |
| **อ่านไฟล์ PDF** | PDF.js (CDN) — แยกข้อความในเบราว์เซอร์ |
| **การประมวลผล** | ทั้งหมดฝั่งผู้ใช้ (client-side) — ไม่มี backend, ไม่ส่งไฟล์ออก |
| **การจัดเก็บสถิติ** | `localStorage` (แยกตามประเภทผู้ใช้) |
| **Deploy** | GitHub Pages |

**กลไกการสกัดข้อมูล (AI / Rule-based Engine):**
- วิเคราะห์ข้อความตามโครงสร้างงานวิจัย (IMRaD)
- จำแนก **22 สาขาวิชา** และ **ระเบียบวิธีวิจัย** ด้วยการให้น้ำหนักคำสำคัญ (weighted keyword scoring)
- รองรับเอกสารทั้งภาษาไทยและอังกฤษ (แปลงปี พ.ศ. → ค.ศ. และดึงชื่อผู้แต่งแบบเต็ม โดยหลีกเลี่ยงชื่อสถาบัน)

---

## 🤖 การประยุกต์ใช้ Kiro

ผลงานนี้พัฒนาด้วย **Kiro** แบบ **Spec-driven Development** ตลอดวงจรการพัฒนา — ตรวจสอบย้อนกลับได้จาก `git log` (14 commits)

| ขั้นตอน | ไฟล์ / หลักฐาน |
|---|---|
| **1. Steering Rules** — กำหนดมาตรฐานการพัฒนา | [`.kiro/steering/rules.md`](./.kiro/steering/rules.md) |
| **2. Specs** — กำหนดข้อกำหนดและเกณฑ์การยอมรับผลงาน | [`.kiro/specs/lit-review-spec.md`](./.kiro/specs/lit-review-spec.md) |
| **3. Agent พัฒนา** — Kiro เขียนโค้ด + ทดสอบตามข้อกำหนด | `git log` (commit history) |
| **4. Deploy** — เผยแพร่และตรวจสอบได้จริง | GitHub Pages |

> **Steering Rules** กำหนด: single-file architecture, การใช้ Tailwind + Three.js r128, ธีม Mahidol Royal Blue, รูปแบบการอ้างอิง `[ผู้แต่ง et al., ปี]`
> **Specs** กำหนด Acceptance Criteria 4 ข้อ: Multi-Paper Ingestion · 3D Knowledge Graph · Comparative Methodology Matrix · Editable Lit-Review Draft

---

## ✨ คุณสมบัติหลัก (Features)

- 📄 **นำเข้าไฟล์งานวิจัยหลายฉบับ** (PDF / TXT) — ลากวางหรือเลือกไฟล์
- 🧠 **สกัดประเด็นสำคัญอัตโนมัติ** — วัตถุประสงค์ · กลุ่มตัวอย่าง · ระเบียบวิธี · ผลการศึกษา · ข้อจำกัด
- 🕸️ **กราฟความรู้ 3 มิติ** — เชื่อมโยงงานวิจัยกับระเบียบวิธี + ค้นหาเชิงแนวคิด (Semantic Search)
- 📊 **ตารางเปรียบเทียบระเบียบวิธีวิจัย** — 5 ประเด็น
- ✍️ **ร่างบททบทวนวรรณกรรม** — สังเคราะห์อัตโนมัติ + อ้างอิงแทรก + ส่งออก APA + แก้ไขได้
- 🔑 **ระบบเข้าใช้แยกประเภท** — อาจารย์ · เจ้าหน้าที่ · นักศึกษา
- 📈 **แดชบอร์ดสถิติการใช้บริการ** — การ์ดสถิติ · กราฟโดนัท · กราฟ 7 วัน · บันทึกล่าสุด

---

## ▶️ วิธีใช้งาน (Quick Start)

1. เปิด [ลิงก์ Live Demo](https://kanchanathsao-crypto.github.io/lika-lit-review-helper-/lika-nexus.html) ในเบราว์เซอร์ (Chrome / Edge / Firefox)
2. เข้าสู่ระบบโดยเลือกประเภทผู้ใช้ (หรือกดข้าม)
3. ไปที่แท็บ **AI Auto-Cataloger** → อัปโหลดไฟล์งานวิจัย **≥ 3 ฉบับ** (หรือกด *โหลดตัวอย่าง 6 สาขา*)
4. ดูผลที่ **3D Knowledge Explorer** (กราฟ) · ตารางเปรียบเทียบ · ร่าง Lit-Review
5. คัดลอกร่าง / ส่งออกการอ้างอิง APA / ดูสถิติที่ **Usage Dashboard**

> ต้องเชื่อมต่ออินเทอร์เน็ตเมื่อเปิดครั้งแรก เพื่อโหลด Tailwind / Three.js / PDF.js จาก CDN

---

## 🧰 เทคโนโลยีที่ใช้ (Tech Stack)

`HTML5` · `Tailwind CSS` · `JavaScript (ES6)` · `Three.js r128` · `OrbitControls` · `PDF.js` · `GitHub Pages` · `Kiro`

---

<div align="center">

**LIKA Nexus** · Research Knowledge Nexus
Mahidol University Kanchanaburi Campus Library · 2026
พัฒนาด้วย 🤖 **Kiro** (Spec-driven Development)

</div>

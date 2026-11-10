# UTSL Forms (เว็บแอปใหม่) — คู่มือติดตั้ง

เปิดได้ทุกเครื่อง (Windows / Mac / iPad / Android / iPhone) ผ่านเบราว์เซอร์ เข้าสู่ระบบด้วยบัญชี Microsoft ของบริษัท (@utsl.co.th)

**การทำงาน**
1. เข้าสู่ระบบด้วย Microsoft
2. **เลือกเลขที่ใบสั่งงาน** — แอปอ่านทะเบียนงาน (`ทะเบียนงาน_IND & PP <ปี>.xlsx` ชีตเดือนปัจจุบัน) แสดงงานของวันนี้ตามช่วงวันในทะเบียน เปลี่ยนวัน/ค้นหาได้ หรือพิมพ์เลขเอง
3. เลือกแบบฟอร์ม (5 แบบ) — **เลขที่ใบสั่งงานและชื่อลูกค้าถูกเติมให้อัตโนมัติ** (ช่องสีเขียวอ่อน แก้ได้)
4. กดบันทึก → สร้าง PDF จากแม่แบบเดิม → บันทึกลง SharePoint `Service/X-Rite/X-Rite_<ปี>/<ลูกค้า>/<เลขที่ใบสั่งงาน>` (สร้างโฟลเดอร์ให้ถ้ายังไม่มี) → ส่งอีเมลจากบัญชีของผู้กรอกถึงตัวเอง + ลูกค้า → ดาวน์โหลด PDF ได้

ผู้ใช้แต่ละคนใช้สิทธิ์ SharePoint/อีเมลของตัวเอง (ไม่มีรหัสผ่านหรือ secret เก็บในแอป)

## ไฟล์ในแอป
```
index.html        หน้าแอป
app.js            โปรแกรม
config.js         ← ตั้งค่า (ใส่ CLIENT_ID ที่นี่)
lib/pdf-lib.min.js
templates/        แม่แบบ PDF 5 ไฟล์ (เปลี่ยนแม่แบบได้ ชื่อไฟล์ต้องเหมือนเดิม)
```
ยังไม่ใส่ `CLIENT_ID` = **โหมดทดลอง**: ใช้รายการงานตัวอย่าง สร้าง/ดาวน์โหลด PDF ได้ แต่ไม่อ่านทะเบียนจริง ไม่บันทึก SharePoint และไม่ส่งอีเมล

---

## ขั้นที่ 1 — วางเว็บ (เลือกอย่างใดอย่างหนึ่ง ต้องเป็น HTTPS)

**ก. Azure Static Web Apps (แนะนำ ฟรี อยู่ใน Microsoft ของบริษัท)** — ฝ่าย IT
1. https://portal.azure.com › Create a resource › **Static Web App** › Plan: Free › Deployment source: **Other**
2. ติดตั้ง SWA CLI แล้วอัปโหลดโฟลเดอร์แอป: `npx @azure/static-web-apps-cli deploy ./utsl-app --deployment-token <token จากหน้า Manage deployment token> --env production`
3. ได้ที่อยู่เว็บ เช่น `https://xxxx.azurestaticapps.net/`

**ข. GitHub Pages (ฟรี)** — สร้าง repository › Upload files (ทั้งโฟลเดอร์) › Settings › Pages › Branch: main › ได้ที่อยู่ `https://<ชื่อ>.github.io/<repo>/`
(โค้ดไม่มีข้อมูลลับ ข้อมูลทั้งหมดอยู่หลังการเข้าสู่ระบบ Microsoft)

## ขั้นที่ 2 — ลงทะเบียนแอปใน Microsoft Entra ID (ฝ่าย IT ~10 นาที)
1. https://entra.microsoft.com › Applications › **App registrations › New registration**
   - Name: `UTSL Forms` · Accounts in this organizational directory only
   - Redirect URI: แพลตฟอร์ม **Single-page application (SPA)** · ใส่ที่อยู่เว็บจากขั้นที่ 1 เช่น `https://xxxx.azurestaticapps.net/`
     (ถ้าเปิดแบบมี `index.html` ท้าย URL ให้เพิ่มอีกบรรทัด `https://xxxx.azurestaticapps.net/index.html`)
2. **API permissions › Add a permission › Microsoft Graph › Delegated permissions** เพิ่ม
   `User.Read` · `Files.ReadWrite.All` · `Sites.ReadWrite.All` · `Mail.Send`
   แล้วกด **Grant admin consent for UNION TSL LIMITED** (ผู้ใช้จะไม่ต้องกดยินยอมเอง)
3. คัดลอก **Application (client) ID** — ไม่ต้องสร้าง Client secret

## ขั้นที่ 3 — ใส่ค่าใน config.js แล้วอัปโหลดใหม่
```js
CLIENT_ID: "<Application (client) ID>",
```
ค่าอื่นตั้งไว้ให้แล้ว (tenant, site Industrial, ไลบรารี General Ind / Service, แบรนด์ XR)

## ขั้นที่ 4 — ใช้งาน
- ส่งลิงก์ให้ทีม · บน iPad/iPhone: Safari › แชร์ › **Add to Home Screen** · Android: Chrome › ⋮ › **Add to Home screen**
- ทดสอบ: เข้าสู่ระบบ › ต้องเห็นงานของวันนี้ตรงกับทะเบียนงาน › กรอก Service Report › ตรวจว่ามีไฟล์ในโฟลเดอร์ใบสั่งงานและได้รับอีเมล

---

## ตั้งค่าเพิ่ม (config.js)
| ต้องการ | แก้ |
|---|---|
| เพิ่มแบรนด์ เช่น S | `BRANDS: { XR: {...}, S: { root: "<โฟลเดอร์แบรนด์>", year: "<ชื่อโฟลเดอร์ปี {YYYY}>" } }` |
| ทะเบียนงานย้ายที่ | `REGISTER.library` / `REGISTER.path` |
| แถวงานเกิน 900 แถว/เดือน | `REGISTER.range: "A1:Q1500"` |

ขึ้นปีใหม่: สร้างโฟลเดอร์ `X-Rite_<ปีใหม่>` และไฟล์ทะเบียนงานปีใหม่ไว้ตามรูปแบบเดิม

## แก้ปัญหา
| อาการ | สาเหตุ / วิธีแก้ |
|---|---|
| AADSTS50011 redirect URI | ที่อยู่เว็บไม่ตรงกับ Redirect URI (ตรวจ / ท้าย URL และ index.html) |
| AADSTS65001 / ขอยินยอม | ยังไม่ได้ Grant admin consent |
| ไม่พบไฟล์ทะเบียนงาน / ชีต | ชื่อไฟล์หรือชื่อชีตเดือนไม่ตรง (ชีตต้องชื่อ January…December) |
| บันทึก SharePoint 403 | ผู้ใช้คนนั้นไม่มีสิทธิ์แก้ไขในไลบรารี Service |
| PDF ใหญ่เกิน 3 MB | อีเมลถึงตัวเองจะแนบลิงก์ SharePoint แทน ส่วนลูกค้าให้ดาวน์โหลดแล้วส่งเอง |
| ล็อกอินบน iPhone วนซ้ำ | Settings › Safari › ปิด "Prevent Cross-Site Tracking" ชั่วคราว หรือเปิดจาก Add to Home Screen |

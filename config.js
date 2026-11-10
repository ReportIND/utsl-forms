/* ===== ตั้งค่าแอป UTSL Forms (แก้ไฟล์นี้ไฟล์เดียว) ===== */
window.APP_CONFIG = {
  /* Application (client) ID ที่ฝ่าย IT ลงทะเบียนให้ — ว่างไว้ = โหมดทดลอง (ไม่ต้องเข้าสู่ระบบ, ไม่อัปโหลด/ไม่ส่งอีเมล) */
  CLIENT_ID: "887e3d8a-ec0b-41be-82d0-8234775f15d0",
  TENANT_ID: "f4dd6a23-2dd0-42a5-966d-54356962f6dc",
  DOMAIN: "utsl.co.th",                       // อนุญาตเฉพาะบัญชีที่ลงท้ายด้วยโดเมนนี้

  SP_HOST: "uniontslth.sharepoint.com",
  SITE: "/sites/Industrial",

  /* ทะเบียนงาน (Excel ใน SharePoint) — {YYYY} = ปี ค.ศ. ของวันที่เลือก · ชีต = ชื่อเดือนภาษาอังกฤษ เช่น November */
  REGISTER: {
    library: "General Ind",
    path: "ทะเบียนงาน_IND/{YYYY}/ทะเบียนงาน_IND & PP {YYYY}.xlsx",
    range: "A1:Q900"
  },

  /* ที่เก็บ PDF: <JOB_LIBRARY>/<root>/<year>/<ชื่อลูกค้า>/<เลขที่ใบสั่งงาน> — แบรนด์อ่านจากเลขที่ใบสั่งงาน */
  JOB_LIBRARY: "Service",
  BRANDS: {
    XR: { root: "X-Rite", year: "X-Rite_{YYYY}" }
  }
};

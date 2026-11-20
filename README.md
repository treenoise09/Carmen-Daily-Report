# Zoho Daily Report

ดับเบิลคลิก `start-daily-report.cmd` เพื่อดาวน์โหลดรายงาน Zoho ทั้งสองและสร้างรายงาน Excel
หาก session หมดอายุ ให้ล็อกอินในเบราว์เซอร์ที่โปรแกรมเปิดขึ้น
ต้องติดตั้ง Node.js, Microsoft Excel และ dependencies ใน `scripts` ก่อนใช้งาน

- `reports/` เก็บรายงานที่สร้างและไฟล์สำรอง โปรแกรมสร้างโฟลเดอร์นี้ให้อัตโนมัติ
- `templates/Tickets_AllZoho.xls` ต้นแบบที่รวมอยู่ในโปรเจกต์ ต้องเก็บไว้เมื่อย้ายโปรเจกต์
- `scripts/downloads/zoho-daily/` เก็บรายงานต้นทางจาก Zoho แยกแต่ละรอบ
- `scripts/playwright/.auth/zoho-profile/` เก็บ session เข้าสู่ระบบ ไม่ควรแชร์หรือส่งเข้า Git

จากโฟลเดอร์ `scripts` ใช้ `npm start` เพื่อดาวน์โหลดและสร้างรายงาน หรือ `npm run merge` เพื่อรวมไฟล์ใน `scripts/input`

ค่าเริ่มต้นทั้งหมดอ้างอิงตำแหน่งโปรเจกต์ ไม่จำเป็นต้องมีโฟลเดอร์ Daily Report บน Desktop
หากต้องการกำหนดเอง ใช้ `DAILY_REPORT_DIR` สำหรับโฟลเดอร์ผลลัพธ์ และ `DAILY_REPORT_TEMPLATE` สำหรับต้นแบบ

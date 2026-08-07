/**
 * ===================================================================
 * ระบบ OP App — Apps Script Backend (รวม 3 โมดูล)
 * ===================================================================
 * 1. บันทึกกะ OP (Shift Log)
 * 2. ความพร้อมใช้งานเครื่องจักร (Machine Readiness) — ทุกหมวด
 * 3. มิเตอร์ไฟฟ้า-ประปา (Meter Readings)
 * ===================================================================
 * วิธี Deploy:
 * 1. สร้าง Google Sheet ใหม่ แล้ว Extensions > Apps Script
 * 2. วางโค้ดนี้ทั้งหมด
 * 3. รันฟังก์ชัน seedMachines() และ seedMeterPoints() ครั้งเดียว
 * 4. Deploy > New deployment > Web app > Execute as: Me, Who has access:
 *    Anyone with the link > Deploy > คัดลอก URL ไปใส่ใน CONFIG.GAS_URL ของ
 *    ไฟล์ op-shift-log.html, op-machine-readiness.html, op-meter-reading.html
 * ===================================================================
 */

const SH_SHIFT = "บันทึกกะ OP";
const SH_MACHINE = "เครื่องจักร";
const SH_ISSUE = "ปัญหาเครื่องจักร";
const SH_CHECKLOG = "บันทึกตรวจเช็กเครื่องจักร";
const SH_METER_POINT = "จุดมิเตอร์";
const SH_METER_LOG = "บันทึกมิเตอร์";

function getSheet(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet){
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function fmtDT(v){
  if (v && typeof v.getFullYear === "function") return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
  return String(v || "");
}
function fmtDate(v){
  if (v && typeof v.getFullYear === "function") return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd");
  return String(v || "");
}

/* ===================== โมดูล 1: บันทึกกะ OP ===================== */

function getShiftSheet(){
  return getSheet(SH_SHIFT, [
    "เวลาบันทึก","กะ","ผู้บันทึก",
    "อุณหภูมิภายนอก(°C)","ความชื้น(%rh)","Setpoint Chiller(°F)",
    "สถานะเครื่องจักร (JSON)",
    "มิเตอร์ไฟฟ้า(kWh)","มิเตอร์ประปา(m3)","มิเตอร์คูลลิ่ง(m3)",
    "Water Treatment","Fire Extinguishing","Medical Gas(psig)",
    "Generator สถานะ","Generator(kW)",
    "Submersible Drain","Submersible Soil","Submersible Sewage",
    "หมายเหตุ"
  ]);
}

function doPost_shift(data){
  const sheet = getShiftSheet();
  sheet.appendRow([
    new Date(), data.shift || "", data.recorder || "",
    data.outsideTemp || "", data.humidity || "", data.chillerSetpoint || "",
    JSON.stringify(data.machineStatus || []),
    data.meterElectric || "", data.meterWater || "", data.meterCooling || "",
    data.waterTreatment || "", data.fireExtinguishing || "", data.medicalGas || "",
    data.generatorStatus || "", data.generatorKw || "",
    data.subDrain || "", data.subSoil || "", data.subSewage || "",
    data.note || ""
  ]);
  return { status: "ok" };
}

function getLastShiftMeters(){
  const sheet = getShiftSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { found: false };
  const last = data[data.length - 1];
  return {
    found: true, time: fmtDT(last[0]),
    meterElectric: last[7], meterWater: last[8], meterCooling: last[9]
  };
}

/* ===================== โมดูล 2: ความพร้อมใช้งานเครื่องจักร ===================== */

function getMachineSheet(){
  return getSheet(SH_MACHINE, ["หมวด","ชื่อเครื่องจักร","จำนวนทั้งหมด"]);
}
function getIssueSheet(){
  return getSheet(SH_ISSUE, ["เวลาแจ้ง","หมวด","ชื่อเครื่องจักร","จำนวนที่ไม่พร้อมใช้","รายละเอียดปัญหา","ผู้รับผิดชอบ/บริษัท","ผู้แจ้ง","สถานะ","เวลาปิดงาน","ผู้ปิดงาน"]);
}
function getCheckLogSheet(){
  return getSheet(SH_CHECKLOG, [
    "เวลาตรวจ","หมวด","ชื่อเครื่องจักร",
    "ความสะอาด","เสียง/การสั่นสะเทือน","การทำงานเมื่อทดสอบ",
    "สถานะโดยรวม","จำนวนที่ไม่พร้อมใช้","รายละเอียดปัญหา","ผู้รับผิดชอบ/บริษัท",
    "ผู้ตรวจ"
  ]);
}

/**
 * บันทึกผลตรวจเช็ก (checksheet) ของเครื่องจักร 1 เครื่อง — ต้องมาจากการสแกน
 * QR หน้าเครื่องเท่านั้น (ฝั่งเว็บบังคับ ไม่ให้เข้าถึงฟอร์มนี้จากการเลือกจาก
 * รายการเฉยๆ) ทุกครั้งที่ตรวจจะถูกบันทึกเป็นประวัติไว้ในชีตนี้เสมอ ไม่ว่าผล
 * จะปกติหรือไม่ปกติ — ถ้าผลรวมคือ "ไม่พร้อมใช้งาน" จะสร้างรายการปัญหาใหม่ใน
 * ชีต "ปัญหาเครื่องจักร" ให้อัตโนมัติด้วย (สถานะเปิด) เพื่อให้ตัวเลขความ
 * พร้อมใช้งานอัปเดตถูกต้องทันที
 */
function doPost_machineChecksheet(data){
  if (!data.category || !data.name) throw new Error("ไม่พบข้อมูลเครื่องจักร");
  if (!data.inspector || !String(data.inspector).trim()) throw new Error("กรุณาระบุชื่อผู้ตรวจ");
  if (!data.overall) throw new Error("กรุณาระบุสถานะโดยรวม");

  const logSheet = getCheckLogSheet();
  logSheet.appendRow([
    new Date(), data.category, data.name,
    data.cleanliness || "", data.noise || "", data.testRun || "",
    data.overall, data.overall === "ไม่พร้อมใช้งาน" ? (Number(data.qty)||1) : 0,
    data.detail || "", data.vendor || "",
    data.inspector.trim()
  ]);

  if (data.overall === "ไม่พร้อมใช้งาน"){
    if (!data.detail || !String(data.detail).trim()) throw new Error("กรุณาระบุรายละเอียดปัญหาเมื่อเครื่องไม่พร้อมใช้งาน");
    const isheet = getIssueSheet();
    isheet.appendRow([
      new Date(), data.category, data.name, Number(data.qty) || 1,
      data.detail.trim(), data.vendor || "", data.inspector.trim(), "เปิด", "", ""
    ]);
  }
  return { status: "ok" };
}

function seedMachines(){
  const sheet = getMachineSheet();
  const existing = sheet.getDataRange().getValues().slice(1).map(r => r[0] + "|" + r[1]);
  const list = MACHINE_SEED_DATA;
  let added = 0;
  list.forEach(m => {
    const key = m.cat + "|" + m.name;
    if (existing.indexOf(key) === -1){ sheet.appendRow([m.cat, m.name, m.total]); added++; }
  });
  Logger.log("เพิ่มเครื่องจักร " + added + " รายการ");
}

// รายชื่อเครื่องจักรตั้งต้นทั้ง 8 หมวด (ดึงจากรายงานความพร้อมใช้งานเครื่องจักร)
const MACHINE_SEED_DATA = [
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Chiller Water Cool", total:8},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Automatic Tube Cleaning System", total:8},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Cooling Tower", total:8},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Primary Chiller Water Pump", total:8},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Secondary Chiller Water Pump", total:9},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Condenser Water Pump", total:8},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Expansion Tank ดาดฟ้า", total:1},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Car Parking Fresh Air", total:2},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Pressurize Fan", total:4},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Car Parking Exhaust Air", total:3},
  {cat:"หมวดที่ 1 ระบบปรับอากาศและระบายอากาศ", name:"Smoke Exhaust Fan", total:11},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Ring Main Unit", total:4},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Transformer (Dry Type)", total:9},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Main Distribution Board", total:9},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Emergency Main Distribution Board", total:8},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Automatic Transfer Switch", total:6},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Capacitor Bank", total:9},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Generator", total:3},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Uninterruptible Power Supply", total:20},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Isolated Power Supply", total:82},
  {cat:"หมวดที่ 2 ระบบไฟฟ้าและไฟฟ้าสื่อสาร", name:"Heliport Lighting", total:1},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Digital Meter System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Two Wire Remote Control", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Chiller Control System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Access Control & CCTV", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Fire Alarm System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Lift System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Electrical System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Lighting Control System (Stair)", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Escalator System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Sanitary System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Fire Protection System", total:2},
  {cat:"หมวดที่ 3 ระบบ Building Automation System (BAS)", name:"Gas System (แก๊สทางการแพทย์)", total:2},
  {cat:"หมวดที่ 4 ระบบสุขาภิบาล", name:"Cold Water Transfer Pump (PC)", total:4},
  {cat:"หมวดที่ 4 ระบบสุขาภิบาล", name:"Booster Pump (รดน้ำต้นไม้)", total:1},
  {cat:"หมวดที่ 4 ระบบสุขาภิบาล", name:"Underground and Roof Tank", total:5},
  {cat:"หมวดที่ 4 ระบบสุขาภิบาล", name:"Hot Water Pump", total:17},
  {cat:"หมวดที่ 4 ระบบสุขาภิบาล", name:"Softener Water Tank", total:3},
  {cat:"หมวดที่ 4 ระบบสุขาภิบาล", name:"ระบบน้ำพุด้านหน้าอาคาร (ปั้ม)", total:3},
  {cat:"หมวดที่ 5 ระบบบ่อพักน้ำเสีย", name:"Submersible Drainage Pump", total:14},
  {cat:"หมวดที่ 5 ระบบบ่อพักน้ำเสีย", name:"Submersible Sewage Pump", total:16},
  {cat:"หมวดที่ 5 ระบบบ่อพักน้ำเสีย", name:"Sewage Pump", total:11},
  {cat:"หมวดที่ 5 ระบบบ่อพักน้ำเสีย", name:"Neutralization Pump", total:8},
  {cat:"หมวดที่ 6 ระบบป้องกันอัคคีภัย ดับเพลิงและรักษาความปลอดภัย", name:"Fire Alarm Control Panel", total:2},
  {cat:"หมวดที่ 6 ระบบป้องกันอัคคีภัย ดับเพลิงและรักษาความปลอดภัย", name:"Engine Fire Pump", total:2},
  {cat:"หมวดที่ 6 ระบบป้องกันอัคคีภัย ดับเพลิงและรักษาความปลอดภัย", name:"Jockey Pump", total:2},
  {cat:"หมวดที่ 6 ระบบป้องกันอัคคีภัย ดับเพลิงและรักษาความปลอดภัย", name:"FM 200", total:5},
  {cat:"หมวดที่ 6 ระบบป้องกันอัคคีภัย ดับเพลิงและรักษาความปลอดภัย", name:"Pre-Action", total:26},
  {cat:"หมวดที่ 7 ระบบขนส่ง", name:"Passenger Lift", total:46},
  {cat:"หมวดที่ 7 ระบบขนส่ง", name:"Dumb Waiter Lift", total:7},
  {cat:"หมวดที่ 7 ระบบขนส่ง", name:"Escalator", total:6},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"Oxygen Control System", total:2},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"Vacuum Pump", total:3},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"Medical Air", total:3},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"Nitrous Oxide Control System", total:1},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"Instrument Air", total:2},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"Carbon dioxide", total:1},
  {cat:"หมวดที่ 8 ระบบแก๊สทางการแพทย์ และระบบแก๊ส LPG", name:"LPG Station", total:1}
];

/**
 * คืนรายชื่อเครื่องจักรทั้งหมด พร้อมคำนวณ "จำนวนไม่พร้อมใช้งาน" สดๆ จาก
 * ผลรวมของปัญหาที่ยังเปิดอยู่ (สถานะ = เปิด) ของเครื่องจักรนั้น — ไม่ต้อง
 * แก้ตัวเลขในชีตด้วยมือ ระบบคำนวณให้อัตโนมัติจากประวัติปัญหา
 */
function getMachinesWithReadiness(){
  const msheet = getMachineSheet();
  const mdata = msheet.getDataRange().getValues();
  mdata.shift();

  const isheet = getIssueSheet();
  const idata = isheet.getDataRange().getValues();
  idata.shift();

  const notReadyMap = {}; // "หมวด|ชื่อ" -> sum จำนวนไม่พร้อมใช้ (เฉพาะปัญหาที่เปิดอยู่)
  idata.forEach(r => {
    if (r[7] === "เปิด"){
      const key = r[1] + "|" + r[2];
      notReadyMap[key] = (notReadyMap[key] || 0) + (Number(r[3]) || 0);
    }
  });

  return mdata.map(r => {
    const key = r[0] + "|" + r[1];
    const notReady = notReadyMap[key] || 0;
    return {
      category: r[0], name: r[1], total: Number(r[2]) || 0,
      notReady: notReady, ready: Math.max(0, (Number(r[2])||0) - notReady)
    };
  });
}

function getOpenIssues(category, name){
  const isheet = getIssueSheet();
  const idata = isheet.getDataRange().getValues();
  idata.shift();
  const out = [];
  idata.forEach((r, i) => {
    if (r[7] === "เปิด" && r[1] === category && r[2] === name){
      out.push({
        rowIndex: i + 2, time: fmtDT(r[0]), category: r[1], name: r[2],
        qty: r[3], detail: r[4], vendor: r[5], reporter: r[6]
      });
    }
  });
  return out;
}

function getAllOpenIssues(){
  const isheet = getIssueSheet();
  const idata = isheet.getDataRange().getValues();
  idata.shift();
  const out = [];
  idata.forEach(r => {
    if (r[7] === "เปิด"){
      out.push({ time: fmtDT(r[0]), category: r[1], name: r[2], qty: r[3], detail: r[4], vendor: r[5], reporter: r[6] });
    }
  });
  out.sort((a,b)=> new Date(a.time) - new Date(b.time));
  return out;
}

function doPost_machineReport(data){
  if (!data.category || !data.name) throw new Error("ไม่พบข้อมูลเครื่องจักร");
  if (!data.detail || !String(data.detail).trim()) throw new Error("กรุณาระบุรายละเอียดปัญหา");
  if (!data.reporter || !String(data.reporter).trim()) throw new Error("กรุณาระบุชื่อผู้แจ้ง");
  const isheet = getIssueSheet();
  isheet.appendRow([
    new Date(), data.category, data.name, Number(data.qty) || 1,
    data.detail.trim(), data.vendor || "", data.reporter.trim(), "เปิด", "", ""
  ]);
  return { status: "ok" };
}

function doPost_machineClose(data){
  const rowIndex = Number(data.rowIndex);
  const closer = (data.closer || "").trim();
  if (!rowIndex) throw new Error("ไม่พบรายการปัญหาที่จะปิด");
  if (!closer) throw new Error("กรุณาระบุชื่อผู้ปิดงาน");
  const isheet = getIssueSheet();
  isheet.getRange(rowIndex, 8, 1, 3).setValues([["ปิด", new Date(), closer]]);
  return { status: "ok" };
}

/* ===================== โมดูล 3: มิเตอร์ไฟฟ้า-ประปา ===================== */

function getMeterPointSheet(){
  return getSheet(SH_METER_POINT, ["รหัส","ชื่อจุดมิเตอร์","หน่วย"]);
}
function getMeterLogSheet(){
  return getSheet(SH_METER_LOG, ["วันที่","รหัสมิเตอร์","ชื่อจุดมิเตอร์","ค่าก่อนหน้า","ค่าปัจจุบัน","ผลต่าง","หน่วย","ผู้บันทึก","หมายเหตุ"]);
}

const METER_SEED_DATA = [
  {code:"M411", name:"METER 411", unit:"kWh"},
  {code:"M412", name:"METER 412", unit:"kWh"},
  {code:"M422", name:"METER 422", unit:"kWh"},
  {code:"WATER", name:"มิเตอร์ประปา", unit:"m3"},
  {code:"COOL", name:"มิเตอร์คูลลิ่ง", unit:"m3"},
  {code:"MAIN28040940", name:"MAIN 28040940", unit:"kWh"},
  {code:"ROOF1", name:"Roof 1", unit:"kWh"},
  {code:"ROOF2", name:"Roof 2", unit:"kWh"},
  {code:"SOLAR", name:"Solar Cell", unit:"kWh"},
  {code:"FUEL", name:"น้ำมันคงเหลือ", unit:"ลิตร"}
];

function seedMeterPoints(){
  const sheet = getMeterPointSheet();
  const existing = sheet.getDataRange().getValues().slice(1).map(r => r[0]);
  let added = 0;
  METER_SEED_DATA.forEach(m => {
    if (existing.indexOf(m.code) === -1){ sheet.appendRow([m.code, m.name, m.unit]); added++; }
  });
  Logger.log("เพิ่มจุดมิเตอร์ " + added + " จุด");
}

function getMeterPoints(){
  const sheet = getMeterPointSheet();
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(r => ({ code: r[0], name: r[1], unit: r[2] }));
}

function getLastMeterReading(code){
  const sheet = getMeterLogSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--){
    if (data[i][1] === code){
      return { found: true, date: fmtDate(data[i][0]), reading: data[i][4] };
    }
  }
  return { found: false };
}

function doPost_meterSubmit(data){
  const code = data.code;
  if (!code) throw new Error("ไม่พบรหัสมิเตอร์");
  const pointSheet = getMeterPointSheet();
  const points = pointSheet.getDataRange().getValues();
  let point = null;
  for (let i = 1; i < points.length; i++){ if (points[i][0] === code){ point = points[i]; break; } }
  if (!point) throw new Error("ไม่พบจุดมิเตอร์นี้ในระบบ");

  const current = Number(data.reading);
  if (isNaN(current)) throw new Error("กรุณากรอกค่ามิเตอร์เป็นตัวเลข");

  const last = getLastMeterReading(code);
  const prev = last.found ? Number(last.reading) : 0;
  const diff = last.found ? (current - prev) : 0;

  const logSheet = getMeterLogSheet();
  logSheet.appendRow([
    new Date(), code, point[1], last.found ? prev : "", current, last.found ? diff : "",
    point[2], data.recorder || "", data.note || ""
  ]);
  return { status: "ok", previous: last.found ? prev : null, diff: last.found ? diff : null };
}

/* ===================== doGet / doPost router ===================== */

function doGet(e){
  const action = e.parameter.action;
  if (action === "last_shift_meters") return jsonOut(getLastShiftMeters());
  if (action === "machines") return jsonOut(getMachinesWithReadiness());
  if (action === "machine_issues") return jsonOut(getOpenIssues(e.parameter.category, e.parameter.name));
  if (action === "all_issues") return jsonOut(getAllOpenIssues());
  if (action === "meter_points") return jsonOut(getMeterPoints());
  if (action === "meter_last") return jsonOut(getLastMeterReading(e.parameter.code));
  return jsonOut({ status:"ok", message:"OP App backend กำลังทำงาน" });
}

function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;
    if (action === "shift_submit") result = doPost_shift(data);
    else if (action === "machine_report") result = doPost_machineReport(data);
    else if (action === "machine_checksheet") result = doPost_machineChecksheet(data);
    else if (action === "machine_close") result = doPost_machineClose(data);
    else if (action === "meter_submit") result = doPost_meterSubmit(data);
    else throw new Error("ไม่รู้จักคำสั่ง: " + action);
    return jsonOut(result);
  }catch(err){
    return jsonOut({ status: "error", message: err.message });
  }
}

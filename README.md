const SPREADSHEET_ID = '1u4dGPFshrkxvEqwKZ2f73QMyxYDTNkiO8lBM-yh-usw';
const FOLDER_ID = '1xeZv0ZoFdIndXPksMCNDunA2inVWBxdd';
const SS = SpreadsheetApp.openById(SPREADSHEET_ID);

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Office Stock Elite')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- 🔐 ระบบ AUTHENTICATION ---
function checkLogin(username, password) {
  const sheet = SS.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === username && data[i][1].toString() === password) {
      return { 
        success: true, 
        user: data[i][0].toString(), // 🚨 เพิ่มบรรทัดนี้ เพื่อให้ฝั่ง HTML รู้จัก username
        role: data[i][3], 
        name: data[i][2], 
        dept: data[i][4] 
      };
    }
  }
  return { success: false };
}

/**
 * ฟังก์ชันเบิกสินค้าแบบ FIFO พร้อมระบบแจ้งเตือน Email & Line
 */
function withdrawProductFIFO(productId, qtyNeeded, user, dept, approverEmail) {
  const lotSheet = SS.getSheetByName("Inventory_Lots");
  const lots = lotSheet.getDataRange().getValues();
  let remaining = parseFloat(qtyNeeded);
  
  // --- 1. ตรวจสอบสต๊อกรวมก่อนเบิกเพื่อป้องกัน Error ---
  let totalInStock = 0;
  for (let i = 1; i < lots.length; i++) {
    if (lots[i][1] == productId) totalInStock += parseFloat(lots[i][2]);
  }
  
  if (totalInStock < remaining) {
    throw new Error("ขออภัย สินค้าในสต๊อกคงเหลือไม่พอ (คงเหลือรวม: " + totalInStock + ")");
  }

  // --- 2. เริ่มการตัดสต๊อกแบบ FIFO ---
  for (let i = 1; i < lots.length; i++) {
    if (lots[i][1] == productId && lots[i][2] > 0) {
      let take = Math.min(remaining, lots[i][2]);
      let newQty = lots[i][2] - take;
      remaining -= take;
      
      lotSheet.getRange(i + 1, 3).setValue(newQty); // อัปเดตคอลัมน์ C (qty_remain)
      if (remaining <= 0) break;
    }
  }
  
  // --- 3. บันทึกลงตาราง Transactions (11 คอลัมน์) ---
  const txSheet = SS.getSheetByName("Transactions");
  const txDate = new Date();
  txSheet.appendRow([
    txDate,         // A: วันที่
    "WITHDRAW",     // B: ประเภท
    productId,      // C: รหัสสินค้า
    qtyNeeded,      // D: จำนวน
    user,           // E: ชื่อผู้เบิก (Username)
    dept,           // F: แผนก
    approverEmail,  // G: ผู้อนุมัติ (Email)
    "",             // H: วันที่คืน (ถ้ามี)
    "",             // I: ID อ้างอิง
    "",             // J: หมายเหตุ
    "รออนุมัติ"      // K: สถานะ
  ]);
  
  // --- 4. ส่งแจ้งเตือนผ่าน Email และ Line ทันที ---
  try {
    const msg = `📦 มีรายการเบิกสินค้าใหม่!\n------------------\n🔹 สินค้า: ${productId}\n🔹 จำนวน: ${qtyNeeded}\n🔹 ผู้เบิก: ${user} (${dept})\n🔹 สถานะ: รอการอนุมัติ\n------------------\nกรุณาเข้าสู่ระบบเพื่อดำเนินการตรวจสอบครับ`;
    
    // เรียกใช้ฟังก์ชัน notifyMultiChannel ที่เราสร้างไว้
    notifyMultiChannel(approverEmail, "รายการเบิกใหม่รอการอนุมัติ", msg);
    
  } catch (err) {
    console.error("Notification Error: " + err.message);
    // เราจะไม่ throw error ตรงนี้เพื่อให้การเบิกสต๊อกยังคงสำเร็จแม้แจ้งเตือนจะขัดข้อง
  }
  
  return { 
    success: true, 
    message: "ส่งคำขอเบิกเรียบร้อย และส่งแจ้งเตือนไปยัง " + approverEmail + " แล้วครับ" 
  };
}

function logTransaction(type, pid, qty, user, dept, status) {
  const sheet = SS.getSheetByName("Transactions");
  sheet.appendRow([new Date(), type, pid, qty, user, dept, status]);
}

// --- ✅ ระบบอนุมัติ (APPROVAL) ---
function getPendingRequests(dept) {
  const sheet = SS.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  // คืนค่ารายการที่สถานะเป็น PENDING ของแผนกนั้นๆ
  return data.filter(r => r[6] === "PENDING" && r[5] === dept);
}

function updateStatus(rowId, newStatus) {
  const sheet = SS.getSheetByName("Transactions");
  // โลจิกสำหรับเปลี่ยนสถานะใน Sheet และแจ้งเตือนผ่าน Email/Line
  sheet.getRange(rowId, 7).setValue(newStatus);
  return true;
}

// --- 📊 ระบบรายงาน (REPORT EXPORT) ---
function exportToPDF() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  // โลจิกการสร้าง PDF จากข้อมูลใน Sheet และเก็บลงใน Drive_ID ของคุณ
  // จะส่ง URL ไฟล์กลับไปให้ User ดาวน์โหลด
}
// ดึงข้อมูลสำหรับทำ Chart ใน Dashboard
function getChartData() {
  return {
    line: {
      labels: ['จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.', 'อา.'],
      data: [12, 25, 18, 30, 15, 5, 8] 
    },
    bar: {
      labels: ['IT', 'HR', 'บัญชี', 'การตลาด', 'ธุรการ'],
      data: [120, 45, 30, 80, 65] 
    },
    pie: {
      labels: ['เครื่องเขียน', 'อุปกรณ์ IT', 'ทำความสะอาด'],
      data: [45, 35, 20] 
    }
  };
}

function registerNewUser(username, password, name, dept) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName("Users"); 
  
  if (!sheet) return { success: false, message: "ไม่พบแผ่นชีทชื่อ 'Users'" };

  const data = sheet.getDataRange().getValues();
  
  // ตรวจสอบความยาว 5 ตัวอักษร
  if (username.length < 5 || password.length < 5) {
    return { success: false, message: "ID และ Password ต้องมีอย่างน้อย 5 ตัวอักษร" };
  }

  // 1. ตรวจสอบว่าชื่อผู้ใช้ซ้ำหรือไม่
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === username.toString().toLowerCase()) {
      return { 
        success: false, 
        message: "ขออภัย! ชื่อผู้ใช้ '" + username + "' นี้มีผู้ใช้งานแล้ว" 
      };
    }
  }

  // 🚨 2. บันทึกข้อมูล (คุณลืมส่วนนี้ไปในโค้ดล่าสุด)
  try {
    // โครงสร้างตามรูป: username, password, name, role, dept
    sheet.appendRow([username, password, name, "user", dept]); 
    return { success: true, message: "ลงทะเบียนสำเร็จเรียบร้อยแล้ว!" };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ดึงรายชื่อสินค้า พร้อมคำนวณสต๊อกคงเหลือ
function getProductsList() {
  const prodSheet = SS.getSheetByName("Products");
  const lotSheet = SS.getSheetByName("Inventory_Lots"); // ดึงยอดจาก Sheet ล็อตสินค้า
  if (!prodSheet || !lotSheet) return [];
  
  const prods = prodSheet.getDataRange().getValues();
  const lots = lotSheet.getDataRange().getValues();
  
  // 1. คำนวณสต๊อกรวมของแต่ละ Product_ID จาก Inventory_Lots
  let stockMap = {};
  for(let i = 1; i < lots.length; i++) {
    let pid = lots[i][1]; // คอลัมน์ Product_ID
    let remainQty = parseFloat(lots[i][2]) || 0; // คอลัมน์ Qty_Remain
    if(!stockMap[pid]) stockMap[pid] = 0;
    stockMap[pid] += remainQty;
  }
  
  // 2. นำสต๊อกที่ได้ ไปจับคู่กับชื่อสินค้า
  let products = [];
  for (let i = 1; i < prods.length; i++) {
    let pid = prods[i][0];
    products.push({ 
      id: pid, 
      name: prods[i][2], 
      unit: prods[i][3],
      stock: stockMap[pid] || 0 // ถ้ายอดไม่มีให้เป็น 0
    });
  }
  return products;
}

// ดึงรายชื่อผู้อนุมัติ ตามแผนกของ User ที่ล็อกอิน
function getApproversList(dept) {
  const sheet = SS.getSheetByName("Approvers");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let approvers = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === dept || dept === 'admin') { // ถ้าเป็นแผนกเดียวกันให้ดึงมา
      approvers.push({ name: data[i][1], email: data[i][2] });
    }
  }
  return approvers;
}

// --- บันทึกการยืม (11 คอลัมน์ เริ่มที่ Date) ---
function recordBorrowing(pid, qty, user, dept, approver, returnDate) {
  const sheet = SS.getSheetByName("Transactions");
  const borrowId = "BRW-" + new Date().getTime();
  
  sheet.appendRow([
    new Date(),    // A (Index 0): date
    "BORROW",      // B (Index 1): type
    pid,           // C (Index 2): product_id
    qty,           // D (Index 3): qty
    user,          // E (Index 4): user (บันทึก Username)
    dept,          // F (Index 5): dept
    approver,      // G (Index 6): approver
    returnDate,    // H (Index 7): return_date
    borrowId,      // I (Index 8): borrow_id
    "PENDING",     // J (Index 9): return_status
    "รออนุมัติ"      // K (Index 10): status
  ]);
  return { success: true, message: "ส่งคำขอยืมอุปกรณ์เรียบร้อย" };
}

// --- ดึงรายการที่ยังไม่คืน (เวอร์ชันอัปเกรด ค้นหาแบบแม่นยำ) ---
function getMyBorrowedItems(username, fullname) {
  const sheet = SS.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  
  // แปลงชื่อที่ส่งมาเป็นพิมพ์เล็กและตัดช่องว่างหัวท้ายทิ้ง
  const u = username ? username.toString().trim().toLowerCase() : "___";
  const f = fullname ? fullname.toString().trim().toLowerCase() : "___";
  
  const result = [];
  
  // วนลูปอ่านข้อมูลทีละบรรทัด
  for(let i = 1; i < data.length; i++) {
    const r = data[i];
    if(!r[1]) continue; // ถ้าบรรทัดนั้นว่างให้ข้ามไป
    
    // ดึงข้อมูลในแถวมาตัดช่องว่างทิ้งเช่นกัน
    const rowType = r[1].toString().trim().toUpperCase(); 
    const rowUser = r[4] ? r[4].toString().trim().toLowerCase() : ""; 
    const rowRetStatus = r[9] ? r[9].toString().trim().toUpperCase() : ""; 
    
    // เงื่อนไข: ต้องเป็น BORROW + ชื่อตรงกับ Username หรือชื่อจริง + สถานะต้องไม่ใช่ RETURNED
    if(rowType === "BORROW" && (rowUser === u || rowUser === f) && rowRetStatus !== "RETURNED") {
      result.push([
        r[0] ? r[0].toString() : "", // 🚨 แปลงวันที่คอลัมน์ A เป็น Text เพื่อไม่ให้ GAS บล็อกข้อมูล
        r[1], r[2], r[3], r[4], r[5], r[6], 
        r[7] ? r[7].toString() : "", // 🚨 แปลงวันที่คอลัมน์ H เป็น Text เพื่อไม่ให้ GAS บล็อกข้อมูล
        r[8], 
        r[9], r[10]
      ]);
    }
  }
  return result;
}

// --- บันทึกการส่งคืน (เวอร์ชันอัปเกรด) ---
function returnItem(borrowId) {
  const sheet = SS.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  
  const targetId = borrowId.toString().trim(); // ตัดช่องว่างรหัสที่จะคืน
  
  for (let i = 1; i < data.length; i++) {
    // เช็คจาก borrow_id คอลัมน์ I (Index 8)
    if (data[i][8] && data[i][8].toString().trim() === targetId) { 
      sheet.getRange(i + 1, 10).setValue("RETURNED"); // ปรับแก้สถานะที่คอลัมน์ J (ที่ 10)
      return { success: true, message: "บันทึกการส่งคืนสำเร็จ" };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลรหัสการยืมนี้" };
}


// --- ดึงประวัติการเบิก (แก้ไขบัควันที่เรียบร้อย) ---
function getMyWithdrawHistory(username, fullname) {
  const sheet = SS.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  
  const u = username ? username.toString().trim().toLowerCase() : "___";
  const f = fullname ? fullname.toString().trim().toLowerCase() : "___";
  
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[1]) continue; // ข้ามแถวว่าง
    
    const rowType = r[1].toString().trim().toUpperCase();
    const rowUser = r[4] ? r[4].toString().trim().toLowerCase() : "";
    
    // ตรวจสอบว่าเป็นรายการเบิก และตรงกับ User ที่ล็อกอิน
    if ((rowType === "WITHDRAW" || rowType === "เบิก") && (rowUser === u || rowUser === f)) {
      result.push([
        r[0] ? r[0].toString() : "", // 🚨 แปลงวันที่เป็น Text ป้องกันระบบบล็อกข้อมูล
        r[2], // Product ID
        r[3], // จำนวน (Qty)
        r[10] ? r[10].toString() : r[6] // สถานะ (Status)
      ]);
    }
  }
  
  return result.reverse(); // สลับให้รายการล่าสุดขึ้นด้านบน
}

// ==========================================
// 📦 ระบบจัดการคลังสินค้า (Admin เท่านั้น)
// ==========================================

// 1. โหลดข้อมูลภาพรวมคลังสินค้า (ปรับให้รองรับ Min Stock ของแต่ละรายการ)
function getAdminStockData() {
  const prodSheet = SS.getSheetByName("Products");
  const lotSheet = SS.getSheetByName("Inventory_Lots");
  
  const prods = prodSheet.getDataRange().getValues();
  const lots = lotSheet.getDataRange().getValues();
  
  let stockMap = {};
  let totalItems = 0;
  
  // คำนวณสต๊อกคงเหลือ
  for(let i = 1; i < lots.length; i++) {
    let pid = lots[i][1];
    let remainQty = parseFloat(lots[i][2]) || 0;
    if(!stockMap[pid]) stockMap[pid] = 0;
    stockMap[pid] += remainQty;
    totalItems += remainQty;
  }
  
  let products = [];
  let lowStockCount = 0;
  
  // ประกอบร่างข้อมูลให้ครบ 6 คอลัมน์
  for (let i = 1; i < prods.length; i++) {
    if(!prods[i][0]) continue;
    let pid = prods[i][0];        // A: id
    let barcode = prods[i][1];    // B: barcode
    let name = prods[i][2];       // C: name
    let unit = prods[i][3];       // D: unit
    let minStock = parseFloat(prods[i][4]) || 0; // E: min_stock
    let category = prods[i][5];   // F: category
    
    let qty = stockMap[pid] || 0;
    
    // เช็คว่าสินค้านี้ ต่ำกว่าหรือเท่ากับจุดสั่งซื้อ (Min Stock) หรือไม่
    if (qty <= minStock) lowStockCount++; 
    
    products.push({ 
      id: pid, barcode: barcode, name: name, unit: unit, minStock: minStock, category: category, stock: qty
    });
  }
  
  return { products: products, summary: { totalSKU: products.length, totalItems: totalItems, lowStock: lowStockCount } };
}

// 2. เติมสต๊อก (รับของเข้าคลัง)
function receiveStockAdmin(pid, qty, cost, adminUser) {
  const lotSheet = SS.getSheetByName("Inventory_Lots");
  const txSheet = SS.getSheetByName("Transactions");
  const date = new Date();
  
  // สร้าง Lot ID ใหม่แบบสุ่มตัวเลขต่อท้ายเพื่อไม่ให้ซ้ำ
  const newLotId = "LOT-" + date.getTime().toString().slice(-5);
  
  // เพิ่มล็อตใหม่ (A:lot_id, B:product_id, C:qty_remain, D:cost, E:date_in)
  lotSheet.appendRow([newLotId, pid, qty, cost, date]);
  
  // บันทึกประวัติลง Transactions (11 คอลัมน์)
  txSheet.appendRow([
    date, "รับเข้า", pid, qty, adminUser, "Admin", adminUser, "", "", "", "สำเร็จ"
  ]);
  
  return { success: true, message: "เพิ่มสต๊อกเรียบร้อยแล้ว" };
}

// 3. จัดการสินค้า (เพิ่ม / แก้ไข / ลบ) แบบ 6 คอลัมน์
function manageProductAdmin(action, id, barcode, name, unit, minStock, category, oldId) {
  const sheet = SS.getSheetByName("Products");
  const data = sheet.getDataRange().getValues();
  
  if (action === "ADD") {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) return { success: false, message: "รหัสสินค้านี้มีอยู่แล้ว" };
    }
    // เพิ่มข้อมูล 6 ช่อง
    sheet.appendRow([id, barcode, name, unit, minStock, category]); 
    return { success: true, message: "เพิ่มสินค้าใหม่สำเร็จ" };
  } 
  else if (action === "EDIT") {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === oldId) {
        // อัปเดตข้อมูล 6 ช่อง
        sheet.getRange(i + 1, 1).setValue(id);
        sheet.getRange(i + 1, 2).setValue(barcode);
        sheet.getRange(i + 1, 3).setValue(name);
        sheet.getRange(i + 1, 4).setValue(unit);
        sheet.getRange(i + 1, 5).setValue(minStock);
        sheet.getRange(i + 1, 6).setValue(category);
        return { success: true, message: "อัปเดตข้อมูลสำเร็จ" };
      }
    }
  }
  else if (action === "DELETE") {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "ลบสินค้าเรียบร้อยแล้ว" };
      }
    }
  }
  return { success: false, message: "เกิดข้อผิดพลาด ไม่พบข้อมูล" };
}

// ==========================================
// ✅ ระบบอนุมัติรายการ (Approval System)
// ==========================================

function getPendingApprovals(name, dept, role) {
  const txSheet = SS.getSheetByName("Transactions");
  const appSheet = SS.getSheetByName("Approvers");
  
  // 1. ตรวจสอบสิทธิ์ว่าเป็นผู้อนุมัติหรือไม่ (Admin ผ่านอัตโนมัติ)
  let isApprover = (role === 'admin');
  if (!isApprover) {
     const appData = appSheet.getDataRange().getValues();
     // เทียบชื่อผู้ใช้กับรายชื่อในตาราง Approvers
     for(let i = 1; i < appData.length; i++) {
        if(appData[i][1].toString().trim() === name.toString().trim()) { 
           isApprover = true;
           break;
        }
     }
  }
  
  // ถ้าไม่มีสิทธิ์ให้ตีกลับ
  if(!isApprover) {
    return { error: true, message: "ขออภัย สิทธิ์นี้เฉพาะ 'ผู้อนุมัติประจำแผนก' หรือ 'Admin' เท่านั้นครับ" };
  }
  
  // 2. ดึงข้อมูลที่รออนุมัติ
  const data = txSheet.getDataRange().getValues();
  const results = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][10] === "รออนุมัติ") {
      // Admin เห็นทั้งหมด / ผู้อนุมัติเห็นเฉพาะคนในแผนกตัวเอง
      if(role === 'admin' || data[i][5] === dept) {
        results.push({
          row: i + 1, // เก็บเลขบรรทัดไว้ใช้ตอนอัปเดตสถานะ
          date: data[i][0] ? data[i][0].toString() : "",
          type: data[i][1],
          pid: data[i][2],
          qty: data[i][3],
          user: data[i][4],
          dept: data[i][5]
        });
      }
    }
  }
  return { error: false, data: results.reverse() }; // นำรายการล่าสุดขึ้นก่อน
}

// ฟังก์ชันสำหรับ อนุมัติ / ไม่อนุมัติ
function processApproval(row, action, type, pid, qty) {
  const txSheet = SS.getSheetByName("Transactions");
  
  if(action === 'APPROVE') {
    txSheet.getRange(row, 11).setValue("สำเร็จ"); // เปลี่ยนสถานะ
    return { success: true, message: "อนุมัติรายการเรียบร้อย" };
  } 
  else {
    txSheet.getRange(row, 11).setValue("ไม่อนุมัติ"); // เปลี่ยนสถานะ
    
    // สำคัญ: หากไม่อนุมัติการ "เบิก" ต้องคืนของกลับเข้าคลัง (เพราะโดนตัด FIFO ไปก่อนหน้าแล้ว)
    if(type === "WITHDRAW" || type === "เบิก") {
       const lotSheet = SS.getSheetByName("Inventory_Lots");
       const returnLotId = "RET-" + new Date().getTime().toString().slice(-5);
       // คืนยอดเป็นล็อตใหม่ (ตั้งต้นทุนเป็น 0 เพราะเป็นของตีกลับ)
       lotSheet.appendRow([returnLotId, pid, qty, 0, new Date()]); 
    }
    return { success: true, message: "ปฏิเสธคำขอ และคืนสต๊อกเรียบร้อยแล้ว" };
  }
}

// ==========================================
// 📊 ระบบออกรายงาน (Report Generator)
// ==========================================

function generateReportData(reportType, startDate, endDate) {
  const txSheet = SS.getSheetByName("Transactions");
  const lotSheet = SS.getSheetByName("Inventory_Lots");
  const prodSheet = SS.getSheetByName("Products");

  const txData = txSheet.getDataRange().getValues();
  const lotData = lotSheet.getDataRange().getValues();
  const prodData = prodSheet.getDataRange().getValues();

  // สร้าง Dictionary เก็บชื่อสินค้าและหาราคาเฉลี่ย
  let prodMap = {};
  for(let i = 1; i < prodData.length; i++) {
    if(prodData[i][0]) prodMap[prodData[i][0]] = { name: prodData[i][2], unit: prodData[i][3], avgCost: 0 };
  }

  // หาราคาเฉลี่ยของสินค้าจากประวัติการรับเข้า (Lot)
  let costAgg = {};
  for(let i = 1; i < lotData.length; i++) {
    let pid = lotData[i][1];
    let qty = parseFloat(lotData[i][2]) || 0;
    let cost = parseFloat(lotData[i][3]) || 0;
    if(!costAgg[pid]) costAgg[pid] = { totalQty: 0, totalValue: 0 };
    costAgg[pid].totalQty += qty;
    costAgg[pid].totalValue += (qty * cost);
  }
  for(let pid in costAgg) {
    if(costAgg[pid].totalQty > 0 && prodMap[pid]) {
      prodMap[pid].avgCost = costAgg[pid].totalValue / costAgg[pid].totalQty;
    }
  }

  // แปลงค่าวันที่สำหรับกรองข้อมูล
  let sDate = startDate ? new Date(startDate) : new Date('2000-01-01');
  let eDate = endDate ? new Date(endDate) : new Date('2100-01-01');
  eDate.setHours(23, 59, 59); // ให้ครอบคลุมถึงสิ้นวัน

  let results = [];

  // 📝 1. รายงานสรุปยอดคงเหลือ
  if (reportType === 'balance') {
    let balanceMap = {};
    for(let i = 1; i < lotData.length; i++) {
      let pid = lotData[i][1];
      let remain = parseFloat(lotData[i][2]) || 0;
      if(!balanceMap[pid]) balanceMap[pid] = 0;
      balanceMap[pid] += remain;
    }
    for(let pid in balanceMap) {
      if(balanceMap[pid] > 0) {
        let pInfo = prodMap[pid] || { name: "ไม่พบชื่อ", unit: "-" };
        let val = balanceMap[pid] * (pInfo.avgCost || 0);
        results.push([pid, pInfo.name, balanceMap[pid], pInfo.unit, val.toFixed(2)]);
      }
    }
  } 
  
  // 📝 2. รายงานรับเข้า
  else if (reportType === 'inbound') {
    for(let i = 1; i < txData.length; i++) {
      let rDate = new Date(txData[i][0]);
      let type = txData[i][1].toString().trim();
      if ((type === "รับ" || type === "รับเข้า") && rDate >= sDate && rDate <= eDate) {
        let pid = txData[i][2];
        let pInfo = prodMap[pid] || { name: "ไม่พบชื่อ", unit: "-" };
        results.push([Utilities.formatDate(rDate, "GMT+7", "dd/MM/yyyy"), pid, pInfo.name, txData[i][3], pInfo.unit, txData[i][4]]);
      }
    }
  }
  
  // 📝 3. รายงานสรุปจำนวนเบิกตามแผนก
  else if (reportType === 'withdraw_qty') {
    let deptMap = {};
    for(let i = 1; i < txData.length; i++) {
      let rDate = new Date(txData[i][0]);
      let type = txData[i][1].toString().trim().toUpperCase();
      let status = txData[i][10] ? txData[i][10].toString().trim() : "";
      
      if ((type === "WITHDRAW" || type === "เบิก") && status === "สำเร็จ" && rDate >= sDate && rDate <= eDate) {
        let dept = txData[i][5] || "ไม่ระบุแผนก";
        let pid = txData[i][2];
        let qty = parseFloat(txData[i][3]) || 0;
        
        let key = dept + "|" + pid;
        if(!deptMap[key]) deptMap[key] = { dept: dept, pid: pid, qty: 0 };
        deptMap[key].qty += qty;
      }
    }
    for(let key in deptMap) {
      let item = deptMap[key];
      let pInfo = prodMap[item.pid] || { name: "ไม่พบชื่อ", unit: "-" };
      results.push([item.dept, item.pid, pInfo.name, item.qty, pInfo.unit]);
    }
    // เรียงตามแผนก
    results.sort((a, b) => a[0].localeCompare(b[0]));
  }
  
  // 📝 4. รายงานสรุปมูลค่าการเบิกรวม
  else if (reportType === 'withdraw_value') {
    let deptValMap = {};
    for(let i = 1; i < txData.length; i++) {
      let rDate = new Date(txData[i][0]);
      let type = txData[i][1].toString().trim().toUpperCase();
      let status = txData[i][10] ? txData[i][10].toString().trim() : "";
      
      if ((type === "WITHDRAW" || type === "เบิก") && status === "สำเร็จ" && rDate >= sDate && rDate <= eDate) {
        let dept = txData[i][5] || "ไม่ระบุแผนก";
        let pid = txData[i][2];
        let qty = parseFloat(txData[i][3]) || 0;
        let avgCost = prodMap[pid] ? (prodMap[pid].avgCost || 0) : 0;
        let value = qty * avgCost;
        
        if(!deptValMap[dept]) deptValMap[dept] = { items: 0, value: 0 };
        deptValMap[dept].items += qty;
        deptValMap[dept].value += value;
      }
    }
    for(let dept in deptValMap) {
      results.push([dept, deptValMap[dept].items, deptValMap[dept].value.toFixed(2)]);
    }
  }

  return { success: true, data: results, reportType: reportType };
}

// --- 🚨 ตั้งค่า LINE Messaging API 🚨 ---
const LINE_ACCESS_TOKEN = 'ใส่_CHANNEL_ACCESS_TOKEN_ของคุณที่นี่';

function notifyMultiChannel(targetEmail, subject, message) {
  // 1. ส่ง Email ผ่าน Service ของ Google
  try {
    MailApp.sendEmail({
      to: targetEmail,
      subject: "🔔 Stock Elite: " + subject,
      htmlBody: `
        <div style="font-family: 'Sarabun', sans-serif; border: 1px solid #e2e8f0; padding: 20px; border-radius: 10px;">
          <h2 style="color: #38bdf8;">Stock Elite System</h2>
          <p style="font-size: 16px;">${message.replace(/\n/g, '<br>')}</p>
          <hr style="border: 0; border-top: 1px solid #eee;">
          <small style="color: #94a3b8;">กรุณาเข้าสู่ระบบเพื่อดำเนินการตรวจสอบ</small>
        </div>`
    });
  } catch(e) { console.error("Email Error: " + e.message); }

  // 2. ส่ง LINE Messaging API (Broadcast หรือ Push)
  // หมายเหตุ: การส่งหาบุคคลเฉพาะเจาะจงต้องใช้ USER_ID แต่ในที่นี้จะใช้ Broadcast เพื่อทดสอบง่ายๆ
  const url = "https://api.line.me/v2/bot/message/broadcast";
  const options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_ACCESS_TOKEN
    },
    "payload": JSON.stringify({
      "messages": [{ "type": "text", "text": "📦 Stock Elite Alert!\n" + message }]
    })
  };
  
  try { UrlFetchApp.fetch(url, options); } catch(e) { console.error("Line Error: " + e.message); }
}

function processReturnDetailed(borrowId, condition, remark) {
  const sheet = SS.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // ค้นหาแถวที่มี BorrowID ตรงกัน (สมมติเก็บในคอลัมน์ I)
    if (data[i][8] == borrowId) {
      const finalStatus = condition === 'ปกติ' ? 'คืนแล้ว' : 'คืนแล้ว (' + condition + ')';
      sheet.getRange(i + 1, 11).setValue(finalStatus); // อัปเดต Column K (Status)
      sheet.getRange(i + 1, 10).setValue(remark);      // อัปเดต Column J (Remark)
      
      // คืนสต๊อกกลับเข้าไปถ้าของไม่ได้หาย
      if (condition !== 'สูญหาย') {
        const pid = data[i][2];
        const qty = data[i][3];
        const lotSheet = SS.getSheetByName("Inventory_Lots");
        lotSheet.appendRow(["RET-" + borrowId, pid, qty, 0, new Date()]);
      }
      
      return { success: true, message: "บันทึกการคืนอุปกรณ์เรียบร้อย" };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลรายการยืม" };
}

// ==========================================
// ระบบจัดการผู้อนุมัติ (Approvers)
// ==========================================

// 1. ดึงข้อมูลผู้อนุมัติทั้งหมด
function getApproversList() {
  const sheet = SS.getSheetByName("Approvers");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      result.push({ dept: data[i][0], name: data[i][1], email: data[i][2] });
    }
  }
  return result;
}

// 2. เพิ่มหรือแก้ไขผู้อนุมัติ
function saveApprover(oldDept, newDept, name, email) {
  const sheet = SS.getSheetByName("Approvers");
  const data = sheet.getDataRange().getValues();
  
  // กรณีแก้ไข (มี oldDept ส่งมา)
  if (oldDept) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == oldDept) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([[newDept, name, email]]);
        return { success: true, message: "อัปเดตข้อมูลผู้อนุมัติสำเร็จ" };
      }
    }
  }
  
  // กรณีเพิ่มใหม่: เช็คว่าแผนกนี้มีคนอนุมัติหรือยัง
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == newDept) {
      return { success: false, message: "แผนกนี้มีผู้อนุมัติแล้ว กรุณาใช้ปุ่มแก้ไขแทนครับ" };
    }
  }
  
  sheet.appendRow([newDept, name, email]);
  return { success: true, message: "เพิ่มผู้อนุมัติสำเร็จ" };
}

// 3. ลบผู้อนุมัติ
function deleteApproverData(dept) {
  const sheet = SS.getSheetByName("Approvers");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == dept) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "ลบข้อมูลสำเร็จ" };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลที่ต้องการลบ" };
}

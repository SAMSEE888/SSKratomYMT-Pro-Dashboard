
// ID ของ Google Sheet ที่ต้องการบันทึกข้อมูล
const SHEET_ID = '11vhg37MbHRm53SSEHLsCI3EBXx5_meXVvlRuqhFteaY';
// ชื่อของชีต (แท็บ) ที่ต้องการบันทึกข้อมูล
const SHEET_NAME = 'SaleForm';

// ฟังก์ชันหลักที่เรียกเมื่อเว็บแอปเปิด
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SSKratomYMT Pro Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันบันทึกข้อมูลจากฟอร์ม
function doPost(postData) {
  try {
    const data = JSON.parse(postData.postData.contents);
    
    // เปิด Spreadsheet
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    // หากชีตว่างเปล่า ให้เพิ่มหัวข้อ
    if (sheet.getLastRow() === 0) {
      const headers = [
        'วันที่', 
        'จำนวนที่ขาย (ขวด)', 
        'ค้างน้ำดิบ (ขวด)', 
        'เคลียร์ยอดค้าง (ขวด)',
        'รายได้',
        'ค่าท่อม',
        'ค่าแชร์',
        'ค่าใช้จ่ายอื่น',
        'เก็บออมเงิน',
        'รายจ่าย',
        'ยอดคงเหลือ (กำไร)',
        'เวลาบันทึก'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    // เตรียมข้อมูลสำหรับบันทึก
    const newRow = [
      new Date(data.date), // วันที่
      data.sold,          // จำนวนที่ขาย
      data.pending,       // ค้างน้ำดิบ
      data.cleared,       // เคลียร์ยอดค้าง
      data.revenue,       // รายได้
      data.pipeFee,       // ค่าท่อม
      data.shareFee,      // ค่าแชร์
      data.otherFee,      // ค่าใช้จ่ายอื่น
      data.saveFee,       // เก็บออมเงิน
      data.expense,       // รายจ่าย
      data.balance,       // ยอดคงเหลือ
      new Date()          // เวลาบันทึก
    ];
    
    // บันทึกข้อมูลลงแถวใหม่
    sheet.appendRow(newRow);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ฟังก์ชันดึงข้อมูลทั้งหมดสำหรับแดชบอร์ด
function getData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    // หากไม่มีข้อมูล
    if (sheet.getLastRow() <= 1) {
      return JSON.stringify([]);
    }
    
    // ดึงข้อมูลทั้งหมด (ข้ามหัวข้อ)
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11);
    const data = dataRange.getValues();
    
    // แปลงข้อมูลเป็น object
    const result = data.map(row => {
      // แปลงวันที่จาก Google Sheets format
      let dateObj;
      if (row[0] instanceof Date) {
        dateObj = row[0];
      } else {
        // หากเป็น string ให้แปลงเป็น Date object
        dateObj = new Date(row[0]);
      }
      
      return {
        date: dateObj.toISOString().split('T')[0], // เก็บเป็น YYYY-MM-DD
        sold: Number(row[1]) || 0,
        pending: Number(row[2]) || 0,
        cleared: Number(row[3]) || 0,
        revenue: Number(row[4]) || 0,
        pipeFee: Number(row[5]) || 0,
        shareFee: Number(row[6]) || 0,
        otherFee: Number(row[7]) || 0,
        saveFee: Number(row[8]) || 0,
        expense: Number(row[9]) || 0,
        balance: Number(row[10]) || 0
      };
    });
    
    return JSON.stringify(result);
    
  } catch (error) {
    return JSON.stringify({
      error: error.message
    });
  }
}

// ฟังก์ชันดาวน์โหลดข้อมูลเป็น CSV
function getSheetDataAsCsv() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    // หากไม่มีข้อมูล
    if (sheet.getLastRow() === 0) {
      return "ไม่มีข้อมูล";
    }
    
    // ดึงข้อมูลทั้งหมด
    const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 11);
    const data = dataRange.getValues();
    
    // แปลงเป็น CSV
    const csvContent = data.map(row => {
      // จัดรูปแบบวันที่ให้เหมาะสม
      if (row[0] instanceof Date) {
        row[0] = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
      return row.map(field => {
        // ใส่เครื่องหมายคำพูดหากมี comma ใน field
        if (typeof field === 'string' && field.includes(',')) {
          return `"${field}"`;
        }
        return field;
      }).join(',');
    }).join('\n');
    
    return csvContent;
    
  } catch (error) {
    return "Error: " + error.message;
  }
}



// ฟังก์ชันสร้างชีตรายงานสรุปทั้งหมด
function createSummaryReports() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // สร้างชีตรายงานสรุปรายวัน
    createDailySummary(ss);
    
    // สร้างชีตรายงานสรุปรายสัปดาห์
    createWeeklySummary(ss);
    
    // สร้างชีตรายงานสรุปรายเดือน
    createMonthlySummary(ss);
    
    // สร้างชีตรายงานสรุปรายปี
    createYearlySummary(ss);
    
    // สร้างชีตรายงานสรุปตามหมวดหมู่ค่าใช้จ่าย
    createExpenseCategorySummary(ss);
    
    // สร้างชีตรายงานเปรียบเทียบ
    createComparisonReport(ss);
    
    // สร้างชีตสถิติและแนวโน้ม
    createTrendAnalysis(ss);
    
    return JSON.stringify({
      success: true,
      message: 'สร้างชีตรายงานสรุปทั้งหมดเรียบร้อยแล้ว'
    });
    
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error.message
    });
  }
}

// สร้างรายงานสรุปรายวัน
function createDailySummary(ss) {
  const sheetName = 'รายงานสรุปรายวัน';
  let summarySheet = ss.getSheetByName(sheetName);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
  } else {
    summarySheet.clear();
  }
  
  // หัวข้อรายงาน
  const headers = [
    'วันที่',
    'ยอดขายรวม (ขวด)',
    'รายได้รวม',
    'รายจ่ายรวม',
    'กำไรสุทธิ',
    'อัตรากำไร (%)',
    'ค่าใช้จ่ายเฉลี่ย/ขวด',
    'กำไรเฉลี่ย/ขวด',
    'จำนวนวันทำการ'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  // ดึงข้อมูลจากชีตหลัก
  const dataSheet = ss.getSheetByName(SHEET_NAME);
  if (dataSheet.getLastRow() <= 1) return;
  
  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11).getValues();
  
  // จัดกลุ่มข้อมูลตามวัน
  const dailyData = {};
  data.forEach(row => {
    const date = new Date(row[0]);
    const dateKey = date.toISOString().split('T')[0];
    
    if (!dailyData[dateKey]) {
      dailyData[dateKey] = {
        date: date,
        sold: 0,
        revenue: 0,
        expense: 0,
        balance: 0,
        count: 0
      };
    }
    
    dailyData[dateKey].sold += Number(row[1]) || 0;
    dailyData[dateKey].revenue += Number(row[4]) || 0;
    dailyData[dateKey].expense += Number(row[9]) || 0;
    dailyData[dateKey].balance += Number(row[10]) || 0;
    dailyData[dateKey].count++;
  });
  
  // คำนวณและบันทึกข้อมูลสรุป
  const summaryData = [];
  Object.keys(dailyData).sort().forEach(dateKey => {
    const day = dailyData[dateKey];
    const profitMargin = day.revenue > 0 ? (day.balance / day.revenue) * 100 : 0;
    const avgExpensePerBottle = day.sold > 0 ? day.expense / day.sold : 0;
    const avgProfitPerBottle = day.sold > 0 ? day.balance / day.sold : 0;
    
    summaryData.push([
      Utilities.formatDate(day.date, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      day.sold,
      day.revenue,
      day.expense,
      day.balance,
      `${profitMargin.toFixed(2)}%`,
      avgExpensePerBottle.toFixed(2),
      avgProfitPerBottle.toFixed(2),
      day.count
    ]);
  });
  
  // บันทึกข้อมูล
  if (summaryData.length > 0) {
    summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    
    // คำนวณรวมทั้งหมด
    const totalRow = summaryData.length + 2;
    summarySheet.getRange(totalRow, 1, 1, headers.length)
      .setValues([[
        'รวมทั้งหมด',
        `=SUM(B2:B${totalRow-1})`,
        `=SUM(C2:C${totalRow-1})`,
        `=SUM(D2:D${totalRow-1})`,
        `=SUM(E2:E${totalRow-1})`,
        `=IF(C${totalRow}>0, E${totalRow}/C${totalRow}*100, 0)`,
        `=IF(B${totalRow}>0, D${totalRow}/B${totalRow}, 0)`,
        `=IF(B${totalRow}>0, E${totalRow}/B${totalRow}, 0)`,
        `=SUM(I2:I${totalRow-1})`
      ]])
      .setFontWeight('bold')
      .setBackground('#e6f3ff');
    
    // จัดรูปแบบ
    summarySheet.setColumnWidths(1, headers.length, 120);
    summarySheet.getRange('C:C').setNumberFormat('#,##0.00');
    summarySheet.getRange('D:D').setNumberFormat('#,##0.00');
    summarySheet.getRange('E:E').setNumberFormat('#,##0.00');
    summarySheet.getRange('F:F').setNumberFormat('0.00%');
    summarySheet.getRange('G:G').setNumberFormat('#,##0.00');
    summarySheet.getRange('H:H').setNumberFormat('#,##0.00');
    
    // เพิ่มการจัดรูปแบบเงื่อนไขสำหรับอัตรากำไร
    const profitMarginRange = summarySheet.getRange(`F2:F${totalRow}`);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(20)
      .setBackground('#c6efce')
      .setRanges([profitMarginRange])
      .build();
    const rules = summarySheet.getConditionalFormatRules();
    rules.push(rule);
    summarySheet.setConditionalFormatRules(rules);
  }
}

// สร้างรายงานสรุปรายสัปดาห์
function createWeeklySummary(ss) {
  const sheetName = 'รายงานสรุปรายสัปดาห์';
  let summarySheet = ss.getSheetByName(sheetName);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
  } else {
    summarySheet.clear();
  }
  
  const headers = [
    'สัปดาห์',
    'ปี',
    'ยอดขายรวม (ขวด)',
    'รายได้รวม',
    'รายจ่ายรวม',
    'กำไรสุทธิ',
    'อัตรากำไร (%)',
    'ยอดขายเฉลี่ย/วัน',
    'รายได้เฉลี่ย/วัน',
    'กำไรเฉลี่ย/วัน',
    'วันทำการ'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  const dataSheet = ss.getSheetByName(SHEET_NAME);
  if (dataSheet.getLastRow() <= 1) return;
  
  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11).getValues();
  
  // จัดกลุ่มข้อมูลตามสัปดาห์
  const weeklyData = {};
  data.forEach(row => {
    const date = new Date(row[0]);
    const weekNumber = getWeekNumber(date);
    const year = date.getFullYear();
    const weekKey = `${year}-W${weekNumber}`;
    
    if (!weeklyData[weekKey]) {
      weeklyData[weekKey] = {
        year: year,
        week: weekNumber,
        sold: 0,
        revenue: 0,
        expense: 0,
        balance: 0,
        days: new Set(),
        count: 0
      };
    }
    
    weeklyData[weekKey].sold += Number(row[1]) || 0;
    weeklyData[weekKey].revenue += Number(row[4]) || 0;
    weeklyData[weekKey].expense += Number(row[9]) || 0;
    weeklyData[weekKey].balance += Number(row[10]) || 0;
    weeklyData[weekKey].days.add(date.toISOString().split('T')[0]);
    weeklyData[weekKey].count++;
  });
  
  // คำนวณและบันทึกข้อมูลสรุป
  const summaryData = [];
  Object.keys(weeklyData).sort().forEach(weekKey => {
    const week = weeklyData[weekKey];
    const profitMargin = week.revenue > 0 ? (week.balance / week.revenue) * 100 : 0;
    const workingDays = week.days.size;
    
    summaryData.push([
      `สัปดาห์ ${week.week}`,
      week.year,
      week.sold,
      week.revenue,
      week.expense,
      week.balance,
      profitMargin.toFixed(2),
      (week.sold / workingDays).toFixed(1),
      (week.revenue / workingDays).toFixed(2),
      (week.balance / workingDays).toFixed(2),
      workingDays
    ]);
  });
  
  if (summaryData.length > 0) {
    summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    
    // จัดรูปแบบ
    summarySheet.setColumnWidths(1, headers.length, 120);
    summarySheet.getRange('D:F').setNumberFormat('#,##0.00');
    summarySheet.getRange('G:G').setNumberFormat('0.00%');
    summarySheet.getRange('H:J').setNumberFormat('#,##0.00');
  }
}

// สร้างรายงานสรุปรายเดือน
function createMonthlySummary(ss) {
  const sheetName = 'รายงานสรุปรายเดือน';
  let summarySheet = ss.getSheetByName(sheetName);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
  } else {
    summarySheet.clear();
  }
  
  const headers = [
    'เดือน',
    'ปี',
    'ยอดขายรวม (ขวด)',
    'รายได้รวม',
    'รายจ่ายรวม',
    'กำไรสุทธิ',
    'อัตรากำไร (%)',
    'ยอดขายเฉลี่ย/วัน',
    'รายได้เฉลี่ย/วัน',
    'กำไรเฉลี่ย/วัน',
    'วันทำการ',
    'เปรียบเทียบกับเดือนก่อน (%)'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  const dataSheet = ss.getSheetByName(SHEET_NAME);
  if (dataSheet.getLastRow() <= 1) return;
  
  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11).getValues();
  
  // จัดกลุ่มข้อมูลตามเดือน
  const monthlyData = {};
  data.forEach(row => {
    const date = new Date(row[0]);
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
    
    if (!monthlyData[monthKey]) {
      monthlyData[monthKey] = {
        year: year,
        month: month,
        sold: 0,
        revenue: 0,
        expense: 0,
        balance: 0,
        days: new Set(),
        count: 0
      };
    }
    
    monthlyData[monthKey].sold += Number(row[1]) || 0;
    monthlyData[monthKey].revenue += Number(row[4]) || 0;
    monthlyData[monthKey].expense += Number(row[9]) || 0;
    monthlyData[monthKey].balance += Number(row[10]) || 0;
    monthlyData[monthKey].days.add(date.toISOString().split('T')[0]);
    monthlyData[monthKey].count++;
  });
  
  // คำนวณและบันทึกข้อมูลสรุป
  const summaryData = [];
  const sortedMonths = Object.keys(monthlyData).sort();
  
  sortedMonths.forEach((monthKey, index) => {
    const month = monthlyData[monthKey];
    const profitMargin = month.revenue > 0 ? (month.balance / month.revenue) * 100 : 0;
    const workingDays = month.days.size;
    
    // คำนวณการเปรียบเทียบกับเดือนก่อน
    let growthRate = 'N/A';
    if (index > 0) {
      const prevMonth = monthlyData[sortedMonths[index - 1]];
      if (prevMonth.balance !== 0) {
        const growth = ((month.balance - prevMonth.balance) / Math.abs(prevMonth.balance)) * 100;
        growthRate = growth.toFixed(1);
      }
    }
    
    const monthNames = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    
    summaryData.push([
      monthNames[month.month - 1],
      month.year,
      month.sold,
      month.revenue,
      month.expense,
      month.balance,
      profitMargin.toFixed(2),
      (month.sold / workingDays).toFixed(1),
      (month.revenue / workingDays).toFixed(2),
      (month.balance / workingDays).toFixed(2),
      workingDays,
      growthRate
    ]);
  });
  
  if (summaryData.length > 0) {
    summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    
    // จัดรูปแบบ
    summarySheet.setColumnWidths(1, headers.length, 120);
    summarySheet.getRange('D:F').setNumberFormat('#,##0.00');
    summarySheet.getRange('G:G').setNumberFormat('0.00%');
    summarySheet.getRange('H:J').setNumberFormat('#,##0.00');
    summarySheet.getRange('L:L').setNumberFormat('0.0"%";-0.0"%";"N/A"');
    
    // เพิ่มการจัดรูปแบบเงื่อนไขสำหรับการเติบโต
    const growthRange = summarySheet.getRange(`L2:L${summaryData.length + 1}`);
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#c6efce')
      .setRanges([growthRange])
      .build();
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#ffcccc')
      .setRanges([growthRange])
      .build();
    
    summarySheet.setConditionalFormatRules([positiveRule, negativeRule]);
  }
}

// สร้างรายงานสรุปตามหมวดหมู่ค่าใช้จ่าย
function createExpenseCategorySummary(ss) {
  const sheetName = 'สรุปค่าใช้จ่ายตามหมวดหมู่';
  let summarySheet = ss.getSheetByName(sheetName);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
  } else {
    summarySheet.clear();
  }
  
  const headers = [
    'หมวดหมู่',
    'ยอดรวม',
    'เปอร์เซ็นต์',
    'เฉลี่ย/วัน',
    'เฉลี่ย/ขวด'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  const dataSheet = ss.getSheetByName(SHEET_NAME);
  if (dataSheet.getLastRow() <= 1) return;
  
  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11).getValues();
  
  // คำนวณยอดรวมแต่ละหมวดหมู่
  const categories = {
    'ค่าท่อม': { total: 0, days: new Set() },
    'ค่าแชร์': { total: 0, days: new Set() },
    'ค่าใช้จ่ายอื่น': { total: 0, days: new Set() },
    'เก็บออมเงิน': { total: 0, days: new Set() }
  };
  
  let totalSold = 0;
  
  data.forEach(row => {
    const date = new Date(row[0]);
    const dateKey = date.toISOString().split('T')[0];
    
    categories['ค่าท่อม'].total += Number(row[5]) || 0;
    categories['ค่าแชร์'].total += Number(row[6]) || 0;
    categories['ค่าใช้จ่ายอื่น'].total += Number(row[7]) || 0;
    categories['เก็บออมเงิน'].total += Number(row[8]) || 0;
    
    categories['ค่าท่อม'].days.add(dateKey);
    categories['ค่าแชร์'].days.add(dateKey);
    categories['ค่าใช้จ่ายอื่น'].days.add(dateKey);
    categories['เก็บออมเงิน'].days.add(dateKey);
    
    totalSold += Number(row[1]) || 0;
  });
  
  // คำนวณเปอร์เซ็นต์และค่าเฉลี่ย
  const totalExpense = Object.values(categories).reduce((sum, cat) => sum + cat.total, 0);
  const totalDays = Math.max(...Object.values(categories).map(cat => cat.days.size));
  
  const summaryData = [];
  Object.keys(categories).forEach(category => {
    const cat = categories[category];
    const percentage = totalExpense > 0 ? (cat.total / totalExpense) * 100 : 0;
    const avgPerDay = cat.total / totalDays;
    const avgPerBottle = totalSold > 0 ? cat.total / totalSold : 0;
    
    summaryData.push([
      category,
      cat.total,
      percentage.toFixed(2),
      avgPerDay.toFixed(2),
      avgPerBottle.toFixed(2)
    ]);
  });
  
  if (summaryData.length > 0) {
    summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    
    // เพิ่มแถวรวม
    const totalRow = summaryData.length + 2;
    summarySheet.getRange(totalRow, 1, 1, headers.length)
      .setValues([[
        'รวมทั้งหมด',
        `=SUM(B2:B${totalRow-1})`,
        '100%',
        `=SUM(D2:D${totalRow-1})`,
        `=SUM(E2:E${totalRow-1})`
      ]])
      .setFontWeight('bold')
      .setBackground('#e6f3ff');
    
    // จัดรูปแบบ
    summarySheet.setColumnWidths(1, headers.length, 150);
    summarySheet.getRange('B:B').setNumberFormat('#,##0.00');
    summarySheet.getRange('C:C').setNumberFormat('0.00%');
    summarySheet.getRange('D:E').setNumberFormat('#,##0.00');
    
    // สร้างแผนภูมิวงกลม
    if (summaryData.length > 0) {
      const chartRange = summarySheet.getRange(`A2:B${summaryData.length + 1}`);
      const chart = summarySheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartRange)
        .setPosition(2, 7, 0, 0)
        .setOption('title', 'สัดส่วนค่าใช้จ่ายตามหมวดหมู่')
        .setOption('pieHole', 0.4)
        .setOption('is3D', true)
        .build();
      
      summarySheet.insertChart(chart);
    }
  }
}

// สร้างรายงานเปรียบเทียบ
function createComparisonReport(ss) {
  const sheetName = 'รายงานเปรียบเทียบ';
  let summarySheet = ss.getSheetByName(sheetName);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
  } else {
    summarySheet.clear();
  }
  
  const headers = [
    'ช่วงเวลา',
    'ยอดขาย (ขวด)',
    'รายได้',
    'รายจ่าย',
    'กำไร',
    'อัตรากำไร (%)',
    'ยอดขาย/วัน',
    'รายได้/วัน',
    'กำไร/วัน'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  const dataSheet = ss.getSheetByName(SHEET_NAME);
  if (dataSheet.getLastRow() <= 1) return;
  
  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11).getValues();
  
  // คำนวณข้อมูลสำหรับแต่ละช่วงเวลา
  const periods = {
    '7 วันที่ผ่านมา': getDateRange(7),
    '30 วันที่ผ่านมา': getDateRange(30),
    '3 เดือนที่ผ่านมา': getDateRange(90),
    '6 เดือนที่ผ่านมา': getDateRange(180),
    '1 ปีที่ผ่านมา': getDateRange(365),
    'ทั้งหมด': null
  };
  
  const summaryData = [];
  const now = new Date();
  
  Object.keys(periods).forEach(periodName => {
    const periodData = periods[periodName] ? 
      data.filter(row => {
        const rowDate = new Date(row[0]);
        return rowDate >= periods[periodName] && rowDate <= now;
      }) : data;
    
    if (periodData.length === 0) return;
    
    const totalSold = periodData.reduce((sum, row) => sum + (Number(row[1]) || 0), 0);
    const totalRevenue = periodData.reduce((sum, row) => sum + (Number(row[4]) || 0), 0);
    const totalExpense = periodData.reduce((sum, row) => sum + (Number(row[9]) || 0), 0);
    const totalProfit = totalRevenue - totalExpense;
    const profitMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;
    
    // คำนวณจำนวนวัน
    const days = periodName === 'ทั้งหมด' ? 
      new Set(periodData.map(row => new Date(row[0]).toISOString().split('T')[0])).size :
      Math.min((now - periods[periodName]) / (1000 * 60 * 60 * 24), periodData.length);
    
    summaryData.push([
      periodName,
      totalSold,
      totalRevenue,
      totalExpense,
      totalProfit,
      profitMargin.toFixed(2),
      (totalSold / days).toFixed(1),
      (totalRevenue / days).toFixed(2),
      (totalProfit / days).toFixed(2)
    ]);
  });
  
  if (summaryData.length > 0) {
    summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    
    // จัดรูปแบบ
    summarySheet.setColumnWidths(1, headers.length, 120);
    summarySheet.getRange('C:E').setNumberFormat('#,##0.00');
    summarySheet.getRange('F:F').setNumberFormat('0.00%');
    summarySheet.getRange('G:I').setNumberFormat('#,##0.00');
  }
}

// สร้างรายงานวิเคราะห์แนวโน้ม
function createTrendAnalysis(ss) {
  const sheetName = 'วิเคราะห์แนวโน้ม';
  let summarySheet = ss.getSheetByName(sheetName);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
  } else {
    summarySheet.clear();
  }
  
  const headers = [
    'เดือน',
    'ยอดขาย',
    'แนวโน้มยอดขาย',
    'รายได้',
    'แนวโน้มรายได้',
    'กำไร',
    'แนวโน้มกำไร',
    'อัตราการเติบโต (%)'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  // ดึงข้อมูลจากรายงานเดือน
  const monthlySheet = ss.getSheetByName('รายงานสรุปรายเดือน');
  if (!monthlySheet || monthlySheet.getLastRow() <= 1) return;
  
  const monthlyData = monthlySheet.getRange(2, 1, monthlySheet.getLastRow() - 1, 6).getValues();
  
  const trendData = [];
  monthlyData.forEach((row, index) => {
    const sales = row[2];
    const revenue = row[3];
    const profit = row[5];
    
    // คำนวณแนวโน้ม (ค่าเฉลี่ยเคลื่อนที่ 3 เดือน)
    let salesTrend = sales;
    let revenueTrend = revenue;
    let profitTrend = profit;
    
    if (index >= 2) {
      salesTrend = (monthlyData[index][2] + monthlyData[index-1][2] + monthlyData[index-2][2]) / 3;
      revenueTrend = (monthlyData[index][3] + monthlyData[index-1][3] + monthlyData[index-2][3]) / 3;
      profitTrend = (monthlyData[index][5] + monthlyData[index-1][5] + monthlyData[index-2][5]) / 3;
    }
    
    // คำนวณอัตราการเติบโต
    let growthRate = 'N/A';
    if (index > 0) {
      const prevProfit = monthlyData[index-1][5];
      if (prevProfit !== 0) {
        growthRate = ((profit - prevProfit) / Math.abs(prevProfit)) * 100;
      }
    }
    
    trendData.push([
      `${row[0]} ${row[1]}`,
      sales,
      salesTrend.toFixed(0),
      revenue,
      revenueTrend.toFixed(2),
      profit,
      profitTrend.toFixed(2),
      growthRate === 'N/A' ? growthRate : growthRate.toFixed(1)
    ]);
  });
  
  if (trendData.length > 0) {
    summarySheet.getRange(2, 1, trendData.length, headers.length).setValues(trendData);
    
    // จัดรูปแบบ
    summarySheet.setColumnWidths(1, headers.length, 130);
    summarySheet.getRange('C:C').setNumberFormat('#,##0');
    summarySheet.getRange('D:E').setNumberFormat('#,##0.00');
    summarySheet.getRange('F:G').setNumberFormat('#,##0.00');
    summarySheet.getRange('H:H').setNumberFormat('0.0"%";-0.0"%";"N/A"');
    
    // เพิ่มแผนภูมิแนวโน้ม
    if (trendData.length >= 3) {
      const chartRange = summarySheet.getRange(`A2:A${trendData.length + 1}`);
      const profitRange = summarySheet.getRange(`F2:G${trendData.length + 1}`);
      
      const chart = summarySheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartRange)
        .addRange(profitRange)
        .setPosition(2, 9, 0, 0)
        .setOption('title', 'แนวโน้มกำไร (เปรียบเทียบกับค่าเฉลี่ยเคลื่อนที่)')
        .setOption('hAxis', { title: 'เดือน' })
        .setOption('vAxis', { title: 'จำนวนเงิน (บาท)' })
        .setOption('series', {
          0: { labelInLegend: 'กำไรจริง' },
          1: { labelInLegend: 'แนวโน้มกำไร' }
        })
        .build();
      
      summarySheet.insertChart(chart);
    }
  }
}

// ฟังก์ชันช่วยเหลือ
function getWeekNumber(date) {
  const firstDayOfYear = new Date(date.getFullYear(), 0, 1);
  const pastDaysOfYear = (date - firstDayOfYear) / 86400000;
  return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
}

function getDateRange(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return date;
}

// ฟังก์ชันเรียกใช้งานจากเว็บแอป
function generateAllReports() {
  try {
    createSummaryReports();
    return JSON.stringify({
      success: true,
      message: 'สร้างรายงานสรุปทั้งหมดเรียบร้อยแล้ว'
    });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error.message
    });
  }
}
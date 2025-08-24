// =================================================================
// UNIVERSAL DATA READER - PHIÊN BẢN NÂNG CẤP
// =================================================================

/**
 * Hàm đọc dữ liệu chính, có thể đọc từ các sheet có tiền tố khác nhau.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Đối tượng spreadsheet.
 * @param {string} sheetPrefix Tiền tố của sheet cần đọc (ví dụ: 'DL_NHAP', 'DL_XUAT').
 * @param {Array<string>} requiredColumns Mảng các cột bắt buộc phải có.
 * @returns {Array<Object>} Mảng các đối tượng dữ liệu đã được xử lý.
 */
function readDataFromPrefixedSheets(spreadsheet, sheetPrefix, requiredColumns) {
  const allSheets = spreadsheet.getSheets();
  const dataSheets = allSheets.filter(sheet => sheet.getName().startsWith(sheetPrefix));
  
  if (dataSheets.length === 0) {
    console.log(`Không tìm thấy sheet nào bắt đầu với "${sheetPrefix}"`);
    return [];
  }

  const combinedData = [];
  for (const sheet of dataSheets) {
    const sheetData = processSingleSheet(sheet, requiredColumns);
    if (sheetData.length > 0) {
      combinedData.push(...sheetData);
    }
  }
  return combinedData;
}

/**
 * Xử lý dữ liệu cho một sheet duy nhất.
 * PHIÊN BẢN NÂNG CẤP: Sẽ bỏ qua hoàn toàn các dòng không có dữ liệu ở cột NGAY_HT.
 */
function processSingleSheet(sheet, requiredColumns) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headerRow = data[0].map(h => h.toString().trim().toUpperCase());
  
  // Tìm vị trí của cột NGAY_HT
  const colIndexNgayHT = headerRow.indexOf('NGAY_HT');

  // Kiểm tra xem có đủ các cột bắt buộc không
  const missingCols = requiredColumns.filter(col => !headerRow.includes(col));
  if (missingCols.length > 0) {
    console.error(`Sheet "${sheet.getName()}" thiếu các cột bắt buộc: ${missingCols.join(', ')}`);
    return [];
  }

  const processedData = [];
  // Lặp từ dòng 2 (index = 1) để bỏ qua tiêu đề
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // **SỬA LỖI**: Kiểm tra điều kiện NGAY_HT trước tiên
    // Lấy giá trị ngày và kiểm tra. Nếu ô rỗng, null, hoặc undefined -> bỏ qua dòng này
    const ngayHTValue = (colIndexNgayHT !== -1) ? row[colIndexNgayHT] : null;
    if (!ngayHTValue) {
      continue; // Bỏ qua dòng này và chuyển sang dòng tiếp theo
    }
    
    const rowData = {
      sheet: sheet.getName(),
      row: i + 1
    };
    
    headerRow.forEach((header, index) => {
      rowData[header] = row[index];
    });
    
    processedData.push(rowData);
  }
  return processedData;
}

/**
 * =================================================================
 * MODULE TÍNH GIÁ XUẤT KHO - PHIÊN BẢN 2.2 (Sửa lỗi & Ghi đủ cột)
 * =================================================================
 */

/**
 * HÀM TRUNG TÂM: Điều phối việc tính giá xuất kho.
 */

function tinhGiaXuatKho(phuongPhap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const validMethods = {
    'BQGQ_THANG': 'Bình quân gia quyền TỰ ĐỘNG THEO THÁNG',
    'BQDD': 'Bình quân di động (sau mỗi lần nhập)',
    'FIFO': 'Nhập trước, Xuất trước',
    'LIFO': 'Nhập sau, Xuất trước'
  };

  if (!validMethods[phuongPhap]) {
    ui.alert('Lỗi', 'Phương pháp không hợp lệ.', ui.ButtonSet.OK);
    return;
  }
  const tenPhuongPhap = validMethods[phuongPhap];
  
  const confirmation = ui.alert('Xác nhận', `Thực hiện tính giá xuất kho theo phương pháp:\n"${tenPhuongPhap}"?\n\nThao tác này sẽ GHI ĐÈ dữ liệu lên cột ĐƠN GIÁ và SỐ TIỀN của các sheet "DL_XUAT".`, ui.ButtonSet.YES_NO);
  if (confirmation !== ui.Button.YES) return;

  ss.toast('Bắt đầu quá trình...', 'Vui lòng chờ', -1);
  try {
    // --- BƯỚC 1: Tải và phân loại dữ liệu ---
    ss.toast('Đang đọc dữ liệu...', 'Bước 1/3');
    
    const sheetDMHH = ss.getSheetByName('DMHH');
    if (!sheetDMHH) throw new Error('Không tìm thấy sheet "DMHH"');
    const dataDMHH = sheetDMHH.getDataRange().getValues();
    const tonDauKyMap = new Map();
    dataDMHH.slice(1).forEach(row => {
      const maKho = row[0]?.toString().trim();
      const maHang = row[1]?.toString().trim();
      if (maKho && maHang) {
        tonDauKyMap.set(`${maKho}|${maHang}`, { sl: parseFloat(row[5]) || 0, gt: parseFloat(row[6]) || 0 });
      }
    });

    const transactionsNhap = readDataFromPrefixedSheets(ss, 'DL_NHAP', ['NGAY_HT', 'MA_KHO', 'MA_HANG', 'SO_LUONG', 'SO_TIEN']).map(t => ({...t, type: 'NHAP'}));
    const transactionsXuat = readDataFromPrefixedSheets(ss, 'DL_XUAT', ['NGAY_HT', 'MA_KHO', 'MA_HANG', 'SO_LUONG']).map(t => ({...t, type: 'XUAT'}));
    
    const allTransactions = [...transactionsNhap, ...transactionsXuat]
      .map(t => ({...t, NGAY_HT: new Date(t.NGAY_HT)}))
      .sort((a, b) => a.NGAY_HT - b.NGAY_HT);

    const itemsMap = new Map();
    for (const trans of allTransactions) {
      const key = `${trans.MA_KHO}|${trans.MA_HANG}`;
      if (!itemsMap.has(key)) {
        itemsMap.set(key, { tonDauKy: tonDauKyMap.get(key) || { sl: 0, gt: 0 }, transactions: [] });
      }
      itemsMap.get(key).transactions.push(trans);
    }

    // --- BƯỚC 2: Tính toán giá xuất kho ---
    ss.toast('Đang tính toán đơn giá...', 'Bước 2/3');
    const xuatKhoUpdates = [];
    for (const [key, itemData] of itemsMap.entries()) {
      let calculatedExports = [];
      switch (phuongPhap) {
        case 'BQGQ_THANG':
          calculatedExports = tinhGia_BinhQuanThang_TuDong(itemData);
          break;
        case 'BQDD':
          calculatedExports = tinhGia_BinhQuanDiDong(itemData);
          break;
        case 'FIFO':
          calculatedExports = tinhGia_FIFO(itemData);
          break;
        case 'LIFO':
          calculatedExports = tinhGia_LIFO(itemData);
          break;
      }
      if (calculatedExports.length > 0) xuatKhoUpdates.push(...calculatedExports);
    }

    // --- BƯỚC 3: Ghi lại dữ liệu ---
    ss.toast(`Đang cập nhật ${xuatKhoUpdates.length} giao dịch...`, 'Bước 3/3');
    const updatesBySheet = new Map();
    xuatKhoUpdates.forEach(u => {
      if (!updatesBySheet.has(u.sheet)) updatesBySheet.set(u.sheet, []);
      updatesBySheet.get(u.sheet).push(u);
    });

    for (const [sheetName, updates] of updatesBySheet.entries()) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
        const colIndexDonGia = headerRow.indexOf('DON_GIA');
        const colIndexSoTien = headerRow.indexOf('SO_TIEN');

        if (colIndexDonGia !== -1 && colIndexSoTien !== -1) {
          updates.forEach(u => {
            sheet.getRange(u.row, colIndexDonGia + 1).setValue(u.DON_GIA);
            sheet.getRange(u.row, colIndexSoTien + 1).setValue(u.SO_TIEN);
          });
        }
      }
    }
    
    ss.toast('Hoàn thành!', 'Thành công', 5);
    ui.alert('Thành công', `Đã cập nhật giá trị cho ${xuatKhoUpdates.length} giao dịch xuất kho.`, ui.ButtonSet.OK);

  } catch (e) {
    console.error("LỖI TÍNH GIÁ XUẤT KHO: " + e.toString() + e.stack);
    ss.toast('Gặp lỗi, vui lòng xem Logs.', 'Lỗi', 10);
    ui.alert('Lỗi', 'Quá trình tính giá gặp lỗi: ' + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 1. Tính giá theo phương pháp Bình quân gia quyền TỰ ĐỘNG THEO THÁNG
 */
function tinhGia_BinhQuanThang_TuDong(itemData) {
    const exportsToUpdate = [];
    if(itemData.transactions.length === 0) return [];

    const transByMonth = {};
    for(const trans of itemData.transactions) {
        const monthKey = `${trans.NGAY_HT.getFullYear()}-${trans.NGAY_HT.getMonth()}`;
        if (!transByMonth[monthKey]) transByMonth[monthKey] = [];
        transByMonth[monthKey].push(trans);
    }

    const sortedMonths = Object.keys(transByMonth).sort();
    let tonDauKySL = itemData.tonDauKy.sl;
    let tonDauKyGT = itemData.tonDauKy.gt;

    for (const monthKey of sortedMonths) {
        const monthTransactions = transByMonth[monthKey];
        let nhapTrongThangSL = 0;
        let nhapTrongThangGT = 0;
        
        monthTransactions.forEach(trans => {
            // **SỬA LỖI**: Dùng thuộc tính 'type' để phân biệt
            if (trans.type === 'NHAP') {
                nhapTrongThangSL += parseFloat(trans.SO_LUONG) || 0;
                nhapTrongThangGT += parseFloat(trans.SO_TIEN) || 0;
            }
        });

        const tongSL = tonDauKySL + nhapTrongThangSL;
        const tongGT = tonDauKyGT + nhapTrongThangGT;
        const donGiaBinhQuan = (tongSL > 0) ? (tongGT / tongSL) : 0;

        let xuatTrongThangSL = 0;
        monthTransactions.forEach(trans => {
            // **SỬA LỖI**: Dùng thuộc tính 'type' để phân biệt
            if (trans.type === 'XUAT') {
                const soLuongXuat = parseFloat(trans.SO_LUONG) || 0;
                const giaTriXuat = soLuongXuat * donGiaBinhQuan;
                exportsToUpdate.push({
                    ...trans,
                    DON_GIA: donGiaBinhQuan, // Thêm đơn giá vào kết quả
                    SO_TIEN: giaTriXuat
                });
                xuatTrongThangSL += soLuongXuat;
            }
        });

        tonDauKySL = tongSL - xuatTrongThangSL;
        tonDauKyGT = tongGT - (xuatTrongThangSL * donGiaBinhQuan);
    }
    return exportsToUpdate;
}

/**
 * 2. Tính giá theo phương pháp Bình quân di động (sau mỗi lần nhập)
 */
function tinhGia_BinhQuanDiDong(itemData) {
    const exportsToUpdate = [];
    let tonSL = itemData.tonDauKy.sl;
    let tonGT = itemData.tonDauKy.gt;
    let donGiaHienTai = (tonSL > 0) ? (tonGT / tonSL) : 0;

    for (const trans of itemData.transactions) {
        if (trans.type === 'NHAP') {
            tonSL += parseFloat(trans.SO_LUONG) || 0;
            tonGT += parseFloat(trans.SO_TIEN) || 0;
            // Cập nhật lại đơn giá bình quân ngay sau khi nhập
            donGiaHienTai = (tonSL > 0) ? (tonGT / tonSL) : 0;
        } else if (trans.type === 'XUAT') {
            const soLuongXuat = parseFloat(trans.SO_LUONG) || 0;
            const giaTriXuat = soLuongXuat * donGiaHienTai;
            exportsToUpdate.push({
                ...trans,
                DON_GIA: donGiaHienTai,
                SO_TIEN: giaTriXuat
            });
            // Giảm tồn kho
            tonSL -= soLuongXuat;
            tonGT -= giaTriXuat;
            // Đảm bảo giá trị tồn không bị âm do làm tròn
            if(tonSL <= 0) tonGT = 0;
        }
    }
    return exportsToUpdate;
}

/**
 * 3. Tính giá theo phương pháp Nhập trước, Xuất trước (FIFO)
 */
function tinhGia_FIFO(itemData) {
    const exportsToUpdate = [];
    const queueNhap = []; // Hàng đợi chứa các lô hàng nhập
    
    // Thêm tồn đầu kỳ vào hàng đợi như một lô hàng đầu tiên
    if (itemData.tonDauKy.sl > 0) {
        const donGiaDauKy = (itemData.tonDauKy.sl > 0) ? (itemData.tonDauKy.gt / itemData.tonDauKy.sl) : 0;
        queueNhap.push({ sl: itemData.tonDauKy.sl, donGia: donGiaDauKy });
    }

    for (const trans of itemData.transactions) {
        if (trans.type === 'NHAP') {
            const donGiaNhap = (trans.SO_LUONG > 0) ? (trans.SO_TIEN / trans.SO_LUONG) : 0;
            queueNhap.push({ sl: parseFloat(trans.SO_LUONG) || 0, donGia: donGiaNhap });
        } else if (trans.type === 'XUAT') {
            let slXuatCanXuLy = parseFloat(trans.SO_LUONG) || 0;
            let gtXuat = 0;
            
            while (slXuatCanXuLy > 0 && queueNhap.length > 0) {
                const loHang = queueNhap[0]; // Lấy lô hàng cũ nhất
                const slCoTheXuat = Math.min(slXuatCanXuLy, loHang.sl);
                
                gtXuat += slCoTheXuat * loHang.donGia;
                loHang.sl -= slCoTheXuat;
                slXuatCanXuLy -= slCoTheXuat;
                
                if (loHang.sl <= 0.00001) { // Sử dụng sai số để tránh lỗi làm tròn
                    queueNhap.shift(); // Xóa lô hàng đã hết
                }
            }
            const donGiaXuat = (trans.SO_LUONG > 0) ? gtXuat / trans.SO_LUONG : 0;
            exportsToUpdate.push({ ...trans, DON_GIA: donGiaXuat, SO_TIEN: gtXuat });
        }
    }
    return exportsToUpdate;
}

/**
 * 4. Tính giá theo phương pháp Nhập sau, Xuất trước (LIFO)
 */
function tinhGia_LIFO(itemData) {
    const exportsToUpdate = [];
    const stackNhap = []; // Ngăn xếp chứa các lô hàng nhập
    
    if (itemData.tonDauKy.sl > 0) {
        const donGiaDauKy = (itemData.tonDauKy.sl > 0) ? (itemData.tonDauKy.gt / itemData.tonDauKy.sl) : 0;
        stackNhap.push({ sl: itemData.tonDauKy.sl, donGia: donGiaDauKy });
    }

    for (const trans of itemData.transactions) {
        if (trans.type === 'NHAP') {
            const donGiaNhap = (trans.SO_LUONG > 0) ? (trans.SO_TIEN / trans.SO_LUONG) : 0;
            stackNhap.push({ sl: parseFloat(trans.SO_LUONG) || 0, donGia: donGiaNhap });
        } else if (trans.type === 'XUAT') {
            let slXuatCanXuLy = parseFloat(trans.SO_LUONG) || 0;
            let gtXuat = 0;
            
            while (slXuatCanXuLy > 0 && stackNhap.length > 0) {
                const loHang = stackNhap[stackNhap.length - 1]; // Lấy lô hàng mới nhất
                const slCoTheXuat = Math.min(slXuatCanXuLy, loHang.sl);
                
                gtXuat += slCoTheXuat * loHang.donGia;
                loHang.sl -= slCoTheXuat;
                slXuatCanXuLy -= slCoTheXuat;

                if (loHang.sl <= 0.00001) {
                    stackNhap.pop(); // Xóa lô hàng đã hết
                }
            }
            const donGiaXuat = (trans.SO_LUONG > 0) ? gtXuat / trans.SO_LUONG : 0;
            exportsToUpdate.push({ ...trans, DON_GIA: donGiaXuat, SO_TIEN: gtXuat });
        }
    }
    return exportsToUpdate;
}


// =================================================================
// SIDEBAR HÀNG HÓA
// =================================================================

/**
 * Mở sidebar chọn hàng hóa.
 */
function moSidebarHangHoa() {
  const html = HtmlService.createHtmlOutputFromFile('sidebarHangHoa')
    .setWidth(450)
    .setTitle('📦 Chọn Hàng hóa');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Lấy danh sách hàng hóa từ sheet DMHH và sử dụng cache để tăng tốc.
 */
function getHangHoaDataForSidebar() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'DANH_SACH_HANG_HOA';

  const cachedData = cache.get(CACHE_KEY);
  if (cachedData != null) {
    console.log('✅ Loaded products from CACHE.');
    return JSON.parse(cachedData);
  }

  console.log('⚠️ Cache miss. Reading products from Sheet "DMHH".');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDMHH = ss.getSheetByName('DMHH');
  if (!sheetDMHH) {
    throw new Error('Không tìm thấy sheet "DMHH"');
  }

  const data = sheetDMHH.getDataRange().getValues();
  const hangHoaList = [];
  // Bắt đầu từ dòng 2 để bỏ qua tiêu đề
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const maKho = row[0]?.toString().trim();
    const maHang = row[1]?.toString().trim();
    if (maKho && maHang) { // Chỉ lấy hàng hóa có đủ mã kho và mã hàng
      hangHoaList.push({
        maKho: maKho,
        maHang: maHang,
        tenHang: row[2]?.toString().trim() || '',
        quyCach: row[3]?.toString().trim() || '',
        dvt: row[4]?.toString().trim() || ''
      });
    }
  }

  // Sắp xếp để dễ tìm kiếm
  hangHoaList.sort((a, b) => a.maKho.localeCompare(b.maKho) || a.maHang.localeCompare(b.maHang));

  // Lưu vào cache trong 15 phút
  cache.put(CACHE_KEY, JSON.stringify(hangHoaList), 900);
  console.log(`✅ Loaded and cached ${hangHoaList.length} products.`);

  return hangHoaList;
}

/**
 * Ghi danh sách hàng hóa đã chọn (5 cột) vào sheet, bắt đầu từ ô đang được chọn.
 * @param {Array<Object>} selectedItems Mảng các đối tượng hàng hóa đã chọn.
 */
function ghiHangHoaVaoSheet(selectedItems) {
  try {
    if (!selectedItems || selectedItems.length === 0) {
      return { success: false, error: 'Không có hàng hóa nào được chọn.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeCell = ss.getActiveCell();
    const sheet = activeCell.getSheet();
    const startRow = activeCell.getRow();
    const startCol = activeCell.getColumn();
    
    // **SỬA LỖI**: Tạo mảng 2 chiều với đầy đủ 5 thông tin
    const outputData = selectedItems.map(item => [
      item.maKho, 
      item.maHang, 
      item.tenHang, 
      item.quyCach || '', // Ghi chuỗi rỗng nếu không có dữ liệu
      item.dvt || ''      // Ghi chuỗi rỗng nếu không có dữ liệu
    ]);
    
    // **SỬA LỖI**: Ghi dữ liệu ra một vùng rộng 5 cột
    sheet.getRange(startRow, startCol, outputData.length, 5).setValues(outputData);

    console.log(`✅ Written ${outputData.length} items (5 columns) to ${sheet.getName()}`);
    // Trả về số lượng đã ghi để hiển thị trên thông báo
    return { success: true, count: outputData.length }; 

  } catch (e) {
    console.error('Error in ghiHangHoaVaoSheet: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Hàm này sẽ được kích hoạt bởi Installable Trigger "On selection change".
 * Nó sẽ kiểm tra và mở sidebar hàng hóa nếu đúng điều kiện.
 * @param {Object} e Đối tượng sự kiện.
 */
function handleHangHoaSelectionChange(e) {
  try {
    const activeRange = e.range;
    const activeSheet = activeRange.getSheet();
    const sheetName = activeSheet.getName();

    // Điều kiện 1: Tên sheet phải bắt đầu bằng "DL_"
    if (!sheetName.startsWith('DL_')) return;

    // Điều kiện 2: Tiêu đề cột phải là "MA_KHO" hoặc "MA_HANG"
    const activeColumn = activeRange.getColumn();
    const headerValue = activeSheet.getRange(1, activeColumn).getValue().toString().trim().toUpperCase();

    const validHeaders = ['MA_KHO', 'MA_HANG'];
    if (!validHeaders.includes(headerValue)) return;

    // Nếu tất cả điều kiện đều đúng -> Mở sidebar
    moSidebarHangHoa();

  } catch (error) {
    console.error('Error in handleHangHoaSelectionChange trigger: ' + error.toString());
  }
}

// =================================================================
// TỰ ĐỘNG ĐIỀN THÔNG TIN HÀNG HÓA KHI NHẬP LIỆU
// =================================================================

/**
 * Hàm này tạo và trả về một Map chứa thông tin hàng hóa để tra cứu nhanh.
 * Dữ liệu được lấy từ cache để tăng tốc độ, nếu cache không có thì mới đọc từ sheet "DMHH".
 * @returns {Map<string, Object>} Một Map với key là "maKho|maHang" và value là {tenHang, quyCach, dvt}.
 */
function getHangHoaLookupMap() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'HANG_HOA_LOOKUP_MAP';

  const cachedMap = cache.get(CACHE_KEY);
  if (cachedMap != null) {
    // Nếu có cache, chuyển chuỗi JSON về lại đối tượng Map
    return new Map(JSON.parse(cachedMap));
  }

  // Nếu không có cache, đọc từ sheet "DMHH"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDMHH = ss.getSheetByName('DMHH');
  if (!sheetDMHH) {
    console.error('Không tìm thấy sheet "DMHH" để tạo lookup map.');
    return new Map();
  }

  const data = sheetDMHH.getDataRange().getValues();
  const hangHoaMap = new Map();

  // Bỏ qua dòng tiêu đề (i = 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const maKho = row[0]?.toString().trim();
    const maHang = row[1]?.toString().trim();
    if (maKho && maHang) {
      const key = `${maKho}|${maHang}`;
      const value = {
        tenHang: row[2]?.toString().trim() || '',
        quyCach: row[3]?.toString().trim() || '',
        dvt: row[4]?.toString().trim() || ''
      };
      hangHoaMap.set(key, value);
    }
  }

  // Lưu Map vào cache dưới dạng chuỗi JSON trong 6 giờ
  cache.put(CACHE_KEY, JSON.stringify(Array.from(hangHoaMap.entries())), 21600);
  console.log(`✅ Created and cached a lookup map for ${hangHoaMap.size} products.`);
  
  return hangHoaMap;
}

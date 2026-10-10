/**
 * Script to generate a realistic, multi-sheet Accounting & Financial Excel workbook
 * using the Excel MCP Server's ToolHandler directly.
 */
import path from 'node:path';
import { PermissionChecker } from '../src/security/permission-checker.js';
import { ExcelService } from '../src/services/excel-service.js';
import { ToolHandler } from '../src/tools/tool-handler.js';
import { Logger } from '../src/utils/logger.js';

async function main() {
  console.log('🚀 Starting sample accounting workbook generation using Excel MCP tools...\n');

  const cwd = process.cwd();
  const filePath = path.join(cwd, 'sample_accounting_report.xlsx');

  // Initialize MCP components
  const permissionChecker = new PermissionChecker({
    allowedPaths: [cwd],
    deniedPaths: [],
    maxFileSize: 50 * 1024 * 1024,
    allowedExtensions: ['.xlsx'],
    permissions: ['read', 'write', 'delete'],
  });

  const logger = new Logger('info');
  const excelService = new ExcelService(permissionChecker, logger);
  const handler = new ToolHandler(excelService, permissionChecker, logger);

  // Helper to execute a tool and assert success
  async function execTool(toolName: string, args: Record<string, unknown>) {
    const res = await handler.executeTool(toolName, args);
    if (!res.success) {
      console.error(`❌ Tool ${toolName} failed:`, res.error);
      throw new Error(`Tool failure: ${toolName}: ${res.error}`);
    }
    console.log(`✅ [${toolName}] Success`);
    return res;
  }

  // 1. Create workbook
  await execTool('excel_create_workbook', {
    filename: filePath,
    outputPath: cwd,
  });

  // Rename initial Sheet1 to BaoCaoKetQuaKinhDoanh
  await execTool('excel_rename_worksheet', {
    filename: filePath,
    oldName: 'Sheet1',
    newName: 'BaoCaoKetQuaKinhDoanh',
  });

  // =========================================================================
  // SHEET 1: BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH (P&L 12 THÁNG)
  // =========================================================================
  console.log('\n--- Generating Sheet 1: BaoCaoKetQuaKinhDoanh ---');

  const pnlHeaders = [
    { cellAddress: 'A1', value: 'Tháng' },
    { cellAddress: 'B1', value: 'Doanh thu thuần' },
    { cellAddress: 'C1', value: 'Giá vốn hàng bán' },
    { cellAddress: 'D1', value: 'Chi phí bán hàng & QLDN' },
    { cellAddress: 'E1', value: 'Lợi nhuận trước thuế' },
    { cellAddress: 'F1', value: 'Thuế TNDN (20%)' },
    { cellAddress: 'G1', value: 'Lợi nhuận sau thuế' },
  ];

  // 12 months data with realistic numbers, including Month 2 with a temporary loss
  const pnlData = [
    { month: 'Tháng 01', rev: 180000000, cogs: 105000000, opex: 42000000 },
    { month: 'Tháng 02', rev: 120000000, cogs:  95000000, opex: 38000000 }, // Loss: -13,000,000
    { month: 'Tháng 03', rev: 210000000, cogs: 120000000, opex: 46000000 },
    { month: 'Tháng 04', rev: 235000000, cogs: 130000000, opex: 48000000 },
    { month: 'Tháng 05', rev: 250000000, cogs: 140000000, opex: 50000000 },
    { month: 'Tháng 06', rev: 280000000, cogs: 155000000, opex: 54000000 },
    { month: 'Tháng 07', rev: 260000000, cogs: 148000000, opex: 52000000 },
    { month: 'Tháng 08', rev: 275000000, cogs: 152000000, opex: 53000000 },
    { month: 'Tháng 09', rev: 290000000, cogs: 160000000, opex: 55000000 },
    { month: 'Tháng 10', rev: 310000000, cogs: 170000000, opex: 58000000 },
    { month: 'Tháng 11', rev: 340000000, cogs: 185000000, opex: 62000000 },
    { month: 'Tháng 12', rev: 400000000, cogs: 215000000, opex: 72000000 },
  ];

  const pnlCells: Array<{ cellAddress: string; value: unknown }> = [...pnlHeaders];

  pnlData.forEach((row, i) => {
    const r = i + 2;
    pnlCells.push({ cellAddress: `A${r}`, value: row.month });
    pnlCells.push({ cellAddress: `B${r}`, value: row.rev });
    pnlCells.push({ cellAddress: `C${r}`, value: row.cogs });
    pnlCells.push({ cellAddress: `D${r}`, value: row.opex });
    // Formulas: E = B - C - D, F = IF(E>0, E*0.2, 0), G = E - F
    pnlCells.push({ cellAddress: `E${r}`, value: `=B${r}-C${r}-D${r}` });
    pnlCells.push({ cellAddress: `F${r}`, value: `=IF(E${r}>0,E${r}*0.2,0)` });
    pnlCells.push({ cellAddress: `G${r}`, value: `=E${r}-F${r}` });
  });

  // Summary labels and formulas at rows 14 & 15
  pnlCells.push({ cellAddress: 'A14', value: 'TỔNG CỘNG NĂM' });
  pnlCells.push({ cellAddress: 'B14', value: '=SUM(B2:B13)' });
  pnlCells.push({ cellAddress: 'C14', value: '=SUM(C2:C13)' });
  pnlCells.push({ cellAddress: 'D14', value: '=SUM(D2:D13)' });
  pnlCells.push({ cellAddress: 'E14', value: '=SUM(E2:E13)' });
  pnlCells.push({ cellAddress: 'F14', value: '=SUM(F2:F13)' });
  pnlCells.push({ cellAddress: 'G14', value: '=SUM(G2:G13)' });

  pnlCells.push({ cellAddress: 'A15', value: 'TRUNG BÌNH THÁNG' });
  pnlCells.push({ cellAddress: 'B15', value: '=AVERAGE(B2:B13)' });
  pnlCells.push({ cellAddress: 'C15', value: '=AVERAGE(C2:C13)' });
  pnlCells.push({ cellAddress: 'D15', value: '=AVERAGE(D2:D13)' });
  pnlCells.push({ cellAddress: 'E15', value: '=AVERAGE(E2:E13)' });
  pnlCells.push({ cellAddress: 'F15', value: '=AVERAGE(F2:F13)' });
  pnlCells.push({ cellAddress: 'G15', value: '=AVERAGE(G2:G13)' });

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    data: pnlCells,
  });

  // Also verify MCP accounting tools: sum & average calculation
  const sumRes = await execTool('excel_financial_sum', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    rangeStart: 'B2',
    rangeEnd: 'B13',
  });
  console.log('   -> Calculated Total Revenue via excel_financial_sum:', sumRes.data);

  const avgRes = await execTool('excel_financial_average', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    rangeStart: 'B2',
    rangeEnd: 'B13',
  });
  console.log('   -> Calculated Average Monthly Revenue via excel_financial_average:', avgRes.data);

  // Format header, borders, and currency
  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    startCell: 'A1',
    endCell: 'G1',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    startCell: 'A1',
    endCell: 'G15',
  });

  // VND Currency format for revenue, cogs, opex, tax, pat
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    startCell: 'B2',
    endCell: 'D15',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    startCell: 'F2',
    endCell: 'G15',
  });
  // Negative Red format for profit before tax (to highlight Month 2 loss)
  await execTool('excel_negative_red_format', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    startCell: 'E2',
    endCell: 'E15',
  });

  // Freeze top row and auto fit columns
  await execTool('excel_freeze_panes', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
    cellAddress: 'A2',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'BaoCaoKetQuaKinhDoanh',
  });

  // =========================================================================
  // SHEET 2: PHÂN TÍCH HIỆU QUẢ DỰ ÁN ĐẦU TƯ (NPV & IRR)
  // =========================================================================
  console.log('\n--- Generating Sheet 2: PhanTichDauTu_NPV_IRR ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'PhanTichDauTu_NPV_IRR',
  });

  const projectCells = [
    { cellAddress: 'A1', value: 'Năm' },
    { cellAddress: 'B1', value: 'Dòng tiền dự án (CF)' },
    { cellAddress: 'C1', value: 'Ghi chú' },
    // Year 0: Initial investment (-750,000,000)
    { cellAddress: 'A2', value: 'Năm 0' },
    { cellAddress: 'B2', value: -750000000 },
    { cellAddress: 'C2', value: 'Vốn đầu tư ban đầu (CAPEX)' },
    // Years 1 to 5: Positive cash flows
    { cellAddress: 'A3', value: 'Năm 1' },
    { cellAddress: 'B3', value: 200000000 },
    { cellAddress: 'C3', value: 'Dòng tiền vận hành năm 1' },
    { cellAddress: 'A4', value: 'Năm 2' },
    { cellAddress: 'B4', value: 240000000 },
    { cellAddress: 'C4', value: 'Dòng tiền vận hành năm 2' },
    { cellAddress: 'A5', value: 'Năm 3' },
    { cellAddress: 'B5', value: 280000000 },
    { cellAddress: 'C5', value: 'Dòng tiền vận hành năm 3' },
    { cellAddress: 'A6', value: 'Năm 4' },
    { cellAddress: 'B6', value: 260000000 },
    { cellAddress: 'C6', value: 'Dòng tiền vận hành năm 4' },
    { cellAddress: 'A7', value: 'Năm 5' },
    { cellAddress: 'B7', value: 190000000 },
    { cellAddress: 'C7', value: 'Dòng tiền vận hành + Thu hồi TSCĐ' },

    // Financial KPI table
    { cellAddress: 'E1', value: 'Chỉ số thẩm định' },
    { cellAddress: 'F1', value: 'Giá trị tính toán' },
    { cellAddress: 'G1', value: 'Đánh giá / Kết luận' },

    { cellAddress: 'E2', value: 'Chi phí vốn (WACC / r)' },
    { cellAddress: 'F2', value: 0.10 }, // 10%
    { cellAddress: 'G2', value: 'Tỷ suất chiết khấu tham chiếu' },

    { cellAddress: 'E3', value: 'NPV (Dòng tiền năm 1-5)' },
    { cellAddress: 'E4', value: 'NPV ròng (Sau trừ vốn ban đầu)' },
    { cellAddress: 'E5', value: 'Tỷ suất hoàn vốn nội bộ (IRR)' },
  ];

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    data: projectCells,
  });

  // Calculate NPV on years 1-5 cashflows at 10%
  const npvRes = await execTool('excel_calculate_npv', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    rate: 0.10,
    valuesRange: 'B3:B7',
  });
  const npvOperating = Number((npvRes.data as any)?.npv);
  const netNPV = npvOperating + (-750000000);

  // Calculate IRR on full project cashflows (Years 0-5)
  const irrRes = await execTool('excel_calculate_irr', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    valuesRange: 'B2:B7',
    guess: 0.1,
  });
  const irrVal = Number((irrRes.data as any)?.irr);

  // Write calculated KPIs back to the summary table
  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    data: [
      { cellAddress: 'F3', value: npvOperating },
      { cellAddress: 'G3', value: 'Hiện giá các dòng tiền tương lai' },
      { cellAddress: 'F4', value: netNPV },
      { cellAddress: 'G4', value: netNPV > 0 ? 'Dự án Khả thi (NPV > 0)' : 'Không khả thi' },
      { cellAddress: 'F5', value: irrVal },
      { cellAddress: 'G5', value: irrVal > 0.10 ? 'IRR > WACC (Chấp thuận đầu tư)' : 'Từ chối' },
    ],
  });

  // Style Sheet 2
  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    startCell: 'A1',
    endCell: 'C1',
  });
  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    startCell: 'E1',
    endCell: 'G1',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    startCell: 'A1',
    endCell: 'C7',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    startCell: 'E1',
    endCell: 'G5',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    startCell: 'B2',
    endCell: 'B7',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
    startCell: 'F3',
    endCell: 'F4',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'PhanTichDauTu_NPV_IRR',
  });

  // =========================================================================
  // SHEET 3: LỊCH TRẢ NỢ VAY NGÂN HÀNG (LOAN AMORTIZATION)
  // =========================================================================
  console.log('\n--- Generating Sheet 3: LichTraNoVay ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'LichTraNoVay',
  });

  const loanInfo = [
    { cellAddress: 'A1', value: 'THÔNG TIN KHOẢN VAY ĐẦU TƯ MÁY MÓC' },
    { cellAddress: 'A2', value: 'Số tiền vay gốc (VNĐ):' },
    { cellAddress: 'B2', value: 600000000 },
    { cellAddress: 'A3', value: 'Lãi suất hàng năm:' },
    { cellAddress: 'B3', value: '9.60%' },
    { cellAddress: 'C3', value: '(0.8% / tháng)' },
    { cellAddress: 'A4', value: 'Thời hạn vay:' },
    { cellAddress: 'B4', value: 12 },
    { cellAddress: 'C4', value: 'tháng (trả góp đều)' },
  ];

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'LichTraNoVay',
    data: loanInfo,
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'LichTraNoVay',
    startCell: 'B2',
    endCell: 'B2',
  });

  // Generate 12-month amortization schedule starting at A6
  await execTool('excel_create_amortization_schedule', {
    filename: filePath,
    worksheet: 'LichTraNoVay',
    startCell: 'A6',
    principal: 600000000,
    annualRate: 0.096,
    numberOfPeriods: 12,
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'LichTraNoVay',
    startCell: 'A6',
    endCell: 'E6',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'LichTraNoVay',
    startCell: 'A6',
    endCell: 'E18',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'LichTraNoVay',
  });

  // =========================================================================
  // SHEET 4: BÁO CÁO PHÂN TÍCH TUỔI NỢ PHẢI THU (AGING REPORT)
  // =========================================================================
  console.log('\n--- Generating Sheet 4: TuoiNoKhachHang ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'TuoiNoKhachHang',
  });

  const agingInvoices = [
    { cellAddress: 'A1', value: 'Mã Hóa Đơn' },
    { cellAddress: 'B1', value: 'Tên Khách Hàng' },
    { cellAddress: 'C1', value: 'Ngày Xuất HĐ' },
    { cellAddress: 'D1', value: 'Số Tiền Phải Thu' },

    // Invoices distributed across buckets (reference date: 2026-10-09)
    // 0-30 days
    { cellAddress: 'A2', value: 'HD-2026-089' },
    { cellAddress: 'B2', value: 'Công ty TNHH Ánh Dương' },
    { cellAddress: 'C2', value: '2026-10-01' },
    { cellAddress: 'D2', value: 45000000 },

    { cellAddress: 'A3', value: 'HD-2026-092' },
    { cellAddress: 'B3', value: 'Tập đoàn Hòa Phát Chi Nhánh 1' },
    { cellAddress: 'C3', value: '2026-09-25' },
    { cellAddress: 'D3', value: 82000000 },

    // 31-60 days
    { cellAddress: 'A4', value: 'HD-2026-071' },
    { cellAddress: 'B4', value: 'Công ty CP Công Nghệ Sao Mai' },
    { cellAddress: 'C4', value: '2026-08-20' },
    { cellAddress: 'D4', value: 35000000 },

    // 61-90 days
    { cellAddress: 'A5', value: 'HD-2026-055' },
    { cellAddress: 'B5', value: 'Công ty TNHH Vận Tải Biển Đông' },
    { cellAddress: 'C5', value: '2026-07-28' },
    { cellAddress: 'D5', value: 64000000 },

    // 91-120 days
    { cellAddress: 'A6', value: 'HD-2026-041' },
    { cellAddress: 'B6', value: 'Doanh nghiệp Tư nhân Nam Hải' },
    { cellAddress: 'C6', value: '2026-06-25' },
    { cellAddress: 'D6', value: 28000000 },

    // 120+ days
    { cellAddress: 'A7', value: 'HD-2026-018' },
    { cellAddress: 'B7', value: 'Công ty CP Xây Dựng Thăng Long' },
    { cellAddress: 'C7', value: '2026-04-10' },
    { cellAddress: 'D7', value: 50000000 },
  ];

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    data: agingInvoices,
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    startCell: 'A1',
    endCell: 'D1',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    startCell: 'A1',
    endCell: 'D7',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    startCell: 'D2',
    endCell: 'D7',
  });

  // Generate Aging Schedule Summary starting at F1
  await execTool('excel_create_aging_report', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    invoiceDateColumn: 'C',
    amountColumn: 'D',
    asOfDate: '2026-10-09',
    outputStartCell: 'F1',
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    startCell: 'F1',
    endCell: 'G1',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    startCell: 'F1',
    endCell: 'G6',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
    startCell: 'G2',
    endCell: 'G6',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'TuoiNoKhachHang',
  });

  // =========================================================================
  // SHEET 5: PHÂN TÍCH BIẾN ĐỘNG NGÂN SÁCH (VARIANCE & BALANCE CHECK)
  // =========================================================================
  console.log('\n--- Generating Sheet 5: CanDoiVaBienDongNganSach ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'CanDoiVaBienDongNganSach',
  });

  const varianceData = [
    { cellAddress: 'A1', value: 'Khoản mục chi phí' },
    { cellAddress: 'B1', value: 'Dự toán (Budget)' },
    { cellAddress: 'C1', value: 'Thực tế (Actual)' },

    { cellAddress: 'A2', value: 'Chi phí tiền lương & nhân sự' },
    { cellAddress: 'B2', value: 250000000 },
    { cellAddress: 'C2', value: 265000000 },

    { cellAddress: 'A3', value: 'Chi phí marketing & quảng cáo' },
    { cellAddress: 'B3', value: 80000000 },
    { cellAddress: 'C3', value: 72000000 },

    { cellAddress: 'A4', value: 'Chi phí thuê văn phòng' },
    { cellAddress: 'B4', value: 50000000 },
    { cellAddress: 'C4', value: 50000000 },

    { cellAddress: 'A5', value: 'Chi phí công nghệ & phần mềm' },
    { cellAddress: 'B5', value: 30000000 },
    { cellAddress: 'C5', value: 35000000 },

    { cellAddress: 'A6', value: 'Chi phí đào tạo & hội thảo' },
    { cellAddress: 'B6', value: 20000000 },
    { cellAddress: 'C6', value: 16000000 },
  ];

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    data: varianceData,
  });

  // Perform variance analysis: Actual (C) vs Budget (B) -> writes to D (Diff) and E (% Diff)
  await execTool('excel_variance_analysis', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    actualRange: 'C2:C6',
    budgetRange: 'B2:B6',
    outputRange: 'D2:E6',
    showPercentage: true,
  });

  // Add variance headers at D1, E1
  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    data: [
      { cellAddress: 'D1', value: 'Chênh lệch (Actual - Budget)' },
      { cellAddress: 'E1', value: '% Biến động' },
    ],
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'A1',
    endCell: 'E1',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'A1',
    endCell: 'E6',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'B2',
    endCell: 'D6',
  });

  // Balance Check section at row 8
  const balanceSection = [
    { cellAddress: 'A8', value: 'BẢNG KIỂM TRA CÂN ĐỐI TÀI SẢN & NGUỒN VỐN' },
    { cellAddress: 'A9', value: 'Khoản mục Tài sản (Nợ)' },
    { cellAddress: 'B9', value: 'Giá trị' },
    { cellAddress: 'D9', value: 'Khoản mục Nguồn vốn (Có)' },
    { cellAddress: 'E9', value: 'Giá trị' },

    { cellAddress: 'A10', value: 'Tiền mặt & Tiền gửi' },
    { cellAddress: 'B10', value: 450000000 },
    { cellAddress: 'D10', value: 'Nợ phải trả người bán' },
    { cellAddress: 'E10', value: 280000000 },

    { cellAddress: 'A11', value: 'Phải thu khách hàng' },
    { cellAddress: 'B11', value: 309000000 },
    { cellAddress: 'D11', value: 'Vay & Nợ thuê tài chính' },
    { cellAddress: 'E11', value: 600000000 },

    { cellAddress: 'A12', value: 'Hàng tồn kho' },
    { cellAddress: 'B12', value: 520000000 },
    { cellAddress: 'D12', value: 'Vốn đầu tư của chủ sở hữu' },
    { cellAddress: 'E12', value: 1200000000 },

    { cellAddress: 'A13', value: 'Tài sản cố định thuần' },
    { cellAddress: 'B13', value: 801000000 },
    { cellAddress: 'D13', value: 'Lợi nhuận sau thuế chưa PP' },
    { cellAddress: 'E13', value: 0 },
  ];

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    data: balanceSection,
  });

  // Calculate Balance Check via tool
  const balCheck = await execTool('excel_check_balance', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    debitRange: 'B10:B13',
    creditRange: 'E10:E13',
  });

  const isBalanced = (balCheck.data as any)?.isBalanced as boolean;
  const totalDebit = Number((balCheck.data as any)?.debitTotal);
  const totalCredit = Number((balCheck.data as any)?.creditTotal);

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    data: [
      { cellAddress: 'A14', value: 'TỔNG TÀI SẢN (NỢ)' },
      { cellAddress: 'B14', value: totalDebit },
      { cellAddress: 'D14', value: 'TỔNG NGUỒN VỐN (CÓ)' },
      { cellAddress: 'E14', value: totalCredit },
      { cellAddress: 'A15', value: 'KẾT LUẬN CÂN ĐỐI:' },
      { cellAddress: 'B15', value: isBalanced ? 'CÂN ĐỐI HOÀN TOÀN (Nợ = Có)' : 'LỆCH CÂN ĐỐI' },
    ],
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'A9',
    endCell: 'B9',
  });
  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'D9',
    endCell: 'E9',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'A9',
    endCell: 'B14',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'D9',
    endCell: 'E14',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'B10',
    endCell: 'B14',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
    startCell: 'E10',
    endCell: 'E14',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'CanDoiVaBienDongNganSach',
  });

  // =========================================================================
  // SHEET 6: KHẤU HAO TÀI SẢN CỐ ĐỊNH (STRAIGHT-LINE & DOUBLE-DECLINING)
  // =========================================================================
  console.log('\n--- Generating Sheet 6: KhauHao_TSCD ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'KhauHao_TSCD',
  });

  await execTool('excel_write_cell', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    cellAddress: 'A1',
    value: 'BẢNG TÍNH KHẤU HAO TSCĐ - PHƯƠNG PHÁP ĐƯỜNG THẲNG (STRAIGHT-LINE)',
  });
  await execTool('excel_apply_title_style', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'A1',
  });

  // Calculate Straight-Line depreciation for 1.2 tỷ machine over 5 years (salvage 120M)
  await execTool('excel_calculate_depreciation', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'A3',
    cost: 1200000000,
    salvageValue: 120000000,
    usefulLife: 5,
    method: 'straight-line',
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'A3',
    endCell: 'D3',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'A3',
    endCell: 'D8',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'B4',
    endCell: 'D8',
  });

  // Double-Declining Balance comparison table at F3
  await execTool('excel_write_cell', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    cellAddress: 'F1',
    value: 'KHẤU HAO THEO SỐ DƯ GIẢM DẦN CÓ ĐIỀU CHỈNH (DDB)',
  });
  await execTool('excel_apply_title_style', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'F1',
  });

  await execTool('excel_calculate_depreciation', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'F3',
    cost: 1200000000,
    salvageValue: 120000000,
    usefulLife: 5,
    method: 'double-declining',
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'F3',
    endCell: 'I3',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'F3',
    endCell: 'I8',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
    startCell: 'G4',
    endCell: 'I8',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'KhauHao_TSCD',
  });

  // =========================================================================
  // SHEET 7: THUẾ TNCN LUỸ TIẾN TỪNG PHẦN (PROGRESSIVE TAX)
  // =========================================================================
  console.log('\n--- Generating Sheet 7: ThueTNCN_LuyTien ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'ThueTNCN_LuyTien',
  });

  const pitTitle: Array<{ cellAddress: string; value: unknown }> = [
    { cellAddress: 'A1', value: 'BẢNG TÍNH THUẾ THU NHẬP CÁ NHÂN (PIT) - BIỂU THUẾ LUỸ TIẾN TỪNG PHẦN' },
    { cellAddress: 'A3', value: 'Mã NV' },
    { cellAddress: 'B3', value: 'Họ và tên nhân sự' },
    { cellAddress: 'C3', value: 'Chức danh' },
    { cellAddress: 'D3', value: 'Thu nhập tính thuế (VND)' },
    { cellAddress: 'E3', value: 'Thuế TNCN phải nộp (VND)' },
    { cellAddress: 'F3', value: 'Thu nhập thực nhận (VND)' },
  ];

  const pitEmployees = [
    { id: 'NV01', name: 'Nguyễn Văn An', title: 'Giám đốc Điều hành', income: 85000000 },
    { id: 'NV02', name: 'Trần Thị Bình', title: 'Kế toán trưởng', income: 45000000 },
    { id: 'NV03', name: 'Lê Hoàng Cường', title: 'Trưởng phòng Kỹ thuật', income: 38000000 },
    { id: 'NV04', name: 'Phạm Thị Dung', title: 'Chuyên viên Phân tích', income: 24000000 },
    { id: 'NV05', name: 'Vũ Minh Đức', title: 'Kỹ sư Phần mềm Senior', income: 30000000 },
    { id: 'NV06', name: 'Đỗ Thu Hằng', title: 'Chuyên viên Nhân sự', income: 16000000 },
    { id: 'NV07', name: 'Bùi Tuấn Kiên', title: 'Kỹ sư Kiểm thử', income: 14000000 },
    { id: 'NV08', name: 'Hoàng Thị Lan', title: 'Nhân viên Hành chính', income: 8000000 },
  ];

  pitEmployees.forEach((emp, i) => {
    const r = i + 4;
    pitTitle.push({ cellAddress: `A${r}`, value: emp.id });
    pitTitle.push({ cellAddress: `B${r}`, value: emp.name });
    pitTitle.push({ cellAddress: `C${r}`, value: emp.title });
    pitTitle.push({ cellAddress: `D${r}`, value: emp.income });
    pitTitle.push({ cellAddress: `F${r}`, value: `=D${r}-E${r}` });
  });

  // Summary row
  pitTitle.push({ cellAddress: 'A12', value: 'TỔNG CỘNG' });
  pitTitle.push({ cellAddress: 'D12', value: '=SUM(D4:D11)' });
  pitTitle.push({ cellAddress: 'E12', value: '=SUM(E4:E11)' });
  pitTitle.push({ cellAddress: 'F12', value: '=SUM(F4:F11)' });

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'ThueTNCN_LuyTien',
    data: pitTitle,
  });

  // Calculate PIT using 7 progressive tax brackets (Vietnam PIT Law)
  const vnPitBrackets = [
    { threshold: 0, rate: 5 },
    { threshold: 5000000, rate: 10 },
    { threshold: 10000000, rate: 15 },
    { threshold: 18000000, rate: 20 },
    { threshold: 32000000, rate: 25 },
    { threshold: 52000000, rate: 30 },
    { threshold: 80000000, rate: 35 },
  ];

  const pitRes = await execTool('excel_calculate_progressive_tax', {
    filename: filePath,
    worksheet: 'ThueTNCN_LuyTien',
    amountRange: 'D4:D11',
    brackets: vnPitBrackets,
    outputRange: 'E4',
  });
  console.log('   -> Calculated Total Progressive Tax:', pitRes.data);

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'ThueTNCN_LuyTien',
    startCell: 'A3',
    endCell: 'F3',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'ThueTNCN_LuyTien',
    startCell: 'A3',
    endCell: 'F12',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'ThueTNCN_LuyTien',
    startCell: 'D4',
    endCell: 'F12',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'ThueTNCN_LuyTien',
  });

  // =========================================================================
  // SHEET 8: TỶ SUẤT HOÀN VỐN NỘI BỘ DÒNG TIỀN BẤT THƯỜNG (XIRR)
  // =========================================================================
  console.log('\n--- Generating Sheet 8: XIRR_DongTienDuAn ---');

  await execTool('excel_add_worksheet', {
    filename: filePath,
    worksheetName: 'XIRR_DongTienDuAn',
  });

  const xirrTable = [
    { cellAddress: 'A1', value: 'PHÂN TÍCH HIỆU QUẢ DỰ ÁN VỚI DÒNG TIỀN KHÔNG ĐỀU THEO NGÀY (XIRR)' },
    { cellAddress: 'A3', value: 'Ngày giao dịch' },
    { cellAddress: 'B3', value: 'Dòng tiền (VND)' },
    { cellAddress: 'C3', value: 'Mô tả cột mốc dự án' },

    { cellAddress: 'A4', value: '2024-01-15' },
    { cellAddress: 'B4', value: -1500000000 },
    { cellAddress: 'C4', value: 'Góp vốn ban đầu khởi công dự án' },

    { cellAddress: 'A5', value: '2024-05-20' },
    { cellAddress: 'B5', value: -300000000 },
    { cellAddress: 'C5', value: 'Bổ sung vốn giai đoạn 1 (mua thiết bị)' },

    { cellAddress: 'A6', value: '2024-11-10' },
    { cellAddress: 'B6', value: 450000000 },
    { cellAddress: 'C6', value: 'Doanh thu nghiệm thu giai đoạn thử nghiệm' },

    { cellAddress: 'A7', value: '2025-04-05' },
    { cellAddress: 'B7', value: 700000000 },
    { cellAddress: 'C7', value: 'Cổ tức phân phối năm đầu tiên' },

    { cellAddress: 'A8', value: '2025-10-30' },
    { cellAddress: 'B8', value: 1250000000 },
    { cellAddress: 'C8', value: 'Thoái vốn / Bán chuyển nhượng dự án' },
  ];

  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    data: xirrTable,
  });

  // Calculate XIRR via tool
  const xirrRes = await execTool('excel_calculate_xirr', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    dateRange: 'A4:A8',
    valuesRange: 'B4:B8',
    guess: 0.1,
  });
  const calculatedXirr = Number((xirrRes.data as any)?.xirr);
  console.log(`   -> Calculated XIRR: ${calculatedXirr.toFixed(2)}%`);

  // Write XIRR metric box
  await execTool('excel_write_batch', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    data: [
      { cellAddress: 'A10', value: 'CHỈ SỐ XIRR TOÀN DỰ ÁN (%/năm):' },
      { cellAddress: 'B10', value: calculatedXirr / 100 },
      { cellAddress: 'C10', value: calculatedXirr > 12 ? 'Dự án sinh lời vượt trội (> WACC 12%)' : 'Dự án đạt tỷ suất kỳ vọng' },
    ],
  });

  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    startCell: 'A3',
    endCell: 'C3',
  });
  await execTool('excel_apply_all_borders', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    startCell: 'A3',
    endCell: 'C8',
  });
  await execTool('excel_vnd_currency_format', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    startCell: 'B4',
    endCell: 'B8',
  });
  await execTool('excel_apply_percentage_format', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    startCell: 'B10',
    endCell: 'B10',
  });
  await execTool('excel_apply_header_style', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
    startCell: 'A10',
    endCell: 'A10',
  });
  await execTool('excel_auto_fit_columns', {
    filename: filePath,
    worksheet: 'XIRR_DongTienDuAn',
  });

  // =========================================================================
  // SAVE WORKBOOK TO DISK
  // =========================================================================
  console.log('\n--- Saving Final Workbook ---');
  await execTool('excel_save_workbook', {
    filename: filePath,
  });

  console.log(`\n🎉 Báo cáo kế toán mẫu đã được tạo thành công tại:\n👉 ${filePath}`);
}

main().catch((err) => {
  console.error('Fatal error during workbook generation:', err);
  process.exit(1);
});

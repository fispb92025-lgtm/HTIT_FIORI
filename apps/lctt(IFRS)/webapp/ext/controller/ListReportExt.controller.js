sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    'use strict';

    return {
        btnExportExcel: async function (_oEvent) {
            try {
                // Lấy dữ liệu từ SmartTable
                const oSmartTable = this.getView().byId("lcttifrs::sap.suite.ui.generic.template.ListReport.view.ListReport::ZDD_LCTTGT_IFRS--listReport");
                if (!oSmartTable) {
                    MessageToast.show("Không tìm thấy bảng dữ liệu.");
                    return;
                }

                const oTable = oSmartTable.getTable();
                if (!oTable) {
                    MessageToast.show("Không tìm thấy bảng dữ liệu.");
                    return;
                }

                const oBinding = oTable?.getBinding("rows");
                if (!oBinding) {
                    MessageToast.show("Không có dữ liệu để export.");
                    return;
                }

                const aContexts = oBinding.getAllCurrentContexts();
                const dataArray = aContexts.map(c => c.getObject());
                if (dataArray.length === 0) {
                    MessageToast.show("Không có dòng dữ liệu để export.");
                    return;
                }

                console.log("Data Array:", dataArray);

                // Lấy thông tin header
                const firstRecord = dataArray[0] || {};
                const period = firstRecord.monat ? `${firstRecord.monat}.${firstRecord.gjahr}` : '2508';

                // Tạo workbook
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('CF IFRS');

                // Cấu hình trang A4 landscape
                worksheet.pageSetup = {
                    paperSize: 9,
                    orientation: 'landscape',
                    margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
                    fitToPage: true,
                    fitToWidth: 1,
                    fitToHeight: 0
                };

                // Set width cột thủ công
                worksheet.getColumn(1).width = 40; // Items
                worksheet.getColumn(2).width = 8;  // Code
                worksheet.getColumn(3).width = 14; // Note
                worksheet.getColumn(4).width = 22; // Current year
                worksheet.getColumn(5).width = 22; // Previous year

                // Row 1-2: CASHFLOW STATEMENT và Company name - Viền ngoài
                worksheet.mergeCells('A1:E2');
                worksheet.getCell('A1').value =
                    "CASHFLOW STATEMENT\nHaiphong Port TIL International Terminal Company";
                worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                worksheet.getCell('A1').font = { bold: true, size: 16 };

                // Tạo viền ngoài cho A1:E2
                for (let r = 1; r <= 2; r++) {
                    for (let c = 1; c <= 5; c++) {
                        const cell = worksheet.getCell(r, c);
                        const border = {};
                        if (r === 1) border.top = { style: 'thin' };
                        if (r === 2) border.bottom = { style: 'thin' };
                        if (c === 1) border.left = { style: 'thin' };
                        if (c === 5) border.right = { style: 'thin' };
                        cell.border = border;
                    }
                }

                worksheet.getRow(1).height = 30;
                worksheet.getRow(2).height = 20;

                // Row 3: Empty
                worksheet.getRow(3).height = 10;

                // Row 4-7: Parameters - Viền ngoài
                const info = [
                    { label: "Period :", value: period },
                    { label: "Actuality :", value: "AC" },
                    { label: "Currency :", value: "LC" },
                    { label: "Comp curr :", value: "USD" }
                ];

                info.forEach((p, i) => {
                    const r = 4 + i;
                    worksheet.getCell(`A${r}`).value = p.label;
                    worksheet.getCell(`B${r}`).value = p.value;
                    worksheet.getCell(`A${r}`).font = { bold: false };
                    worksheet.getCell(`A${r}`).alignment = { horizontal: 'center', vertical: 'middle' };
                    worksheet.getCell(`B${r}`).alignment = { horizontal: 'left', vertical: 'middle' };
                    worksheet.getRow(r).height = 18;
                });

                // Chỉ tạo viền ngoài cho A4:E7
                for (let r = 4; r <= 7; r++) {
                    for (let c = 1; c <= 5; c++) {
                        const cell = worksheet.getCell(r, c);
                        const border = {};
                        if (r === 4) border.top = { style: 'thin' };
                        if (r === 7) border.bottom = { style: 'thin' };
                        if (c === 1) border.left = { style: 'thin' };
                        if (c === 5) border.right = { style: 'thin' };
                        cell.border = border;
                    }
                }

                // Row 8: Empty
                worksheet.getRow(8).height = 10;

                // Row 9: Headers - Style: chữ trắng, nền xanh đậm, căn trái
                const headerRow = worksheet.getRow(9);
                headerRow.values = ['Items', 'Code', 'Note', 'Current year', 'Previous year'];
                headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } }; // Chữ trắng
                headerRow.alignment = { horizontal: "left", vertical: "middle" }; // Căn trái
                headerRow.eachCell(cell => {
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FF1F4E78" } // Nền xanh đậm
                    };
                    cell.border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" }
                    };
                });
                worksheet.getRow(9).height = 25;

                // Điền dữ liệu từ dataArray
                dataArray.forEach((row, index) => {
                    const dataRow = index + 10;

                    worksheet.getCell(`A${dataRow}`).value = row.HierarchyNode_TXT || '';
                    worksheet.getCell(`B${dataRow}`).value = row.HierarchyNode || '';
                    worksheet.getCell(`C${dataRow}`).value = '';
                    worksheet.getCell(`D${dataRow}`).value = Number(row.sokynay) || 0;
                    worksheet.getCell(`E${dataRow}`).value = Number(row.sokytruoc) || 0;

                    // Định dạng số
                    worksheet.getCell(`D${dataRow}`).numFmt = '#,##0.00';
                    worksheet.getCell(`E${dataRow}`).numFmt = '#,##0.00';

                    // Border cho data rows
                    ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                        const cell = worksheet.getCell(`${col}${dataRow}`);
                        cell.border = {
                            top: { style: "thin" },
                            left: { style: "thin" },
                            bottom: { style: "thin" },
                            right: { style: "thin" }
                        };
                    });

                    // Alignment - TẤT CẢ CĂN TRÁI
                    ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                        const cell = worksheet.getCell(`${col}${dataRow}`);
                        cell.alignment = {
                            horizontal: 'left',
                            vertical: 'middle',
                            wrapText: col === 'A' ? true : false
                        };
                    });

                    // In đậm cả hàng nếu Zfont === 'X'
                    if (row.Zfont === 'X') {
                        ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                            const cell = worksheet.getCell(`${col}${dataRow}`);
                            cell.font = {
                                name: "Times New Roman",
                                size: 11,
                                bold: true,
                                color: { argb: "FF000000" }
                            };
                        });
                    }

                    worksheet.getRow(dataRow).height = 20;
                });

                // Font Times New Roman toàn bộ
                worksheet.eachRow({ includeEmpty: true }, (row) => {
                    row.eachCell({ includeEmpty: true }, (cell) => {
                        const existingFont = cell.font || {};
                        cell.font = {
                            name: "Times New Roman",
                            size: existingFont.size || 11,
                            bold: existingFont.bold || false,
                            color: existingFont.color || { argb: "FF000000" } // Mặc định đen
                        };
                    });
                });

                // Xuất file
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });

                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                const now = new Date();
                const timestamp = now.toLocaleTimeString('en-GB', { hour12: false }).replace(/:/g, '') + '_' +
                    now.toLocaleDateString('en-GB').replace(/\//g, '');
                link.download = `CF_IFRS_${timestamp}.xlsx`;
                link.click();
                window.URL.revokeObjectURL(url);

                MessageToast.show("Xuất Excel thành công!");

            } catch (error) {
                console.error("Export Error:", error);
                MessageToast.show("Lỗi khi export Excel: " + error.message);
            }
        }
    };
});
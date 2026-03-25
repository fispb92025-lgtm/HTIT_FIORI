sap.ui.define([
    "sap/m/MessageToast",
    // "bangcandoiketoan/project1/thirdparty/exceljs"
], function (MessageToast, ExcelJS) {
    "use strict";

    return {
        btnExportExcel: async function (_oEvent) {
            try {
                const oSmartTable = this.getView().byId("bangcandoiketoan.project1::sap.suite.ui.generic.template.ListReport.view.ListReport::ZDD_BS--listReport");
                if (!oSmartTable) {
                    MessageToast.show("Không tìm thấy bảng dữ liệu.");
                    return;
                }

                const oTable = oSmartTable.getTable();
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

                // Lấy thông tin period từ data
                const gjahr = dataArray[0].gjahr || '2025';
                const monat = dataArray[0].monat || '03';
                const period = monat + '.' + gjahr;
                const previousYear = (parseInt(gjahr) - 1).toString();

                // --- Tạo workbook và worksheet ---
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet("BS_IFRS");

                // --- Cấu hình trang ---
                worksheet.pageSetup = {
                    paperSize: 9, // A4
                    orientation: "portrait",
                    fitToPage: true,
                    fitToWidth: 1,
                    fitToHeight: 0
                };

                // === BỔ SUNG 2 CỘT MỚI ===
                worksheet.columns = [
                    { header: "", key: "", width: 7 },                        // A (empty)
                    { header: "Acct", key: "HierarchyNode", width: 15 },      // B
                    { header: "Description", key: "HierarchyNode_TXT", width: 38 }, // C
                    { header: "", key: "socuoiky", width: 17 },               // D - Closing balance
                    { header: "", key: "sodauky", width: 17 },                // E - Opening balance
                    { header: "", key: "socuoiky_vnd", width: 25 },           // F - Closing balance VND
                    { header: "", key: "sodauky_vnd", width: 25 },            // G - Opening balance VND
                    { header: "", key: "socuoiky_qd", width: 25 },            // H - Converted closing balance
                    { header: "", key: "sodauky_qd", width: 25 },             // I - Converted opening balance
                    { header: "", key: "closing_ifrs", width: 25 },           // J - Closing balance IFRS (mới)
                    { header: "", key: "opening_ifrs", width: 25 }            // K - Opening balance IFRS (mới)
                ];

                // --- Header ---
                worksheet.mergeCells("B1", "E1");
                worksheet.getCell("B1").value = "Balance Sheet";
                worksheet.getCell("B1").font = { bold: true, size: 14 };
                worksheet.getCell("B1").alignment = { horizontal: "center", vertical: "middle" };
                worksheet.getCell("B1").border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    right: { style: "thin" }
                };

                worksheet.mergeCells("B2", "E2");
                worksheet.getCell("B2").value = "Haiphong Port TIL International Terminal Company";
                worksheet.getCell("B2").font = { bold: true, size: 14 };
                worksheet.getCell("B2").alignment = { horizontal: "center", vertical: "middle" };
                worksheet.getCell("B2").border = {
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" }
                };

                // Row 3 empty

                worksheet.getCell("B4").value = "Period :";
                worksheet.getCell("C4").value = period;

                worksheet.getCell("B5").value = "Actuality :";
                worksheet.getCell("C5").value = "AC";

                worksheet.getCell("B6").value = "Currency :";
                worksheet.getCell("C6").value = "LC";

                worksheet.getCell("B7").value = "Comp curr :";
                worksheet.getCell("C7").value = "USD";

                // Định dạng font phần thông tin
                for (let i = 4; i <= 7; i++) {
                    worksheet.getCell(`B${i}`).font = { bold: true };
                    worksheet.getRow(i).height = 18;
                }

                // Viền bao quanh vùng thông tin (B4:E7)
                for (let r = 4; r <= 7; r++) {
                    for (let c = 2; c <= 5; c++) {
                        const cell = worksheet.getCell(r, c);
                        cell.border = {
                            top: r === 4 ? { style: "thin" } : {},
                            bottom: r === 7 ? { style: "thin" } : {},
                            left: c === 2 ? { style: "thin" } : {},
                            right: c === 5 ? { style: "thin" } : {}
                        };
                    }
                }

                // Row 8 empty

                // Header dòng 10 - BỔ SUNG TIÊU ĐỀ 2 CỘT MỚI
                worksheet.getCell("B10").value = "Acct";
                worksheet.getCell("C10").value = "Description";
                worksheet.getCell("D10").value = "Closing balance";
                worksheet.getCell("E10").value = "Opening balance";
                worksheet.getCell("F10").value = "Closing balance VND";
                worksheet.getCell("G10").value = "Opening balance VND";
                worksheet.getCell("H10").value = "Converted closing balance";
                worksheet.getCell("I10").value = "Converted opening balance";
                worksheet.getCell("J10").value = "Closing balance IFRS";   // Cột mới J
                worksheet.getCell("K10").value = "Opening balance IFRS";   // Cột mới K

                const headerRow = worksheet.getRow(10);
                headerRow.font = { bold: true };
                headerRow.alignment = { horizontal: "center", vertical: "middle" };

                headerRow.eachCell(cell => {
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FF1F4E78" }
                    };
                    cell.font = {
                        bold: true,
                        color: { argb: "FFFFFFFF" }
                    };
                    cell.border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" }
                    };
                });

                // === ĐỔ DỮ LIỆU - BỔ SUNG 2 CỘT MỚI ===
                let startRow = 11;
dataArray.forEach((row, index) => {
                    // === LOG ĐỂ XEM TẤT CẢ CÁC TRƯỜNG TRONG DỮ LIỆU ===
                    if (index === 0) {  // Chỉ in dòng đầu tiên để tránh spam console
                        console.log("=== DỮ LIỆU MẪU (dòng đầu tiên) ===");
                        console.log(row);
                        console.log("Các trường có sẵn:", Object.keys(row));
                        console.log("=====================================");
                    }
                    // ===========================================

                    const excelRow = worksheet.getRow(startRow + index);

                    excelRow.getCell(2).value = row.HierarchyNode || '';
                    excelRow.getCell(3).value = row.HierarchyNode_TXT || '';
                    excelRow.getCell(4).value = Number(row.socuoiky) || 0;
                    excelRow.getCell(5).value = Number(row.sodauky) || 0;
                    excelRow.getCell(6).value = Number(row.socuoiky_vnd) || 0;
                    excelRow.getCell(7).value = Number(row.sodauky_vnd) || 0;
                    excelRow.getCell(8).value = Number(row.socuoiky_qd) || 0;
                    excelRow.getCell(9).value = Number(row.sodauky_qd) || 0;

                    // === 2 CỘT MỚI IFRS ===
                    excelRow.getCell(10).value = Number(row.socuoiky_ifrs || 0);  // J - Closing balance IFRS
                    excelRow.getCell(11).value = Number(row.sodauky_ifrs || 0);   // K - Opening balance IFRS
                    if (row.Zfont === 'X') {
                        excelRow.font = { bold: true };
                    }

                    // Number Format cho tất cả các cột số (từ D đến K)
                    for (let col = 4; col <= 11; col++) {  // Đã mở rộng đến cột K
                        excelRow.getCell(col).numFmt =
                            '_(* #,###.00_);_(* -#,###.00;_(* "0,00"_);_(@_)';
                    }

                    // Border cho toàn bộ dòng
                    excelRow.eachCell(cell => {
                        cell.border = {
                            top: { style: "thin" },
                            left: { style: "thin" },
                            bottom: { style: "thin" },
                            right: { style: "thin" }
                        };
                    });
                });

                // --- Xuất file ---
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], {
                    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                });

                const url = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `BS_IFRS_${period.replace('.', '_')}.xlsx`;
                a.click();
                window.URL.revokeObjectURL(url);

                MessageToast.show("Xuất Excel thành công!");
            } catch (error) {
                console.error(error);
                MessageToast.show("Lỗi khi export Excel: " + error.message);
            }
        }
    };
});
///test
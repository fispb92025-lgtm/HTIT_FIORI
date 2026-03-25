sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    "use strict";

    return {
        btnExportExcel: async function (_oEvent) {
            try {
                const oSmartTable = this.getView().byId("baocaoketquakinhdoanh.project2::sap.suite.ui.generic.template.ListReport.view.ListReport::ZDD_pl--listReport");
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

                // Lấy thông tin period từ data (giả sử tất cả row có cùng gjahr và monat)
                const gjahr = dataArray[0].gjahr || '2025';
                const monat = dataArray[0].monat || '03';
                const period = monat + '.' + gjahr;
                const previousYear = (parseInt(gjahr) - 1).toString();

                // --- Tạo workbook và worksheet ---
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet("PL_IFRS");

                // --- Cấu hình trang ---
                worksheet.pageSetup = {
                    paperSize: 9, // A4
                    orientation: "portrait",
                    fitToPage: true,
                    fitToWidth: 1,
                    fitToHeight: 0
                };

                // --- Định nghĩa cột với header và key ---
                worksheet.columns = [
                    { header: "", key: "", width: 7 },
                    { header: "", key: "HierarchyNode", width: 15 },
                    { header: "", key: "HierarchyNode_TXT", width: 40 },
                    { header: "", key: "actual_current", width: 25 },
                    { header: "", key: "ytd_current", width: 25 },
                    { header: "", key: "actual_previous", width: 25 },
                    { header: "", key: "ytd_previous", width: 25 }
                ];

                // --- Header ---
                worksheet.mergeCells("B1", "G1");
                worksheet.getCell("B1").value = "Profit & Loss";
                worksheet.getCell("B1").font = { bold: true, size: 14 };
                worksheet.getCell("B1").alignment = { horizontal: "center", vertical: "middle" };
                worksheet.getCell("B1").border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    right: { style: "thin" }
                };

                worksheet.mergeCells("B2", "G2");
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
                // worksheet.getCell("D4").value = "Là kỳ chạy báo cáo";
                // worksheet.getCell("D4").font = { color: { argb: "FF0000" }, italic: true };

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

                // Viền bao quanh vùng thông tin (B4:G7)
                for (let r = 4; r <= 7; r++) {
                    for (let c = 2; c <= 7; c++) {
                        const cell = worksheet.getCell(r, c);
                        cell.border = {
                            top: r === 4 ? { style: "thin" } : {},
                            bottom: r === 7 ? { style: "thin" } : {},
                            left: c === 2 ? { style: "thin" } : {},
                            right: c === 7 ? { style: "thin" } : {}
                        };
                    }
                }

                // Row 8 empty

                // worksheet.mergeCells("D9", "E9");
                // worksheet.getCell("D9").value = gjahr;
                // worksheet.getCell("D9").alignment = { horizontal: "center" };

                // worksheet.mergeCells("F9", "G9");
                // worksheet.getCell("F9").value = previousYear;
                // worksheet.getCell("F9").alignment = { horizontal: "center" };

                worksheet.getCell("B10").value = "Acct";
                worksheet.getCell("C10").value = "Description";
                worksheet.getCell("D10").value = "Actual - Current period";
                worksheet.getCell("E10").value = "YTD - Current year";
                worksheet.getCell("F10").value = "Actual - Previous period";
                worksheet.getCell("G10").value = "YTD - Previous year";

                // Định dạng header row 10
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
                        color: { argb: "FFFFFFFF" } // màu chữ trắng
                    };
                    cell.border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" }
                    };
                });

                let startRow = 11;
                dataArray.forEach((row, index) => {
                    const excelRow = worksheet.getRow(startRow + index);

                    excelRow.getCell(2).value = row.HierarchyNode || '';          // Acct
                    excelRow.getCell(3).value = row.HierarchyNode_TXT || '';      // Description

                    excelRow.getCell(4).value = Number(row.sokynay) || 0;       // Actual - Current period
                    excelRow.getCell(5).value = Number(row.lkkynay) || 0;      // YTD - Current year
                    excelRow.getCell(6).value = Number(row.sokytruoc) || 0;     // Actual - Previous period
                    excelRow.getCell(7).value = Number(row.lkkytruoc) || 0;    // YTD - Previous year

                    // Nếu có dòng tiêu đề (zfont = 'X')
                    if (row.Zfont === 'X') {
                        excelRow.font = { bold: true };
                    }

                    workbook.calcProperties.fullCalcOnLoad = true; // đảm bảo Excel tự làm mới định dạng
                    worksheet.properties.defaultRowHeight = 18;

                    // Định dạng số
                    excelRow.getCell(4).numFmt = '_(* #,###.00_);_(* -#,###.00;_(* "0,00"_);_(@_)';
                    excelRow.getCell(5).numFmt = '_(* #,###.00_);_(* -#,###.00;_(* "0,00"_);_(@_)';
                    excelRow.getCell(6).numFmt = '_(* #,###.00_);_(* -#,###.00;_(* "0,00"_);_(@_)';
                    excelRow.getCell(7).numFmt = '_(* #,###.00_);_(* -#,###.00;_(* "0,00"_);_(@_)';

                    // Kẻ border cho từng ô
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
                a.download = `PL_IFRS_${period.replace('.', '_')}.xlsx`;
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
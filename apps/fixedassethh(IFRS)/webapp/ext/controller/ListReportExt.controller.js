sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    "use strict";

    return {
        btnExportExcel: async function (_oEvent) {
            try {
                // Lấy dữ liệu từ SmartTable
                const oSmartTable = this.getView().byId("bcthtgtscd.fixedasset::sap.suite.ui.generic.template.ListReport.view.ListReport::FixedAsset--listReport");

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

                // Log dữ liệu để kiểm tra
                console.log("Data Array:", dataArray.map(row => ({
                    ItemCode: row.ItemCode,
                    COL02: row.COL02,
                    COL03: row.COL03,
                    COL04: row.COL04,
                    COL05: row.COL05,
                    COL06: row.COL06,
                    COL07: row.COL07
                })));

                // 🔹 Lấy dữ liệu từ dòng đầu tiên trong dataArray

                const firstRecord = dataArray[0] || {};

                const companyCode = firstRecord.CompanyCode || '6710';
                const fromPeriod = firstRecord.FromPeriod || '';
                const toPeriod = firstRecord.ToPeriod || '';
                const fromFiscalYear = firstRecord.FromFiscalYear || '';
                const toFiscalYear = firstRecord.ToFiscalYear || '';
                const preparedby = firstRecord.Preparedby || '';
                const chiefAccountant = firstRecord.ChiefAccountant || '';

                const formattedFromPeriod = fromPeriod.toString().padStart();
                const formattedToPeriod = toPeriod.toString().padStart();

                // --- Tạo workbook và worksheet ---
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet("Form");


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
                    { header: "", key: "", width: 30 },
                    { header: "", key: "col02", width: 30 },
                    { header: "", key: "col03", width: 30 },
                    { header: "", key: "col04", width: 30 },
                    { header: "", key: "col05", width: 30 },
                    { header: "", key: "col06", width: 30 },
                    { header: "", key: "col07", width: 30 }
                ];

                // --- Header ---
                // --- Header ---
                worksheet.mergeCells("A1:G1");
                worksheet.getCell("A1").value = {
                    richText: [
                        { text: "CÔNG TY TNHH CẢNG QUỐC TẾ TIL CẢNG HẢI PHÒNG", font: { name: "Times New Roman", bold: true, italic: false, color: { argb: "FF000000" } } },
                        { text: "\nHAIPHONG PORT TIL INTERNATIONAL TERMINAL COMPANY LIMITED", font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } }
                    ]
                };
                worksheet.getCell("A1").alignment = { horizontal: "left", vertical: "middle", wrapText: true };
                worksheet.getRow(1).height = 40;

                worksheet.mergeCells("A2:G2");
                worksheet.getCell("A2").value = {
                    richText: [
                        { text: "Bến 3&4 Cảng nước sâu Lạch Huyện, Khu phố Đôn Lương, Đặc khu Cát Hải, Thành phố Hải Phòng, Việt Nam", font: { name: "Times New Roman", bold: true, italic: false, color: { argb: "FF000000" } } },
                        { text: "\nBerth No. 3&4 Lach Huyen Deep-sea Port, Don Luong Quarter, Cat Hai Special Zone, Hai Phong City, Vietnam", font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } }
                    ]
                };
                worksheet.getCell("A2").alignment = { horizontal: "left", vertical: "middle", wrapText: true };
                worksheet.getRow(2).height = 40;

                worksheet.mergeCells("A4:G4");
                worksheet.getCell("A4").value = "BÁO CÁO TÌNH HÌNH TĂNG GIẢM TSCĐ";
                worksheet.getCell("A4").font = { bold: true, size: 14 };
                worksheet.getCell("A4").alignment = { horizontal: "center" };

                worksheet.mergeCells("A5:G5");
                worksheet.getCell("A5").value = "(Fixed Asset Register - VAS)";
                worksheet.getCell("A5").font = { italic: true, color: { argb: "FF0000FF" } };
                worksheet.getCell("A5").alignment = { horizontal: "center" };

                worksheet.mergeCells("A6:G6");
                worksheet.getCell("A6").value = {
                    richText: [
                        { text: `Từ kỳ ${formattedFromPeriod}/${fromFiscalYear} đến kỳ ${formattedToPeriod}/${toFiscalYear}`, font: { name: "Times New Roman", italic: false, color: { argb: "FF000000" } } },
                        { text: `/From period ${formattedFromPeriod}/${fromFiscalYear} To period ${formattedToPeriod}/${toFiscalYear}`, font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } }
                    ]
                };
                worksheet.getCell("A6").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                worksheet.getRow(6).height = 30; // Tăng chiều cao để chứa văn bản song ngữ

                // --- Tiêu đề cột (A8:G8) với định dạng song ngữ ---
                const headers = [
                    { cell: "A8", vietnamese: "Hạng mục", english: "Category" },
                    { cell: "B8", vietnamese: "Nhà cửa, vật kiến trúc", english: "Houses, buildings" },
                    { cell: "C8", vietnamese: "Máy móc và thiết bị", english: "Machinery and equipment" },
                    { cell: "D8", vietnamese: "Phương tiện vận tải, truyền dẫn", english: "Means of transport, transmission" },
                    { cell: "E8", vietnamese: "Thiết bị, dụng cụ quản lý", english: "Management equipment and tools" },
                    { cell: "F8", vietnamese: "Tài sản cố định khác", english: "Other fixed assets" },
                    { cell: "G8", vietnamese: "Cộng", english: "Total" }
                ];

                headers.forEach(header => {
                    worksheet.getCell(header.cell).value = {
                        richText: [
                            { text: header.vietnamese, font: { name: "Times New Roman", bold: true, italic: false, color: { argb: "FF000000" } } },
                            { text: `\n${header.english}`, font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "FF0000FF" } } }
                        ]
                    };
                    worksheet.getCell(header.cell).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                    worksheet.getCell(header.cell).border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" }
                    };
                });

                // Định dạng dòng tiêu đề cột
                const titleRow = worksheet.getRow(8);
                titleRow.height = 40; // Tăng chiều cao để chứa văn bản song ngữ
                titleRow.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

                // --- Danh sách các hàng cần điền ---
                const categories = [
                    { row: 9, vietnamese: "Nguyên giá", english: "Acquisition", itemCode: "0100" },
                    { row: 10, vietnamese: "Số đầu kỳ", english: "Beginning", itemCode: "0101" },
                    { row: 11, vietnamese: "Mua trong kỳ", english: "Purchased during the period", itemCode: "0102" },
                    { row: 12, vietnamese: "Đầu tư XDCB hoàn thành", english: "Completed asset under contruction", itemCode: "0103" },
                    { row: 13, vietnamese: "Tăng từ cổ đông góp vốn", english: "Increase from equity shareholders", itemCode: "0104" },
                    { row: 14, vietnamese: "Tăng từ chuyển giao nợ", english: "Increase from debt transfer", itemCode: "0105" },
                    { row: 15, vietnamese: "Tăng khác", english: "Other increases", itemCode: "0106" },
                    { row: 16, vietnamese: "Chuyển sang bất động sản đầu tư", english: "Switch to investment real estate", itemCode: "0107" },
                    { row: 17, vietnamese: "Thanh lý, nhượng bán", english: "Liquidation, sale", itemCode: "0108" },
                    { row: 18, vietnamese: "Xóa sổ", english: "Write off", itemCode: "0109" },
                    { row: 19, vietnamese: "Giảm khác", english: "Other decreases", itemCode: "0110" },
                    { row: 20, vietnamese: "Số cuối kỳ", english: "Ending", itemCode: "0111" },
                    { row: 21, vietnamese: "Trong đó:", english: "In there:", itemCode: "0112" },
                    { row: 22, vietnamese: "Đã khấu hao hết nhưng vẫn còn sử dụng", english: "Fully depreciated but still in use", itemCode: "0113" },
                    { row: 24, vietnamese: "Giá trị hao mòn", english: "Depreciation", itemCode: "0114" },
                    { row: 25, vietnamese: "Số đầu kỳ", english: "Beginning", itemCode: "0115" },
                    { row: 26, vietnamese: "Khấu hao trong kỳ", english: "Depreciation during the period", itemCode: "0116" },
                    { row: 27, vietnamese: "Tăng khác", english: "Other increases", itemCode: "0117" },
                    { row: 28, vietnamese: "Thanh lý, nhượng bán", english: "Liquidation, sale", itemCode: "0118" },
                    { row: 29, vietnamese: "Xóa sổ", english: "Write off", itemCode: "0119" },
                    { row: 30, vietnamese: "Giảm khác", english: "Other decreases", itemCode: "0120" },
                    { row: 31, vietnamese: "Số cuối kỳ", english: "Ending", itemCode: "0121" },
                    { row: 33, vietnamese: "Giá trị còn lại", english: "Net Book Value", itemCode: "0122" },
                    { row: 34, vietnamese: "Số đầu kỳ", english: "Beginning", itemCode: "0123" },
                    { row: 35, vietnamese: "Số cuối kỳ", english: "Ending", itemCode: "0124" }
                ];

                // Điền tiêu đề hàng với định dạng song ngữ
                categories.forEach(cat => {
                    worksheet.getCell(`A${cat.row}`).value = {
                        richText: [
                            { text: cat.vietnamese, font: { name: "Times New Roman", italic: false, color: { argb: "FF000000" } } },
                            { text: `\n${cat.english}`, font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } },
                        ]
                    };
                    worksheet.getCell(`A${cat.row}`).alignment = { vertical: "middle", wrapText: true };
                    worksheet.getCell(`A${cat.row}`).border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" }
                    };
                });

                // Điền dữ liệu từ dataArray vào các hàng tương ứng
                dataArray.forEach(row => {
                    const itemCode = row.ItemCode;
                    const matchedCategory = categories.find(cat => cat.itemCode === itemCode);
                    if (matchedCategory) {
                        const rowIndex = matchedCategory.row;
                        worksheet.getCell(`B${rowIndex}`).value = Number(row.COL02) || 0;
                        worksheet.getCell(`C${rowIndex}`).value = Number(row.COL03) || 0;
                        worksheet.getCell(`D${rowIndex}`).value = Number(row.COL04) || 0;
                        worksheet.getCell(`E${rowIndex}`).value = Number(row.COL05) || 0;
                        worksheet.getCell(`F${rowIndex}`).value = Number(row.COL06) || 0;
                        worksheet.getCell(`G${rowIndex}`).value = Number(row.COL07) || 0;

                        // Định dạng số
                        worksheet.getCell(`B${rowIndex}`).numFmt = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';
                        worksheet.getCell(`C${rowIndex}`).numFmt = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';
                        worksheet.getCell(`D${rowIndex}`).numFmt = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';
                        worksheet.getCell(`E${rowIndex}`).numFmt = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';
                        worksheet.getCell(`F${rowIndex}`).numFmt = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';
                        worksheet.getCell(`G${rowIndex}`).numFmt = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';

                        // Kẻ viền cho ô dữ liệu
                        worksheet.getRow(rowIndex).eachCell(cell => {
                            cell.border = {
                                top: { style: "thin" },
                                left: { style: "thin" },
                                bottom: { style: "thin" },
                                right: { style: "thin" }
                            };
                        });
                    } else {
                        console.warn(`ItemCode không hợp lệ hoặc không tìm thấy: ${itemCode}`);
                    }
                });

                // --- Kẻ viền toàn bảng (bao gồm header và dữ liệu) ---
                const startRow = 8;  // Dòng tiêu đề
                const endRow = 35;   // Dòng cuối của bảng
                const startCol = 1;  // Cột A
                const endCol = 7;    // Cột G

                for (let r = startRow; r <= endRow; r++) {
                    for (let c = startCol; c <= endCol; c++) {
                        const cell = worksheet.getCell(r, c);
                        cell.border = {
                            top: { style: "thin" },
                            left: { style: "thin" },
                            bottom: { style: "thin" },
                            right: { style: "thin" }
                        };
                    }
                }


                // --- Footer song ngữ ---
                worksheet.mergeCells("E38:G38");
                worksheet.getCell("E38").value = {
                    richText: [
                        { text: "Ngày … tháng … năm …", font: { name: "Times New Roman", italic: false, color: { argb: "FF000000" } } },
                        { text: "\nDate …… Month …… Year ……", font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } },
                    ]
                };
                worksheet.getCell("E38").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                worksheet.getRow(38).height = 40;

                // --- Người lập biểu ---
                worksheet.mergeCells("A40:C40");
                worksheet.getCell("A40").value = {
                    richText: [
                        { text: "Người lập biểu", font: { name: "Times New Roman", bold: true, italic: false, color: { argb: "FF000000" } } },
                        { text: "\nPrepared by", font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "FF0000FF" } } },
                    ]
                };
                worksheet.getCell("A40").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                worksheet.getRow(40).height = 50;

                worksheet.mergeCells("A41:C41");
                worksheet.getCell("A41").value = {
                    richText: [
                        { text: "(Ký, họ tên/", font: { name: "Times New Roman", italic: false, color: { argb: "FF000000" } } },
                        { text: "Signature, full name", font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } },
                        { text: ")", font: { italic: true } }
                    ]
                };
                worksheet.getCell("A41").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                worksheet.getRow(41).height = 40;

                // --- Kế toán trưởng ---
                worksheet.mergeCells("E40:G40");
                worksheet.getCell("E40").value = {
                    richText: [
                        { text: "Kế toán trưởng", font: { name: "Times New Roman", bold: true, italic: false, color: { argb: "FF000000" } } },
                        { text: "\nChief Accountant", font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "FF0000FF" } } },
                    ]
                };
                worksheet.getCell("E40").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                worksheet.getRow(40).height = 50;

                worksheet.mergeCells("E41:G41");
                worksheet.getCell("E41").value = {
                    richText: [
                        { text: "(Ký, họ tên/", font: { name: "Times New Roman", italic: false, color: { argb: "FF000000" } } },
                        { text: "Signature, full name", font: { name: "Times New Roman", italic: true, color: { argb: "FF0000FF" } } },
                        { text: ")", font: { italic: true } }
                    ]
                };
                worksheet.getCell("E41").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
                worksheet.getRow(41).height = 40;

                // --- Hiển thị tên người ký dưới phần ký tên ---
                // Người lập biểu (Prepared by)
                worksheet.mergeCells("A45:C45");
                worksheet.getCell("A45").value = preparedby || "";
                worksheet.getCell("A45").alignment = { horizontal: "center", vertical: "middle" };
                worksheet.getCell("A45").font = { name: "Times New Roman", size: 11, bold: true };
                worksheet.getRow(45).height = 40;

                // Kế toán trưởng (Chief Accountant)
                worksheet.mergeCells("E45:G45");
                worksheet.getCell("E45").value = chiefAccountant || "";
                worksheet.getCell("E45").alignment = { horizontal: "center", vertical: "middle" };
                worksheet.getCell("E45").font = { name: "Times New Roman", size: 11, bold: true };
                worksheet.getRow(45).height = 40;

                // --- Xuất file ---
                const period = `${fromPeriod}.${fromFiscalYear}`;
                // --- Áp dụng font Times New Roman cho toàn bộ sheet ---
                worksheet.eachRow({ includeEmpty: true }, (row) => {
                    row.eachCell({ includeEmpty: true }, (cell) => {
                        const existingFont = cell.font || {};
                        cell.font = {
                            name: "Times New Roman",
                            size: existingFont.size || 11,
                            bold: existingFont.bold,
                            italic: existingFont.italic,
                            color: existingFont.color
                        };
                    });
                });

                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], {
                    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                });

                const url = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `Bao_Cao_Tinh_Hinh_Tang_Giam_TSCD_${period.replace('.', '_')}.xlsx`;
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
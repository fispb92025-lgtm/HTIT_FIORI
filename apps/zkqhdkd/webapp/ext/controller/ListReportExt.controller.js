sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, exceljs) {
    'use strict';

    return {
        EXCEL: async function (oEvent) {
            // Lay du lieu man hinh

            try {
                const oSmartTable = this.getView().byId("zkqhdkd::sap.suite.ui.generic.template.ListReport.view.ListReport::ZC_KQHDKD--listReport");
                const oTable = oSmartTable.getTable();
                const oBinding = oTable.getBinding("rows");
                var aData = oBinding.getContexts(0, oBinding.getLength()).map(function (oContext) {
                    return oContext.getObject();
                });
                if (!aData || aData.length === 0) {
                    MessageToast.show("Không có dữ liệu để export.");
                    return;
                }
            } catch (error) {
                MessageToast.show("Không có dữ liệu để export.");
                return;
            }


            if (!aData || aData.length === 0) {
                MessageToast.show("Không có dữ liệu để export.");
                sap.ui.core.BusyIndicator.hide(); // Ẩn spinner
                return;
            }
            // Tham số
            const oFilterBar = this.byId("zkqhdkd::sap.suite.ui.generic.template.ListReport.view.ListReport::ZC_KQHDKD--listReportFilter"); // ID của SmartFilterBar
            const oFilterData = oFilterBar.getFilterData();
            var loaibc; // xử lý theo tham số lọc
            const thuyetMinh = "Thuyết minh";
            const maSo = "Mã số";
            const chiTieu = "Chỉ tiêu";
            const tenNL = oFilterData.tenNL;
            const tenKT = oFilterData.tenKT;
            const tenGD = oFilterData.tenGD;
            const gjahr = oFilterData.gjahr;

            var thang;
            var luyKethang;
            var namNay;
            var namTruoc;
            var header;
            var reportName;
            var header_namthang;
            var numberId = 3;
            var trueValue;
            var rowData;
            var lbl_namnay = 'Năm nay';
            var lbl_namtruoc = 'Năm trước';
            var lbl_merge;
            var lbl_merge_lk;
            var colWidth;
            var currentLine;
            var endline;
            var nguoiLapBieu = 'Người lập biểu';
            var keToanTruong = 'Kế toán trưởng';

            var giamDoc;
            if (gjahr < 2026) {
                giamDoc = 'Giám đốc';
            } else {
                giamDoc = 'Người đại diện theo pháp luật';
            }

            var hoTen = '(Ký, họ tên)';
            var dongDau = '(Ký, họ tên, đóng dấu)';
            var ghichu = 'Ghi chú: ';
            var dong1 = '(1). Những chỉ tiêu không có số liệu có thể không phải trình bày nhưng không được đánh lại số thứ tự chỉ tiêu và "Mã số" ';
            var dong2 = '(2) Số liệu trong các chỉ tiêu có dấu (*) được ghi bằng số âm dưới hình thức ghi trong ngoặc đơn (…). ';
            var dong3 = '(3) Đối với doanh nghiệp có kỳ kế toán năm là năm dương lịch (X) thì “Số cuối năm” có thể ghi là ';
            var header_mauso = 'Mẫu số B 02a - DN';
            var header_banhanh = '(Ban hành theo Thông tư số 200/2014/TT-BTC';
            var header_ngay = 'Ngày 22/12/2014 của Bộ Tài chính)';
            var header_mauso_new = 'Mẫu số B 02 - DN';
            var header_banhanh_new = '(Kèm theo Thông tư số 99/2025/TT-BTC';
            var header_ngay_new = 'ngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính)';
            var header_congty = aData[0].CCname;
            var header_diachi = aData[0].CCadrr;
            var header_dvt = 'Đơn vị tính:VND';
            const type = oFilterData.type;
            const monat = oFilterData.monat;

            const accountingFormat = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)'; // format số kiểu accounting
            const workbook = new exceljs.Workbook();
            const worksheet = workbook.addWorksheet('Sheet1');
            worksheet.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
            worksheet.views = [
                { state: 'normal', zoomScale: 70 } // Zoom khi mở file
            ];
            //header
            // Check loai form
            switch (type) {
                case "01":
                    lbl_merge = `Tháng ${monat}`;
                    lbl_merge_lk = `Lũy kế từ đầu năm đến cuối tháng ${monat}`;
                    reportName = ' BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH';
                    header_namthang = `Tháng ${monat} năm ${gjahr}`;
                    loaibc = 1;
                    break;
                case "02":
                    lbl_merge = `Quý 1`;
                    lbl_merge_lk = `Lũy kế từ đầu năm đến cuối quý này`;
                    lbl_merge = `Tháng ${monat}`;
                    header_namthang = `Quý 1 năm ${gjahr}`;
                    reportName = ' BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH';
                    loaibc = 1;
                    break;
                case "03":
                    lbl_merge = `Quý 2`;
                    lbl_merge_lk = `Lũy kế từ đầu năm đến cuối quý này`;
                    header_namthang = `Quý 2 năm ${gjahr}`;
                    reportName = ' BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH';
                    loaibc = 1;
                    break;
                case "04":
                    lbl_merge = `Quý 3`;
                    lbl_merge_lk = `Lũy kế từ đầu năm đến cuối quý này`;
                    header_namthang = `Quý 3 năm ${gjahr}`;
                    reportName = ' BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH';
                    loaibc = 1;
                    break;
                case "05":
                    lbl_merge = `Quý 4`;
                    lbl_merge_lk = `Lũy kế từ đầu năm đến cuối quý này`;
                    header_namthang = `Quý 4 năm ${gjahr}`;
                    reportName = ' BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH';
                    loaibc = 1;
                    break;
                case "06":
                    lbl_merge = `6 tháng đầu năm ${gjahr}`;
                    header_namthang = `Kỳ kế toán từ ngày 01/01/${gjahr} đến ngày 30/06/${gjahr}`;
                    reportName = 'BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH GIỮA NIÊN ĐỘ';
                    loaibc = 2;
                    break;
                case "07":
                    lbl_merge = `6 tháng cuối năm ${gjahr}`;
                    header_namthang = `Kỳ kế toán từ ngày 01/07/${gjahr} đến ngày 31/12/${gjahr}`;
                    reportName = 'BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH GIỮA NIÊN ĐỘ';
                    loaibc = 2;
                    break;
                case "08":
                    lbl_merge = `Năm`;
                    header_namthang = `Năm ${gjahr}`;
                    reportName = ' BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH';
                    loaibc = 2;
                    break;
            }
            //
            if (loaibc == 1) { // option tháng và quý
                // Header
                worksheet.mergeCells('A2:D2');
                worksheet.getCell('A2').value = header_congty;
                worksheet.getCell('A2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'left' } };

                worksheet.mergeCells('A3:D3');
                worksheet.getCell('A3').value = header_diachi;
                worksheet.getCell('A3').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'left' } };

                worksheet.mergeCells('F2:G2');

                worksheet.mergeCells('F3:G3');
                worksheet.mergeCells('F4:G4');

                if (gjahr < 2026) {
                    worksheet.getCell('F2').value = header_mauso;
                    worksheet.getCell('F2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('F3').value = header_banhanh;
                    worksheet.getCell('F3').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('F4').value = header_ngay;
                    worksheet.getCell('F4').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };
                } else {
                    if (type === "08") {
                        worksheet.getCell('F2').value = header_mauso_new;
                        worksheet.getCell('F2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };
                    } else {
                        worksheet.getCell('F2').value = header_mauso;
                        worksheet.getCell('F2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };
                    }

                    worksheet.getCell('F3').value = header_banhanh_new;
                    worksheet.getCell('F3').style = { font: { name: 'Times New Roman', size: 10, italic: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('F4').value = header_ngay_new;
                    worksheet.getCell('F4').style = { font: { name: 'Times New Roman', size: 10, italic: true }, alignment: { horizontal: 'center' } };
                }

                worksheet.mergeCells('A6:G6');
                worksheet.getCell('A6').value = reportName;
                worksheet.getCell('A6').style = { font: { name: 'Times New Roman', size: 14, bold: true }, alignment: { horizontal: 'center' } };

                worksheet.mergeCells('A7:G7');
                worksheet.getCell('A7').value = header_namthang;
                worksheet.getCell('A7').style = { font: { name: 'Times New Roman', size: 11, bold: false }, alignment: { horizontal: 'center' } };

                worksheet.getCell('G8').value = header_dvt;
                worksheet.getCell('G8').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'right' } };
                //
                worksheet.mergeCells('A9:A10');
                worksheet.mergeCells('B9:B10');
                worksheet.mergeCells('C9:C10');
                worksheet.mergeCells('D9:E9');
                worksheet.mergeCells('F9:G9');

                worksheet.getCell('A9').value = chiTieu;
                worksheet.getCell('B9').value = maSo;
                worksheet.getCell('C9').value = thuyetMinh;
                worksheet.getCell('D9').value = lbl_merge;
                worksheet.getCell('F9').value = lbl_merge_lk;
                // worksheet.getCell('D10').value = lbl_namnay;
                // worksheet.getCell('E10').value = lbl_namtruoc;
                // worksheet.getCell('F10').value = lbl_namnay;
                // worksheet.getCell('G10').value = lbl_namtruoc;

                // header = [chiTieu, maSo, thuyetMinh, namNay, namTruoc, namNay, namTruoc];
                header = [chiTieu, maSo, thuyetMinh, lbl_namnay, lbl_namtruoc, lbl_namnay, lbl_namtruoc];
                colWidth = [56.56, 10, 16, 27, 27, 27, 27];
                currentLine = 11;
            } else if (loaibc == 2) { // option 6 tháng

                worksheet.mergeCells('A2:C2');
                worksheet.getCell('A2').value = header_congty;
                worksheet.getCell('A2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'left' } };

                worksheet.mergeCells('A3:C3');
                worksheet.getCell('A3').value = header_diachi;
                worksheet.getCell('A3').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'left' } };

                worksheet.mergeCells('D2:E2');
                worksheet.mergeCells('D3:E3');
                worksheet.mergeCells('D4:E4');

                if (gjahr < 2026) {
                    worksheet.getCell('D2').value = header_mauso;
                    worksheet.getCell('D2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('D3').value = header_banhanh;
                    worksheet.getCell('D3').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('D4').value = header_ngay;
                    worksheet.getCell('D4').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };
                } else {
                    if (type === "08") {
                        worksheet.getCell('D2').value = header_mauso_new;
                        worksheet.getCell('D2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };
                    } else {
                        worksheet.getCell('D2').value = header_mauso;
                        worksheet.getCell('D2').style = { font: { name: 'Times New Roman', size: 10, bold: true }, alignment: { horizontal: 'center' } };
                    }

                    worksheet.getCell('D3').value = header_banhanh_new;
                    worksheet.getCell('D3').style = { font: { name: 'Times New Roman', size: 10, italic: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('D4').value = header_ngay_new;
                    worksheet.getCell('D4').style = { font: { name: 'Times New Roman', size: 10, italic: true }, alignment: { horizontal: 'center' } };
                }

                worksheet.mergeCells('A6:E6');
                worksheet.getCell('A6').value = reportName;
                worksheet.getCell('A6').style = { font: { name: 'Times New Roman', size: 14, bold: true }, alignment: { horizontal: 'center' } };
                worksheet.mergeCells('A7:E7');
                worksheet.getCell('A7').value = header_namthang;
                worksheet.getCell('A7').style = { font: { name: 'Times New Roman', size: 10, bold: false }, alignment: { horizontal: 'center' } };

                worksheet.getCell('E8').value = header_dvt;
                worksheet.getCell('E8').style = { font: { name: 'Times New Roman', size: 10, italic: true }, alignment: { horizontal: 'right' } };

                // Nhãn động


                worksheet.mergeCells('A9:A10');
                worksheet.mergeCells('B9:B10');
                worksheet.mergeCells('C9:C10');
                worksheet.mergeCells('D9:E9');
                worksheet.getCell('A9').value = chiTieu;
                worksheet.getCell('B9').value = maSo;
                worksheet.getCell('C9').value = thuyetMinh;
                worksheet.getCell('D9').value = lbl_merge;
                // worksheet.getCell('D10').value = lbl_namnay;
                // worksheet.getCell('E10').value = lbl_namtruoc;

                // header = [chiTieu, maSo, thuyetMinh, namNay, namTruoc];
                header = [chiTieu, maSo, thuyetMinh, lbl_namnay, lbl_namtruoc];
                colWidth = [56.56, 10, 16, 27, 27, 27, 27];
                currentLine = 11;
            } else { //option năm 
                // header = [chiTieu, maSo, thuyetMinh, namNay, namTruoc];
                header = [chiTieu, maSo, thuyetMinh, lbl_namnay, lbl_namtruoc];
                colWidth = [56.56, 10, 16, 27, 27, 27, 27];
                currentLine = 10;
            }

            header.forEach((value, i) => {
                const cell = worksheet.getCell(currentLine - 1, i + 1);
                cell.value = value;
                cell.style = {
                    font: { name: 'Times New Roman', size: 11, bold: true, color: { argb: '000000' } },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } // bạn có thể đổi màu nếu cần
                };

                const cell9 = worksheet.getCell(9, i + 1);
                cell9.style = {
                    font: { name: 'Times New Roman', size: 11, bold: true, color: { argb: '000000' } },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } // bạn có thể đổi màu nếu cần
                };

                const cellSTT = worksheet.getCell(currentLine, i + 1);
                cellSTT.value = i + 1;
                cellSTT.style = {
                    font: { name: 'Times New Roman', size: 11, bold: true, color: { argb: '000000' } },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } // bạn có thể đổi màu nếu cần
                };
            });

            colWidth.forEach((width, i) => worksheet.getColumn(i + 1).width = width);
            currentLine = currentLine + 1; //bắt đầu đổ data từ đây
            endline = currentLine;
            aData.forEach((item, i) => {
                endline = endline + 1;
                if (loaibc == 1) {
                    rowData = [item.HierarchyNode_TXT, item.HierarchyNode, '', item.sokynay, item.sokytruoc, item.lkkynay, item.lkkytruoc];
                } else {
                    rowData = [item.HierarchyNode_TXT, item.HierarchyNode, '', item.sokynay, item.sokytruoc];
                }

                rowData.forEach((value, j) => {
                    if (j >= numberId) {
                        try {
                            trueValue = Number(value);
                        } catch (error) {
                            trueValue = value;
                        }
                    } else {
                        trueValue = value;
                    }
                    worksheet.getCell(currentLine + i, j + 1).value = trueValue;
                    worksheet.getCell(currentLine + i, j + 1).style = {
                        font: {
                            name: 'Times New Roman', size: 11,
                            bold: String(item.type).toUpperCase() === '1',
                            italic: String(item.type).toUpperCase() === '2' && (j == 0 || j == 2)
                        },
                        alignment: { horizontal: j === 1 ? 'center' : j >= numberId ? 'right' : 'left' },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                    if (j >= numberId) {
                        worksheet.getCell(currentLine + i, j + 1).numFmt = accountingFormat;
                    }

                });
            });
            endline = endline + 2.

            if (loaibc == 1) {
                worksheet.getCell(endline, 7).value = '....ngày....tháng....năm....';
                worksheet.getCell(endline, 7).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'right' },
                };
                // worksheet.mergeCells(endline +1 , 4, endline + 1, 5);
                worksheet.getCell(endline + 1, 1).value = nguoiLapBieu;
                worksheet.getCell(endline + 1, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(endline + 1, 4, endline + 1, 5);
                worksheet.getCell(endline + 1, 4).value = keToanTruong;
                worksheet.getCell(endline + 1, 4).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };
                //  worksheet.mergeCells(endline + 1, 4, endline + 1, 5);
                worksheet.getCell(endline + 1, 7).value = giamDoc;
                worksheet.getCell(endline + 1, 7).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };
                worksheet.getCell(endline + 2, 1).value = hoTen;
                worksheet.getCell(endline + 2, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: false },
                    alignment: { horizontal: 'center' },
                };

                worksheet.getCell(endline + 7, 1).value = tenNL;
                worksheet.getCell(endline + 7, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(endline + 2, 4, endline + 2, 5);
                worksheet.getCell(endline + 2, 4).value = hoTen;
                worksheet.getCell(endline + 2, 4).style = {
                    font: { name: 'Times New Roman', size: 10, bold: false },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(endline + 7, 4, endline + 7, 5);
                worksheet.getCell(endline + 7, 4).value = tenKT;
                worksheet.getCell(endline + 7, 4).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };


                //  worksheet.mergeCells(endline + 1, 4, endline + 1, 5);
                worksheet.getCell(endline + 2, 7).value = dongDau;
                worksheet.getCell(endline + 2, 7).style = {
                    font: { name: 'Times New Roman', size: 10, bold: false },
                    alignment: { horizontal: 'center' },
                };

                worksheet.getCell(endline + 7, 7).value = tenGD;
                worksheet.getCell(endline + 7, 7).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                endline = endline + 9.
                worksheet.getCell(endline, 1).value = ghichu;
                worksheet.getCell(endline, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'left' },
                };
                worksheet.getCell(endline + 1, 1).value = dong1;
                worksheet.getCell(endline + 1, 1).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'left' },
                };
                worksheet.getCell(endline + 2, 1).value = dong2;
                worksheet.getCell(endline + 2, 1).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'left' },
                };
                worksheet.getCell(endline + 3, 1).value = dong3;
                worksheet.getCell(endline + 3, 1).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'left' },
                };
            }
            else {
                worksheet.getCell(endline, 5).value = '....ngày....tháng....năm....';
                worksheet.getCell(endline, 5).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'right' },
                };
                // worksheet.mergeCells(endline +1 , 4, endline + 1, 5);
                worksheet.getCell(endline + 1, 1).value = nguoiLapBieu;
                worksheet.getCell(endline + 1, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };
                worksheet.mergeCells(endline + 1, 2, endline + 1, 3);
                worksheet.getCell(endline + 1, 2).value = keToanTruong;
                worksheet.getCell(endline + 1, 2).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };
                //  worksheet.mergeCells(endline + 1, 4, endline + 1, 5);
                worksheet.getCell(endline + 1, 5).value = giamDoc;
                worksheet.getCell(endline + 1, 5).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };
                worksheet.getCell(endline + 2, 1).value = hoTen;
                worksheet.getCell(endline + 2, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: false },
                    alignment: { horizontal: 'center' },
                };

                worksheet.getCell(endline + 7, 1).value = tenNL;
                worksheet.getCell(endline + 7, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(endline + 2, 2, endline + 2, 3);
                worksheet.getCell(endline + 2, 2).value = hoTen;
                worksheet.getCell(endline + 2, 2).style = {
                    font: { name: 'Times New Roman', size: 10, bold: false },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(endline + 7, 2, endline + 7, 3);
                worksheet.getCell(endline + 7, 2).value = tenKT;
                worksheet.getCell(endline + 7, 2).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                //  worksheet.mergeCells(endline + 1, 4, endline + 1, 5);
                worksheet.getCell(endline + 2, 5).value = dongDau;
                worksheet.getCell(endline + 2, 5).style = {
                    font: { name: 'Times New Roman', size: 10, bold: false },
                    alignment: { horizontal: 'center' },
                };

                worksheet.getCell(endline + 7, 5).value = tenGD;
                worksheet.getCell(endline + 7, 5).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                endline = endline + 9.
                worksheet.getCell(endline, 1).value = ghichu;
                worksheet.getCell(endline, 1).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'left' },
                };
                worksheet.getCell(endline + 1, 1).value = dong1;
                worksheet.getCell(endline + 1, 1).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'left' },
                };
                worksheet.getCell(endline + 2, 1).value = dong2;
                worksheet.getCell(endline + 2, 1).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'left' },
                };
                worksheet.getCell(endline + 3, 1).value = dong3;
                worksheet.getCell(endline + 3, 1).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'left' },
                };
            }
            try {
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'Bao_cao_ket_qua_hoat_dong_kinh_doanh.xlsx';
                link.click();
                window.URL.revokeObjectURL(url);
                MessageToast.show("Export Successful!");
            } catch (error) {
                MessageToast.show("Export failed: " + error.message);
            }
        }
    };
});
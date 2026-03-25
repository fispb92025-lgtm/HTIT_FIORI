sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    'use strict';

    return {
        EXCEL: async function (oEvent) {

            const oSmartTable = this.getView().byId("zbcdkt::sap.suite.ui.generic.template.ListReport.view.ListReport::ZI_BCDKT--listReport");
            const oTable = oSmartTable.getTable();
            // const oBinding = oTable.getBinding("rows");
            const oBinding = oTable.getBinding("rows");

            sap.ui.core.BusyIndicator.show();

            try {
                var aData = oBinding.getContexts(0, oBinding.getLength()).map(function (oContext) {
                    return oContext.getObject();
                });
                if (!aData || aData.length === 0) {
                    sap.ui.core.BusyIndicator.hide();
                    MessageToast.show("Không có dữ liệu để export.");
                    return;
                }
            } catch (error) {
                sap.ui.core.BusyIndicator.hide();
                MessageToast.show("Không có dữ liệu để export.");
                return;
            }

            const oFilterBar = this.byId("zbcdkt::sap.suite.ui.generic.template.ListReport.view.ListReport::ZI_BCDKT--listReportFilter"); // ID của SmartFilterBar
            const oFilterData = oFilterBar.getFilterData();

            //Base Url
            const sBaseUrl = "/sap/bc/http/sap/zhttp_bcdkt";

            var sUrl = "";
            // Get data Raw
            sUrl = `${sBaseUrl}?`
                + `&companycode=${oFilterData.bukrs}`
                + `&ledger=${oFilterData.rldnr}`
                + `&fiscalyear=${oFilterData.gjahr}`
                + `&thang=${oFilterData.monat}`
                + `&loaibc=${oFilterData.type}`
                + `&chitiet=${oFilterData.ShowDetail}`;

            // const aData = await this._callHttpService(sUrl);

            // if (!aData || aData.length === 0) {
            //     MessageToast.show("Không có dữ liệu để export.");
            //     sap.ui.core.BusyIndicator.hide(); // Ẩn spinner
            //     return;
            // }

            const TaiSan = aData.filter(item => item.HierarchyNode < "300");
            const NguonVon = aData.filter(item => item.HierarchyNode >= "300");

            var lastLine = 0;
            const showDetail = oFilterData.ShowDetail;
            const type = oFilterData.type;
            const monat = oFilterData.monat;
            const gjahr = oFilterData.gjahr;
            const tenNL = oFilterData.tenNL;
            const tenKT = oFilterData.tenKT;
            const tenGD = oFilterData.tenGD;

            var reportName;
            var SoDK = 'Số đầu kỳ';
            var SoCK = 'Số cuối kỳ';
            var header;
            var header2;
            var lineWidth;
            var numberId;
            var rowData;
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
            var ghiChu = 'Ghi chú';
            var dong1 = ' 1). Những chỉ tiêu không có số liệu có thể không phải trình bày nhưng không được đánh lại số thứ tự chỉ tiêu và "Mã số" ';
            var dong2 = ' (2) Số liệu trong các chỉ tiêu có dấu (*) được ghi bằng số âm dưới hình thức ghi trong ngoặc đơn (…). ';
            var dong3 = ' (3) Đối với doanh nghiệp có kỳ kế toán năm là năm dương lịch (X) thì “Số cuối năm” có thể ghi là ';
            var dong4 = ' “31.12.X”; “Số đầu năm” có thể ghi là “01.01.X”. ';
            var subName1;
            var subName2;
            var gw_period;
            var lw_period;
            var tenCongTy = aData[0].CCname;
            var dcCongTy = aData[0].CCadrr;
            var trueValue;
            const accountingFormat = '_(* #,##_);_(* (#,##);_(* "-"_);_(@_)';

            const today = new Date();
            const day = today.getDate().toString().padStart(2, '0');           // → ngày (1–31)
            const month = (today.getMonth() + 1).toString().padStart(2, '0');    // → tháng (0–11 nên cần +1)
            const year = today.getFullYear();      // → năm

            var ngayThang = `......, Ngày ${day} tháng ${month} năm ${year}`;
            //Xử lý header
            // tên report
            if (type == 4 && showDetail !== "true") {
                if (gjahr < 2026) {
                    reportName = 'BẢNG CÂN ĐỐI KẾ TOÁN GIỮA NIÊN ĐỘ';
                } else {
                    reportName = 'BÁO CÁO TÌNH HÌNH TÀI CHÍNH GIỮA NIÊN ĐỘ';
                }
            } else {
                if (gjahr < 2026) {
                    reportName = 'BẢNG CÂN ĐỐI KẾ TOÁN';
                } else {
                    reportName = 'BÁO CÁO TÌNH HÌNH TÀI CHÍNH';
                }
            }

            switch (type) {
                case "02":
                    gw_period = 3;
                    lw_period = 'Quý 1';
                    break;
                case "03":
                    gw_period = 6;
                    lw_period = 'Quý 2';
                    break;
                case "05":
                    gw_period = 9;
                    lw_period = 'Quý 3';
                    break;
                case "06":
                    gw_period = 12;
                    lw_period = 'Quý 4';
                    break;
                case "07":
                    gw_period = 16.
                    break;
                case "01":
                    gw_period = monat;
                    break;
            }

            let lw_month_next = parseInt(gw_period, 10) + 1;
            let lw_year_next = gjahr;
            if (lw_month_next > 12) {
                lw_month_next = 1;
                lw_year_next++;
            }

            let gw_budat_last = `${lw_year_next}${String(lw_month_next).padStart(2, '0')}01`; // ví dụ: 20250101
            let lw_date = gw_budat_last;

            if (type === "07") {
                subName1 = `Năm ${gjahr}`;
                SoDK = 'Số đầu năm';
                SoCK = 'Số cuối năm';
            } else if (type === "01") {
                // Chuyển chuỗi ngày về dạng Date để trừ 1 ngày
                let d = new Date(`${lw_year_next}-${String(lw_month_next).padStart(2, '0')}-01`);
                d.setDate(d.getDate() - 1);
                lw_date = d;

                subName1 = `Tại ngày ${String(d.getDate()).padStart(2, '0')} tháng ${String(d.getMonth() + 1).padStart(2, '0')} năm ${d.getFullYear()}`;
            } else if (type === "04") {
                let d = new Date(`${lw_year_next}-${String(lw_month_next).padStart(2, '0')}-01`);
                d.setDate(d.getDate() - 1);
                lw_date = d;

                // subName1 = lw_period;
                // subName2 = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;

                subName1 = `Tại ngày 30/6/${gjahr}`;
            } else {
                subName1 = `${lw_period}/${gjahr}`;
            }

            if (type == "02") {
                subName2 = `Tại ngày 31 tháng 03 năm ${gjahr}`;
            };
            if (type == "03") {
                subName2 = `Tại ngày 30 tháng 06 năm ${gjahr}`;
            };
            if (type == "05") {
                subName2 = `Tại ngày 30 tháng 09 năm ${gjahr}`;
            };
            if (type == "06") {
                subName2 = `Tại ngày 31 tháng 12 năm ${gjahr}`;
            };
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Sheet1');
            worksheet.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
            if (showDetail == "true") {
                header = ['TÀI SẢN', 'Mã số', 'Thuyết minh', 'Mã NCC/KH', 'GL account', SoCK, SoDK];
                header2 = ['NGUỒN VỐN', '', '', '', '', '', ''];
                lineWidth = [51, 24.33, 33.67, 31.67, 31.67, 27, 27];
                numberId = 5;
                worksheet.views = [
                    { state: 'normal', zoomScale: 70 } // Zoom khi mở file
                ];
            } else {
                header = ['TÀI SẢN', 'Mã số', 'Thuyết minh', SoCK, SoDK];
                header2 = ['NGUỒN VỐN', '', '', '', ''];
                lineWidth = [51, 11, 10.44, 27, 27];
                numberId = 3;
                worksheet.views = [
                    { state: 'normal', zoomScale: 100 } // Zoom khi mở file
                ];
            };
            // Company header
            if (showDetail !== "true") {
                worksheet.mergeCells('A1:C1');
                worksheet.getCell('A1').value = tenCongTy;
                worksheet.getCell('A1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'left', wrapText: true } };

                worksheet.mergeCells('A2:C3');
                worksheet.getCell('A2').value = dcCongTy;
                worksheet.getCell('A2').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'left', wrapText: true } };

                worksheet.mergeCells('D1:E1');

                worksheet.mergeCells('D2:E2');
                worksheet.mergeCells('D3:E3');

                if (gjahr < 2026) {
                    worksheet.getCell('D1').value = 'Mẫu số B 01a - DN';
                    worksheet.getCell('D1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('D2').value = '(Ban hành theo Thông tư số 200/2014/TT-BTC ';
                    worksheet.getCell('D2').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('D3').value = 'Ngày 22/12/2014 của Bộ Tài chính) ';
                    worksheet.getCell('D3').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };
                } else {
                    if (type === "07") {
                        worksheet.getCell('D1').value = 'Mẫu số B 01 - DN';
                        worksheet.getCell('D1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };
                    } else {
                        worksheet.getCell('D1').value = 'Mẫu số B 01a - DN';
                        worksheet.getCell('D1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };
                    }

                    worksheet.getCell('D2').value = '(Kèm theo Thông tư số 99/2025/TT-BTC';
                    worksheet.getCell('D2').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('D3').value = 'ngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính) ';
                    worksheet.getCell('D3').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'center' } };
                }

                worksheet.mergeCells('A5:E5');
                worksheet.getCell('A5').value = reportName;
                worksheet.getCell('A5').style = { font: { name: 'Times New Roman', size: 16, bold: true }, alignment: { horizontal: 'center' } };

                worksheet.mergeCells('A6:E6');
                worksheet.getCell('A6').value = subName1;
                worksheet.getCell('A6').style = { font: { name: 'Times New Roman', size: 11 }, alignment: { horizontal: 'center' } };

                worksheet.mergeCells('A7:E7');
                worksheet.getCell('A7').value = subName2;
                worksheet.getCell('A7').style = { font: { name: 'Times New Roman', size: 11 }, alignment: { horizontal: 'center' } };

                worksheet.getCell('E9').value = 'ĐVT: VND';
                worksheet.getCell('E9').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'right' } };

            } else {
                worksheet.mergeCells('A1:C1');
                worksheet.getCell('A1').value = tenCongTy;
                worksheet.getCell('A1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'left' } };

                worksheet.mergeCells('A2:C2');
                worksheet.getCell('A2').value = dcCongTy;
                worksheet.getCell('A2').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'left' } };

                worksheet.mergeCells('F1:G1');

                worksheet.mergeCells('F2:G2');
                worksheet.mergeCells('F3:G3');

                if (gjahr < 2026) {

                    worksheet.getCell('F1').value = 'Mẫu số B 01a - DN';
                    worksheet.getCell('F1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('F2').value = '(Ban hành theo Thông tư số 200/2014/TT-BTC ';
                    worksheet.getCell('F2').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('F3').value = 'Ngày 22/12/2014 của Bộ Tài chính) ';
                    worksheet.getCell('F3').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };
                } else {
                    if (type === "07") {
                        worksheet.getCell('F1').value = 'Mẫu số B 01 - DN';
                        worksheet.getCell('F1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };
                    } else {
                        worksheet.getCell('F1').value = 'Mẫu số B 01a - DN';
                        worksheet.getCell('F1').style = { font: { name: 'Times New Roman', size: 11, bold: true }, alignment: { horizontal: 'center' } };
                    }

                    worksheet.getCell('F2').value = '(Kèm theo Thông tư số 99/2025/TT-BTC';
                    worksheet.getCell('F2').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'center' } };

                    worksheet.getCell('F3').value = 'ngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính) ';
                    worksheet.getCell('F3').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'center' } };
                }

                worksheet.mergeCells('A5:G5');
                worksheet.getCell('A5').value = reportName;
                worksheet.getCell('A5').style = { font: { name: 'Times New Roman', size: 16, bold: true }, alignment: { horizontal: 'center' } };

                worksheet.mergeCells('A6:G6');
                worksheet.getCell('A6').value = subName1;
                worksheet.getCell('A6').style = { font: { name: 'Times New Roman', size: 11 }, alignment: { horizontal: 'center' } };

                worksheet.mergeCells('A7:G7');
                worksheet.getCell('A7').value = subName2;
                worksheet.getCell('A7').style = { font: { name: 'Times New Roman', size: 11 }, alignment: { horizontal: 'center' } };

                worksheet.getCell('G9').value = 'ĐVT: VND';
                worksheet.getCell('G9').style = { font: { name: 'Times New Roman', size: 11, italic: true }, alignment: { horizontal: 'right' } };
            };

            //     Tai san
            header.forEach((header, i) => {
                worksheet.getCell(10, i + 1).value = header;
                worksheet.getCell(10, i + 1).style = {
                    font: { name: 'Times New Roman', size: 11, bold: true },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    fill: {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFFF' }
                    }
                };
                worksheet.getCell(11, i + 1).value = i + 1;
                worksheet.getCell(11, i + 1).style = {
                    font: { name: 'Times New Roman', size: 11, bold: true },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    fill: {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFFF' }
                    }
                };
            });
            // Column widths
            lineWidth.forEach((width, i) => worksheet.getColumn(i + 1).width = width);

            // Data Tài sản
            lastLine = 12;
            TaiSan.forEach((item, i) => {
                if (showDetail == "true") {
                    rowData = [item.HierarchyNode_TXT, item.HierarchyNode, item.HierarchyNode_TXT, item.kunnr, item.glaccount, item.SoCK, item.SoDK];
                } else {
                    rowData = [item.HierarchyNode_TXT, item.HierarchyNode, item.glaccount, item.SoCK, item.SoDK];
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
                    worksheet.getCell(lastLine + i, j + 1).value = trueValue;
                    worksheet.getCell(lastLine + i, j + 1).style = {
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
                        worksheet.getCell(lastLine + i, j + 1).numFmt = accountingFormat;
                    }

                });
            });

            // Nguồn vốn
            lastLine = lastLine + TaiSan.length;
            header2.forEach((header, i) => {
                worksheet.getCell(lastLine, i + 1).value = header;
                worksheet.getCell(lastLine, i + 1).style = {
                    font: { name: 'Times New Roman', size: 11, bold: true },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    fill: {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFFF' }
                    }
                };
            });

            lastLine = lastLine + 1;
            NguonVon.forEach((item, i) => {
                if (showDetail == "true") {
                    rowData = [item.HierarchyNode_TXT, item.HierarchyNode, item.HierarchyNode_TXT, item.kunnr, item.glaccount, item.SoCK, item.SoDK];
                } else {
                    rowData = [item.HierarchyNode_TXT, item.HierarchyNode, item.glaccount, item.SoCK, item.SoDK];
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

                    worksheet.getCell(lastLine + i, j + 1).value = trueValue;
                    worksheet.getCell(lastLine + i, j + 1).style = {
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
                        worksheet.getCell(lastLine + i, j + 1).numFmt = accountingFormat;
                    }
                });
            });

            // Chân ký
            lastLine = lastLine + NguonVon.length + 1;
            if (showDetail == "true") {
                // dòng 1
                worksheet.mergeCells(lastLine, 5, lastLine, 6);
                worksheet.getCell(lastLine, 5).value = ngayThang // Ngày tháng
                worksheet.getCell(lastLine, 5).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(lastLine + 2, 5, lastLine + 2, 6);
                worksheet.getCell(lastLine + 2, 5).value = dongDau;
                worksheet.getCell(lastLine + 2, 5).style = {
                    font: { name: 'Times New Roman', size: 10 },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(lastLine + 1, 5, lastLine + 1, 6);
                worksheet.getCell(lastLine + 1, 5).value = giamDoc // Giám đốc
                worksheet.getCell(lastLine + 1, 5).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(lastLine + 6, 5, lastLine + 6, 6);
                worksheet.getCell(lastLine + 6, 5).value = tenGD // Giám đốc
                worksheet.getCell(lastLine + 6, 5).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };
            } else {
                // dòng 1
                worksheet.mergeCells(lastLine, 4, lastLine, 5);
                worksheet.getCell(lastLine, 4).value = ngayThang // Ngày tháng
                worksheet.getCell(lastLine, 4).style = {
                    font: { name: 'Times New Roman', size: 10, italic: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(lastLine + 2, 4, lastLine + 2, 5);
                worksheet.getCell(lastLine + 2, 4).value = dongDau;
                worksheet.getCell(lastLine + 2, 4).style = {
                    font: { name: 'Times New Roman', size: 10 },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(lastLine + 1, 4, lastLine + 1, 5);
                worksheet.getCell(lastLine + 1, 4).value = giamDoc // Giám đốc
                worksheet.getCell(lastLine + 1, 4).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

                worksheet.mergeCells(lastLine + 6, 4, lastLine + 6, 5);
                worksheet.getCell(lastLine + 6, 4).value = tenGD // Giám đốc
                worksheet.getCell(lastLine + 6, 4).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'center' },
                };

            }
            // dòng 2
            worksheet.getCell(lastLine + 1, 1).value = nguoiLapBieu; // người lập biểu
            worksheet.getCell(lastLine + 1, 1).style = {
                font: { name: 'Times New Roman', size: 10, bold: true },
                alignment: { horizontal: 'center' },
            };

            worksheet.mergeCells(lastLine + 1, 2, lastLine + 1, 3);
            worksheet.getCell(lastLine + 1, 2).value = keToanTruong // kế toán trưởng
            worksheet.getCell(lastLine + 1, 2).style = {
                font: { name: 'Times New Roman', size: 10, bold: true },
                alignment: { horizontal: 'center' },
            };

            // dòng 3
            worksheet.getCell(lastLine + 2, 1).value = hoTen;
            worksheet.getCell(lastLine + 2, 1).style = {
                font: { name: 'Times New Roman', size: 10 },
                alignment: { horizontal: 'center' },
            };

            worksheet.mergeCells(lastLine + 2, 2, lastLine + 2, 3);
            worksheet.getCell(lastLine + 2, 2).value = hoTen;
            worksheet.getCell(lastLine + 2, 2).style = {
                font: { name: 'Times New Roman', size: 10 },
                alignment: { horizontal: 'center' },
            };

            // chu ky

            worksheet.getCell(lastLine + 6, 1).value = tenNL; // người lập biểu
            worksheet.getCell(lastLine + 6, 1).style = {
                font: { name: 'Times New Roman', size: 10, bold: true },
                alignment: { horizontal: 'center' },
            };

            worksheet.mergeCells(lastLine + 6, 2, lastLine + 6, 3);
            worksheet.getCell(lastLine + 6, 2).value = tenKT // kế toán trưởng
            worksheet.getCell(lastLine + 6, 2).style = {
                font: { name: 'Times New Roman', size: 10, bold: true },
                alignment: { horizontal: 'center' },
            };


            lastLine = lastLine + 8;
            worksheet.getCell(lastLine, 1).value = ghiChu;
            worksheet.getCell(lastLine, 1).style = {
                font: { name: 'Times New Roman', size: 10, bold: true, italic: true },
                alignment: { horizontal: 'left' },
            };
            worksheet.getCell(lastLine + 1, 1).value = dong1;
            worksheet.getCell(lastLine + 1, 1).style = {
                font: { name: 'Times New Roman', size: 10, italic: true },
                alignment: { horizontal: 'left' },
            };
            worksheet.getCell(lastLine + 2, 1).value = dong2;
            worksheet.getCell(lastLine + 2, 1).style = {
                font: { name: 'Times New Roman', size: 10, italic: true },
                alignment: { horizontal: 'left' },
            };
            worksheet.getCell(lastLine + 3, 1).value = dong3;
            worksheet.getCell(lastLine + 3, 1).style = {
                font: { name: 'Times New Roman', size: 10, italic: true },
                alignment: { horizontal: 'left' },
            };
            worksheet.getCell(lastLine + 4, 1).value = dong4;
            worksheet.getCell(lastLine + 4, 1).style = {
                font: { name: 'Times New Roman', size: 10, italic: true },
                alignment: { horizontal: 'left' },
            };

            // Export file
            try {
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'Bao_cao_tinh_hinh_tai_chinh.xlsx';
                link.click();
                window.URL.revokeObjectURL(url);
                sap.ui.core.BusyIndicator.hide(); // Ẩn spinner
                MessageToast.show("Export Successful!");
            } catch (error) {
                sap.ui.core.BusyIndicator.hide(); // Ẩn spinner
                MessageToast.show("Export failed: " + error.message);
            }
            //END xủ lý dl
        },
        _callHttpService: async function (sUrl) {

            try {
                const res = await fetch(sUrl, {
                    method: "GET",
                    headers: {
                        "Content-Type": "application/json",
                        "Accept": "*/*",
                        "Accept-Encoding": "gzip, deflate, br"
                    }
                });
                if (!res.ok) throw new Error(res.statusText);
                return await res.json();
            } catch (e) {
                console.error("Lỗi fetch", e);
                MessageToast.show("Gọi API thất bại: " + e.message);
                sap.ui.core.BusyIndicator.hide();
                return {};
            }
        },

    };
});
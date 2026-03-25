sap.ui.define([
    "sap/m/MessageToast",
    "exceljs",
], function (MessageToast, ExcelJS) {
    'use strict';

    return {
        exportExcel: async function (oEvent) {
            const oSmartTable = this.getView().byId("zaccpaysummary::sap.suite.ui.generic.template.ListReport.view.ListReport::ZC_ACCPAY_SUMMARY--listReport");
            if (!oSmartTable) {
                MessageToast.show("Không tìm thấy bảng dữ liệu.");
                return;
            }
            const oTable = oSmartTable.getTable();
            if (!oTable) {
                MessageToast.show("Không tìm thấy bảng dữ liệu.");
                return;
            }
            const oBinding = oTable.getBinding("rows");
            if (!oBinding) {
                MessageToast.show("Không tìm thấy dữ liệu để export.");
                return;
            }


            // const aContexts = oBinding.getCurrentContexts();
            // const dataArray = aContexts.map(c => c.getObject());

            const aContexts = oBinding.getContexts(0, oBinding.getLength());
            const dataArray = aContexts.map(c => c.getObject());

            if (dataArray.length === 0) {
                MessageToast.show("Không có dòng dữ liệu để export.");
                return;
            }

            const filteredData = dataArray;

            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Sheet1');

            // Set page setup for A4 landscape
            worksheet.pageSetup = {
                paperSize: 9, // A4
                orientation: 'landscape',
                margins: {
                    left: 0.5, right: 0.5, top: 0.5, bottom: 0.5,
                    header: 0.1, footer: 0.1
                },
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 0 // auto height
            };

            // Default company information
            let companyNameVN = 'CÔNG TY TNHH CẢNG QUỐC TẾ TIL CẢNG HẢI PHÒNG';
            let companyNameEN = 'HAIPHONG PORT TIL INTERNATIONAL TERMINAL COMPANY LIMITED';
            let companyAddressVN = 'Bến 3&4 Cảng nước sâu Lạch Huyện, Khu phố Đôn Lương, Đặc khu Cát Hải, Thành phố Hải Phòng, Việt Nam';
            let companyAddressEN = 'Berth No. 3&4 Lach Huyen Deep-sea Port, Don Luong Quarter, Cat Hai Special Zone, Hai Phong City, Vietnam';
            const companyCode = filteredData[0]?.RBUKRS || '6710'; // Default to 6710 if not available

            // Fetch company information from API
            try {
                const response = await fetch(`/sap/bc/http/sap/zhttp_common_core?name=companycode&companycode=${companyCode}`, {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json',
                        'Cookie': 'sap-usercontext=sap-client=100'
                    }
                });

                if (response.ok) {
                    const data = await response.json();
                    companyNameVN = data?.Companycodename || companyNameVN;
                    companyAddressVN = data?.Companycodeaddr || companyAddressVN;
                }
            } catch (error) {
                console.error('Failed to fetch company data:', error);
            }


            // ===== HEADER CÔNG TY =====
            worksheet.mergeCells('A2:Q2');
            worksheet.getCell('A2').value = companyNameVN;
            worksheet.getCell('A2').style = {
                font: { name: 'Times New Roman', size: 12, bold: true },
                alignment: { horizontal: 'left' }
            };

            worksheet.mergeCells('A3:Q3');
            worksheet.getCell('A3').value = companyNameEN;
            worksheet.getCell('A3').style = {
                font: { name: 'Times New Roman', size: 12, bold: true, color: { argb: 'FF0070C0' } },
                alignment: { horizontal: 'left' }
            };

            worksheet.mergeCells('A4:Q4');
            worksheet.getCell('A4').value = companyAddressVN;
            worksheet.getCell('A4').style = {
                font: { name: 'Times New Roman', size: 12, bold: true },
                alignment: { horizontal: 'left' }
            };

            worksheet.mergeCells('A5:Q5');
            worksheet.getCell('A5').value = companyAddressEN;
            worksheet.getCell('A5').style = {
                font: { name: 'Times New Roman', size: 12, bold: true, color: { argb: 'FF0070C0' } },
                alignment: { horizontal: 'left' }
            };

            // Report title
            worksheet.mergeCells('A7:Q7');
            worksheet.getCell('A7').value = 'BÁO CÁO TỔNG HỢP CÔNG NỢ PHẢI TRẢ';
            worksheet.getCell('A7').style = {
                font: { name: 'Times New Roman', size: 14, bold: true },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            // report title eng
            worksheet.mergeCells('A8:Q8');
            worksheet.getCell('A8').value = 'ACCOUNTS PAYABLE SUMMARY REPORT';
            worksheet.getCell('A8').style = {
                font: { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FF0070C0' } },
                alignment: { horizontal: 'center' }
            };

            // Helper function to format date to dd/mm/yyyy
            function formatDate(date) {
                if (!date) return '';
                let d = date;
                if (typeof date === 'string') {
                    d = new Date(date);
                }
                if (isNaN(d)) return '';
                const day = String(d.getDate()).padStart(2, '0');
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const year = d.getFullYear();
                return `${day}/${month}/${year}`;
            }

            // Date range
            // note: get data from p_start_date and p_end_date in the model aData
            const firstRow = filteredData[0] || {}; // get the first row to extract date range
            const startDate = formatDate(firstRow.p_start_date);
            const endDate = formatDate(firstRow.p_end_date);

            // Date range
            worksheet.mergeCells('A10:Q10');
            worksheet.getCell('A10').value = {
                richText: [
                    { text: 'Từ ngày ', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: '/ From', font: { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FF0070C0' } } },
                    { text: ': ', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: startDate + '  ', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: 'Đến ngày ', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: '/ To', font: { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FF0070C0' } } },
                    { text: ': ', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: endDate, font: { name: 'Times New Roman', size: 14, bold: true } }
                ]
            };
            worksheet.getCell('A10').alignment = { horizontal: 'center', vertical: 'middle' };

            // Lấy danh sách tài khoản duy nhất từ dữ liệu
            const accounts = [...new Set(filteredData.map(item => item.AccountNumber).filter(Boolean))];

            // Nếu chỉ có 1 tài khoản thì hiển thị, nếu nhiều thì để trống
            const accountNumber = accounts.length === 1 ? accounts[0] : '';

            worksheet.mergeCells('A9:Q9');
            worksheet.getCell('A9').value = {
                richText: [
                    { text: 'Tài khoản /', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: ' Account', font: { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FF0070C0' } } },
                    { text: ': ', font: { name: 'Times New Roman', size: 14, bold: true } },
                    { text: accountNumber, font: { name: 'Times New Roman', size: 14, bold: true } }
                ]
            };
            worksheet.getCell('A9').alignment = { horizontal: 'center' };

            // // Currency info
            // worksheet.mergeCells('A11:Q11');
            // worksheet.getCell('A11').value = {
            //     richText: [
            //         { text: 'Loại tiền ', font: { name: 'Times New Roman', size: 14, bold: true } },
            //         { text: '/ Currency', font: { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FF0070C0' } } },
            //         { text: ': ', font: { name: 'Times New Roman', size: 14, bold: true } },
            //         { text: firstRow.COMPANYCODECURRENCY || '', font: { name: 'Times New Roman', size: 14, bold: true } }
            //     ]
            // };
            worksheet.getCell('A11').alignment = { horizontal: 'center', vertical: 'middle' };

            // Table headers
            const headersVN = ['STT', 'Mã đối tượng', 'Tên đối tượng', 'Tài khoản', 'Đầu kỳ', '', '', '', 'Phát sinh trong kỳ', '', '', '', 'Cuối kỳ', '', '', '', 'Ghi chú'];
            const headersEN = ['No.', 'Partner code', 'Partner name', 'Account', 'Opening balance', '', '', '', 'Arising', '', '', '', 'Closing balance', '', '', '', 'Remark'];
            const subHeaders = ['', '', '', '', 'Nợ / Debit', '', 'Có / Credit', '', 'Nợ / Debit', '', 'Có / Credit', '', 'Nợ / Debit', '', 'Có / Credit', '', ' '];
            const currencyHeaders = ['', '', '', '', 'Ngoại tệ', 'VND', 'Ngoại tệ', 'VND', 'Ngoại tệ', 'VND', 'Ngoại tệ', 'VND', 'Ngoại tệ', 'VND', 'Ngoại tệ', 'VND', ''];
            const subHeaders2 = ['A', 'B', 'C', 'D', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', 'E'];

            // Thêm 5 dòng
            worksheet.spliceRows(13, 0, headersVN);
            worksheet.spliceRows(14, 0, headersEN);
            worksheet.spliceRows(15, 0, subHeaders);
            worksheet.spliceRows(16, 0, currencyHeaders);
            worksheet.spliceRows(17, 0, subHeaders2);

            // Merge các nhóm cột
            worksheet.mergeCells('E13:H13'); worksheet.mergeCells('E14:H14');
            worksheet.mergeCells('I13:L13'); worksheet.mergeCells('I14:L14');
            worksheet.mergeCells('M13:P13'); worksheet.mergeCells('M14:P14');
            worksheet.mergeCells('E15:F15'); worksheet.mergeCells('G15:H15');
            worksheet.mergeCells('I15:J15'); worksheet.mergeCells('K15:L15');
            worksheet.mergeCells('M15:N15'); worksheet.mergeCells('O15:P15');

            // Style chung cho headers
            [13, 14, 15, 16, 17].forEach(r => {
                worksheet.getRow(r).eachCell(cell => {
                    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    cell.border = { top: { style: 'medium' }, bottom: { style: 'medium' }, left: { style: 'medium' }, right: { style: 'medium' } };
                    cell.font = { name: 'Times New Roman', size: 10, color: { argb: 'FF000000' } }; // mặc định đen
                });
            });

            // Xử lý màu riêng cho row 14 (EN -> xanh)
            headersEN.forEach((text, colIndex) => {
                const cell = worksheet.getCell(14, colIndex + 1);
                cell.font = { name: 'Times New Roman', size: 10, color: { argb: 'FF0070C0' } };
            });
            // Xử lý màu riêng cho row 15 (Nợ / Debit, Có / Credit - VN đen, EN xanh)
            subHeaders.forEach((text, colIndex) => {
                const cell = worksheet.getCell(15, colIndex + 1);
                if (text.includes('/')) {
                    const [vn, en] = text.split('/');
                    cell.value = {
                        richText: [
                            { text: vn.trim() + ' ', font: { name: 'Times New Roman', size: 10 } },
                            { text: ' / ', font: { name: 'Times New Roman', size: 10 } },
                            { text: en.trim(), font: { name: 'Times New Roman', size: 10, color: { argb: 'FF0070C0' } } }
                        ]
                    };
                }
            });

            // Row 13 (Đầu kỳ, Phát sinh trong kỳ, Cuối kỳ) - viền trên đậm
            worksheet.getRow(13).eachCell(cell => {
                cell.border = {
                    top: { style: 'medium' },
                    bottom: { style: 'none' },
                    left: { style: 'medium' },
                    right: { style: 'medium' }
                };
            });

            // Row 14 
            worksheet.getRow(14).eachCell((cell, colNumber) => {
                if (colNumber >= 5 && colNumber <= 16) {
                    cell.border = {
                        top: { style: 'none' },
                        bottom: { style: 'medium' },
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                } else {
                    cell.border = {
                        top: { style: 'none' },
                        bottom: { style: 'none' },
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                }
            });

            // Row 15 (Nợ / Debit, Có / Credit) 
            worksheet.getRow(15).eachCell((cell, colNumber) => {
                if (colNumber >= 5 && colNumber <= 16) {
                    cell.border = {
                        top: { style: 'none' },
                        bottom: { style: 'medium' },
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                } else {
                    cell.border = {
                        top: { style: 'none' },
                        bottom: { style: 'none' },
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                }
            });

            // Row 16 (Ngoại tệ, VND) 
            worksheet.getRow(16).eachCell((cell, colNumber) => {
                if (colNumber >= 5 && colNumber <= 16) {
                    cell.border = {
                        top: { style: 'none' },
                        bottom: { style: 'medium' },
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                } else {
                    cell.border = {
                        top: { style: 'none' },
                        bottom: { style: 'none' },
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                }
            });

            // Row 17 (A B C D 1 2 3 ... 12 E)
            worksheet.getRow(17).eachCell((cell, colNumber) => {
                const isLeftPart = colNumber <= 4;   // A đến D (col 1-4)
                const isRightPart = colNumber === 17; // Q (col 17)

                cell.border = {
                    top: (isLeftPart || isRightPart) ? { style: 'medium' } : { style: 'none' },
                    bottom: { style: 'medium' },
                    left: { style: 'medium' },
                    right: { style: 'medium' }
                };
            });

            // Giữ viền trái/phải đầy đủ cho cột A-D và Q ở tất cả các row header
            ['A13', 'B13', 'C13', 'D13', 'Q13',
                'A14', 'B14', 'C14', 'D14', 'Q14',
                'A15', 'B15', 'C15', 'D15', 'Q15',
                'A16', 'B16', 'C16', 'D16', 'Q16',
                'A17', 'B17', 'C17', 'D17', 'Q17'].forEach(cellStr => {
                    const cell = worksheet.getCell(cellStr);
                    cell.border = {
                        ...(cell.border || {}),
                        left: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                });


            // Set column widths
            const columnWidths = [8, 13, 50, 13, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 35];
            columnWidths.forEach((width, index) => {
                worksheet.getColumn(index + 1).width = width;
            });

            // ===== DATA bắt đầu từ dòng 18 =====
            filteredData.forEach((item, idx) => {
                const rowIndex = 18 + idx;
                const rowData = [
                    idx + 1,
                    item.BP || '',
                    item.BP_NAME || '',
                    item.AccountNumber || '',
                    Number(item.OPEN_DEBIT_TRAN || 0) !== 0 ? Number(item.OPEN_DEBIT_TRAN) : '',
                    Number(item.OPEN_DEBIT || 0) !== 0 ? Number(item.OPEN_DEBIT) : '',
                    Number(item.OPEN_CREDIT_TRAN || 0) !== 0 ? Number(item.OPEN_CREDIT_TRAN) : '',
                    Number(item.OPEN_CREDIT || 0) !== 0 ? Number(item.OPEN_CREDIT) : '',
                    Number(item.TOTAL_DEBIT_TRAN || 0) !== 0 ? Number(item.TOTAL_DEBIT_TRAN) : '',
                    Number(item.TOTAL_DEBIT || 0) !== 0 ? Number(item.TOTAL_DEBIT) : '',
                    Number(item.TOTAL_CREDIT_TRAN || 0) !== 0 ? Number(item.TOTAL_CREDIT_TRAN) : '',
                    Number(item.TOTAL_CREDIT || 0) !== 0 ? Number(item.TOTAL_CREDIT) : '',
                    Number(item.END_DEBIT_TRAN || 0) !== 0 ? Number(item.END_DEBIT_TRAN) : '',
                    Number(item.END_DEBIT || 0) !== 0 ? Number(item.END_DEBIT) : '',
                    Number(item.END_CREDIT_TRAN || 0) !== 0 ? Number(item.END_CREDIT_TRAN) : '',
                    Number(item.END_CREDIT || 0) !== 0 ? Number(item.END_CREDIT) : '',
                    ''
                ];

                worksheet.addRow(rowData);
                worksheet.getRow(rowIndex).eachCell((cell, colIndex) => {
                    cell.style = {
                        font: { name: 'Times New Roman', size: 10 },
                        alignment: { horizontal: (colIndex >= 5 && colIndex <= 16) ? 'right' : 'left', vertical: 'middle', wrapText: true },
                        border: { top: { style: 'medium' }, bottom: { style: 'medium' }, left: { style: 'medium' }, right: { style: 'medium' } }
                    };

                    // Chỉ áp dụng định dạng số cho các cột tiền (5-16)
                    if (colIndex >= 5 && colIndex <= 16) {
                        // Cột Ngoại tệ: colIndex lẻ (5,7,9,11,13,15) → có 2 thập phân
                        // Cột VND: colIndex chẵn (6,8,10,12,14,16) → không thập phân
                        if (colIndex % 2 === 1) {
                            cell.numFmt = '#,##0.00';  // Ngoại tệ: 2 chữ số thập phân
                        } else {
                            cell.numFmt = '#,##0';      // VND: không thập phân
                        }
                    }
                });
            });

            filteredData.forEach((item, index) => {
                const rowIndex = 18 + index;
                const row = worksheet.getRow(rowIndex);

                // Cột chữ (C,Q) -> căn trái
                ['C', 'Q'].forEach(col => {
                    row.getCell(col).alignment = { horizontal: 'left', vertical: 'middle' };
                });

                // Cột chữ (A,B,D) -> căn giữa
                ['A', 'B', 'D'].forEach(col => {
                    row.getCell(col).alignment = { horizontal: 'center', vertical: 'middle' };
                });

                // Cột số (E to P) -> căn phải
                ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'].forEach(col => {
                    row.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' };
                });
            });

            // ===== TỔNG CỘNG =====
            const totalRowIndex = 18 + filteredData.length;

            worksheet.mergeCells(`A${totalRowIndex}:D${totalRowIndex}`);
            worksheet.getCell(`A${totalRowIndex}`).value = 'Tổng ';
            worksheet.getCell(`A${totalRowIndex}`).style = {
                font: { name: 'Times New Roman', size: 10, bold: true },
                alignment: { horizontal: 'center' },
            };

            ['A', 'B', 'C', 'D'].forEach(col => {
                worksheet.getCell(`${col}${totalRowIndex}`).border = {
                    top: { style: 'medium' },
                    bottom: { style: 'medium' },
                    left: { style: 'medium' },
                    right: { style: 'medium' }
                };
            });


            for (let col = 5; col <= 16; col++) {
                let total = 0;
                for (let row = 18; row < totalRowIndex; row++) {
                    const val = worksheet.getCell(row, col).value;

                    if (typeof val === 'number') {
                        total += val;
                    } else if (typeof val === 'string' && !isNaN(parseFloat(val))) {
                        total += parseFloat(val);
                    } else if (val === '' || val === null) {
                        total += 0;
                    }
                }

                worksheet.getCell(totalRowIndex, col).value = total || '';
                worksheet.getCell(totalRowIndex, col).style = {
                    font: { name: 'Times New Roman', size: 10, bold: true },
                    alignment: { horizontal: 'right' },
                    border: { top: { style: 'medium' }, bottom: { style: 'medium' }, left: { style: 'medium' }, right: { style: 'medium' } },
                    numFmt: (col % 2 === 1) ? '#,##0.00' : '#,##0'  // col lẻ: ngoại tệ có thập phân, chẵn: VND không
                };
            }

            // Kẻ viền cho cột Ghi chú (Q)
            worksheet.getCell(totalRowIndex, 17).style = {
                border: {
                    top: { style: 'medium' },
                    bottom: { style: 'medium' },
                    left: { style: 'medium' },
                    right: { style: 'medium' }
                },
                alignment: { horizontal: 'left', vertical: 'middle' }
            };

            // Add signature section
            const lastDataRow = 18 + filteredData.length;
            const signatureRow = lastDataRow + 2;

            // Date signature
            // worksheet.mergeCells(`E${signatureRow}:J${signatureRow}`);
            // Set the date signature cell to today's date format 'ngày dd tháng mm năm yyyy'
            const today = new Date();
            const formattedDateVN = `Ngày ${today.getDate()} tháng ${today.getMonth() + 1} năm ${today.getFullYear()}`;
            const formattedDateEN = `Date ${today.getDate()} month ${today.getMonth() + 1} year ${today.getFullYear()}`;

            worksheet.getCell(`Q${signatureRow}`).value = {
                richText: [
                    { text: formattedDateVN + '\n', font: { name: 'Times New Roman', size: 10, italic: true, color: { argb: 'FF000000' } } },
                    { text: formattedDateEN, font: { name: 'Times New Roman', size: 10, italic: true, color: { argb: 'FF0070C0' } } }
                ]
            };
            worksheet.getCell(`Q${signatureRow}`).alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };

            // Signature titles
            // Người lập biểu / Preparer
            worksheet.getCell(`B${signatureRow + 1}`).value = {
                richText: [
                    { text: "Người lập biểu\n", font: { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } } },
                    { text: "Preparer", font: { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF0070C0' } } }
                ]
            };
            worksheet.getCell(`B${signatureRow + 1}`).alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };

            // Kế toán trưởng / Chief Accountant
            worksheet.getCell(`Q${signatureRow + 1}`).value = {
                richText: [
                    { text: "Kế toán trưởng\n", font: { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } } },
                    { text: "Chief Accountant", font: { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF0070C0' } } }
                ]
            };
            worksheet.getCell(`Q${signatureRow + 1}`).alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };

            //Hide any empty rows if they exist
            for (let i = 18; i <= worksheet.lastRow.number; i++) {
                const row = worksheet.getRow(i);
                if (row.values.every(cell => cell === undefined || cell === null || cell === '' || cell === 0 || cell === ' ')) {
                    // If all cells in the row are empty, hide the row
                    row.hidden = true; // hide the row instead of deleting
                }
            }

            // // hide columns 'Nhóm đối tượng'
            // worksheet.getColumn(3).hidden = true;

            //Generate and download file
            try {
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });

                // Create download link
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'Export_VND_So_Tong_Hop_Cong_No_Phai_Tra.xlsx';
                link.click();
                // Clean up
                window.URL.revokeObjectURL(url);
                MessageToast.show("Export Successful!");
            } catch (error) {
                MessageToast.show("Error exporting Excel file: " + error.message);
            }
        },
        //
        //                       _oo0oo_
        //                      o8888888o
        //                      88" . "88
        //                      (| -_- |)
        //                      0\  =  /0
        //                    ___/`---'\___
        //                  .' \\|     |// '.
        //                 / \\|||  :  |||// \
        //                / _||||| -:- |||||- \
        //               |   | \\\  -  /// |   |
        //               | \_|  ''\---/''  |_/ |
        //               \  .-\__  '-'  ___/-. /
        //             ___'. .'  /--.--\  `. .'___
        //          ."" '<  `.___\_<|>_/___.' >' "".
        //         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
        //         \  \ `_.   \_ __\ /__ _/   .-` /  /
        //     =====`-.____`.___ \_____/___.-`___.-'=====
        //                       `=---='
        //
        //
        //     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        //
        // exportNew: async function (oEvent) {
        //     const oSmartTable = this.getView().byId("zaccpaysummary::sap.suite.ui.generic.template.ListReport.view.ListReport::ZC_ACCPAY_SUMMARY--listReport");
        //     if (!oSmartTable) {
        //         MessageToast.show("Không tìm thấy bảng dữ liệu.");
        //         return;
        //     }
        //     const oTable = oSmartTable.getTable();
        //     if (!oTable) {
        //         MessageToast.show("Không tìm thấy bảng dữ liệu.");
        //         return;
        //     }
        //     const oBinding = oTable.getBinding("rows");
        //     if (!oBinding) {
        //         MessageToast.show("Không tìm thấy dữ liệu để export.");
        //         return;
        //     }

        //     const aContexts = oBinding.getCurrentContexts();
        //     const dataArray = aContexts.map(c => c.getObject());

        //     if (dataArray.length === 0) {
        //         MessageToast.show("Không có dòng dữ liệu để export.");
        //         return;
        //     }

        //     const filteredData = dataArray;

        //     const workbook = new ExcelJS.Workbook();
        //     const worksheet = workbook.addWorksheet('Sheet1');

        //     // Set page setup for A4 landscape
        //     worksheet.pageSetup = {
        //         paperSize: 9, // A4
        //         orientation: 'landscape',
        //         margins: {
        //             left: 0.5, right: 0.5, top: 0.5, bottom: 0.5,
        //             header: 0.1, footer: 0.1
        //         },
        //         fitToPage: true,
        //         fitToWidth: 1,
        //         fitToHeight: 0 // auto height
        //     };

        //     // Default company information
        //     let companyName = 'DEMO CÔNG TY CỔ PHẦN CASLA';
        //     let companyAddress = 'DEMO Khu CN Châu Sơn, Phường Châu Sơn, Thành phố Phủ Lý, Tỉnh Hà Nam, Việt Nam';

        //     // Fetch company information from API
        //     try {
        //         const response = await fetch(`/sap/bc/http/sap/zhttp_common_core?name=companycode&companycode=${filteredData[0].RBUKRS}`, {
        //             method: 'GET',
        //             headers: {
        //                 'Content-Type': 'application/json',
        //                 'Cookie': 'sap-usercontext=sap-client=100'
        //             }
        //         });

        //         if (response.ok) {
        //             const data = await response.json();
        //             companyName = data?.Companycodename || companyName;
        //             companyAddress = data?.Companycodeaddr || companyAddress;
        //         }
        //     } catch (error) {
        //         console.error('Failed to fetch company data:', error);
        //     }
        //     // Header company information
        //     worksheet.mergeCells('A1:L1');
        //     worksheet.getCell('A1').value = `Đơn vị: ${companyName}`;
        //     worksheet.getCell('A1').style = {
        //         font: { name: 'Times New Roman', size: 11, bold: true },
        //         alignment: { horizontal: 'left', vertical: 'middle' }
        //     };

        //     worksheet.mergeCells('T1:AB1');
        //     worksheet.getCell('T1').value = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
        //     worksheet.getCell('T1').style = {
        //         font: { name: 'Times New Roman', size: 11, bold: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     // Second line of header
        //     worksheet.mergeCells('A2:L2');
        //     worksheet.getCell('A2').value = companyAddress;
        //     worksheet.getCell('A2').style = {
        //         font: { name: 'Times New Roman', size: 10 },
        //         alignment: { horizontal: 'left', vertical: 'middle' }
        //     };

        //     worksheet.mergeCells('T2:AB2');
        //     worksheet.getCell('T2').value = 'Độc lập - Tự do - Hạnh phúc';
        //     worksheet.getCell('T2').style = {
        //         font: { name: 'Times New Roman', size: 10, bold: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     // Report title
        //     worksheet.mergeCells('A4:AB4');
        //     worksheet.getCell('A4').value = 'SỔ TỔNG HỢP CÔNG NỢ PHẢI TRẢ';
        //     worksheet.getCell('A4').style = {
        //         font: { name: 'Times New Roman', size: 16, bold: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     // Date range
        //     const firstRow = filteredData[0] || {};
        //     const formatDate = (date) => {
        //         if (!date) return '';
        //         let d = date;
        //         if (typeof date === 'string') {
        //             d = new Date(date);
        //         }
        //         if (isNaN(d)) return '';
        //         const day = String(d.getDate()).padStart(2, '0');
        //         const month = String(d.getMonth() + 1).padStart(2, '0');
        //         const year = d.getFullYear();
        //         return `${day}/${month}/${year}`;
        //     };

        //     const startDate = formatDate(firstRow.p_start_date);
        //     const endDate = formatDate(firstRow.p_end_date);

        //     worksheet.mergeCells('A5:AB5');
        //     worksheet.getCell('A5').value = `Từ ${startDate} Đến ${endDate}`;
        //     worksheet.getCell('A5').style = {
        //         font: { name: 'Times New Roman', size: 12, bold: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     // Currency info
        //     worksheet.mergeCells('AA6:AB6');
        //     worksheet.getCell('AA6').value = `Loại tiền: ${firstRow.RHCUR || ''}`;
        //     worksheet.getCell('AA6').style = {
        //         font: { name: 'Times New Roman', size: 10 },
        //         alignment: { horizontal: 'right', vertical: 'middle' }
        //     };

        //     // Create complex header structure (rows 7-9)
        //     // Row 7 - Main headers with merging
        //     worksheet.mergeCells('A7:A8');
        //     worksheet.getCell('A7').value = 'Mã đối tượng';

        //     worksheet.mergeCells('B7:B8');
        //     worksheet.getCell('B7').value = 'Tên đối tượng';

        //     worksheet.mergeCells('C7:C8');
        //     worksheet.getCell('C7').value = 'Nhóm đối tượng';

        //     worksheet.mergeCells('D7:D8');
        //     worksheet.getCell('D7').value = 'Tài khoản';

        //     worksheet.mergeCells('E7:H7');
        //     worksheet.getCell('E7').value = 'Dư Nợ đầu Kỳ';

        //     worksheet.mergeCells('I7:L7');
        //     worksheet.getCell('I7').value = 'Dư Có đầu kỳ';

        //     worksheet.mergeCells('M7:P7');
        //     worksheet.getCell('M7').value = 'Tổng PS Nợ trong kỳ';

        //     worksheet.mergeCells('Q7:T7');
        //     worksheet.getCell('Q7').value = 'Tổng PS Có trong kỳ';

        //     worksheet.mergeCells('U7:X7');
        //     worksheet.getCell('U7').value = 'Dư Nợ cuối kỳ';

        //     worksheet.mergeCells('Y7:AB7');
        //     worksheet.getCell('Y7').value = 'Dư Có cuối kỳ';

        //     // Row 8 - Currency subheaders
        //     worksheet.mergeCells('E8:F8');
        //     worksheet.getCell('E8').value = 'Nguyên tệ';

        //     worksheet.mergeCells('G8:H8');
        //     worksheet.getCell('G8').value = 'VND';

        //     worksheet.mergeCells('I8:J8');
        //     worksheet.getCell('I8').value = 'Nguyên tệ';

        //     worksheet.mergeCells('K8:L8');
        //     worksheet.getCell('K8').value = 'VND';

        //     worksheet.mergeCells('M8:N8');
        //     worksheet.getCell('M8').value = 'Nguyên tệ';

        //     worksheet.mergeCells('O8:P8');
        //     worksheet.getCell('O8').value = 'VND';

        //     worksheet.mergeCells('Q8:R8');
        //     worksheet.getCell('Q8').value = 'Nguyên tệ';

        //     worksheet.mergeCells('S8:T8');
        //     worksheet.getCell('S8').value = 'VND';

        //     worksheet.mergeCells('U8:V8');
        //     worksheet.getCell('U8').value = 'Nguyên tệ';

        //     worksheet.mergeCells('W8:X8');
        //     worksheet.getCell('W8').value = 'VND';

        //     worksheet.mergeCells('Y8:Z8');
        //     worksheet.getCell('Y8').value = 'Nguyên tệ';

        //     worksheet.mergeCells('AA8:AB8');
        //     worksheet.getCell('AA8').value = 'VND';

        //     // Apply styles to all header cells
        //     for (let row = 7; row <= 8; row++) {
        //         for (let col = 1; col <= 28; col++) {
        //             const cell = worksheet.getCell(row, col);
        //             if (cell.value) {
        //                 cell.style = {
        //                     font: { name: 'Times New Roman', size: 9, bold: true },
        //                     alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        //                     border: {
        //                         top: { style: 'thin' },
        //                         bottom: { style: 'thin' },
        //                         left: { style: 'thin' },
        //                         right: { style: 'thin' }
        //                     }
        //                 };
        //                 // if E7, I7, M7, Q7, U7, Y7 are filled with color
        //                 if (row === 7 && (col === 5 || col === 9)) {
        //                     cell.style.fill = {
        //                         type: 'pattern',
        //                         pattern: 'solid',
        //                         fgColor: { argb: 'CCECFF' } // light blue for header
        //                     };
        //                 } else if (row === 7 && (col === 13 || col === 17)) {
        //                     cell.style.fill = {
        //                         type: 'pattern',
        //                         pattern: 'solid',
        //                         fgColor: { argb: 'ffcccc' } // light red for header
        //                     };
        //                 } else if (row === 7 && (col === 21 || col === 25)) {
        //                     cell.style.fill = {
        //                         type: 'pattern',
        //                         pattern: 'solid',
        //                         fgColor: { argb: 'ffffcc' } // light yellow for header
        //                     };
        //                 }
        //             }
        //         }
        //     }

        //     // Set column widths
        //     const columnWidths = [12, 25, 15, 12, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5, 18, 5];
        //     columnWidths.forEach((width, index) => {
        //         worksheet.getColumn(index + 1).width = width;
        //     });

        //     // Set header row heights
        //     worksheet.getRow(7).height = 25;
        //     worksheet.getRow(8).height = 25;

        //     // Add data rows starting from row 10
        //     filteredData.forEach((item, index) => {
        //         const rowIndex = 9 + index;
        //         const rowData = [
        //             item.BP || '',
        //             item.BP_NAME || '',
        //             item.BP_GR_TITLE || '',
        //             item.AccountNumber || '',
        //             Number(item.OPEN_DEBIT_TRAN || 0) !== 0 ? Number(item.OPEN_DEBIT_TRAN) : '',
        //             Number(item.OPEN_DEBIT_TRAN || 0) === 0 ? '' : item.RHCUR || '',
        //             Number(item.OPEN_DEBIT || 0) !== 0 ? Number(item.OPEN_DEBIT) : '',
        //             Number(item.OPEN_DEBIT || 0) === 0 ? '' : item.COMPANYCODECURRENCY || '',
        //             Number(item.OPEN_CREDIT_TRAN || 0) !== 0 ? Number(item.OPEN_CREDIT_TRAN) : '',
        //             Number(item.OPEN_CREDIT_TRAN || 0) === 0 ? '' : item.RHCUR || '',
        //             Number(item.OPEN_CREDIT || 0) !== 0 ? Number(item.OPEN_CREDIT) : '',
        //             Number(item.OPEN_CREDIT || 0) === 0 ? '' : item.COMPANYCODECURRENCY || '',
        //             Number(item.TOTAL_DEBIT_TRAN || 0) !== 0 ? Number(item.TOTAL_DEBIT_TRAN) : '',
        //             Number(item.TOTAL_DEBIT_TRAN || 0) === 0 ? '' : item.RHCUR || '',
        //             Number(item.TOTAL_DEBIT || 0) !== 0 ? Number(item.TOTAL_DEBIT) : '',
        //             Number(item.TOTAL_DEBIT || 0) === 0 ? '' : item.COMPANYCODECURRENCY || '',
        //             Number(item.TOTAL_CREDIT_TRAN || 0) !== 0 ? Number(item.TOTAL_CREDIT_TRAN) : '',
        //             Number(item.TOTAL_CREDIT_TRAN || 0) === 0 ? '' : item.RHCUR || '',
        //             Number(item.TOTAL_CREDIT || 0) !== 0 ? Number(item.TOTAL_CREDIT) : '',
        //             Number(item.TOTAL_CREDIT || 0) === 0 ? '' : item.COMPANYCODECURRENCY || '',
        //             Number(item.END_DEBIT_TRAN || 0) !== 0 ? Number(item.END_DEBIT_TRAN) : '',
        //             Number(item.END_DEBIT_TRAN || 0) === 0 ? '' : item.RHCUR || '',
        //             Number(item.END_DEBIT || 0) !== 0 ? Number(item.END_DEBIT) : '',
        //             Number(item.END_DEBIT || 0) === 0 ? '' : item.COMPANYCODECURRENCY || '',
        //             Number(item.END_CREDIT_TRAN || 0) !== 0 ? Number(item.END_CREDIT_TRAN) : '',
        //             Number(item.END_CREDIT_TRAN || 0) === 0 ? '' : item.RHCUR || '',
        //             Number(item.END_CREDIT || 0) !== 0 ? Number(item.END_CREDIT) : '',
        //             Number(item.END_CREDIT || 0) === 0 ? '' : item.COMPANYCODECURRENCY || ''
        //         ];

        //         rowData.forEach((value, colIndex) => {
        //             const cell = worksheet.getCell(rowIndex, colIndex + 1);
        //             cell.value = value;
        //             cell.style = {
        //                 font: { name: 'Times New Roman', size: 9 },
        //                 alignment: { horizontal: 'left', vertical: 'middle' },
        //                 border: {
        //                     top: { style: 'thin' },
        //                     bottom: { style: 'thin' },
        //                 }
        //             };
        //             // from colIndex 4 to 28, even indexes are top, bottom, left border, odd indexes are right border
        //             if (colIndex % 2 === 0 && colIndex < 28 && colIndex >= 4) {
        //                 cell.style.border.left = { style: 'thin' };
        //             } else {
        //                 cell.style.border.right = { style: 'thin' };
        //             }
        //             if (colIndex < 4) {
        //                 cell.style.border.right = { style: 'thin' };
        //                 cell.style.border.left = { style: 'thin' };
        //             }

        //             // Right align and format numeric columns
        //             if (colIndex >= 4) {
        //                 cell.style.alignment.horizontal = 'right';
        //                 if (typeof value === 'number' && value !== 0) {
        //                     cell.numFmt = '#,##0.00';
        //                 }
        //             }
        //             if (colIndex = 1) {
        //                 cell.style.alignment.wrapText = true;
        //             }


        //         });
        //     });

        //     // Add total row
        //     const totalRowIndex = 9 + filteredData.length;
        //     worksheet.getCell(totalRowIndex, 4).value = 'Tổng cộng:';
        //     worksheet.getCell(totalRowIndex, 4).style = {
        //         font: { name: 'Times New Roman', size: 9, bold: true },
        //         alignment: { horizontal: 'left', vertical: 'middle' },
        //         border: {
        //             top: { style: 'thin' },
        //             bottom: { style: 'thin' },
        //             left: { style: 'thin' },
        //             right: { style: 'thin' }
        //         }
        //     };

        //     // Calculate totals for numeric columns (5-28)
        //     for (let colIndex = 5; colIndex <= 28; colIndex += 2) {
        //         let total = 0;
        //         for (let rowIndex = 9; rowIndex < totalRowIndex; rowIndex++) {
        //             const cellValue = worksheet.getCell(rowIndex, colIndex).value;
        //             if (typeof cellValue === 'number') {
        //                 total += cellValue;
        //             }
        //         }
        //         worksheet.getCell(totalRowIndex, colIndex).value = total !== 0 ? total : '';
        //         worksheet.getCell(totalRowIndex, colIndex).style = {
        //             font: { name: 'Times New Roman', size: 9, bold: true },
        //             alignment: { horizontal: 'right', vertical: 'middle' },
        //             border: {
        //                 top: { style: 'thin' },
        //                 bottom: { style: 'thin' },
        //                 left: { style: 'thin' },
        //                 // right: { style: 'thin' }
        //             },
        //             numFmt: '#,##0.00'
        //         };
        //         // Find currency value from any cell above in this column
        //         let currencyValue = '';
        //         for (let searchRow = totalRowIndex - 1; searchRow >= 9; searchRow--) {
        //             const cellValue = worksheet.getCell(searchRow, colIndex + 1).value;
        //             if (cellValue && cellValue.trim() !== '') {
        //                 currencyValue = cellValue;
        //                 break; // Use the first non-empty currency value found
        //             }
        //         }
        //         worksheet.getCell(totalRowIndex, colIndex + 1).value = currencyValue;
        //         worksheet.getCell(totalRowIndex, colIndex + 1).style = {
        //             font: { name: 'Times New Roman', size: 9, bold: true },
        //             alignment: { horizontal: 'left', vertical: 'middle' },
        //             border: {
        //                 top: { style: 'thin' },
        //                 bottom: { style: 'thin' },
        //                 // left: { style: 'thin' },
        //                 right: { style: 'thin' }
        //             }
        //         };
        //     }

        //     // Add signature section
        //     const signatureRow = totalRowIndex + 3;
        //     const today = new Date();
        //     const formattedDate = `Ngày ${today.getDate()} tháng ${today.getMonth() + 1} năm ${today.getFullYear()}`;

        //     worksheet.mergeCells(`J${signatureRow}:AB${signatureRow}`);
        //     worksheet.getCell(`J${signatureRow}`).value = formattedDate;
        //     worksheet.getCell(`J${signatureRow}`).style = {
        //         font: { name: 'Times New Roman', size: 10, italic: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     worksheet.mergeCells(`A${signatureRow + 1}:I${signatureRow + 1}`);
        //     worksheet.getCell(`A${signatureRow + 1}`).value = 'Người lập biểu';
        //     worksheet.getCell(`A${signatureRow + 1}`).style = {
        //         font: { name: 'Times New Roman', size: 10, bold: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     worksheet.mergeCells(`J${signatureRow + 1}:AB${signatureRow + 1}`);
        //     worksheet.getCell(`J${signatureRow + 1}`).value = 'Kế toán trưởng';
        //     worksheet.getCell(`J${signatureRow + 1}`).style = {
        //         font: { name: 'Times New Roman', size: 10, bold: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     worksheet.mergeCells(`A${signatureRow + 2}:I${signatureRow + 2}`);
        //     worksheet.getCell(`A${signatureRow + 2}`).value = '(Ký, họ tên)';
        //     worksheet.getCell(`A${signatureRow + 2}`).style = {
        //         font: { name: 'Times New Roman', size: 9, italic: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     worksheet.mergeCells(`J${signatureRow + 2}:AB${signatureRow + 2}`);
        //     worksheet.getCell(`J${signatureRow + 2}`).value = '(Ký, họ tên)';
        //     worksheet.getCell(`J${signatureRow + 2}`).style = {
        //         font: { name: 'Times New Roman', size: 9, italic: true },
        //         alignment: { horizontal: 'center', vertical: 'middle' }
        //     };

        //     // Generate and download file
        //     const buffer = await workbook.xlsx.writeBuffer();
        //     const blob = new Blob([buffer], {
        //         type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        //     });

        //     // hide columns 'Nhóm đối tượng'
        //     worksheet.getColumn(3).hidden = true;

        //     //Generate and download file
        //     try {
        //         const buffer = await workbook.xlsx.writeBuffer();
        //         const blob = new Blob([buffer], {
        //             type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        //         });

        //         // Create download link
        //         const url = window.URL.createObjectURL(blob);
        //         const link = document.createElement('a');
        //         link.href = url;
        //         link.download = 'Export_So_Tong_Hop_Cong_No_Phai_Tra.xlsx';
        //         link.click();

        //         // Clean up
        //         window.URL.revokeObjectURL(url);

        //         MessageToast.show("Export Successful!");
        //     } catch (error) {
        //         MessageToast.show("Error exporting Excel file: " + error.message);
        //     }
        // }
    };
});
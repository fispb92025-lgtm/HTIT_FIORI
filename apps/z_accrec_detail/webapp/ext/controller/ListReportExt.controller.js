sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    'use strict';

    // Constants for better maintainability
    const CONSTANTS = {
        SMART_TABLE_ID: "zaccrecdetail::sap.suite.ui.generic.template.ListReport.view.ListReport::ZC_ACCREC_DETAIL--listReport",
        PAPER_SIZE: 9,
        MARGINS: {
            left: 0.5, right: 0.5,
            top: 0.5, bottom: 0.5,
            header: 0.1, footer: 0.1
        },
        FONT_NAME: 'Times New Roman',
        DEFAULT_COMPANY: {
            name: 'CÔNG TY TNHH CẢNG QUỐC TẾ TIL CẢNG HẢI PHÒNG',
            nameUS: 'HAIPHONG PORT TIL INTERNATIONAL TERMINAL COMPANY LIMITED',
            addressVN: 'Bến 3&4 Cảng nước sâu Lạch Huyện, Khu phố Đôn Lương, Đặc khu Cát Hải, Thành phố Hải Phòng, Việt Nam',
            addressUS: 'Berth No. 3&4 Lach Huyen Deep-sea Port, Don Luong Quarter, Cat Hai Special Zone, Hai Phong City, Vietnam'
        },
        API_ENDPOINT: '/sap/bc/http/sap/zhttp_common_core',
        MIME_TYPE: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        REQUIRED_FIELDS: ['CompanyCode', 'GLAccountNumber'],
        BORDER_STYLES: {
            thin: 'medium',
            dotted: 'dotted',
            none: 'none'
        },
        VND_COLUMNS: {
            DATE: { index: 1, width: 15 },
            DOC_DATE: { index: 2, width: 15 },
            DOC_NUMBER: { index: 3, width: 17 },
            INVOICE_NUMBER: { index: 4, width: 18 }, // CỘT MỚI: Số hóa đơn
            DESCRIPTION: { index: 5, width: 35 },
            OFFSET_ACCOUNT: { index: 6, width: 20 },
            DEBIT_FC: { index: 7, width: 22 },
            DEBIT_VND: { index: 8, width: 22 },
            CREDIT_FC: { index: 9, width: 22 },
            CREDIT_VND: { index: 10, width: 22 },
            BAL_DEBIT_FC: { index: 11, width: 22 },
            BAL_DEBIT_VND: { index: 12, width: 22 },
            BAL_CREDIT_FC: { index: 13, width: 22 },
            BAL_CREDIT_VND: { index: 14, width: 22 }
        }
    };

    return {
        exportExcel: async function (oEvent) {
            try {
                const exportData = await this._getExportData();
                if (!exportData) return;

                const companyInfo = await this._fetchCompanyInfo(exportData[0].CompanyCode);
                const workbook = this._createWorkbook();

                this._generateVNDWorksheets(workbook, exportData, companyInfo);
                await this._downloadWorkbook(workbook, 'VND_So_chi_tiet_cong_no_phai_thu');
            } catch (error) {
                this._handleError("Failed to generate Excel file", error);
            }
        },

        _getExportData: async function () {
            const smartTable = this._getSmartTable();
            if (!smartTable) return null;

            const table = smartTable.getTable();
            if (!table) {
                MessageToast.show("Không tìm thấy bảng dữ liệu.");
                return null;
            }

            const binding = table.getBinding("rows");
            if (!binding) {
                MessageToast.show("Không tìm thấy dữ liệu để export.");
                return null;
            }

            const contexts = binding.getContexts(0, binding.getLength());
            const dataArray = contexts.map(context => context.getObject());

            if (dataArray.length === 0) {
                MessageToast.show("Không có dòng dữ liệu để export.");
                return null;
            }
// Here
            dataArray.push({});

            const dataArrayNew = [];
            let previousRow = {}

                dataArray.forEach((row, index) => {
                const isSameGroup =
                    row.CompanyCode === previousRow.CompanyCode &&
                    row.BusinessPartner === previousRow.BusinessPartner &&
                    row.GLAccountNumber === previousRow.GLAccountNumber;

                if (!isSameGroup) {
                    if (previousRow && Object.keys(previousRow).length > 0) {
                        dataArrayNew.push(previousRow);
                    }

                    previousRow = row;
                } else {
                    // Parse current JSON safely
                    const currentJson = JSON.parse(row.LineItemsJson);
                    const currentItems = currentJson.LINE_ITEMS || [];

                    const prevItems = JSON.parse(previousRow.LineItemsJson).LINE_ITEMS;
                    const mergedItems = prevItems.concat(currentItems);

                    previousRow.PostingDateFrom = row.PostingDateFrom
                    previousRow.PostingDateTo = row.PostingDateTo

                    previousRow.OpeningCreditBalance = Number(previousRow.OpeningCreditBalance) + Number(row.OpeningCreditBalance);
                    previousRow.OpeningCreditBalanceTran = Number(previousRow.OpeningCreditBalanceTran) + Number(row.OpeningCreditBalanceTran);
                    previousRow.OpeningDebitBalance = Number(previousRow.OpeningDebitBalance) + Number(row.OpeningDebitBalance);
                    previousRow.OpeningDebitBalanceTran = Number(previousRow.OpeningDebitBalanceTran) + Number(row.OpeningDebitBalanceTran);

                    if (previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance > 0) {
                        previousRow.OpeningCreditBalance = previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance;
                        previousRow.OpeningDebitBalance = 0;
                    } else if (previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance < 0) {
                        previousRow.OpeningDebitBalance = Math.abs(previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance);
                        previousRow.OpeningCreditBalance = 0;
                    } else if (previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance == 0) {
                           previousRow.OpeningDebitBalance = 0; 
                           previousRow.OpeningCreditBalance = 0; 
                    }

                    if (previousRow.OpeningCreditBalanceTran - previousRow.OpeningDebitBalanceTran > 0) {
                        previousRow.OpeningCreditBalanceTran = previousRow.OpeningCreditBalanceTran - previousRow.OpeningDebitBalanceTran;
                        previousRow.OpeningDebitBalanceTran = 0;
                    } else if (previousRow.OpeningCreditBalanceTran - previousRow.OpeningDebitBalanceTran < 0) {
                        previousRow.OpeningDebitBalanceTran = Math.abs(previousRow.OpeningCreditBalanceTran - previousRow.OpeningDebitBalanceTran);
                        previousRow.OpeningCreditBalanceTran = 0;
                    } else if (previousRow.OpeningCreditBalanceTran - previousRow.OpeningDebitBalanceTran == 0) {
                        previousRow.OpeningDebitBalanceTran = 0;
                        previousRow.OpeningCreditBalanceTran = 0;
                    }

                    // Update previous row’s LineItemsJson
                    //previousRow.LineItemsJson = JSON.stringify({ LINE_ITEMS: mergedItems });

                    previousRow.CreditAmountDuringPeriod = Number(previousRow.CreditAmountDuringPeriod) + Number(row.CreditAmountDuringPeriod);
                    previousRow.CreditAmountDuringPeriodTran = Number(previousRow.CreditAmountDuringPeriodTran) + Number(row.CreditAmountDuringPeriodTran);
                    previousRow.DebitAmountDuringPeriod = Number(previousRow.DebitAmountDuringPeriod) + Number(row.DebitAmountDuringPeriod);
                    previousRow.DebitAmountDuringPeriodTran = Number(previousRow.DebitAmountDuringPeriodTran) + Number(row.DebitAmountDuringPeriodTran);

                    previousRow.ClosingCredit = Number(previousRow.ClosingCredit) + Number(row.ClosingCredit);
                    previousRow.ClosingCreditTran = Number(previousRow.ClosingCreditTran) + Number(row.ClosingCreditTran);
                    previousRow.ClosingDebit = Number(previousRow.ClosingDebit) + Number(row.ClosingDebit);
                    previousRow.ClosingDebitTran = Number(previousRow.ClosingDebitTran) + Number(row.ClosingDebitTran);

                    if (previousRow.ClosingCredit - previousRow.ClosingDebit > 0) {
                        previousRow.ClosingCredit = previousRow.ClosingCredit - previousRow.ClosingDebit;
                        previousRow.ClosingDebit = 0;
                    } else if (previousRow.ClosingCredit - previousRow.ClosingDebit < 0) {
                        previousRow.ClosingDebit = Math.abs(previousRow.ClosingCredit - previousRow.ClosingDebit);
                        previousRow.ClosingCredit = 0;
                     } else if (previousRow.ClosingCredit - previousRow.ClosingDebit == 0) {
                        previousRow.ClosingDebit = 0;
                        previousRow.ClosingCredit = 0;
                    }

                    if (previousRow.ClosingCreditTran - previousRow.ClosingDebitTran > 0) {
                        previousRow.ClosingCreditTran = previousRow.ClosingCreditTran - previousRow.ClosingDebitTran;
                        previousRow.ClosingDebitTran = 0;
                    } else if (previousRow.ClosingCreditTran - previousRow.ClosingDebitTran < 0) {
                        previousRow.ClosingDebitTran = Math.abs(previousRow.ClosingCreditTran - previousRow.ClosingDebitTran);
                        previousRow.ClosingCreditTran = 0;
                    } else if (previousRow.ClosingCreditTran - previousRow.ClosingDebitTran == 0) {
                        previousRow.ClosingDebitTran = 0;
                        previousRow.ClosingCreditTran = 0;
                    }
                    // 
                    let running_balance_vnd = previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance;
                    let running_balance_tran = previousRow.OpeningCreditBalanceTran - previousRow.OpeningDebitBalanceTran;
                    let lines = {};
                    mergedItems.forEach(item => {
                        const lines = this._normalizeItemFields(item);

                        const debit_vnd  = Number(lines.debit_amount || 0);
                        const credit_vnd = Number(lines.credit_amount || 0);
                        const debit_tran = Number(lines.debit_amount_tran || 0);
                        const credit_tran= Number(lines.credit_amount_tran || 0);

                        running_balance_vnd += credit_vnd - debit_vnd;
                        running_balance_tran += credit_tran - debit_tran;

                        item.ClosingDebit = running_balance_vnd < 0 ? Math.abs(running_balance_vnd) : 0;
                        item.ClosingCredit = running_balance_vnd > 0 ? running_balance_vnd : 0;
                        item.ClosingDebitTran = running_balance_tran < 0 ? Math.abs(running_balance_tran) : 0;
                        item.ClosingCreditTran = running_balance_tran > 0 ? running_balance_tran : 0;
                    });
                    previousRow.LineItemsJson = JSON.stringify({ LINE_ITEMS: mergedItems });

                    //
                    previousRow.TransactionCurrency = row.TransactionCurrency !== 'VND'
                        ? row.TransactionCurrency
                        : previousRow.TransactionCurrency;
                }
            });

            return dataArrayNew;
        },

        _getSmartTable: function () {
            const smartTable = this.getView().byId(CONSTANTS.SMART_TABLE_ID);
            if (!smartTable) {
                MessageToast.show("Không tìm thấy bảng dữ liệu.");
                return null;
            }
            return smartTable;
        },

        _validateRecord: function (record) {
            return CONSTANTS.REQUIRED_FIELDS.every(field => record[field]);
        },

        _fetchCompanyInfo: async function (companyCode) {
            let companyInfo = {
                nameVN: CONSTANTS.DEFAULT_COMPANY.name,
                nameUS: CONSTANTS.DEFAULT_COMPANY.nameUS,
                addressVN: CONSTANTS.DEFAULT_COMPANY.addressVN,
                addressUS: CONSTANTS.DEFAULT_COMPANY.addressUS
            };

            try {
                const response = await fetch(`${CONSTANTS.API_ENDPOINT}?name=companycode&companycode=${companyCode}`, {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json',
                        'Cookie': 'sap-usercontext=sap-client=100'
                    }
                });

                if (response.ok) {
                    const data = await response.json();
                    companyInfo.nameVN = data?.Companycodename || companyInfo.nameVN;
                    companyInfo.addressVN = data?.Companycodeaddr || companyInfo.addressVN;
                    companyInfo.nameUS = data?.CompanycodenameEN || companyInfo.nameUS;
                    companyInfo.addressUS = data?.CompanycodeaddrEN || companyInfo.addressUS;
                }
            } catch (error) {
                console.error('Failed to fetch company data:', error);
            }

            return companyInfo;
        },

        _createWorkbook: function () {
            return new ExcelJS.Workbook();
        },

        _generateVNDWorksheets: function (workbook, exportData, companyInfo) {
            exportData.forEach((record, i) => {
                if (!this._validateRecord(record)) return;
                const name = `${record.CompanyCode}-${record.GLAccountNumber}-${record.BusinessPartner} `;
                const worksheet = this._createWorksheet(workbook, record, name);
                this._createVNDHeader(worksheet, record);
                this._createVNDTable(worksheet, record);
                this._createVNDFooter(worksheet);
                this._applyFontToWorksheet(worksheet);
            });
        },

        _createWorksheet: function (workbook, record, name1) {
            const name = name1;
           // const worksheetName = `${record.CompanyCode}-${record.GLAccountNumber}-${record.BusinessPartner}`;
            return workbook.addWorksheet(name, {
                pageSetup: {
                    paperSize: CONSTANTS.PAPER_SIZE,
                    orientation: 'landscape',
                    fitToPage: true,
                    fitToWidth: 1,
                    fitToHeight: 0,
                    margins: CONSTANTS.MARGINS
                }
            });
        },

        _createVNDHeader: function (worksheet, record) {
            this._setCompanyInfo(worksheet);
            this._setTitle(worksheet);
            this._setAccountInfo(worksheet, record);
        },

        _setCompanyInfo: function (worksheet) {
            const ranges = ["A2:J2", "A3:J3", "A4:J4", "A5:J5"];
            const values = [
                "CÔNG TY TNHH CẢNG QUỐC TẾ TIL CẢNG HẢI PHÒNG",
                "HAIPHONG PORT TIL INTERNATIONAL TERMINAL COMPANY LIMITED",
                "Bến 3&4 Cảng nước sâu Lạch Huyện, Khu phố Đôn Lương, Đặc khu Cát Hải, Thành phố Hải Phòng, Việt Nam",
                "Berth No. 3&4 Lach Huyen Deep-sea Port, Don Luong Quarter, Cat Hai Special Zone, Hai Phong City, Vietnam"
            ];
            const fonts = [
                { size: 12, bold: true },
                { size: 12, bold: true, color: { argb: "FF0070C0" } },
                { size: 12, bold: true },
                { size: 12, bold: true, color: { argb: "FF0070C0" } }
            ];

            ranges.forEach((range, i) => {
                worksheet.mergeCells(range);
                const cell = worksheet.getCell(range.split(":")[0]);
                cell.value = values[i];
                cell.font = fonts[i];
                cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
            });
        },

        _setTitle: function (worksheet) {
            worksheet.mergeCells("A7:M7");
            worksheet.getCell("A7").value = "SỔ CHI TIẾT CÔNG NỢ PHẢI THU";
            worksheet.getCell("A7").font = { bold: true, size: 14 };
            worksheet.getCell("A7").alignment = { horizontal: "center" };

            worksheet.mergeCells("A8:M8");
            worksheet.getCell("A8").value = "DETAILED ACCOUNTS RECEIVABLE LEDGER";
            worksheet.getCell("A8").font = { bold: true, size: 14, color: { argb: "FF0070C0" } };
            worksheet.getCell("A8").alignment = { horizontal: "center" };
        },

        _setAccountInfo: function (worksheet, record) {
            const info = [
                ["Tài khoản / ", "Account", record.GLAccountNumber || ""],
                ["Đối tượng / ", "Partner", `${record.BusinessPartner} - ${record.BusinessPartnerName}`],
                ["Từ ngày / ", "From", this._formatDate(record.PostingDateFrom) || "", " Đến ngày / ", "To", this._formatDate(record.PostingDateTo) || ""]
            ];
            const rows = ["A9:M9", "A10:M10", "A11:M11"];

            rows.forEach((range, i) => {
                worksheet.mergeCells(range);
                const cell = worksheet.getCell(range.split(":")[0]);
                if (i < 2) {
                    cell.value = {
                        richText: [
                            { text: info[i][0], font: { bold: true } },
                            { text: info[i][1], font: { bold: true, color: { argb: "FF0070C0" } } },
                            { text: ": ", font: { bold: true } },
                            { text: info[i][2], font: { bold: true } }
                        ]
                    };
                } else {
                    cell.value = {
                        richText: [
                            { text: info[i][0], font: { bold: true } },
                            { text: info[i][1], font: { bold: true, color: { argb: "FF0070C0" } } },
                            { text: ": ", font: { bold: true } },
                            { text: info[i][2], font: { bold: true } },
                            { text: info[i][3], font: { bold: true } },
                            { text: info[i][4], font: { bold: true, color: { argb: "FF0070C0" } } },
                            { text: ": ", font: { bold: true } },
                            { text: info[i][5], font: { bold: true } }
                        ]
                    };
                }
                cell.font = { bold: true, size: 14 };
                cell.alignment = { horizontal: "center" };
            });
        },

        _createVNDTable: function (worksheet, record) {
            const headerRow = 14;
            this._createVNDTableHeaders(worksheet, headerRow);

            let currentRow = headerRow + 4;
            currentRow = this._addVNDOpeningBalance(worksheet, record, currentRow);
            currentRow = this._addVNDLineItems(worksheet, record, currentRow);
            currentRow = this._addVNDSummaryRows(worksheet, record, currentRow);
            this._setVNDColumnWidths(worksheet);
        },

        _createVNDTableHeaders: function (worksheet, headerRow) {
            const headers = [
                // Ngày ghi sổ
                {
                    range: `A${headerRow}:A${headerRow + 2}`, value: {
                        richText: [
                            { text: "Ngày ghi sổ\n", font: { bold: false } },
                            { text: "Posting date", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },

                // Chứng từ
                {
                    range: `B${headerRow}:C${headerRow}`, value: {
                        richText: [
                            { text: "Chứng từ\n", font: { bold: false } },
                            { text: "Voucher", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                {
                    range: `B${headerRow + 1}:B${headerRow + 2}`, value: {
                        richText: [
                            { text: "Ngày / ", font: { bold: false } },
                            { text: "Date", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                {
                    range: `C${headerRow + 1}:C${headerRow + 2}`, value: {
                        richText: [
                            { text: "Số / ", font: { bold: false } },
                            { text: "Number", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },

                // Số hóa đơn (CỘT MỚI)
                {
                    range: `D${headerRow}:D${headerRow + 2}`, value: {
                        richText: [
                            { text: "Số hóa đơn\n", font: { bold: false } },
                            { text: "Invoice Number", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },

                // Diễn giải
                {
                    range: `E${headerRow}:E${headerRow + 2}`, value: {
                        richText: [
                            { text: "Diễn giải\n", font: { bold: false } },
                            { text: "Description", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },

                // TK đối ứng
                {
                    range: `F${headerRow}:F${headerRow + 2}`, value: {
                        richText: [
                            { text: "TK đối ứng\n", font: { bold: false } },
                            { text: "Contra Accounts", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },

                // SỐ TIỀN
                {
                    range: `G${headerRow}:J${headerRow}`, value: {
                        richText: [
                            { text: "Số tiền / ", font: { bold: false } },
                            { text: "Amount", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                {
                    range: `G${headerRow + 1}:H${headerRow + 1}`, value: {
                        richText: [
                            { text: "Nợ / ", font: { bold: false } },
                            { text: "Debit", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                {
                    range: `I${headerRow + 1}:J${headerRow + 1}`, value: {
                        richText: [
                            { text: "Có / ", font: { bold: false } },
                            { text: "Credit", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                { range: `G${headerRow + 2}:G${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `H${headerRow + 2}:H${headerRow + 2}`, value: 'VND' },
                { range: `I${headerRow + 2}:I${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `J${headerRow + 2}:J${headerRow + 2}`, value: 'VND' },

                // SỐ DƯ
                {
                    range: `K${headerRow}:N${headerRow}`, value: {
                        richText: [
                            { text: "Số dư / ", font: { bold: false } },
                            { text: "Balance", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                {
                    range: `K${headerRow + 1}:L${headerRow + 1}`, value: {
                        richText: [
                            { text: "Nợ / ", font: { bold: false } },
                            { text: "Debit", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                {
                    range: `M${headerRow + 1}:N${headerRow + 1}`, value: {
                        richText: [
                            { text: "Có / ", font: { bold: false } },
                            { text: "Credit", font: { bold: false, color: { argb: "FF0070C0" } } }
                        ]
                    }
                },
                { range: `K${headerRow + 2}:K${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `L${headerRow + 2}:L${headerRow + 2}`, value: 'VND' },
                { range: `M${headerRow + 2}:M${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `N${headerRow + 2}:N${headerRow + 2}`, value: 'VND' }
            ];

            headers.forEach(h => {
                worksheet.mergeCells(h.range);
                const cell = worksheet.getCell(h.range.split(':')[0]);
                cell.value = h.value;
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.border = this._getAllBorders(CONSTANTS.BORDER_STYLES.thin);
            });

            worksheet.getRow(headerRow).height = 40;
            worksheet.getRow(headerRow + 1).height = 22;
            worksheet.getRow(headerRow + 2).height = 20;

            const signRow = headerRow + 3;
            ['A', 'B', 'C', 'D', 'E', 'F'].forEach((val, idx) => {
                const cell = worksheet.getCell(signRow, idx + 1);
                cell.value = val;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = this._getAllBorders(CONSTANTS.BORDER_STYLES.thin);
            });

            for (let i = 1; i <= 8; i++) {
                const colIndex = 6 + i; // bắt đầu từ cột G (index 7)
                const cell = worksheet.getCell(signRow, colIndex);
                cell.value = i;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = this._getAllBorders(CONSTANTS.BORDER_STYLES.thin);
            }
        },

        _addVNDOpeningBalance: function (worksheet, record, currentRow) {
            // Lấy giá trị ngoại tệ (FC) và VND riêng biệt
            const openingDebitFC = Number(record.OpeningDebitBalanceTran) || 0;  // Nợ ngoại tệ
            const openingCreditFC = Number(record.OpeningCreditBalanceTran) || 0;  // Có ngoại tệ
            const openingDebitVND = Number(record.OpeningDebitBalance) || 0;  // Nợ VND
            const openingCreditVND = Number(record.OpeningCreditBalance) || 0;  // Có VND

            const balanceFC = ( openingDebitFC * -1 ) + openingCreditFC;   // Số dư ngoại tệ
            const balanceVND = ( openingDebitVND * -1 ) + openingCreditVND;  // Số dư VND

            // Ghi dòng chữ "Số dư đầu kỳ"
            worksheet.getCell(currentRow, 5).value = {
                richText: [
                    { text: "Số dư đầu kỳ /", font: { bold: true, name: CONSTANTS.FONT_NAME } },
                    { text: " Opening balance", font: { bold: true, name: CONSTANTS.FONT_NAME, color: { argb: "FF0070C0" } } }
                ]
            };

            // Chỉ hiển thị số nếu có giá trị khác 0
            if (balanceVND < 0) {
                // Trường hợp Nợ
                if (balanceFC !== 0) {
                    this._setCurrencyCell(worksheet, currentRow, 11, Math.abs(balanceFC), true, true);   // Cột K: Nợ ngoại tệ
                }
                this._setCurrencyCell(worksheet, currentRow, 12, Math.abs(balanceVND), true, false);   // Cột L: Nợ VND
            } else if (balanceVND > 0) {
                // Trường hợp Có
                if (balanceFC !== 0) {
                    this._setCurrencyCell(worksheet, currentRow, 13, Math.abs(balanceFC), true, true);   // Cột M: Có ngoại tệ
                }
                this._setCurrencyCell(worksheet, currentRow, 14, Math.abs(balanceVND), true, false);   // Cột N: Có VND
            }
            // Nếu balanceVND == 0 nhưng balanceFC != 0 (rất hiếm) → vẫn hiển thị ngoại tệ

            this._applyRowBorders(worksheet, currentRow, 14, CONSTANTS.BORDER_STYLES.thin, CONSTANTS.BORDER_STYLES.dotted);
            return currentRow + 1;
        },

        _addVNDLineItems: function (worksheet, record, currentRow) {
            const lineItems = this._parseLineItems(record.LineItemsJson);

            if (lineItems.length > 0) {
                console.log("=== DỮ LIỆU LINE ITEMS CHI TIẾT (3 dòng đầu) ===");
                lineItems.slice(0, 3).forEach((item, i) => {
                    console.log(`Dòng ${i + 1}:`, item);
                    console.log("Các key có sẵn:", Object.keys(item));
                });
                console.log("================================================");
            }

            lineItems.forEach(item => {
                const normalizedItem = this._normalizeItemFields(item);
                this._populateVNDLineItem(worksheet, normalizedItem, currentRow);
                this._applyRowBorders(worksheet, currentRow, 14, CONSTANTS.BORDER_STYLES.dotted, CONSTANTS.BORDER_STYLES.dotted);
                currentRow++;
            });

            return currentRow;
        },

        _addVNDSummaryRows: function (worksheet, record, currentRow) {
            // Tách riêng ngoại tệ (Tran) và VND
            const debitFC = Number(record.DebitAmountDuringPeriodTran) || 0;  // Nợ ngoại tệ trong kỳ
            const debitVND = Number(record.DebitAmountDuringPeriod) || 0;  // Nợ VND trong kỳ
            const creditFC = Number(record.CreditAmountDuringPeriodTran) || 0;  // Có ngoại tệ trong kỳ
            const creditVND = Number(record.CreditAmountDuringPeriod) || 0;  // Có VND trong kỳ

            // Dòng "Số phát sinh trong kỳ"
            worksheet.getCell(currentRow, 5).value = {
                richText: [
                    { text: "Số phát sinh trong kỳ /", font: { bold: true, name: CONSTANTS.FONT_NAME } },
                    { text: " Arising amount", font: { bold: true, name: CONSTANTS.FONT_NAME, color: { argb: "FF0070C0" } } }
                ]
            };

            this._setCurrencyCell(worksheet, currentRow, 7, debitFC, true, true);   // Cột G: Nợ ngoại tệ
            this._setCurrencyCell(worksheet, currentRow, 8, debitVND, true, false);  // Cột H: Nợ VND
            this._setCurrencyCell(worksheet, currentRow, 9, creditFC, true, true);   // Cột I: Có ngoại tệ
            this._setCurrencyCell(worksheet, currentRow, 10, creditVND, true, false);  // Cột J: Có VND

            this._applyRowBorders(worksheet, currentRow, 14);
            currentRow++;

            // Dòng "Số dư cuối kỳ"
            const closingDebitFC = Number(record.ClosingDebitTran) || 0;
            const closingDebitVND = Number(record.ClosingDebit) || 0;
            const closingCreditFC = Number(record.ClosingCreditTran) || 0;
            const closingCreditVND = Number(record.ClosingCredit) || 0;

            worksheet.getCell(currentRow, 5).value = {
                richText: [
                    { text: "Số dư cuối kỳ /", font: { bold: true, name: CONSTANTS.FONT_NAME } },
                    { text: " Closing balance", font: { bold: true, name: CONSTANTS.FONT_NAME, color: { argb: "FF0070C0" } } }
                ]
            };

            this._setCurrencyCell(worksheet, currentRow, 11, closingDebitFC, true, true);   // Cột K: Nợ ngoại tệ
            this._setCurrencyCell(worksheet, currentRow, 12, closingDebitVND, true, false);  // Cột L: Nợ VND
            this._setCurrencyCell(worksheet, currentRow, 13, closingCreditFC, true, true);   // Cột M: Có ngoại tệ
            this._setCurrencyCell(worksheet, currentRow, 14, closingCreditVND, true, false);  // Cột N: Có VND

            this._applyRowBorders(worksheet, currentRow, 14, CONSTANTS.BORDER_STYLES.dotted, CONSTANTS.BORDER_STYLES.thin);
            return currentRow + 1;
        },

        _parseLineItems: function (lineItemsJson) {
            try {
                if (!lineItemsJson) return [];
                const parsed = JSON.parse(lineItemsJson);
                const items = parsed.LINE_ITEMS || parsed;
                return Array.isArray(items) ? items : (items ? Object.values(items) : []);
            } catch (error) {
                console.error('Failed to parse line items:', error);
                return [];
            }
        },

        _normalizeItemFields: function (item) {
            return Object.fromEntries(Object.entries(item).map(([k, v]) => [k.toLowerCase(), v]));
        },

        _formatctAcount: function (ctAcount) {
            return ctAcount ? ctAcount.replace(/^0+/, '') || '0' : '';
        },

        _populateVNDLineItem: function (worksheet, normalizedItem, currentRow) {
            worksheet.getCell(currentRow, 1).value = this._formatDate(normalizedItem.posting_date);
            worksheet.getCell(currentRow, 1).alignment = { horizontal: 'center' };
            worksheet.getCell(currentRow, 2).value = this._formatDate(normalizedItem.document_date);
            worksheet.getCell(currentRow, 2).alignment = { horizontal: 'center' };
            worksheet.getCell(currentRow, 3).value = normalizedItem.document_number || '';
            worksheet.getCell(currentRow, 3).alignment = { horizontal: 'center' };
            // CỘT MỚI: Số hóa đơn
            worksheet.getCell(currentRow, 4).value = normalizedItem.accountingdocumentheadertext || '';
            worksheet.getCell(currentRow, 4).alignment = { horizontal: 'center' };
            worksheet.getCell(currentRow, 5).value = normalizedItem.item_text || '';
            worksheet.getCell(currentRow, 5).alignment = { horizontal: 'left' };
            worksheet.getCell(currentRow, 6).value = this._formatctAcount(normalizedItem.contra_account);
            worksheet.getCell(currentRow, 6).alignment = { horizontal: 'center' };

            this._setCurrencyCell(worksheet, currentRow, 7, normalizedItem.debit_amount_tran || 0);
            this._setCurrencyCell(worksheet, currentRow, 8, normalizedItem.debit_amount || 0);
            this._setCurrencyCell(worksheet, currentRow, 9, normalizedItem.credit_amount_tran || 0);
            this._setCurrencyCell(worksheet, currentRow, 10, normalizedItem.credit_amount || 0);

            this._setCurrencyCell(worksheet, currentRow, 11, normalizedItem.closingdebit_tran || 0);
            this._setCurrencyCell(worksheet, currentRow, 12, normalizedItem.closingdebit || 0);
            this._setCurrencyCell(worksheet, currentRow, 13, normalizedItem.closingcredit_tran || 0);
            this._setCurrencyCell(worksheet, currentRow, 14, normalizedItem.closingcredit || 0);
        },

        _setCurrencyCell: function (worksheet, row, col, value, isBold = false) {
            const cell = worksheet.getCell(row, col);
            const numValue = Number(value) || 0;
            cell.value = numValue !== 0 ? numValue : '';
            if (numValue !== 0) {
                cell.numFmt = (col % 2 === 1) ? '#,##0.00' : '#,##0'; // Ngoại tệ (lẻ): có thập phân, VND (chẵn): không
            }
            cell.alignment = { horizontal: 'right' };
            if (isBold) cell.font = { bold: true };
        },

        _applyRowBorders: function (worksheet, row, maxCol, topStyle = CONSTANTS.BORDER_STYLES.dotted, bottomStyle = CONSTANTS.BORDER_STYLES.dotted) {
            for (let col = 1; col <= maxCol; col++) {
                const cell = worksheet.getCell(row, col);
                cell.border = {
                    top: { style: topStyle },
                    bottom: { style: bottomStyle },
                    left: { style: CONSTANTS.BORDER_STYLES.thin },
                    right: { style: CONSTANTS.BORDER_STYLES.thin }
                };
            }
        },

        _getAllBorders: function (style) {
            return { top: { style }, bottom: { style }, left: { style }, right: { style } };
        },

        _setVNDColumnWidths: function (worksheet) {
            Object.values(CONSTANTS.VND_COLUMNS).forEach(col => {
                worksheet.getColumn(col.index).width = col.width;
            });
        },

        _createVNDFooter: function (worksheet) {
            let lastRow = worksheet.lastRow ? worksheet.lastRow.number + 3 : 30;

            // Người lập biểu / Preparer
            worksheet.getCell(`C${lastRow}`).value = {
                richText: [
                    { text: "Người lập biểu\n", font: { size: 11, bold: true } },
                    { text: "Preparer", font: { size: 11, bold: true, color: { argb: "FF0070C0" } } }
                ]
            };
            worksheet.getCell(`C${lastRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

            // Ngày cố định 26/12/2025
            worksheet.getCell(`K${lastRow}`).value = {
                richText: [
                    { text: "Ngày 26 tháng 12 năm 2025\n", font: { size: 11 } },
                    { text: "Date 26 month 12 year 2025", font: { size: 11, color: { argb: "FF0070C0" } } }
                ]
            };
            worksheet.getCell(`K${lastRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

            // Kế toán trưởng / Chief accountant
            worksheet.getCell(`K${lastRow + 1}`).value = {
                richText: [
                    { text: "Kế toán trưởng\n", font: { size: 11, bold: true } },
                    { text: "Chief accountant", font: { size: 11, bold: true, color: { argb: "FF0070C0" } } }
                ]
            };
            worksheet.getCell(`K${lastRow + 1}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        },

        _applyFontToWorksheet: function (worksheet) {
            worksheet.eachRow({ includeEmpty: true }, row => {
                row.eachCell({ includeEmpty: true }, cell => {
                    // Font chính của cell
                    const baseFont = cell.font || {};
                    cell.font = {
                        name: CONSTANTS.FONT_NAME,
                        size: baseFont.size || 11,
                        bold: baseFont.bold || false,
                        color: baseFont.color,
                        italic: baseFont.italic
                    };

                    // Nếu dùng richText → sửa từng phần bên trong
                    if (cell.value && cell.value.richText) {
                        cell.value.richText.forEach(part => {
                            const partBase = part.font || {};
                            part.font = {
                                name: CONSTANTS.FONT_NAME,
                                size: partBase.size || 11,
                                bold: partBase.bold || false,
                                color: partBase.color,
                                italic: partBase.italic
                            };
                        });
                    }
                });
            });
        },

        _formatDate: function (dateValue) {
            if (!dateValue) return '';
            if (typeof dateValue === 'string' && dateValue.length === 8) {
                return `${dateValue.substring(6, 8)}/${dateValue.substring(4, 6)}/${dateValue.substring(0, 4)}`;
            }
            const date = new Date(dateValue);
            if (isNaN(date.getTime())) return dateValue;
            return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
        },

        _downloadWorkbook: async function (workbook, baseFilename) {
            const buffer = await workbook.xlsx.writeBuffer();
            const filename = `Export_${baseFilename}_20251226.xlsx`;
            const blob = new Blob([buffer], { type: CONSTANTS.MIME_TYPE });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
            window.URL.revokeObjectURL(url);
        },

        _handleError: function (message, error) {
            console.error(message, error);
            MessageToast.show(`Error: ${message}`);
        }
    };
});
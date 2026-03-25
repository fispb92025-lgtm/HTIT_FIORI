sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    'use strict';

    const CONSTANTS = {
        SMART_TABLE_ID: "zaccpaydetail::sap.suite.ui.generic.template.ListReport.view.ListReport::ZC_ACCPAY_DETAIL--listReport",
        PAPER_SIZE: 9,
        MARGINS: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.1, footer: 0.1 },
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
        BORDER_STYLES: { thin: 'medium', dotted: 'dotted' },
        VND_COLUMNS: {
            DATE: { index: 1, width: 15 },
            DOC_DATE: { index: 2, width: 15 },
            DOC_NUMBER: { index: 3, width: 17 },
            DESCRIPTION: { index: 4, width: 35 },
            OFFSET_ACCOUNT: { index: 5, width: 20 },
            DEBIT_FC: { index: 6, width: 22 },
            DEBIT_VND: { index: 7, width: 22 },
            CREDIT_FC: { index: 8, width: 22 },
            CREDIT_VND: { index: 9, width: 22 },
            BAL_DEBIT_FC: { index: 10, width: 22 },
            BAL_DEBIT_VND: { index: 11, width: 22 },
            BAL_CREDIT_FC: { index: 12, width: 22 },
            BAL_CREDIT_VND: { index: 13, width: 22 }
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
                await this._downloadWorkbook(workbook, 'VND_So_chi_tiet_cong_no_phai_tra');
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

            dataArray.push({});

            const dataArrayNew = [];
            let previousRow = {};

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
                    }  else if (previousRow.OpeningCreditBalance - previousRow.OpeningDebitBalance == 0) {
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
                    headers: { 'Content-Type': 'application/json', 'Cookie': 'sap-usercontext=sap-client=100' }
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

        _createWorkbook: function () { return new ExcelJS.Workbook(); },

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
            const values = [CONSTANTS.DEFAULT_COMPANY.name, CONSTANTS.DEFAULT_COMPANY.nameUS,
            CONSTANTS.DEFAULT_COMPANY.addressVN, CONSTANTS.DEFAULT_COMPANY.addressUS];
            const fonts = [{ size: 12, bold: true },
            { size: 12, bold: true, color: { argb: "FF0070C0" } },
            { size: 12, bold: true },
            { size: 12, bold: true, color: { argb: "FF0070C0" } }];

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
            worksheet.getCell("A7").value = "SỔ CHI TIẾT CÔNG NỢ PHẢI TRẢ";
            worksheet.getCell("A7").font = { bold: true, size: 14 };
            worksheet.getCell("A7").alignment = { horizontal: "center" };

            worksheet.mergeCells("A8:M8");
            worksheet.getCell("A8").value = "DETAILED ACCOUNTS PAYABLE LEDGER";
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
                { range: `A${headerRow}:A${headerRow + 2}`, value: { richText: [{ text: "Ngày ghi sổ\n", font: {} }, { text: "Posting date", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `B${headerRow}:C${headerRow}`, value: { richText: [{ text: "Chứng từ\n", font: {} }, { text: "Voucher", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `B${headerRow + 1}:B${headerRow + 2}`, value: { richText: [{ text: "Ngày / ", font: {} }, { text: "Date", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `C${headerRow + 1}:C${headerRow + 2}`, value: { richText: [{ text: "Số / ", font: {} }, { text: "Number", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `D${headerRow}:D${headerRow + 2}`, value: { richText: [{ text: "Diễn giải\n", font: {} }, { text: "Description", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `E${headerRow}:E${headerRow + 2}`, value: { richText: [{ text: "TK đối ứng\n", font: {} }, { text: "Contra Accounts", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `F${headerRow}:I${headerRow}`, value: { richText: [{ text: "Số tiền / ", font: {} }, { text: "Amount", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `F${headerRow + 1}:G${headerRow + 1}`, value: { richText: [{ text: "Nợ / ", font: {} }, { text: "Debit", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `H${headerRow + 1}:I${headerRow + 1}`, value: { richText: [{ text: "Có / ", font: {} }, { text: "Credit", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `F${headerRow + 2}:F${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `G${headerRow + 2}:G${headerRow + 2}`, value: 'VND' },
                { range: `H${headerRow + 2}:H${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `I${headerRow + 2}:I${headerRow + 2}`, value: 'VND' },
                { range: `J${headerRow}:M${headerRow}`, value: { richText: [{ text: "Số dư / ", font: {} }, { text: "Balance", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `J${headerRow + 1}:K${headerRow + 1}`, value: { richText: [{ text: "Nợ / ", font: {} }, { text: "Debit", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `L${headerRow + 1}:M${headerRow + 1}`, value: { richText: [{ text: "Có / ", font: {} }, { text: "Credit", font: { color: { argb: "FF0070C0" } } }] } },
                { range: `J${headerRow + 2}:J${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `K${headerRow + 2}:K${headerRow + 2}`, value: 'VND' },
                { range: `L${headerRow + 2}:L${headerRow + 2}`, value: 'Ngoại tệ' },
                { range: `M${headerRow + 2}:M${headerRow + 2}`, value: 'VND' }
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
            ['A', 'B', 'C', 'D', 'E'].forEach((v, i) => {
                const cell = worksheet.getCell(signRow, i + 1);
                cell.value = v;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = this._getAllBorders(CONSTANTS.BORDER_STYLES.thin);
            });

            for (let i = 1; i <= 8; i++) {
                const cell = worksheet.getCell(signRow, 5 + i);
                cell.value = i;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = this._getAllBorders(CONSTANTS.BORDER_STYLES.thin);
            }
        },

        _addVNDOpeningBalance: function (worksheet, record, currentRow) {
            const openingDebitFC = Number(record.OpeningDebitBalanceTran || 0);
            const openingCreditFC = Number(record.OpeningCreditBalanceTran || 0);
            const openingDebitVND = Number(record.OpeningDebitBalance || 0);
            const openingCreditVND = Number(record.OpeningCreditBalance || 0);

            const balanceFC = openingCreditFC - openingDebitFC;
            const balanceVND = openingCreditVND - openingDebitVND;

            worksheet.getCell(currentRow, 4).value = {
                richText: [
                    { text: "Số dư đầu kỳ /", font: { bold: true, name: CONSTANTS.FONT_NAME } },
                    { text: " Opening balance", font: { bold: true, name: CONSTANTS.FONT_NAME, color: { argb: "FF0070C0" } } }
                ]
            };

            if (balanceVND > 0) {
                if (balanceFC > 0) this._setCurrencyCell(worksheet, currentRow, 12, balanceFC, true, true);
                this._setCurrencyCell(worksheet, currentRow, 13, balanceVND, true, false);
            } else if (balanceVND < 0) {
                if (balanceFC < 0) this._setCurrencyCell(worksheet, currentRow, 10, Math.abs(balanceFC), true, true);
                this._setCurrencyCell(worksheet, currentRow, 11, Math.abs(balanceVND), true, false);
            }

            this._applyRowBorders(worksheet, currentRow, 13, CONSTANTS.BORDER_STYLES.thin, CONSTANTS.BORDER_STYLES.dotted);
            return currentRow + 1;
        },

        _addVNDLineItems: function (worksheet, record, currentRow) {
            const lineItems = this._parseLineItems(record.LineItemsJson);
            lineItems.forEach(item => {
                const norm = this._normalizeItemFields(item);
                this._populateVNDLineItem(worksheet, norm, currentRow);
                this._applyRowBorders(worksheet, currentRow, 13, CONSTANTS.BORDER_STYLES.dotted, CONSTANTS.BORDER_STYLES.dotted);
                currentRow++;
            });
            return currentRow;
        },

        _addVNDSummaryRows: function (worksheet, record, currentRow) {
            const debitFC = Number(record.DebitAmountDuringPeriodTran) || 0;
            const debitVND = Number(record.DebitAmountDuringPeriod) || 0;
            const creditFC = Number(record.CreditAmountDuringPeriodTran) || 0;
            const creditVND = Number(record.CreditAmountDuringPeriod) || 0;

            worksheet.getCell(currentRow, 4).value = {
                richText: [
                    { text: "Số phát sinh trong kỳ /", font: { bold: true, name: CONSTANTS.FONT_NAME } },
                    { text: " Arising amount", font: { bold: true, name: CONSTANTS.FONT_NAME, color: { argb: "FF0070C0" } } }
                ]
            };

            this._setCurrencyCell(worksheet, currentRow, 6, debitFC, true, true);
            this._setCurrencyCell(worksheet, currentRow, 7, debitVND, true, false);
            this._setCurrencyCell(worksheet, currentRow, 8, creditFC, true, true);
            this._setCurrencyCell(worksheet, currentRow, 9, creditVND, true, false);

            this._applyRowBorders(worksheet, currentRow, 13);
            currentRow++;

            const closingDebitFC = Number(record.ClosingDebitTran) || 0;
            const closingDebitVND = Number(record.ClosingDebit) || 0;
            const closingCreditFC = Number(record.ClosingCreditTran) || 0;
            const closingCreditVND = Number(record.ClosingCredit) || 0;

            worksheet.getCell(currentRow, 4).value = {
                richText: [
                    { text: "Số dư cuối kỳ /", font: { bold: true, name: CONSTANTS.FONT_NAME } },
                    { text: " Closing balance", font: { bold: true, name: CONSTANTS.FONT_NAME, color: { argb: "FF0070C0" } } }
                ]
            };

            this._setCurrencyCell(worksheet, currentRow, 10, closingDebitFC, true, true);
            this._setCurrencyCell(worksheet, currentRow, 11, closingDebitVND, true, false);
            this._setCurrencyCell(worksheet, currentRow, 12, closingCreditFC, true, true);
            this._setCurrencyCell(worksheet, currentRow, 13, closingCreditVND, true, false);

            this._applyRowBorders(worksheet, currentRow, 13, CONSTANTS.BORDER_STYLES.dotted, CONSTANTS.BORDER_STYLES.thin);
            return currentRow + 1;
        },

        _parseLineItems: function (json) {
            if (!json) return [];
            try {
                const parsed = JSON.parse(json);
                const items = parsed.LINE_ITEMS || parsed;
                return Array.isArray(items) ? items : Object.values(items);
            } catch {
                return [];
            }
        },

        _normalizeItemFields: function (item) {
            return Object.fromEntries(Object.entries(item).map(([k, v]) => [k.toLowerCase(), v]));
        },

        _formatctAcount: function (acc) {
            return acc ? acc.replace(/^0+/, '') || '0' : '';
        },

        _populateVNDLineItem: function (worksheet, item, row) {
            worksheet.getCell(row, 1).value = this._formatDate(item.posting_date);
            worksheet.getCell(row, 1).alignment = { horizontal: 'center' };
            worksheet.getCell(row, 2).value = this._formatDate(item.document_date);
            worksheet.getCell(row, 2).alignment = { horizontal: 'center' };
            worksheet.getCell(row, 3).value = item.document_number || '';
            worksheet.getCell(row, 3).alignment = { horizontal: 'center' };
            worksheet.getCell(row, 4).value = item.item_text || '';
            worksheet.getCell(row, 4).alignment = { horizontal: 'left' };
            worksheet.getCell(row, 5).value = this._formatctAcount(item.contra_account);
            worksheet.getCell(row, 5).alignment = { horizontal: 'center' };

            this._setCurrencyCell(worksheet, row, 6, item.debit_amount_tran || 0);
            this._setCurrencyCell(worksheet, row, 7, item.debit_amount || 0);
            this._setCurrencyCell(worksheet, row, 8, item.credit_amount_tran || 0);
            this._setCurrencyCell(worksheet, row, 9, item.credit_amount || 0);
            this._setCurrencyCell(worksheet, row, 10, item.closingdebit_tran || item.ClosingDebitTran || 0);
            this._setCurrencyCell(worksheet, row, 11, item.closingdebit || item.ClosingDebit || 0);
            this._setCurrencyCell(worksheet, row, 12, item.closingcredit_tran || item.ClosingCreditTran || 0);
            this._setCurrencyCell(worksheet, row, 13, item.closingcredit || item.ClosingCredit || 0);
        },

        _setCurrencyCell: function (worksheet, row, col, value, bold = false) {
            const cell = worksheet.getCell(row, col);
            const num = Number(value) || 0;
            cell.value = num !== 0 ? num : '';
            if (num !== 0) cell.numFmt = (col % 2 === 0) ? '#,##0.00' : '#,##0';
            cell.alignment = { horizontal: 'right' };
            if (bold) cell.font = { bold: true };
        },

        _applyRowBorders: function (worksheet, row, maxCol, top = CONSTANTS.BORDER_STYLES.dotted, bottom = CONSTANTS.BORDER_STYLES.dotted) {
            for (let c = 1; c <= maxCol; c++) {
                worksheet.getCell(row, c).border = {
                    top: { style: top },
                    bottom: { style: bottom },
                    left: { style: CONSTANTS.BORDER_STYLES.thin },
                    right: { style: CONSTANTS.BORDER_STYLES.thin }
                };
            }
        },

        _getAllBorders: function (style) {
            return { top: { style }, bottom: { style }, left: { style }, right: { style } };
        },

        _setVNDColumnWidths: function (worksheet) {
            Object.values(CONSTANTS.VND_COLUMNS).forEach(c => worksheet.getColumn(c.index).width = c.width);
        },

        _createVNDFooter: function (worksheet) {
            let lastRow = worksheet.lastRow ? worksheet.lastRow.number + 3 : 30;
            const today = new Date();
            const d = String(today.getDate()).padStart(2, '0');
            const m = String(today.getMonth() + 1).padStart(2, '0');
            const y = today.getFullYear();

            worksheet.getCell(`C${lastRow}`).value = {
                richText: [
                    { text: "Người lập biểu\n", font: { size: 11, bold: true } },
                    { text: "Preparer", font: { size: 11, bold: true, color: { argb: "FF0070C0" } } }
                ]
            };
            worksheet.getCell(`C${lastRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

            worksheet.getCell(`K${lastRow}`).value = {
                richText: [
                    { text: `Ngày ${d} tháng ${m} năm ${y}\n`, font: { size: 11 } },
                    { text: `Date ${d} month ${m} year ${y}`, font: { size: 11, color: { argb: "FF0070C0" } } }
                ]
            };
            worksheet.getCell(`K${lastRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

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
                    const baseFont = cell.font || {};
                    cell.font = {
                        name: CONSTANTS.FONT_NAME,
                        size: baseFont.size || 11,
                        bold: baseFont.bold || false,
                        color: baseFont.color,
                        italic: baseFont.italic
                    };

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

        _formatDate: function (v) {
            if (!v) return '';
            if (typeof v === 'string' && v.length === 8) return `${v.substring(6, 8)}/${v.substring(4, 6)}/${v.substring(0, 4)}`;
            const d = new Date(v);
            return isNaN(d.getTime()) ? v : `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
        },

        _downloadWorkbook: async function (workbook, base) {
            const buffer = await workbook.xlsx.writeBuffer();
            const filename = `Export_${base}_${new Date().getFullYear()}${(new Date().getMonth() + 1).toString().padStart(2, '0')}${new Date().getDate().toString().padStart(2, '0')}.xlsx`;
            const blob = new Blob([buffer], { type: CONSTANTS.MIME_TYPE });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
            URL.revokeObjectURL(url);
        },

        _handleError: function (msg, err) {
            console.error(msg, err);
            MessageToast.show(`Error: ${msg}`);
        }
    };
});
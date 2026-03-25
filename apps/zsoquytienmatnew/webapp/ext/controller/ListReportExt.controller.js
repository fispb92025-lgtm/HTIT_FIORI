sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/odata/v2/ODataModel",
    "sap/m/MessageToast",
    "exceljs",
    "sap/ui/core/Fragment",
    "sap/ui/model/json/JSONModel",
    "zsoquytienmat/template/soquytienmat_trancurr.base64",
    "zsoquytienmat/template/soquytienmat_cocdcurr.base64",
    "zsoquytienmat/template/soquytienmat_trancurr_new.base64",
    "zsoquytienmat/template/soquytienmat_cocdcurr_new.base64"
], function (Controller, ODataModel, MessageToast, ExcelJS, Fragment, JSONModel, Base64TranCurr, Base64LocalCurr, Base64TranCurrNew, Base64LocalCurrNew) {
    "use strict";

    return {
        onInit: function () { },

        ExportToExcel: async function (oEvent) {
            const view = this.getView();

            if (!this._signerData) {
                this._signerData = new JSONModel({ bookkeeper: "", accountant: "", director: "" });
            }

            const signerModel = this._signerData;

            if (!this._signerDialog) {
                this._signerDialog = await Fragment.load({
                    name: "zsoquytienmat.ext.fragments.SignerDialog",
                    controller: this
                });

                view.addDependent(this._signerDialog);
                this._signerDialog.setModel(signerModel, "signer");
            }

            this._signerDialog.open();
        },

        onSignerDialogConfirm: function () {
            const signerData = this._signerData.getData();

            this._signerDialog.close();
            this._executeExportExcel(signerData); // Gọi hàm export chính
        },

        onSignerDialogCancel: function () {
            this._signerDialog.close();
        },

        // Đọc toàn bộ collection OData V2 (UI5 v2 ODataModel)
        readAllPages: function (oModel, sPath, mParams = {}) {
            return new Promise((resolve, reject) => {
                const all = [];
                const baseUrlParams = Object.assign({}, mParams.urlParameters); // giữ các tham số gốc (nếu có)

                // tách tham số từ __next (chỉ giữ $skip/$skiptoken)
                const extractNextParams = (sNext) => {
                    try {
                        const serviceUrl = (oModel.sServiceUrl || "").replace(/\/$/, "");
                        if (sNext.startsWith(serviceUrl)) sNext = sNext.slice(serviceUrl.length); // bỏ prefix
                        const qIndex = sNext.indexOf("?");
                        if (qIndex < 0) return {};
                        const qs = sNext.slice(qIndex + 1);
                        const params = {};
                        qs.split("&").forEach(p => {
                            if (!p) return;
                            const [k, v] = p.split("=");
                            const key = decodeURIComponent(k || "");
                            const val = decodeURIComponent(v || "");
                            if (key === "$skiptoken" || key === "$skip") params[key] = val;
                        });
                        return params;
                    } catch {
                        return {};
                    }
                };

                const readPage = (nextParams = {}) => {
                    oModel.read(sPath, {
                        ...mParams,
                        urlParameters: { ...baseUrlParams, ...nextParams },  // <== truyền token vào đây
                        success: (oData /*, oResponse */) => {
                            const rows = Array.isArray(oData?.results) ? oData.results : (oData ? [oData] : []);
                            all.push(...rows);

                            const sNext = oData && oData.__next;              // OData V2: __next
                            if (sNext) {
                                const np = extractNextParams(sNext);
                                readPage(np);                                    // đọc trang kế tiếp
                            } else {
                                resolve(all);                                    // hết trang
                            }
                        },
                        error: reject
                    });
                };

                // trang đầu
                readPage();
            });
        },

        _getBusy: function () {
            if (!this._busy) {
                this._busy = new sap.m.BusyDialog({
                    text: "Đang xử lý...",
                    title: "Processing"
                });
            }
            return this._busy;
        },

        _executeExportExcel: async function (signerData) {
            /*------------------------------------------------------------------------------------*/
            const view = this.getView();
            const smartFilterBar = view.byId("zsoquytienmat::sap.suite.ui.generic.template.ListReport.view.ListReport::SoQuyTienMat--listReportFilter");
            const filterData = smartFilterBar?.getFilterData() || {};

            /*------------------------------------------------------------------------------------*/
            const sfb = view.byId("zsoquytienmat::sap.suite.ui.generic.template.ListReport.view.ListReport::SoQuyTienMat--listReportFilter");

            const aFilters = sfb.getFilters(); // <-- đã có PostingDate (BT) sẵn
            const oModel = view.getModel();


            // oModel.read("/SoQuyTienMat", {
            //     filters: aFilters,
            //     // urlParameters: { "$select": "CompanyCode,AccountingDocument,FiscalYear,PostingDate" },
            //     urlParameters: { "$expand": "to_Items" },   // ví dụ nếu có nav
            //     success: (oData/*, oResponse*/) => {
            //         // với EntitySet: oData.results là mảng (có thể [] nếu không có dữ liệu)
            //         const rows = Array.isArray(oData?.results) ? oData.results : (oData ? [oData] : []);
            //         console.log("rows:", rows);
            //     },
            //     error: (e) => console.error(e)
            // });

            // try {
            // const dataRaw = await this.readAllPages(oModel, "/SoQuyTienMat", {
            //     filters: aFilters,
            //     // urlParameters: { "$inlinecount": "allpages" } // nếu muốn đếm tổng số bản ghi
            // });
            // console.log("Tổng dòng:", dataRaw.length, dataRaw);
            // // ví dụ: đưa vào JSONModel phụ
            // this.getView().setModel(new sap.ui.model.json.JSONModel(dataRaw), "allData");


            // await this.exportSoQuyTienMatGrouped(dataRaw, signerData);

            // } catch (e) {
            //     console.error("Đọc toàn bộ trang lỗi:", e);
            // }

            const busy = this._getBusy();
            busy.setText("Processing...");
            busy.open();
            try {
                const dataRaw = await this.readAllPages(oModel, "/SoQuyTienMat", { filters: aFilters });

                // busy.setText("Đang tạo file Excel...");
                await this.exportSoQuyTienMatGrouped(dataRaw, signerData);

                sap.m.MessageToast.show("Export thành công!");
            } catch (e) {
                sap.m.MessageBox.error("Có lỗi khi xử lý: " + (e.message || e));
            } finally {
                busy.close();
            }
        },

        exportSoQuyTienMatGrouped: async function (dataRaw, signerData) {
            // Định dạng tiền tệ
            const MONEY_FMT = '#,##0;(#,##0)';

            // >>> LẤY CỜ TỪ DÒNG ĐẦU
            let flagDifferentCurrency = this.detectDifferentCurrency(dataRaw);
            const firstRow = Array.isArray(dataRaw)
                ? dataRaw[0]
                : (dataRaw?.results?.[0] || dataRaw?.items?.[0] || null);

            if (firstRow && firstRow.GLAccount === "11120100") {
                flagDifferentCurrency = "X";
            }

            // 0) Chuẩn hoá list từ dataRaw (array | {results} | {items})
            const toList = Array.isArray(dataRaw)
                ? dataRaw
                : (dataRaw?.results || dataRaw?.items || []);

            // 1) Map record OData -> item dùng để in
            const N = v => Number(v || 0);
            const fmtDate = (val) => {
                if (!val) return "";
                const d = (val instanceof Date) ? val : new Date(val);
                if (isNaN(d)) return "";
                const dd = String(d.getDate()).padStart(2, "0");
                const mm = String(d.getMonth() + 1).padStart(2, "0");
                const yy = d.getFullYear();
                return `${dd}/${mm}/${yy}`;
            };

            const mapped = toList.map((r, i) => {
                console.log(`Row ${i}: businesspartner = ${r.businesspartner}, Doituong = ${r.Doituong}`); // Debug
                return {
                    CompanyCode: r.CompanyCode || r.Companycode || r.COMPANYCODE || "",
                    TransactionCurrency: r.TransactionCurrency || r.Currency || "",
                    CompanyCodeCurrency: r.CompanyCodeCurrency || "",
                    stt: i + 1,
                    PostingDate: fmtDate(r.PostingDate || r.Postingdate),
                    DocumentDate: fmtDate(r.DocumentDate || r.Documentdate),
                    AccountingDocument: r.AccountingDocument || r.Accountingdocument || r.BELNR || "",
                    objectCode: r.businesspartner || "", // Mã đối tượng (cột E)
                    objectName: r.Doituong || "", // Tên đối tượng (cột F)
                    SoHieuCTThu: r.SoHieuCTThu || "",
                    SoHieuCTChi: r.SoHieuCTChi || "",
                    Diengiai: r.Diengiai || "", // Diễn giải (cột I)
                    DebitAmountInCoCode: N(r.DebitAmountInCoCode || r.Debitamountincocode),
                    CreditAmountInCoCode: N(r.CreditAmountInCoCode || r.Creditamountincocode),
                    BalanceInCoCode: N(r.BalanceInCoCode || r.Balanceincocode),
                    DebitAmountInTrans: N(r.DebitAmountInTrans || r.Debitamountintrans),
                    CreditAmountInTrans: N(r.CreditAmountInTrans || r.Creditamountintrans),
                    BalanceInTrans: N(r.BalanceInTrans || r.Balanceintrans),
                    StartingBalanceInCoCode: N(r.StartingBalanceInCoCode || r.Startingbalanceincocode),
                    StartingBalanceInTrans: N(r.StartingBalanceInTrans || r.Startingbalanceintrans),
                    GhiChu: r.GhiChu || r.Ghichu || "",
                    CompanyCodeName: r.CompanyCodeName || "",
                    CompanyCodeAddr: r.CompanyCodeAddr || "",
                    GLAccoutName: r.GLAccoutName || "",
                    PeriodText: r.PeriodText || "",
                    GLAccount: r.GLAccount || "",
                };
            });

            // 2) Group theo CompanyCode
            const groups = mapped.reduce((m, it) => {
                (m[it.CompanyCode] ||= []).push(it);
                return m;
            }, {});

            const view = this.getView();

            const smartFilterBar = view.byId("listReportFilter")
                || view.byId("smartFilterBar")
                || view.byId("template::SmartFilterBar");

            const aFilters = smartFilterBar.getFilters();

            let oPostingDateFilter;

            aFilters.forEach(oFilter => {
                if (oFilter.aFilters) {
                    oFilter.aFilters.forEach(oInner => {
                        try {
                            if (oInner.aFilters[0].sPath === "PostingDate") {
                                oPostingDateFilter = oInner.aFilters[0];
                            }
                        } catch (e) {

                        }
                    });
                }
            });

            // const oFromDate = oPostingDateFilter.oValue1.oDate;
            // const oToDate = oPostingDateFilter.oValue2.oDate;

            const oFromDate = oPostingDateFilter?.oValue1 || null;
            const oToDate = oPostingDateFilter?.oValue2 || null;

            const oFinalDate = oToDate || oFromDate;

            if (!oFinalDate) {
                sap.m.MessageToast.show("Vui lòng chọn Posting Date");
                return;
            }

            const lastPostYear = oFinalDate.getFullYear();

            let TemplateBase64

            if (lastPostYear < 2026)
                TemplateBase64 =
                    (flagDifferentCurrency || '').toUpperCase() === 'X'
                        ? Base64TranCurr
                        : Base64LocalCurr;
            else {
                TemplateBase64 =
                    (flagDifferentCurrency || '').toUpperCase() === 'X'
                        ? Base64TranCurrNew
                        : Base64LocalCurrNew;
            }

            // 3) Load workbook & kiểm tra template
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(this.base64ToUint8Array(TemplateBase64));
            const template = wb.getWorksheet(1);
            console.log("Template headers:", template.getRow(12).values); // Debug: In header row
            console.log("H12 value from template:", template.getCell("H12").value); // Debug: Kiểm tra giá trị H12

            const pristine = wb.addWorksheet("_PRISTINE_");
            this.cloneSheet(template, pristine);
            wb.removeWorksheet(template.id);

            const existingNames = new Set(wb.worksheets.map(s => s.name));
            existingNames.add("_PRISTINE_");

            // Các tham số vị trí dành cho sổ quỹ
            const START_ROW = 13;
            const numCols = 12; // 12 cột: A-L

            for (const [ccode, rows] of Object.entries(groups)) {
                const ws = wb.addWorksheet(this.safeSheetName(ccode || "NO_CC", existingNames));
                this.cloneSheet(pristine, ws);

                // Áp dụng định dạng tiền tệ cho cột J, K, L (cột 10, 11, 12 - Thu, Chi, Tồn)
                [10, 11, 12].forEach((colIdx) => ws.getColumn(colIdx).numFmt = MONEY_FMT);

                // ======= HEADER: Period + Company info =======
                const first = rows[0] || {};
                const last = rows[rows.length - 1] || {};
                const openingCoCode = Number(first.StartingBalanceInCoCode ?? first.BalanceInCoCode ?? 0);
                const openingTrans = Number(first.StartingBalanceInTrans ?? first.BalanceInTrans ?? 0);

                const Companycodename = first.CompanyCodeName || "";
                const Companycodeaddr = first.CompanyCodeAddr || "";
                const GLAcctname = first.GLAccount || "";
                const PeriodText = first.PeriodText || "";

                let endingCoCode;
                let endingTrans;

                if (flagDifferentCurrency === "X") {
                    const totalDebitTr = rows.reduce((s, r) => s + (r.DebitAmountInTrans || 0), 0);
                    const totalCreditTr = rows.reduce((s, r) => s + (r.CreditAmountInTrans || 0), 0);
                    endingTrans = totalDebitTr - totalCreditTr;
                    endingCoCode = Number(last.BalanceInCoCode ?? 0);
                } else {
                    endingCoCode = Number(last.BalanceInCoCode ?? 0);
                    endingTrans = Number(last.BalanceInTrans ?? 0);
                }

                let currency = first.TransactionCurrency || first.CompanyCodeCurrency || "";
                if (first.GLAccount === "11120100") {
                    currency = "USD";
                }

                // Cập nhật header
                ws.getCell("A1").value = {
                    richText: [
                        { text: "CÔNG TY TNHH CẢNG QUỐC TẾ TIL CẢNG HẢI PHÒNG" + "\n", font: { name: "Times New Roman", size: 14, bold: true, color: { argb: "000000" }, bold: true } },
                        { text: "HAIPHONG PORT TIL INTERNATIONAL TERMINAL COMPANY LIMITED", font: { name: "Times New Roman", size: 12, bold: true, color: { argb: "0070C0" } } }
                    ]
                };
                ws.getCell("A1").alignment = { horizontal: "left", vertical: "middle", wrapText: true };

                ws.getCell("A2").value = {
                    richText: [
                        { text: "Bến số 3&4 Cảng nước sâu Lạch Huyện, Khu phố Đôn Lương, Đặc khu Cát Hải, Thành phố Hải Phòng, Việt Nam" + "\n", font: { name: "Times New Roman", size: 14, bold: true, color: { argb: "000000" } } },
                        { text: "Berth No. 3&4 Lach Huyen Deep-sea Port, Don Luong Quarter, Cat Hai Special Zone, Hai Phong City, Vietnam", font: { name: "Times New Roman", size: 12, bold: true, color: { argb: "0070C0" } } }
                    ]
                };
                ws.getCell("A2").alignment = { horizontal: "left", vertical: "middle", wrapText: true };

                // Di chuyển "Sổ quỹ tiền mặt" sang cột H và merge G5:H5
                this.safeMerge(ws, "G5:H5");
                ws.getCell("H5").value = {
                    richText: [
                        { text: `SỔ QUỸ TIỀN MẶT: ${GLAcctname}\n`, font: { name: "Times New Roman", size: 14, bold: true, color: { argb: "000000" } } },
                        { text: `CASH BOOK: ${GLAcctname}`, font: { name: "Times New Roman", size: 12, bold: true, color: { argb: "0070C0" } } }
                    ]
                };
                ws.getCell("H5").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

                // Di chuyển "Loại quỹ" sang cột H và merge G6:H6
                this.safeMerge(ws, "G6:H6");
                ws.getCell("H6").value = {
                    richText: [
                        { text: `Loại quỹ: ${currency}\n`, font: { name: "Times New Roman", size: 14, bold: true, color: { argb: "000000" } } },
                        { text: `Currency: ${currency}`, font: { name: "Times New Roman", size: 12, bold: true, color: { argb: "0070C0" } } }
                    ]
                };
                ws.getCell("H6").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

                // ======= CHÈN DATA =======
                const n = rows.length;
                if (n > 1) ws.duplicateRow(START_ROW, n - 1, true);

                let tDebitCo = 0, tCreditCo = 0, tDebitTr = 0, tCreditTr = 0;
                const templateRow = ws.getRow(START_ROW);

                for (let i = 0; i < n; i++) {
                    const it = rows[i];
                    const r = ws.getRow(START_ROW + i);

                    let rowVals;
                    if (flagDifferentCurrency === "X") {
                        rowVals = [
                            it.stt, it.PostingDate, it.DocumentDate, it.AccountingDocument,
                            it.objectCode, // Mã đối tượng (cột E)
                            it.objectName, // Tên đối tượng (cột F)
                            it.SoHieuCTThu || "", // Số hiệu chứng từ Thu (cột G)
                            it.SoHieuCTChi || "", // Số hiệu chứng từ Chi (cột H)
                            it.Diengiai, // Diễn giải (cột I)
                            it.DebitAmountInTrans, // Thu/Receipt (cột J)
                            it.CreditAmountInTrans, // Chi/Payment (cột K)
                            it.BalanceInTrans, // Tồn/Balance (cột L)
                            it.GhiChu // Ghi chú (cột M, dịch sang phải)
                        ];
                    } else {
                        rowVals = [
                            it.stt, it.PostingDate, it.DocumentDate, it.AccountingDocument,
                            it.objectCode, // Mã đối tượng (cột E)
                            it.objectName, // Tên đối tượng (cột F)
                            it.SoHieuCTThu || "", // Số hiệu chứng từ Thu (cột G)
                            it.SoHieuCTChi || "", // Số hiệu chứng từ Chi (cột H)
                            it.Diengiai, // Diễn giải (cột I)
                            it.DebitAmountInCoCode, // Thu/Receipt (cột J)
                            it.CreditAmountInCoCode, // Chi/Payment (cột K)
                            it.BalanceInCoCode, // Tồn/Balance (cột L)
                            it.GhiChu // Ghi chú (cột M, dịch sang phải)
                        ];
                    }

                    console.log(`Row ${i} values:`, rowVals); // Debug
                    for (let j = 0; j < rowVals.length; j++) {
                        const cell = r.getCell(j + 1);
                        cell.value = rowVals[j] ?? "";
                        if (typeof cell.value === "number") {
                            if (flagDifferentCurrency === "X") {
                                // Áp dụng định dạng khác cho cột STT (cột 1)
                                if (j === 0) { // Cột STT
                                    cell.numFmt = '#,##0'; // Không hiển thị .00
                                } else {
                                    cell.numFmt = '#,##0.00'; // Các cột khác giữ 2 chữ số thập phân
                                }
                            } else {
                                if (Number.isInteger(cell.value)) {
                                    cell.numFmt = '#,##0';
                                } else {
                                    cell.numFmt = '#,##0.00';
                                }
                            }
                        }
                        const styleT = templateRow.getCell(j + 1).style || {};
                        if (!Object.keys(cell.style || {}).length && Object.keys(styleT).length) {
                            cell.style = JSON.parse(JSON.stringify(styleT));
                        }
                    }
                    r.height = templateRow.height;
                    r.commit();

                    if (flagDifferentCurrency === "X") {
                        tDebitTr += it.DebitAmountInTrans || 0;
                        tCreditTr += it.CreditAmountInTrans || 0;
                    } else {
                        tDebitCo += it.DebitAmountInCoCode || 0;
                        tCreditCo += it.CreditAmountInCoCode || 0;
                    }
                }

                // ======= CHÈN TỔNG CỘNG + SỐ DƯ CUỐI KỲ =======
                const lastDataRow = START_ROW + n - 1;
                const afterDataRow = lastDataRow + 1;

                // Tổng cộng
                this.safeMerge(ws, `A${afterDataRow}:I${afterDataRow}`); // Dịch sang cột I
                const totalRow = ws.getRow(afterDataRow);
                totalRow.getCell("A").value = {
                    richText: [
                        { text: "Tổng cộng", font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "000000" } } },
                        { text: " / Total", font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "0070C0" } } }
                    ]
                };
                totalRow.getCell("A").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

                totalRow.getCell("J").value = flagDifferentCurrency === "X" ? tDebitTr : tDebitCo; // Dịch sang cột J
                totalRow.getCell("K").value = flagDifferentCurrency === "X" ? tCreditTr : tCreditCo; // Dịch sang cột K
                if (flagDifferentCurrency === "X") {
                    totalRow.getCell("J").numFmt = '#,##0.00';
                    totalRow.getCell("K").numFmt = '#,##0.00';
                } else {
                    totalRow.getCell("J").numFmt = MONEY_FMT;
                    totalRow.getCell("K").numFmt = MONEY_FMT;
                }
                totalRow.commit();

                // Số dư cuối kỳ
                this.safeMerge(ws, `A${afterDataRow + 1}:I${afterDataRow + 1}`); // Dịch sang cột I
                const endRow = ws.getRow(afterDataRow + 1);
                endRow.getCell("A").value = {
                    richText: [
                        { text: "Số dư cuối kỳ", font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "000000" } } },
                        { text: " / Ending balance", font: { name: "Times New Roman", bold: true, italic: true, color: { argb: "0070C0" } } }
                    ]
                };
                endRow.getCell("A").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

                endRow.getCell("L").value = flagDifferentCurrency === "X" ? endingTrans : endingCoCode; // Dịch sang cột L
                if (flagDifferentCurrency === "X") {
                    endRow.getCell("L").numFmt = '#,##0.00';
                } else {
                    endRow.getCell("L").numFmt = MONEY_FMT;
                }
                endRow.commit();

                this.writeFooter(ws, afterDataRow, flagDifferentCurrency, signerData, lastPostYear);

                // Tự động điều chỉnh độ rộng cột, mở rộng cột Diễn giải (cột 9)
                ws.columns.forEach((column, index) => {
                    let maxLength = 0;
                    column.eachCell({ includeEmpty: true }, cell => {
                        const length = (cell.value ? cell.value.toString().length : 10);
                        maxLength = Math.max(maxLength, length);
                    });
                    if (index === 8) { // Cột I (index 8 vì bắt đầu từ 0)
                        column.width = Math.max(maxLength + 2, 30); // Đặt độ rộng tối thiểu là 30 cho cột Diễn giải
                    } else {
                        column.width = Math.min(Math.max(maxLength + 2, 10), 20); // Giới hạn độ rộng từ 10 đến 20 cho các cột khác
                    }
                });

                // Khôi phục và hiển thị Số dư đầu kỳ (dịch sang cột J)
                const cellH12 = ws.getCell("H12");
                console.log("H12 value:", cellH12.value); // Debug: Kiểm tra giá trị H12 sau khi clone
                if (cellH12 && cellH12.value && (typeof cellH12.value === "object" || typeof cellH12.value === "string")) {
                    ws.getCell("J12").value = flagDifferentCurrency === "X" ? openingTrans : openingCoCode;
                    console.log("Setting J12 value:", flagDifferentCurrency === "X" ? openingTrans : openingCoCode); // Debug: Kiểm tra giá trị gán
                    if (flagDifferentCurrency === "X") {
                        ws.getCell("J12").numFmt = '#,##0.00';
                    } else {
                        ws.getCell("J12").numFmt = MONEY_FMT;
                    }
                } else {
                    // Nếu H12 không chứa richText, đặt số dư đầu kỳ mặc định
                    ws.getCell("J12").value = flagDifferentCurrency === "X" ? openingTrans : openingCoCode;
                    console.log("H12 not richText, setting J12 value:", flagDifferentCurrency === "X" ? openingTrans : openingCoCode); // Debug
                    if (flagDifferentCurrency === "X") {
                        ws.getCell("J12").numFmt = '#,##0.00';
                    } else {
                        ws.getCell("J12").numFmt = MONEY_FMT;
                    }
                }

                // Áp dụng thiết lập in ấn
                const lastRowToPrint = afterDataRow + 14;
                if (ws.getCell(`A${lastRowToPrint + 1}`).value) {
                    lastRowToPrint += 1;
                }
                this._applyPrinting(wb, ws, lastRowToPrint);

                // Freeze header
                ws.views = [{ state: "frozen", ySplit: START_ROW - 1 }];
            }

            // Dọn workbook
            wb.removeWorksheet(pristine.id);

            // Xuất
            try {
                const outBuffer = await wb.xlsx.writeBuffer();
                const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement("a");
                link.href = url;
                link.download = `SoQuyTienMat.xlsx`;
                link.click();
                URL.revokeObjectURL(url);
            } catch (err) {
                // MessageToast.show("Export lỗi: " + err.message);
            }
        },


        // tìm dòng chứa chuỗi (trong A..P), không phân biệt hoa/thường
        findRowByText: function (ws, text, fromRow = 1, toRow) {
            const needle = String(text).toLowerCase();
            const end = toRow || ws.rowCount;
            for (let r = fromRow; r <= end; r++) {
                const row = ws.getRow(r);
                for (let c = 1; c <= 16; c++) {
                    const v = row.getCell(c).value;
                    if (typeof v === "string" && v.toLowerCase().includes(needle)) return r;
                }
            }
            return -1;
        },

        clearRow(ws, rowIndex, lastCol = 16) {
            const row = ws.getRow(rowIndex);
            for (let c = 1; c <= lastCol; c++) row.getCell(c).value = null;
            row.commit();
        },

        base64ToUint8Array: function (b64) {
            const bin = atob(b64);
            const bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            return bytes;
        },

        // đảm bảo tên sheet hợp lệ, duy nhất
        safeSheetName: function (name, existsSet) {
            let n = String(name).substring(0, 31).replace(/[\\/?*\][:]/g, "_");
            let i = 1;
            while (existsSet.has(n)) n = (String(name).substring(0, 28) + "_" + i++).substring(0, 31);
            existsSet.add(n);
            return n;
        },

        // copy style/merge/width từ sheet mẫu sang sheet mới
        cloneSheet: function (template, target) {
            // column widths
            template.columns.forEach((col, idx) => {
                target.getColumn(idx + 1).width = col.width;
            });
            // merged cells
            template.model.merges?.forEach(range => target.mergeCells(range));
            // rows + styles
            const max = template.actualRowCount;
            for (let r = 1; r <= max; r++) {
                const src = template.getRow(r);
                const dst = target.getRow(r);
                dst.height = src.height;
                src.eachCell({ includeEmpty: true }, (cell, c) => {
                    const dcell = dst.getCell(c);
                    dcell.value = cell.value;          // copy text tĩnh, placeholder
                    dcell.style = JSON.parse(JSON.stringify(cell.style || {}));
                    if (cell.numFmt) dcell.numFmt = cell.numFmt;
                });
                dst.commit();
            }

            // ===== NEW: copy các thiết lập in ấn =====
            // target.pageSetup    = JSON.parse(JSON.stringify(template.pageSetup || {}));
            // if (template.headerFooter)
            //     target.headerFooter = JSON.parse(JSON.stringify(template.headerFooter));
            // if (template.views)
            //     target.views = template.views.map(v => JSON.parse(JSON.stringify(v)));
            // if (template.properties)
            //     target.properties = JSON.parse(JSON.stringify(template.properties));

        },

        safeMerge: function (ws, range) {
            this.unmergeIfExists(ws, range)
            // range kiểu "A17:B17"
            const r = range.toUpperCase();
            const merged = (ws.model && ws.model.merges ? ws.model.merges.map(x => x.toUpperCase()) : []);
            if (!merged.includes(r)) {
                ws.mergeCells(range);
            }
            // nếu sợ còn merge chồng lấn khác, có thể dùng:
            // try { ws.mergeCells(range); } catch (e) {
            // if (!/already merged/i.test(e.message)) throw e;
            // }
        },

        unmergeIfExists: function (sheet, range) {
            try {
                if (typeof sheet.unMergeCells === "function") {
                    sheet.unMergeCells(range); // ExcelJS >= 4.3
                } else if (sheet._merges?.[range]) {
                    delete sheet._merges[range]; // ExcelJS < 4.3
                }
            } catch (e) {
                console.warn("Unmerge failed:", range, e.message);
            }
        },

        _applyPrinting: function (wb, ws, lastRowToPrint) {
            // Xóa mọi thiết lập scale trước
            delete ws.pageSetup.scale;

            ws.pageSetup = {
                paperSize: 9,              // A4
                orientation: 'landscape',  // xoay ngang
                fitToPage: true,           // Bật chế độ fit to page
                fitToWidth: 1,             // Vừa đúng 1 trang ngang
                fitToHeight: 0,            // Chiều cao tự động (nhiều trang nếu cần)
                horizontalCentered: true,
                verticalCentered: false,
                margins: { left: 0.7, right: 0.7, top: 0.7, bottom: 0.7, header: 0.1, footer: 0.1 }, // Tăng margins
                printArea: `A1:H${lastRowToPrint}`,  // Khu vực in
                printTitlesRow: '1:10'     // Tiêu đề cố định
            };

            // Overwrite defined names để Excel chắc chắn nhận mới
            wb.definedNames.remove('Print_Area');
            wb.definedNames.remove('Print_Titles');

            const area = `'${ws.name}'!$A$1:$H$${lastRowToPrint}`;
            const titles = `'${ws.name}'!$1:$10`;
            wb.definedNames.add('Print_Area', area);
            wb.definedNames.add('Print_Titles', titles);
        },




        detectDifferentCurrency: function (dataRaw) {
            // dataRaw có thể là array hoặc {results}/{items}
            const first = Array.isArray(dataRaw)
                ? dataRaw[0]
                : (dataRaw?.results?.[0] || dataRaw?.items?.[0] || null);

            if (!first) return ""; // không có dữ liệu ⇒ coi như cùng tiền tệ

            // cố gắng bắt mọi biến thể tên field
            const pick = (obj, ...keys) => {
                for (const k of keys) {
                    if (obj && obj[k] != null && obj[k] !== "") return String(obj[k]);
                }
                return "";
            };

            const cc = pick(first, "CompanyCodeCurrency", "Companycodecurrency", "COMPANYCODECURRENCY", "CompanyCodeCurr")
                .trim().toUpperCase();
            const tr = pick(first, "TransactionCurrency", "Transactioncurrency", "TRANSACTIONCURRENCY", "TransCurr")
                .trim().toUpperCase();

            return (cc && tr && cc !== tr) ? "X" : "";
        },
        writeFooter: function (sheet, afterDataRow, flagDifferentCurrency, signerData, lastPostYear) {
            // A..M (13 cột) = nội tệ; A..P (16 cột) = khác tiền tệ
            const LAST_COL = flagDifferentCurrency === "X" ? 13 : 10; //
            const SIGN_COL = flagDifferentCurrency === "X" ? 9 : 12; // M=13, J=10

            const rowAt = (i) => sheet.getRow(afterDataRow + i);

            // Helpers
            const clearRange = (r, c1, c2) => {
                const row = sheet.getRow(r);
                for (let c = c1; c <= c2; c++) row.getCell(c).value = null;
            };
            const mergeRow = (r, c1, c2) => {
                // unmerge nếu có & clear toàn bộ trước khi merge
                try { sheet.unMergeCells(r, c1, r, c2); } catch { }
                clearRange(r, c1, c2);
                sheet.mergeCells(r, c1, r, c2);
            };

            // ----- 1) "- Sổ này có ... "
            const r1 = afterDataRow + 3;
            mergeRow(r1, 1, LAST_COL);                 // A..LAST_COL
            const cellR1 = sheet.getRow(r1).getCell(1);
            cellR1.value = "- Sổ này có ... trang, đánh số từ trang 01 đến trang ...";
            cellR1.font = { name: "Times New Roman", italic: false };
            cellR1.alignment = { horizontal: "left", wrapText: true };
            rowAt(3 - 0).commit();

            // ----- 2) "- Ngày mở sổ: ..."
            const r2 = afterDataRow + 4;
            mergeRow(r2, 1, LAST_COL);
            const cellR2 = sheet.getRow(r2).getCell(1);
            cellR2.value = "- Ngày mở sổ: ";
            cellR2.font = { name: "Times New Roman", italic: false };
            cellR2.alignment = { horizontal: "left" };
            rowAt(4 - 0).commit();

            // ----- 3) "Ngày ... tháng ... năm ..."
            const r3 = afterDataRow + 5;
            mergeRow(r3, SIGN_COL, LAST_COL);          // cột ký
            const cellR3 = sheet.getRow(r3).getCell(SIGN_COL);
            cellR3.value = "Ngày ... tháng ... năm ...";
            cellR3.font = { name: "Times New Roman", italic: true };
            cellR3.alignment = { horizontal: "center", vertical: "middle" };


            rowAt(5 - 0).commit();

            // ----- Dòng 1: Chức danh tiếng Việt -----
            const r4 = afterDataRow + 6;
            mergeRow(r4, 1, 4);                        // A..D  (Người giữ sổ)
            mergeRow(r4, 6, 8);                        // F..H  (Kế toán trưởng)
            mergeRow(r4, SIGN_COL, LAST_COL);          // cột ký (Giám đốc)

            const rowVN = sheet.getRow(r4);
            rowVN.getCell(1).value = "Người ghi sổ";
            rowVN.getCell(1).font = { name: "Times New Roman", bold: true };
            rowVN.getCell(1).alignment = { horizontal: "center", vertical: "middle" };

            rowVN.getCell(6).value = "Kế toán trưởng";
            rowVN.getCell(6).font = { name: "Times New Roman", bold: true };
            rowVN.getCell(6).alignment = { horizontal: "center", vertical: "middle" };

            if (lastPostYear < 2026) {
                rowVN.getCell(SIGN_COL).value = "Tổng giám đốc";
            } else {
                rowVN.getCell(SIGN_COL).value = "Người đại diện theo pháp luật";
            }

            rowVN.getCell(SIGN_COL).font = { name: "Times New Roman", bold: true };
            rowVN.getCell(SIGN_COL).alignment = { horizontal: "center", vertical: "middle" };
            rowVN.commit();

            // ----- Dòng 2: Chức danh tiếng Anh -----
            const r5 = afterDataRow + 7;
            mergeRow(r5, 1, 4);
            mergeRow(r5, 6, 8);
            mergeRow(r5, SIGN_COL, LAST_COL);

            const rowEN = sheet.getRow(r5);

            // Treasurer
            rowEN.getCell(1).value = "Book Keeper";
            rowEN.getCell(1).font = {
                name: "Times New Roman",
                bold: true,
                color: { argb: "0070C0" }
            };
            rowEN.getCell(1).alignment = { horizontal: "center", vertical: "middle" };

            // Chief Accountant
            rowEN.getCell(6).value = "Chief Accountant";
            rowEN.getCell(6).font = {
                name: "Times New Roman",
                bold: true,
                color: { argb: "0070C0" }
            };
            rowEN.getCell(6).alignment = { horizontal: "center", vertical: "middle" };

            // Director
            if (lastPostYear < 2026) {
                rowEN.getCell(SIGN_COL).value = "General Director";
            } else {
                rowEN.getCell(SIGN_COL).value = "Legal Representative";
            }
            rowEN.getCell(SIGN_COL).font = {
                name: "Times New Roman",
                bold: true,
                color: { argb: "0070C0" }
            };
            rowEN.getCell(SIGN_COL).alignment = { horizontal: "center", vertical: "middle" };

            rowEN.commit();

            // ----- Dòng 3: (Ký, họ tên) -----
            const r6 = afterDataRow + 14;
            mergeRow(r6, 1, 4);
            mergeRow(r6, 6, 8);
            mergeRow(r6, SIGN_COL, LAST_COL);

            const rowSign = sheet.getRow(r6);
            rowSign.getCell(1).value = "(Ký, họ tên)";
            rowSign.getCell(1).font = { name: "Times New Roman", italic: true, size: 10 };
            rowSign.getCell(1).alignment = { horizontal: "center", vertical: "middle" };

            rowSign.getCell(6).value = "(Ký, họ tên)";
            rowSign.getCell(6).font = { name: "Times New Roman", italic: true, size: 10 };
            rowSign.getCell(6).alignment = { horizontal: "center", vertical: "middle" };

            rowSign.getCell(SIGN_COL).value = "(Ký, họ tên)";
            rowSign.getCell(SIGN_COL).font = { name: "Times New Roman", italic: true, size: 10 };
            rowSign.getCell(SIGN_COL).alignment = { horizontal: "center", vertical: "middle" };
            rowSign.commit();

            const row6 = sheet.getRow(r6);
            row6.getCell(1).value = signerData.bookkeeper || "";
            row6.getCell(1).font = { name: "Times New Roman", bold: true };
            row6.getCell(1).alignment = { horizontal: "center" };

            row6.getCell(6).value = signerData.accountant || "";
            row6.getCell(6).font = { name: "Times New Roman", bold: true };
            row6.getCell(6).alignment = { horizontal: "center" };

            row6.getCell(SIGN_COL).value = signerData.director || "";
            row6.getCell(SIGN_COL).font = { name: "Times New Roman", bold: true };
            row6.getCell(SIGN_COL).alignment = { horizontal: "center" };
            row6.commit();



        }

    }
});

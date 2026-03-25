sap.ui.define([
    "sap/m/MessageToast",
    "exceljs"
], function (MessageToast, ExcelJS) {
    'use strict';

    return {
        /**
         * Exports financial cash flow report to Excel format
         * Creates a Vietnamese cash flow statement with proper formatting
         * @param {sap.ui.base.Event} oEvent - The event object from UI5
         */
        export_excel: async function (oEvent) {
            try {
                // Step 1: Data validation and preparation
                const filteredData = this._validateAndPrepareData();
                if (!filteredData) return;

                // console.log('✅Filtered Data: ', filteredData);

                // Step 2: Create and setup workbook
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet("Cash Flow Report");

                // Step 3: Configure worksheet properties
                this._setupWorksheetConfiguration(worksheet);

                // Step 4: Build report content
                this._buildReportHeader(worksheet, filteredData);
                this._buildTableStructure(worksheet, filteredData);
                const dataEndRow = this._populateDataRows(worksheet, filteredData);

                // Step 5: Apply styling and formatting
                this._applyTableBorders(worksheet, dataEndRow);
                this._addSignatureSection(worksheet, dataEndRow, filteredData);
                this._ensureTimesNewRomanFont(worksheet);

                // Step 6: Download the file
                await this._downloadExcelFile(workbook);
                MessageToast.show("Excel file exported successfully!");

            } catch (error) {
                this._handleExportError(error);
            }
        },

        /**
         * Validates model data and prepares filtered dataset
         * @returns {Array|null} Filtered data array or null if invalid
         * @private
         */
        _validateAndPrepareData: function () {
            // const oModel = this.getView().getModel();
            // const aData = oModel.oData;

            // // Check for data availability
            // if (!aData || aData.length === 0) {
            //     MessageToast.show("No data available to export.");
            //     return null;
            // }

            // // Convert to array and filter if needed
            // //const filteredData = Object.values(aData);

            // //Uncomment below for stricter filtering
            // const filteredData = Object.values(aData).filter(item =>
            //     item && item.HierarchyNode_TXT !== undefined
            // );

            const oSmartTable = this.getView().byId("zbclcttgt1::sap.suite.ui.generic.template.ListReport.view.ListReport::ZDD_LCTTGT--listReport");
            if (!oSmartTable) {
                MessageToast.show("Không tìm thấy bảng dữ liệu.");
                return;
            };
            const oTable = oSmartTable.getTable();
            if (!oTable) {
                MessageToast.show("Không tìm thấy bảng dữ liệu.");
                return;
            };
            const oBinding = oTable.getBinding("rows");
            if (!oBinding) {
                MessageToast.show("Không tìm thấy dữ liệu để export.");
                return;
            };

            const aContexts = oBinding.getAllCurrentContexts();
            const dataArray = aContexts.map(c => c.getObject());

            if (dataArray.length === 0) {
                MessageToast.show("Không có dòng dữ liệu để export.");
                return;
            };

            // Get table filtered data
            const filteredData = dataArray;

            return filteredData;
        },

        /**
         * Configures basic worksheet properties including layout and dimensions
         * @param {ExcelJS.Worksheet} worksheet - The worksheet to configure
         * @private
         */
        _setupWorksheetConfiguration: function (worksheet) {
            // Column width configuration for optimal display
            worksheet.columns = [
                { width: 62 }, // A - Chi tiêu (Item description) - wider for long text
                { width: 15 }, // B - Mã số (Item code)
                { width: 15 }, // C - Thuyết minh (Notes)
                { width: 20 }, // D - Kỳ này (Current period)
                { width: 20 }  // E - Kỳ trước (Previous period)
            ];

            // Worksheet properties for professional appearance
            worksheet.properties.outlineLevelRow = 1;
            worksheet.properties.outlineLevelCol = 1;

            // Page setup for A4 portrait printing
            worksheet.pageSetup = {
                paperSize: 9, // A4 standard
                orientation: 'portrait',
                pageOrder: 'downThenOver',
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 0
            };

            // Row height configuration for better readability
            this._setOptimalRowHeights(worksheet);
        },

        /**
         * Sets optimal row heights for different sections of the report
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @private
         */
        _setOptimalRowHeights: function (worksheet) {
            const rowHeights = {
                1: 25,  // Company header
                2: 40,  // Address (needs extra height for text wrapping)
                3: 15,  // Empty spacing row
                4: 15,  // Empty spacing row
                5: 30,  // Main report title
                6: 20,  // Report subtitle
                7: 20,  // Period and currency info
                8: 25,  // Table headers row 1
                9: 25,  // Table headers row 2
                10: 20  // Column number indicators
            };

            Object.entries(rowHeights).forEach(([rowNum, height]) => {
                worksheet.getRow(parseInt(rowNum)).height = height;
            });
        },

        /**
         * Builds the report header section including company info and title
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - Data array for extracting period info
         * @private
         */
        _buildReportHeader: function (worksheet, filteredData) {
            // Company information section
            this._addCompanyInformation(worksheet, filteredData);

            // Report title and subtitle
            this._addReportTitles(worksheet, filteredData);

            // Period and currency information
            this._addPeriodAndCurrencyInfo(worksheet, filteredData);
        },

        /**
         * Adds company name, address and form template information
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - Data for extracting report type
         * @private
         */
        _addCompanyInformation: function (worksheet, filteredData) {
            const firstItem = filteredData.length > 0 ? filteredData[0] : {};

            if (firstItem.type_rp == 1) {
                // Company name (merged across A1:B1)
                worksheet.mergeCells('A1:B1');
                this._setCellWithFormatting(worksheet, 'A1',
                    'CÔNG TY TNHH CẢNG QUỐC TẾ TIL CẢNG HẢI PHÒNG', {
                    font: { bold: true, name: 'Times New Roman' }
                });


                // Form template reference (merged across C1:E1)
                worksheet.mergeCells('C1:E1');
                this._setCellWithFormatting(worksheet, 'C1',
                    'Mẫu số B03a - DN', {
                    font: { bold: true, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

                // Company address (merged across A2:B2)
                worksheet.mergeCells('A2:B2');
                this._setCellWithFormatting(worksheet, 'A2',
                    'Bến 3&4 Cảng nước sâu Lạch Huyện, Khu phố Đôn Lương, Đặc khu Cát Hải, Thành phố Hải Phòng, Việt Nam', {
                    font: { name: 'Times New Roman' },
                    alignment: { horizontal: 'left', wrapText: true, vertical: 'middle' }
                });

                let thongTu;

                if (firstItem.gjahr < 2026) {
                    thongTu = `(Ban hành theo Thông tư số 200/2014/TT-BTC\nngày 22/12/2014 của Bộ Tài chính)`
                } else {
                    thongTu = `(Kèm theo Thông tư số 99/2025/TT-BTC\nngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính)`
                }

                // Legal reference (merged across C2:E2)
                worksheet.mergeCells('C2:E2');
                this._setCellWithFormatting(worksheet, 'C2',
                    thongTu, {
                    font: { name: 'Times New Roman', size: 10 },
                    alignment: { horizontal: 'center', wrapText: true, vertical: 'middle' }
                });

            } else {
                // Company name (merged across A1:B1)
                worksheet.mergeCells('A1:B1');
                this._setCellWithFormatting(worksheet, 'A1',
                    'HAIPHONG PORT TIL INTERNATIONAL TERMINAL COMPANY', {
                    font: { bold: true, name: 'Times New Roman' }
                });

                // Company address (merged across A2:B2)
                worksheet.mergeCells('A2:B2');
                this._setCellWithFormatting(worksheet, 'A2',
                    'Berth No. 3&4 Lach Huyen Deep-sea Port, Don Luong Quarter, Cat Hai Special Zone, Hai Phong City, Vietnam', {
                    font: { name: 'Times New Roman' },
                    alignment: { horizontal: 'left', wrapText: true, vertical: 'middle' }
                });
            }
        },

        /**
         * Adds main report title and subtitle
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - Data for extracting report type
         * @private
         */
        _addReportTitles: function (worksheet, filteredData) {
            const firstItem = filteredData.length > 0 ? filteredData[0] : {};

            if (firstItem.type_rp == 1) {
                // Main report title
                worksheet.mergeCells('A5:E5');
                this._setCellWithFormatting(worksheet, 'A5',
                    'BÁO CÁO LƯU CHUYỂN TIỀN TỆ', {
                    font: { bold: true, size: 16, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

                // Report subtitle with methodology note
                worksheet.mergeCells('A6:E6');
                this._setCellWithFormatting(worksheet, 'A6',
                    '(Theo phương pháp gián tiếp) (*)', {
                    font: { italic: true, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

            } else {
                // Main report title
                worksheet.mergeCells('A5:E5');
                this._setCellWithFormatting(worksheet, 'A5',
                    'CASH FLOW STATEMENT', {
                    font: { bold: true, size: 16, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

                // Report subtitle with methodology note
                worksheet.mergeCells('A6:E6');
                this._setCellWithFormatting(worksheet, 'A6',
                    '(Direct method) (*)', {
                    font: { italic: true, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });
            }
        },

        /**
         * Adds period information and currency unit
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - Data for extracting period and currency
         * @private
         */
        _addPeriodAndCurrencyInfo: function (worksheet, filteredData) {
            // Extract period information from first data item
            const firstItem = filteredData.length > 0 ? filteredData[0] : {};
            const periodInfo = this._extractPeriodInfo(firstItem);

            if (firstItem.type_rp == 1) {
                // Period information (merged across A7:C7)
                worksheet.mergeCells('A7:E7');
                this._setCellWithFormatting(worksheet, 'A7',
                    `Kỳ: ${periodInfo.type} - Năm: ${periodInfo.gjahr}`, {
                    font: { name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

                // Currency unit information (merged across D7:E7)
                worksheet.mergeCells('D8:E8');
                this._setCellWithFormatting(worksheet, 'D8',
                    `Đơn vị tiền tệ: ${periodInfo.currency_code}`, {
                    font: { italic: true, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

            } else {
                // Period information (merged across A7:C7)
                worksheet.mergeCells('A7:E7');
                this._setCellWithFormatting(worksheet, 'A7',
                    `Period: ${periodInfo.type}.${periodInfo.gjahr}`, {
                    font: { name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });

                // Currency unit information (merged across D7:E7)
                worksheet.mergeCells('D8:E8');
                this._setCellWithFormatting(worksheet, 'D8',
                    `Express in: ${periodInfo.currency_code}`, {
                    font: { italic: true, name: 'Times New Roman' },
                    alignment: { horizontal: 'center' }
                });
            }
        },

        /**
         * Extracts period and currency information from data item
         * @param {Object} dataItem - First data item containing period info
         * @returns {Object} Period information object
         * @private
         */
        _extractPeriodInfo: function (dataItem) {
            return {
                type: dataItem.type || '',
                gjahr: dataItem.gjahr || '',
                currency_code: dataItem.currency_code || ''
            };
        },

        /**
         * Builds the complete table structure with headers and column indicators
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - The data to populate
         * @private
         */
        _buildTableStructure: function (worksheet, filteredData) {
            this._addTableHeaders(worksheet, filteredData);
            this._addColumnNumberIndicators(worksheet);
        },

        /**
         * Adds table headers with proper merging and formatting
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - The data to populate
         * @private
         */
        _addTableHeaders: function (worksheet, filteredData) {
            const headerFont = { bold: true, name: 'Times New Roman' };
            const centerAlignment = { horizontal: 'center', vertical: 'middle' };
            const firstItem = filteredData.length > 0 ? filteredData[0] : {};

            let mainHeaders;

            if (firstItem.type_rp == 1) {
                // Main column headers (spanning two rows each)
                mainHeaders = [
                    { range: 'A9:A9', text: 'Chỉ tiêu' },
                    { range: 'B9:B9', text: 'Mã số' },
                    { range: 'C9:C9', text: 'Thuyết minh' },
                    { range: 'D9:D9', text: 'Kỳ này' },
                    { range: 'E9:E9', text: 'Kỳ trước' }
                ];

            } else {
                // Main column headers (spanning two rows each)
                mainHeaders = [
                    { range: 'A9:A9', text: 'Description' },
                    { range: 'B9:B9', text: 'Codes' },
                    { range: 'C9:C9', text: 'Notes' },
                    { range: 'D9:D9', text: 'This period' },
                    { range: 'E9:E9', text: 'Previous period' }
                ];
            }
            mainHeaders.forEach(header => {
                worksheet.mergeCells(header.range);
                this._setCellWithFormatting(worksheet, header.range.split(':')[0], header.text, {
                    font: headerFont,
                    alignment: centerAlignment
                });
            });

            const row8 = worksheet.getRow(8);
            row8.height = 18, 5;
            const row9 = worksheet.getRow(9);
            row9.height = 18, 5;
        },

        /**
         * Adds column number indicators (1, 2, 3, 4, 5)
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @private
         */
        _addColumnNumberIndicators: function (worksheet) {
            const columnCells = ['A10', 'B10', 'C10', 'D10', 'E10'];
            const font = { name: 'Times New Roman' };
            const alignment = { horizontal: 'center' };

            columnCells.forEach((cell, index) => {
                this._setCellWithFormatting(worksheet, cell, (index + 1).toString(), {
                    font: font,
                    alignment: alignment
                });
            });
            const row10 = worksheet.getRow(10);
            row10.height = 18, 5;
        },

        /**
         * Populates data rows with financial information
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {Array} filteredData - The data to populate
         * @returns {number} The row number after the last data row
         * @private
         */
        _populateDataRows: function (worksheet, filteredData) {
            let currentRow = 11; // Start after headers
            const font = { name: 'Times New Roman' };

            filteredData.forEach((item) => {
                // Set row height for better readability
                worksheet.getRow(currentRow).height = 30;

                // Populate each column with appropriate formatting
                this._populateDataRowColumns(worksheet, currentRow, item, font);

                currentRow++;
            });

            return currentRow;
        },

        /**
         * Populates individual columns for a data row
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} rowNum - Current row number
         * @param {Object} item - Data item
         * @param {Object} font - Font configuration
         * @private
         */
        _populateDataRowColumns: function (worksheet, rowNum, item, font) {
            // Ngưỡng số ký tự / dòng
            const CHAR_PER_LINE = 74;
            // Tính số dòng cần thêm
            const lines = Math.ceil(item.HierarchyNode_TXT.length / CHAR_PER_LINE);
            const row = worksheet.getRow(rowNum);
            row.height = lines * 18, 5;
            // Column A: Chi tiêu (with conditional bold formatting)
            const isHeaderItem = ((item.Zfont == 3) || (item.Zfont == 'X')); // isNaN(item.CHI_TIEU.charAt(0));
            this._setCellWithFormatting(worksheet, `A${rowNum}`, item.HierarchyNode_TXT, {
                font: { name: 'Times New Roman', bold: isHeaderItem },
                alignment: { horizontal: 'left', vertical: 'middle', wrapText: true }
            });

            // Column B: Mã số (with null handling)
            //const masoValue = item.MA_SO === 0 ? ' ' : item.MA_SO;
            this._setCellWithFormatting(worksheet, `B${rowNum}`, item.HierarchyNode, //MA_SO,
                { font: { name: 'Times New Roman', bold: isHeaderItem } });

            // Column C: Thuyết minh
            this._setCellWithFormatting(worksheet, `C${rowNum}`, '',
                { font: { name: 'Times New Roman', bold: isHeaderItem } });

            // Column D: Kỳ này (with number formatting)
            this._setFinancialCell(worksheet, `D${rowNum}`, item.sokynay,
                { name: 'Times New Roman', bold: isHeaderItem });

            // Column E: Kỳ trước (with number formatting)
            this._setFinancialCell(worksheet, `E${rowNum}`, item.sokytruoc,
                { name: 'Times New Roman', bold: isHeaderItem });
        },

        /**
         * Sets a financial cell with proper number formatting
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {string} cellRef - Cell reference
         * @param {*} value - Cell value
         * @param {Object} font - Font configuration
         * @private
         */
        _setFinancialCell: function (worksheet, cellRef, value, font = {}) {
            // Handle null/undefined/zero values
            //const displayValue = (value === null || value === undefined || value === 0.00) ? ' ' : value;

            const cell = worksheet.getCell(cellRef);
            //cell.value = displayValue;
            // Nếu là null/undefined → để trống
            cell.value = Number(value);
            cell.font = font;
            cell.numFmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';// '#,##0.00'; 

            //cell.numFmt = '#.##0'; // Number format with comma separators
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
        },

        /**
         * Applies borders to the entire table area
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} endRow - Last row of the table
         * @private
         */
        _applyTableBorders: function (worksheet, endRow) {
            const borderStyle = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            // Apply borders to all table cells (from row 8 to endRow-1, columns A-E)
            for (let row = 9; row < endRow; row++) {
                for (let col = 1; col <= 5; col++) {
                    worksheet.getCell(row, col).border = borderStyle;
                }
            }
        },

        /**
         * Adds signature section with preparer, accountant, and director
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} endRow - Row after last data row
         * @private
         */
        _addSignatureSection: function (worksheet, endRow, filteredData) {
            const footerStartRow = endRow + 3; // Add spacing after table

            // Configure footer row heights
            this._setFooterRowHeights(worksheet, footerStartRow);

            // Add signature blocks
            this._addPreparerSignature(worksheet, footerStartRow);
            this._addAccountantSignature(worksheet, footerStartRow);
            this._addDirectorSignature(worksheet, footerStartRow, filteredData);
        },

        /**
         * Sets heights for footer rows
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} startRow - Starting row for footer
         * @private
         */
        _setFooterRowHeights: function (worksheet, startRow) {
            worksheet.getRow(startRow).height = 25;     // Date row
            worksheet.getRow(startRow + 1).height = 20; // Title row
            worksheet.getRow(startRow + 2).height = 20; // Signature instruction row
        },

        /**
         * Adds preparer signature block
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} startRow - Starting row for footer
         * @private
         */
        _addPreparerSignature: function (worksheet, startRow) {
            this._setCellWithFormatting(worksheet, `A${startRow + 1}`, 'Người lập biểu', {
                font: { bold: true, name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });

            this._setCellWithFormatting(worksheet, `A${startRow + 2}`, '(Ký, họ tên)', {
                font: { name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });
        },

        /**
         * Adds accountant signature block
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} startRow - Starting row for footer
         * @private
         */
        _addAccountantSignature: function (worksheet, startRow) {
            // Merge B and C columns for accountant section
            worksheet.mergeCells(`B${startRow + 1}:C${startRow + 1}`);
            this._setCellWithFormatting(worksheet, `B${startRow + 1}`, 'Kế toán trưởng', {
                font: { bold: true, name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });

            worksheet.mergeCells(`B${startRow + 2}:C${startRow + 2}`);
            this._setCellWithFormatting(worksheet, `B${startRow + 2}`, '(Ký, họ tên)', {
                font: { name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });
        },

        /**
         * Adds director signature block with current date
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {number} startRow - Starting row for footer
         * @private
         */
        _addDirectorSignature: function (worksheet, startRow, filteredData) {
            const firstItem = filteredData.length > 0 ? filteredData[0] : {};
            let director;

            if (firstItem.gjahr < 2026) {
                director = 'Tổng giám đốc';
            } else {
                director = 'Người đại diện theo pháp luật';
            }
            
            // Generate current date string
            const currentDate = new Date();
            const dateString = this._formatCurrentDate(currentDate);

            // Date line
            worksheet.mergeCells(`D${startRow}:E${startRow}`);
            this._setCellWithFormatting(worksheet, `D${startRow}`, dateString, {
                font: { name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });

            // Director title
            worksheet.mergeCells(`D${startRow + 1}:E${startRow + 1}`);
            this._setCellWithFormatting(worksheet, `D${startRow + 1}`, director, {
                font: { bold: true, name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });

            // Signature instruction
            worksheet.mergeCells(`D${startRow + 2}:E${startRow + 2}`);
            this._setCellWithFormatting(worksheet, `D${startRow + 2}`, '(Ký, họ tên, đóng dấu)', {
                font: { name: 'Times New Roman' },
                alignment: { horizontal: 'center' }
            });
        },

        /**
         * Formats current date for Vietnamese report format
         * @param {Date} date - Date to format
         * @returns {string} Formatted date string
         * @private
         */
        _formatCurrentDate: function (date) {
            const day = date.getDate();
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            return `……… ngày ${day} tháng ${month} năm ${year}`;
        },

        /**
         * Ensures all cells use Times New Roman font
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @private
         */
        _ensureTimesNewRomanFont: function (worksheet) {
            worksheet.eachRow((row) => {
                row.eachCell((cell) => {
                    if (!cell.font || !cell.font.name) {
                        cell.font = { name: 'Times New Roman' };
                    }
                });
            });
        },

        /**
         * Downloads the Excel file to user's computer
         * @param {ExcelJS.Workbook} workbook - The workbook to download
         * @private
         */
        _downloadExcelFile: async function (workbook) {
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            // Create and trigger download
            const url = window.URL.createObjectURL(blob);
            const downloadLink = document.createElement('a');
            downloadLink.href = url;
            downloadLink.download = 'BC_Luu_Chuyen_Tien_Te_Gian_Tiep.xlsx';

            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);
            window.URL.revokeObjectURL(url);
        },

        /**
         * Handles export errors with user-friendly messages
         * @param {Error} error - The error object
         * @private
         */
        _handleExportError: function (error) {
            console.error("Excel export error:", error);
            MessageToast.show("Error exporting Excel file: " + error.message);
        },

        /**
         * Helper method to set cell value with formatting options
         * @param {ExcelJS.Worksheet} worksheet - The worksheet
         * @param {string} cellRef - Cell reference (e.g., 'A1')
         * @param {*} value - Cell value
         * @param {Object} formatting - Formatting options (font, alignment, etc.)
         * @private
         */
        _setCellWithFormatting: function (worksheet, cellRef, value, formatting = {}) {
            const cell = worksheet.getCell(cellRef);
            cell.value = value;

            // Apply formatting options
            if (formatting.font) cell.font = formatting.font;
            if (formatting.alignment) cell.alignment = formatting.alignment;
            if (formatting.numFmt) cell.numFmt = formatting.numFmt;
            if (formatting.border) cell.border = formatting.border;
        }
    };
});
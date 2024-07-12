import * as FUniver from "@univerjs/facade";
import * as UniverJS from "@univerjs/core";
import * as ExcelJS from 'exceljs';
import { DEFAULT_BORDER_COLOR } from '../utils/enum';

const waitUserSelectExcelFile = (
    onSelect: (workbook: ExcelJS.Workbook) => void,
    onError: (error: Error) => void
) => {
    try {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = ".xls, .xlsx";

        input.click();

        input.onchange = () => {
            const file = input.files?.[0];
            if (!file) {
                onError(new Error('No file selected.'));
                return;
            }

            const reader = new FileReader();
            reader.readAsArrayBuffer(file);

            reader.onload = () => {
                if (reader.result instanceof ArrayBuffer) {
                    const data = new Uint8Array(reader.result);
                    const workbook = new ExcelJS.Workbook();
                    workbook.xlsx.load(data).then(() => {
                        onSelect(workbook);
                    }).catch(error => {
                        onError(error);
                    });
                } else {
                    onError(new Error('Reader result is not an ArrayBuffer.'));
                }
            };
        };
    } catch (error) {
        onError(error instanceof Error ? error : new Error('Unknown error occurred.'));
    }
};

/**
 * Parse Excel worksheet to extract relevant information.
 * @param sheet The Excel worksheet to parse.
 * @returns An object containing information about the worksheet.
 */
const parseExcelUniverSheetInfo = (sheet: ExcelJS.Worksheet): UniverJS.IWorksheetData => {
    const sheetId = sheet.name;
    const name = sheet.name;
    // const type = SheetTypes.GRID;
    const rowCount = sheet.rowCount + 100;
    const columnCount = sheet.columnCount + 100;
    const defaultColumnWidth = 72;
    const defaultRowHeight = 17.5;
    const scrollTop = 200;
    const scrollLeft = 100;
    const selections = ['A2'];
    const hidden = UniverJS.BooleanNumber.FALSE;
    // const status = 1;
    const showGridlines = 1;
    const rowHeaderWidth = 46;
    const columnHeaderHeight = 20;
    const rightToLeft = UniverJS.BooleanNumber.FALSE;
    const zoomRatio = 1
    const freeze: UniverJS.IFreeze = {
        xSplit: 0, // 水平方向分割的位置
        ySplit: 0, // 垂直方向分割的位置
        startRow: 0, // 冻结区域左上角单元格的行索引，设置为0表示从第一行开始冻结
        startColumn: 0, // 冻结区域左上角单元格的列索引
    };
    const mergeDataBefore: UniverJS.IRange[] = []; // Merged cells
    const cellData: UniverJS.IObjectMatrixPrimitiveType<UniverJS.ICellData> = {}; // Cell data
    const rowData: UniverJS.IObjectArrayPrimitiveType<Partial<UniverJS.IRowData>> = {}; // Row data
    const columnData: UniverJS.IObjectArrayPrimitiveType<Partial<UniverJS.IColumnData>> = {}; // Column data
    sheet.eachRow({ includeEmpty: true }, (row: any) => {
        row.eachCell({ includeEmpty: true }, (cell: any) => {
            if (cell.isMerged) {
                // 如果当前单元格是合并单元格，则将其范围添加到 mergeData 数组中
                const masterCell = cell.master;
                mergeDataBefore.push({
                    startRow: masterCell.fullAddress.row - 1,
                    startColumn: masterCell.fullAddress.col - 1,
                    endRow: cell.fullAddress.row - 1,
                    endColumn: cell.fullAddress.col - 1
                });
            }
        });
    });

    // 创建一个对象，以便根据起始行和列的组合查找合并范围
    const mergedRangesMap: { [key: string]: UniverJS.IRange } = {};

    // 遍历合并范围数组并整理数据
    mergeDataBefore.forEach(range => {
        const key = `${range.startRow}-${range.startColumn}`;
        if (!mergedRangesMap[key]) {
            mergedRangesMap[key] = range;
        } else {
            // 更新现有范围的结束行和列
            mergedRangesMap[key].endRow = Math.max(mergedRangesMap[key].endRow, range.endRow);
            mergedRangesMap[key].endColumn = Math.max(mergedRangesMap[key].endColumn, range.endColumn);
        }
    });

    // 将整理后的合并范围转换为数组
    const mergeData: UniverJS.IRange[] = Object.values(mergedRangesMap);

    console.log(sheet.name)

    //行高
    for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
        const row = sheet.getRow(rowIndex)
        rowData[rowIndex - 1] = { h: row.height ? row.height * (1 + 1 / 3) : undefined };
    }

    //列宽
    for (let colIndex = 1; colIndex <= sheet.columnCount; colIndex++) {
        const column = sheet.getColumn(colIndex)
        columnData[colIndex - 1] = { w: column.width ? column.width * 8 : undefined };
    }

    for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
        const row = sheet.getRow(rowIndex)
        for (let colIndex = 1; colIndex <= sheet.columnCount; colIndex++) {
            const cell = row.getCell(colIndex)
            cellData[rowIndex - 1] = cellData[rowIndex - 1] || {}
            if (cell.model) {
                if (cell.model.formula) {
                    cellData[rowIndex - 1][colIndex - 1] = { f: '=' + cell.model.formula, v: cell.model.result }
                    continue
                }
            }
            if (cell.value) {
                if (cell.isMerged && cell !== cell.master) {
                    cellData[rowIndex - 1][colIndex - 1] = {};
                } else {
                    cellData[rowIndex - 1][colIndex - 1] = { v: cell.value?.toString() };
                }
            } else {
                cellData[rowIndex - 1][colIndex - 1] = {};
            }

            //没有样式
            if (cell.style && Object.keys(cell.style).length === 0) {
                continue
            }

            let cellStyle: UniverJS.IStyleData = {}

            //字体名字，如宋体
            if (cell.style.font?.name) {
                cellStyle.ff = cell.style.font.name
            }

            //字体大小
            if (cell.style.font?.size) {
                cellStyle.fs = cell.style.font.size
            }

            //字体斜体
            if (cell.style.font?.italic) {
                cellStyle.it = 1
            }

            //字体加粗
            if (cell.style.font?.bold) {
                cellStyle.bl = 1
            }

            //下划线
            let cellStyleITextDecoration: UniverJS.ITextDecoration = { s: 0, };
            let cellStyleTextDecoration: UniverJS.TextDecoration = 12
            if (cell.style.font?.underline) {
                cellStyleITextDecoration.s = 1
                if (cell.font.underline === 'double') {
                    cellStyleTextDecoration = 10
                } else if (cell.font.underline === 'singleAccounting') {
                    cellStyleTextDecoration = 14
                } else if (cell.font.underline === 'doubleAccounting') {
                    cellStyleTextDecoration = 15
                }
                cellStyleITextDecoration.t = cellStyleTextDecoration
            }

            //删除线
            let cellStyleStrikeITextDecoration: UniverJS.ITextDecoration = { s: 0, };
            if (cell.style.font?.strike) {
                cellStyleStrikeITextDecoration.s = 1
            }

            //上划线
            // let cellStyleOverlineITextDecoration: UniverJS.ITextDecoration = { s: 0,};

            //背景颜色
            let cellStyleBackground: UniverJS.IColorStyle = { rgb: '#ffffff' }
            if (cell.style.fill) {
                if (cell.style.fill.type === 'pattern') {
                    if (cell.style.fill.fgColor?.argb) {
                        const argb = cell.style.fill.fgColor.argb
                        cellStyleBackground.rgb = '#' + argb.slice(-6);
                    }
                }
            }

            //边框
            let cellStyleBorder: UniverJS.IBorderData = {}
            //左边框
            if (cell.style.border?.left) {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#000000' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                if (cell.style.border.left.style === 'thin') {
                    cellStyleIBorderStyleData.s = 1
                } else if (cell.style.border.left.style === 'hair') {
                    cellStyleIBorderStyleData.s = 2
                } else if (cell.style.border.left.style === 'dotted') {
                    cellStyleIBorderStyleData.s = 3
                } else if (cell.style.border.left.style === 'dashed') {
                    cellStyleIBorderStyleData.s = 4
                } else if (cell.style.border.left.style === 'dashDot') {
                    cellStyleIBorderStyleData.s = 5
                } else if (cell.style.border.left.style === 'dashDotDot') {
                    cellStyleIBorderStyleData.s = 6
                } else if (cell.style.border.left.style === 'double') {
                    cellStyleIBorderStyleData.s = 7
                } else if (cell.style.border.left.style === 'medium') {
                    cellStyleIBorderStyleData.s = 8
                } else if (cell.style.border.left.style === 'mediumDashed') {
                    cellStyleIBorderStyleData.s = 9
                } else if (cell.style.border.left.style === 'mediumDashDot') {
                    cellStyleIBorderStyleData.s = 10
                } else if (cell.style.border.left.style === 'mediumDashDotDot') {
                    cellStyleIBorderStyleData.s = 11
                } else if (cell.style.border.left.style === 'slantDashDot') {
                    cellStyleIBorderStyleData.s = 12
                } else {
                    cellStyleIBorderStyleData.s = 13
                }
                if (cell.style.border.left.color?.argb) {
                    const argb = cell.style.border.left.color.argb
                    cellStyleIColorStyle.rgb = '#' + argb.slice(-6);
                }
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.l = cellStyleIBorderStyleData
            } else {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: DEFAULT_BORDER_COLOR }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                cellStyleIBorderStyleData.s = 1
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.l = cellStyleIBorderStyleData
            }

            //上边框
            if (cell.style.border?.top) {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#000000' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                if (cell.style.border.top.style === 'thin') {
                    cellStyleIBorderStyleData.s = 1
                } else if (cell.style.border.top.style === 'hair') {
                    cellStyleIBorderStyleData.s = 2
                } else if (cell.style.border.top.style === 'dotted') {
                    cellStyleIBorderStyleData.s = 3
                } else if (cell.style.border.top.style === 'dashed') {
                    cellStyleIBorderStyleData.s = 4
                } else if (cell.style.border.top.style === 'dashDot') {
                    cellStyleIBorderStyleData.s = 5
                } else if (cell.style.border.top.style === 'dashDotDot') {
                    cellStyleIBorderStyleData.s = 6
                } else if (cell.style.border.top.style === 'double') {
                    cellStyleIBorderStyleData.s = 7
                } else if (cell.style.border.top.style === 'medium') {
                    cellStyleIBorderStyleData.s = 8
                } else if (cell.style.border.top.style === 'mediumDashed') {
                    cellStyleIBorderStyleData.s = 9
                } else if (cell.style.border.top.style === 'mediumDashDot') {
                    cellStyleIBorderStyleData.s = 10
                } else if (cell.style.border.top.style === 'mediumDashDotDot') {
                    cellStyleIBorderStyleData.s = 11
                } else if (cell.style.border.top.style === 'slantDashDot') {
                    cellStyleIBorderStyleData.s = 12
                } else {
                    cellStyleIBorderStyleData.s = 13
                }
                if (cell.style.border.top.color?.argb) {
                    const argb = cell.style.border.top.color.argb
                    cellStyleIColorStyle.rgb = '#' + argb.slice(-6);
                }
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.t = cellStyleIBorderStyleData
            } else {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#d6d8db' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                cellStyleIBorderStyleData.s = 1
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.t = cellStyleIBorderStyleData
            }

            //右边框
            if (cell.style.border?.right) {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#000000' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                if (cell.style.border.right.style === 'thin') {
                    cellStyleIBorderStyleData.s = 1
                } else if (cell.style.border.right.style === 'hair') {
                    cellStyleIBorderStyleData.s = 2
                } else if (cell.style.border.right.style === 'dotted') {
                    cellStyleIBorderStyleData.s = 3
                } else if (cell.style.border.right.style === 'dashed') {
                    cellStyleIBorderStyleData.s = 4
                } else if (cell.style.border.right.style === 'dashDot') {
                    cellStyleIBorderStyleData.s = 5
                } else if (cell.style.border.right.style === 'dashDotDot') {
                    cellStyleIBorderStyleData.s = 6
                } else if (cell.style.border.right.style === 'double') {
                    cellStyleIBorderStyleData.s = 7
                } else if (cell.style.border.right.style === 'medium') {
                    cellStyleIBorderStyleData.s = 8
                } else if (cell.style.border.right.style === 'mediumDashed') {
                    cellStyleIBorderStyleData.s = 9
                } else if (cell.style.border.right.style === 'mediumDashDot') {
                    cellStyleIBorderStyleData.s = 10
                } else if (cell.style.border.right.style === 'mediumDashDotDot') {
                    cellStyleIBorderStyleData.s = 11
                } else if (cell.style.border.right.style === 'slantDashDot') {
                    cellStyleIBorderStyleData.s = 12
                } else {
                    cellStyleIBorderStyleData.s = 13
                }
                if (cell.style.border.right.color?.argb) {
                    const argb = cell.style.border.right.color.argb
                    cellStyleIColorStyle.rgb = '#' + argb.slice(-6);
                }
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.r = cellStyleIBorderStyleData
            } else {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#d6d8db' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                cellStyleIBorderStyleData.s = 1
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.r = cellStyleIBorderStyleData
            }

            //下边框
            if (cell.style.border?.bottom) {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#000000' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                if (cell.style.border.bottom.style === 'thin') {
                    cellStyleIBorderStyleData.s = 1
                } else if (cell.style.border.bottom.style === 'hair') {
                    cellStyleIBorderStyleData.s = 2
                } else if (cell.style.border.bottom.style === 'dotted') {
                    cellStyleIBorderStyleData.s = 3
                } else if (cell.style.border.bottom.style === 'dashed') {
                    cellStyleIBorderStyleData.s = 4
                } else if (cell.style.border.bottom.style === 'dashDot') {
                    cellStyleIBorderStyleData.s = 5
                } else if (cell.style.border.bottom.style === 'dashDotDot') {
                    cellStyleIBorderStyleData.s = 6
                } else if (cell.style.border.bottom.style === 'double') {
                    cellStyleIBorderStyleData.s = 7
                } else if (cell.style.border.bottom.style === 'medium') {
                    cellStyleIBorderStyleData.s = 8
                } else if (cell.style.border.bottom.style === 'mediumDashed') {
                    cellStyleIBorderStyleData.s = 9
                } else if (cell.style.border.bottom.style === 'mediumDashDot') {
                    cellStyleIBorderStyleData.s = 10
                } else if (cell.style.border.bottom.style === 'mediumDashDotDot') {
                    cellStyleIBorderStyleData.s = 11
                } else if (cell.style.border.bottom.style === 'slantDashDot') {
                    cellStyleIBorderStyleData.s = 12
                } else {
                    cellStyleIBorderStyleData.s = 13
                }
                if (cell.style.border.bottom.color?.argb) {
                    const argb = cell.style.border.bottom.color.argb
                    cellStyleIColorStyle.rgb = '#' + argb.slice(-6);
                }
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.b = cellStyleIBorderStyleData
            } else {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#d6d8db' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                cellStyleIBorderStyleData.s = 1
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                cellStyleBorder.b = cellStyleIBorderStyleData
            }

            //对角线
            if (cell.style.border?.diagonal) {
                let cellStyleIColorStyle: UniverJS.IColorStyle = { rgb: '#000000' }
                let cellStyleIBorderStyleData: UniverJS.IBorderStyleData = { s: 0, cl: cellStyleIColorStyle }
                if (cell.style.border.diagonal.style === 'thin') {
                    cellStyleIBorderStyleData.s = 1
                } else if (cell.style.border.diagonal.style === 'hair') {
                    cellStyleIBorderStyleData.s = 2
                } else if (cell.style.border.diagonal.style === 'dotted') {
                    cellStyleIBorderStyleData.s = 3
                } else if (cell.style.border.diagonal.style === 'dashed') {
                    cellStyleIBorderStyleData.s = 4
                } else if (cell.style.border.diagonal.style === 'dashDot') {
                    cellStyleIBorderStyleData.s = 5
                } else if (cell.style.border.diagonal.style === 'dashDotDot') {
                    cellStyleIBorderStyleData.s = 6
                } else if (cell.style.border.diagonal.style === 'double') {
                    cellStyleIBorderStyleData.s = 7
                } else if (cell.style.border.diagonal.style === 'medium') {
                    cellStyleIBorderStyleData.s = 8
                } else if (cell.style.border.diagonal.style === 'mediumDashed') {
                    cellStyleIBorderStyleData.s = 9
                } else if (cell.style.border.diagonal.style === 'mediumDashDot') {
                    cellStyleIBorderStyleData.s = 10
                } else if (cell.style.border.diagonal.style === 'mediumDashDotDot') {
                    cellStyleIBorderStyleData.s = 11
                } else if (cell.style.border.diagonal.style === 'slantDashDot') {
                    cellStyleIBorderStyleData.s = 12
                } else {
                    cellStyleIBorderStyleData.s = 13
                }
                if (cell.style.border.diagonal.color?.argb) {
                    const argb = cell.style.border.diagonal.color.argb
                    cellStyleIColorStyle.rgb = '#' + argb.slice(-6);
                }
                cellStyleIBorderStyleData.cl = cellStyleIColorStyle
                if (cell.style.border.diagonal.up === false && cell.style.border.diagonal.down === true) {
                    cellStyleBorder.tl_br = cellStyleIBorderStyleData
                } else if (cell.style.border.diagonal.up === true && cell.style.border.diagonal.down === false) {
                    cellStyleBorder.bl_tr = cellStyleIBorderStyleData
                } else if (cell.style.border.diagonal.up === true && cell.style.border.diagonal.down === true) {
                    cellStyleBorder.tl_br = cellStyleIBorderStyleData
                    cellStyleBorder.bl_tr = cellStyleIBorderStyleData
                }
            }

            //字体颜色
            let cellStyleForeground: UniverJS.IColorStyle = { rgb: '#000000' }
            if (cell.style.font?.color) {
                if (cell.style.font.color?.argb) {
                    const argb = cell.style.font.color.argb
                    cellStyleForeground.rgb = '#' + argb.slice(-6);
                }
            }

            //上标，下标
            let cellStyleBaselineOffset: UniverJS.BaselineOffset = 1
            if (cell.style.font?.vertAlign) {
                if (cell.style.font.vertAlign === 'subscript') {
                    cellStyleBaselineOffset = 2
                } else {
                    cellStyleBaselineOffset = 3
                }
            }

            // 文本旋转样式
            const csat = cell.style.alignment?.textRotation
            let cellStyleTextRotation: UniverJS.ITextRotation = { a: 0, v: 0 };
            if (csat) {
                if (csat === 'vertical') {
                    cellStyleTextRotation.v = 1;
                } else {
                    cellStyleTextRotation.a = -csat;
                }
            }

            // 读取方向
            let cellStyleTextDirection: UniverJS.TextDirection = 0
            const csar = cell.style.alignment?.readingOrder
            if (csar) {
                if (csar === 'ltr') {
                    cellStyleTextDirection = 1;
                } else {
                    cellStyleTextDirection = 2;
                }
            }

            //文本折行策略
            let cellStyleWrapStrategy: UniverJS.WrapStrategy = 0
            const csaw = cell.style.alignment?.wrapText
            if (csaw) {
                if (csaw) {
                    cellStyleWrapStrategy = 3;
                }
            }

            //水平对齐方式
            let cellStyleHorizontalAlign: UniverJS.HorizontalAlign = 0
            const csah = cell.style.alignment?.horizontal
            if (csah) {
                if (csah === 'left') {
                    cellStyleHorizontalAlign = 1;
                } else if (csah === 'center') {
                    cellStyleHorizontalAlign = 2;
                } else if (csah === 'right') {
                    cellStyleHorizontalAlign = 3;
                } else if (csah === 'justify') {
                    cellStyleHorizontalAlign = 4;
                }
                //文本折行策略的截断
                if (csah === 'fill') {
                    cellStyleWrapStrategy = 2
                }
            }

            //垂直对齐方式
            let cellStyleVerticalAlign: UniverJS.VerticalAlign = 0
            const csav = cell.style.alignment?.vertical
            if (csav) {
                if (csav === 'top') {
                    cellStyleVerticalAlign = 1;
                } else if (csav === 'middle') {
                    cellStyleVerticalAlign = 2;
                } else if (csav === 'bottom') {
                    cellStyleVerticalAlign = 3;
                }
            }

            //单元格填充（上下左右）
            // let cellStylePaddingData: UniverJS.IPaddingData = {}

            cellStyle.tr = cellStyleTextRotation
            cellStyle.td = cellStyleTextDirection
            cellStyle.ht = cellStyleHorizontalAlign
            cellStyle.vt = cellStyleVerticalAlign
            cellStyle.tb = cellStyleWrapStrategy
            // cellStyle.pd = cellStylePaddingData
            cellStyle.ul = cellStyleITextDecoration
            cellStyle.st = cellStyleStrikeITextDecoration
            // cellStyle.ol
            cellStyle.bg = cellStyleBackground
            cellStyle.bd = cellStyleBorder
            cellStyle.cl = cellStyleForeground
            cellStyle.va = cellStyleBaselineOffset
            cellData[rowIndex - 1][colIndex - 1].s = cellStyle
        }
    }

    const sheetData: UniverJS.IWorksheetData = {
        /**
       * Id of the worksheet. This should be unique and immutable across the lifecycle of the worksheet.
       */
        id: sheetId,
        /** Name of the sheet. */
        name: name,
        tabColor: 'white',
        /**
         * Determine whether the sheet is hidden.
         *
         * @remarks
         * See {@link BooleanNumber| the BooleanNumber enum} for more details.
         *
         * @defaultValue `BooleanNumber.FALSE`
         */
        hidden: hidden,
        freeze: freeze,
        rowCount: rowCount,
        columnCount: columnCount,
        zoomRatio: zoomRatio,
        scrollTop: scrollTop,
        scrollLeft: scrollLeft,
        defaultColumnWidth: defaultColumnWidth,
        defaultRowHeight: defaultRowHeight,
        /** All merged cells in this worksheet. */
        mergeData: mergeData,
        /** A matrix storing cell contents by row and column index. */
        cellData: cellData,
        rowData: rowData,
        columnData: columnData,
        rowHeader: {
            width: rowHeaderWidth,
            hidden: hidden,
        },
        columnHeader: {
            height: columnHeaderHeight,
            hidden: hidden,
        },
        showGridlines: showGridlines,
        /** @deprecated */
        selections: selections,
        rightToLeft: rightToLeft,
    };
    return sheetData;
};

export const importExcel = (univerAPI: FUniver.FUniver) => {

    const univerWorkbook = univerAPI.getActiveWorkbook()
    if (!univerWorkbook) {
        console.log('univerWorkbook is null')
        return
    }

    waitUserSelectExcelFile((workbook: ExcelJS.Workbook) => {
        const unid = univerWorkbook.getId()
        univerAPI.disposeUnit(unid)
        let sheetInfos: UniverJS.IWorksheetData[] = [];
        workbook.eachSheet((worksheet: any) => {
            const sheetInfo: UniverJS.IWorksheetData = parseExcelUniverSheetInfo(worksheet);
            sheetInfos = [...sheetInfos, sheetInfo];
        });
        let workbookInfo: UniverJS.IWorkbookData = {
            id: 'someUniqueId',
            rev: 1,
            name: 'Workbook Name',
            appVersion: '1.0.0',
            locale: UniverJS.LocaleType.ZH_CN,
            styles: {},
            sheetOrder: sheetInfos.map((sheet) => sheet.id),
            sheets: sheetInfos.reduce((acc, sheetInfo) => {
                acc[sheetInfo.id] = sheetInfo;
                return acc;
            }, {} as { [sheetId: string]: Partial<UniverJS.IWorksheetData> }),
            resources: []
        };
        univerAPI.createUniverSheet(workbookInfo)
    }, (error) => {
        console.log('Error selecting or processing file:', error);
    })

}
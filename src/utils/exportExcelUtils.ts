import * as UniverJS from "@univerjs/core";
import * as ExcelJS from 'exceljs';
import { borderMap } from '../utils/map'
import { DEFAULT_BORDER_COLOR } from '../utils/enum';
import * as FUniver from "@univerjs/facade";

//格式化时间
const formatDate = (date: Date) => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

//下载文件
const download_file = (buffer: ExcelJS.Buffer, fileName: string) => {
    let fileURL = window.URL.createObjectURL(new Blob([buffer]));
    let fileLink = document.createElement("a");
    fileLink.href = fileURL;
    fileLink.setAttribute("download", fileName);
    document.body.appendChild(fileLink);
    fileLink.click();
}

//下面是导出的函数
export const exportExcel = async (univerAPI: FUniver.FUniver | null) => {

    if (!univerAPI) return
    const univerWorkbook = univerAPI.getActiveWorkbook()
    if (!univerWorkbook) return;
    const workbook = new ExcelJS.Workbook();
    const sheets = univerWorkbook.getSheets()

    const sheetMap = univerWorkbook.getSnapshot().sheets

    const createColor = (rgb: UniverJS.Nullable<string>): ExcelJS.Color => {
        const colorValue = rgb ? 'FF' + rgb.slice(-6) : 'FFFFFFFF'; // 默认颜色是白色，如果rgb为null或undefined
        return {
            argb: colorValue,
            theme: -1 // 非正确主题
        };
    };

    sheets.forEach((univerSheet: FUniver.FWorksheet) => {
        const sheetId = univerSheet.getSheetId()
        const excelSheet = workbook.addWorksheet(sheetId);

        //合并单元格部分
        const mergeDataList = sheetMap[sheetId].mergeData

        for (let row = 0; row < 20 + 1; row++) {
            for (let col = 0; col < 20 + 1; col++) {
                const range = univerSheet.getRange(row, col)
                const univerCell = range?.getCellData();
                const excelCell = excelSheet.getCell(row + 1, col + 1);

                // console.log(row + 1, col + 1)
                // console.log(univerCell)
                if (univerCell) {
                    if (univerCell?.v) {
                        excelCell.value = univerCell.v
                    } else {
                        excelCell.value = '';
                    }
                    if (univerCell?.s) {
                        if (typeof univerCell.s !== 'string' && univerCell.s !== null && univerCell.s !== undefined) {
                            const styleData: UniverJS.IStyleData = univerCell.s;
                            //字体
                            if (styleData?.ff) {
                                //underline
                                let underline: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' = false
                                if (styleData.ul?.s === 1) {
                                    underline = true
                                    if (styleData.ul.t === 12) {
                                        underline = 'single'
                                    } else if (styleData.ul.t === 10) {
                                        underline = 'double'
                                    } else if (styleData.ul.t === 14) {
                                        underline = 'singleAccounting'
                                    } else if (styleData.ul.t === 15) {
                                        underline = 'doubleAccounting'
                                    }
                                }
                                //vertAlign
                                if (styleData?.va === 2) {
                                    excelCell.font = {
                                        vertAlign: 'subscript'
                                    }
                                } else if (styleData?.va === 3) {
                                    excelCell.font = {
                                        vertAlign: 'superscript'
                                    }
                                }

                                const fontColor: Partial<ExcelJS.Color> = createColor(styleData.cl?.rgb)

                                excelCell.font = {
                                    ...excelCell.font,
                                    name: styleData.ff,//fontFamily
                                    size: styleData.fs || 11,//fontSize
                                    bold: styleData.bl === 1,//bold
                                    italic: styleData.it === 1,//italic
                                    underline,//underline
                                    strike: styleData.st?.s === 1,//strikethrough
                                    color: fontColor,
                                };
                            }

                            //对齐方式
                            // horizontal: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
                            let horizontal: ExcelJS.Alignment['horizontal'] | undefined = undefined;
                            if (styleData?.ht === 1) {
                                horizontal = 'left'
                            } else if (styleData?.ht === 2) {
                                horizontal = 'center'
                            } else if (styleData?.ht === 3) {
                                horizontal = 'right'
                            } else if (styleData?.ht === 4) {
                                horizontal = 'justify'
                            }

                            // vertical: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
                            let vertical: ExcelJS.Alignment['vertical'] | undefined = undefined;
                            if (styleData?.ht === 1) {
                                vertical = 'top'
                            } else if (styleData?.ht === 2) {
                                vertical = 'middle'
                            } else if (styleData?.ht === 3) {
                                vertical = 'bottom'
                            }

                            // wrapText: boolean;
                            let wrapText = false
                            if (styleData?.tb === 3) {
                                wrapText = true
                            }

                            // readingOrder: 'rtl' | 'ltr';
                            let readingOrder: ExcelJS.Alignment['readingOrder'] | undefined = undefined;
                            if (styleData?.td === 1) {
                                readingOrder = 'ltr'
                            } else if (styleData?.td === 2) {
                                readingOrder = 'rtl'
                            }

                            // textRotation: number | 'vertical';
                            let textRotation: ExcelJS.Alignment['textRotation'] | undefined = undefined;
                            if (styleData?.tr) {
                                if (styleData?.tr.a) {
                                    textRotation = -styleData?.tr.a
                                }
                                if (styleData?.tr.v === 1) {
                                    textRotation = 'vertical'
                                }
                            }

                            const alignment: Partial<ExcelJS.Alignment> = {
                                horizontal: horizontal,
                                vertical: vertical,
                                wrapText: wrapText,
                                // shrinkToFit: ,
                                // indent: ,
                                readingOrder: readingOrder,
                                textRotation: textRotation,
                            }
                            excelCell.alignment = alignment

                            //边框
                            let top: ExcelJS.Borders['top'] | undefined = undefined;
                            if (styleData?.bd?.t) {
                                if (styleData.bd.t.cl.rgb !== DEFAULT_BORDER_COLOR || styleData.bd.t.s !== 1) {
                                    top = {
                                        style: borderMap[styleData.bd.t.s],
                                        color: createColor(styleData.bd.t.cl.rgb),
                                    }
                                }
                            }
                            let left: ExcelJS.Borders['left'] | undefined = undefined;
                            if (styleData?.bd?.l) {
                                if (styleData.bd.l.cl.rgb !== DEFAULT_BORDER_COLOR || styleData.bd.l.s !== 1) {
                                    left = {
                                        style: borderMap[styleData.bd.l.s],
                                        color: createColor(styleData.bd.l.cl.rgb),
                                    }
                                }
                            }
                            let right: ExcelJS.Borders['right'] | undefined = undefined;
                            if (styleData?.bd?.r) {
                                if (styleData.bd.r.cl.rgb !== DEFAULT_BORDER_COLOR || styleData.bd.r.s !== 1) {
                                    right = {
                                        style: borderMap[styleData.bd.r.s],
                                        color: createColor(styleData.bd.r.cl.rgb),
                                    }
                                }
                            }
                            let bottom: ExcelJS.Borders['bottom'] | undefined = undefined;
                            if (styleData?.bd?.b) {
                                if (styleData.bd.b.cl.rgb !== DEFAULT_BORDER_COLOR || styleData.bd.b.s !== 1) {
                                    bottom = {
                                        style: borderMap[styleData.bd.b.s],
                                        color: createColor(styleData.bd.b.cl.rgb),
                                    }
                                }
                            }

                            let diagonal: ExcelJS.BorderDiagonal | undefined = undefined;
                            if (styleData?.bd?.tl_br || styleData?.bd?.bl_tr) {
                                let diagonalStyle: Partial<ExcelJS.BorderDiagonal> = {};
                                if (styleData.bd.tl_br) {
                                    diagonalStyle = {
                                        style: borderMap[styleData.bd.tl_br.s],
                                        color: createColor(styleData.bd.tl_br.cl.rgb),
                                        up: false,
                                        down: true,
                                    };
                                }
                                if (styleData.bd.bl_tr) {
                                    diagonalStyle = {
                                        style: borderMap[styleData.bd.bl_tr.s],
                                        color: createColor(styleData.bd.bl_tr.cl.rgb),
                                        up: true,
                                        down: false,
                                    };
                                }
                                if (styleData.bd.tl_br && styleData.bd.bl_tr) {
                                    diagonalStyle = {
                                        style: borderMap[styleData.bd.bl_tr.s],
                                        color: createColor(styleData.bd.bl_tr.cl.rgb),
                                        up: true,
                                        down: true,
                                    };
                                }
                                diagonal = diagonalStyle as ExcelJS.BorderDiagonal;
                            }

                            const border: Partial<ExcelJS.Borders> = {
                                top: top,
                                left: left,
                                right: right,
                                bottom: bottom,
                                diagonal: diagonal,
                            }
                            excelCell.border = border;
                        }
                    }
                }
            }
        }

        mergeDataList?.forEach((mergeData: any) => {
            const mergeCell: ExcelJS.Location = {
                top: mergeData.startRow + 1,
                left: mergeData.startColumn + 1,
                bottom: mergeData.endRow + 1,
                right: mergeData.endColumn + 1,
            };
            excelSheet.mergeCells(mergeCell)
        })
    })
    // 写入文件
    const buffer = await workbook.xlsx.writeBuffer();
    //下载文件
    download_file(buffer, formatDate(new Date()) + "导出excel.xlsx");
}

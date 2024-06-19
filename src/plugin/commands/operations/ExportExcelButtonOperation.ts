import type { ICommand } from '@univerjs/core';
import { CommandType } from '@univerjs/core';
import type { IAccessor } from '@wendellhu/redi';
import * as UniverJS from "@univerjs/core";
import * as ExcelJS from 'exceljs';

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
const excelExport = async (univerWorkbook: any) => {
    const workbook = new ExcelJS.Workbook();
    const sheetMap = univerWorkbook.getWorksheets()
    sheetMap.forEach((sheet: any) => {
        const worksheet = workbook.addWorksheet(sheet.getName());
        for (let row = 0; row < sheet.getRowCount(); row++) {
            const rowData: any[] = [];
            for (let col = 0; col < sheet.getColumnCount(); col++) {
                const cell = sheet.getCell(row, col);
                if (cell) {
                    rowData[col] = cell.v
                }
            }
            worksheet.addRow(rowData);
        }
    })

    // 写入文件
    const buffer = await workbook.xlsx.writeBuffer();
    //下载文件
    download_file(buffer, formatDate(new Date()) + "导出excel.xlsx");
}

export const ExportExcelButtonOperation: ICommand = {
    id: 'custom-menu.operation.ExportExcel',
    type: CommandType.OPERATION,
    handler: async (accessor: IAccessor) => {
        const univer = accessor.get(UniverJS.IUniverInstanceService);
        const univerWorkbook = univer.getCurrentUnitForType<UniverJS.Workbook>(UniverJS.UniverInstanceType.UNIVER_SHEET)
        excelExport(univerWorkbook);
        return true;
    },
};
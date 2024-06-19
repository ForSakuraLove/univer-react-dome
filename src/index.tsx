import '@univerjs/design/lib/index.css';
import "@univerjs/ui/lib/index.css";
import "@univerjs/docs-ui/lib/index.css";
import "@univerjs/sheets-ui/lib/index.css";
import "@univerjs/sheets-formula/lib/index.css";

import { Univer, LocaleType, UniverInstanceType, Tools } from '@univerjs/core';
import { defaultTheme } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import React, { forwardRef, useEffect, useRef, useState } from 'react';
import { UniverSheetsCustomMenuPlugin } from "./plugin";
import DesignZhCN from '@univerjs/design/locale/zh-CN';
import UIZhCN from '@univerjs/ui/locale/zh-CN';
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN';
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN';
import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN';
// import ImportExcelButtonPlugin from '../../plugins/ImportExcelButton';
// import ExportExcelButtonPlugin from '../../plugins/ExportExcelButton';

// eslint-disable-next-line react/display-name
const UniverSheet = forwardRef(() => {
    const univerRef = useRef(null);
    const workbookRef = useRef(null);
    const containerRef = useRef(null);
    const [univeData, setUniveData] = useState({});

    const handleImportExcel: any = (data: any) => {
        setUniveData(data);
    }

    // ImportExcelButtonPlugin.setOnImportExcelCallback(handleImportExcel);

    /**
     * Initialize univer instance and workbook instance
     * @param data {IWorkbookData} document see https://univer.work/api/core/interfaces/IWorkbookData.html
     */
    const init = () => {
        if (!containerRef.current) {
            throw Error('container not initialized');
        }
        const univer = new Univer({
            theme: defaultTheme,
            locale: LocaleType.ZH_CN,
            locales: {
                [LocaleType.ZH_CN]: Tools.deepMerge(
                    SheetsZhCN,
                    DocsUIZhCN,
                    SheetsUIZhCN,
                    UIZhCN,
                    DesignZhCN,
                ),
            },
        });
        univerRef.current = univer;

        // core plugins
        univer.registerPlugin(UniverRenderEnginePlugin);
        univer.registerPlugin(UniverFormulaEnginePlugin);
        univer.registerPlugin(UniverUIPlugin, {
            container: containerRef.current,
        });

        // doc plugins
        univer.registerPlugin(UniverDocsPlugin, {
            hasScroll: false,
        });
        univer.registerPlugin(UniverDocsUIPlugin);

        // sheet plugins
        univer.registerPlugin(UniverSheetsPlugin);
        univer.registerPlugin(UniverSheetsUIPlugin);
        univer.registerPlugin(UniverSheetsFormulaPlugin);

        // custom plugins
        univer.registerPlugin(UniverSheetsCustomMenuPlugin);
        // univer.registerPlugin(ImportExcelButtonPlugin);
        // univer.registerPlugin(ExportExcelButtonPlugin);

        // create workbook instance
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, univeData);
    };

    /**
     * Destroy univer instance and workbook instance
     */
    const destroyUniver = () => {
        univerRef.current = null;
        workbookRef.current = null;
    };

    useEffect(() => {
        init();
        return () => {
            destroyUniver();
        };
    }, [univeData]);

    return (
        <div ref={containerRef} className="univer-container" style={{ height: '100%' }} />
    );

});

export default UniverSheet;

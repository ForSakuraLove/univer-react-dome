import '@univerjs/design/lib/index.css';
import './index.css';

import { Univer, LocaleType, UniverInstanceType } from '@univerjs/core';
import { defaultTheme } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import React, { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { zhCN, enUS } from 'univer:locales'
import ImportExcelButtonPlugin from '../../plugins/ImportExcelButton';
import ExportExcelButtonPlugin from '../../plugins/ExportExcelButton';

// eslint-disable-next-line react/display-name
const UniverSheet = forwardRef(({ data }, ref) => {
    const univerRef = useRef(null);
    const workbookRef = useRef(null);
    const containerRef = useRef(null);
    const [univeData, setUniveData] = useState({});

    useImperativeHandle(ref, () => ({
        getData,
    }));

    const handleImportExcel: any = (data: any) => {
        console.log(data);
        // 更新组件状态以触发刷新
        // setUniveData(data);
    }

    /**
     * Initialize univer instance and workbook instance
     * @param data {IWorkbookData} document see https://univer.work/api/core/interfaces/IWorkbookData.html
     */
    const init = (data = {}) => {
        if (!containerRef.current) {
            throw Error('container not initialized');
        }
        const univer = new Univer({
            theme: defaultTheme,
            locale: LocaleType.EN_US,
            locales: {
                [LocaleType.ZH_CN]: zhCN,
                [LocaleType.EN_US]: enUS,
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
        univer.registerPlugin(ImportExcelButtonPlugin, handleImportExcel);
        univer.registerPlugin(ExportExcelButtonPlugin);

        // create workbook instance
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, data);
    };

    /**
     * Destroy univer instance and workbook instance
     */
    const destroyUniver = () => {
        // univerRef.current?.dispose();
        univerRef.current = null;
        workbookRef.current = null;
    };

    /**
     * Get workbook data
     */
    const getData = () => {
        if (!workbookRef.current) {
            throw new Error('Workbook is not initialized');
        }
        return workbookRef.current.save();
    };

    useEffect(() => {
        init(data);
        return () => {
            destroyUniver();
        };
    }, [data]);

    return <div ref={containerRef} className="univer-container" />;
});

export default UniverSheet;

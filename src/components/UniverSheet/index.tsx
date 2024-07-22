import { Univer, LocaleType, UniverInstanceType, Tools } from '@univerjs/core';
import { FUniver } from "@univerjs/facade";
import { defaultTheme } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import React, { forwardRef, useEffect, useRef, useImperativeHandle } from 'react';
import DesignZhCN from '@univerjs/design/locale/zh-CN';
import UIZhCN from '@univerjs/ui/locale/zh-CN';
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN';
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN';
import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN';
import './index.sass'

import '@univerjs/design/lib/index.css';
import "@univerjs/ui/lib/index.css";
import "@univerjs/docs-ui/lib/index.css";
import "@univerjs/sheets-ui/lib/index.css";
import "@univerjs/sheets-formula/lib/index.css";

const UniverSheet = forwardRef(({ data }: any, ref) => {
    const univerRef = useRef(null);
    const containerRef = useRef(null);
    const fUniverRef = useRef(null);

    useImperativeHandle(ref, () => ({
        univerAPI: fUniverRef,
    }));

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

        univer.createUnit(UniverInstanceType.UNIVER_SHEET, data);

        const univerAPI = FUniver.newAPI(univer);
        const workBook = univerAPI.getActiveWorkbook()
        if (!workBook) return
        // setTimeout(() => {
        //     console.log(false)
        //     workBook.setEditable(false)
        // }, 1000)
        fUniverRef.current = univerAPI
    };

    const destroyUniver = () => {
        univerRef.current = null;
    };

    useEffect(() => {
        init();
        destroyUniver();
    }, []);

    return <div ref={containerRef} className="univer-container" />;

});

export default UniverSheet;

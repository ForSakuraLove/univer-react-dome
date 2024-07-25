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
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverDrawingPlugin } from '@univerjs/drawing';
import { UniverDrawingUIPlugin } from '@univerjs/drawing-ui';
import { UniverSheetsDrawingPlugin } from '@univerjs/sheets-drawing';
import { UniverSheetsDrawingUIPlugin } from '@univerjs/sheets-drawing-ui';
import { UniverFindReplacePlugin } from '@univerjs/find-replace';
import { UniverSheetsFindReplacePlugin } from '@univerjs/sheets-find-replace';
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter';
import { UniverSheetsFilterUIPlugin } from '@univerjs/sheets-filter-ui';
import { UniverSheetsSortPlugin } from '@univerjs/sheets-sort';
import { UniverSheetsSortUIPlugin } from '@univerjs/sheets-sort-ui';
import { UniverDataValidationPlugin } from '@univerjs/data-validation';
import { UniverSheetsDataValidationPlugin } from '@univerjs/sheets-data-validation';
import { UniverSheetsConditionalFormattingPlugin } from '@univerjs/sheets-conditional-formatting';
import { UniverSheetsConditionalFormattingUIPlugin } from '@univerjs/sheets-conditional-formatting-ui';
import { IThreadCommentMentionDataService, UniverThreadCommentUIPlugin } from '@univerjs/thread-comment-ui';
import { UniverThreadCommentPlugin } from '@univerjs/thread-comment';
import { UniverSheetsThreadCommentBasePlugin } from '@univerjs/sheets-thread-comment-base';
import { UniverSheetsThreadCommentPlugin } from '@univerjs/sheets-thread-comment';
import { UniverSheetsZenEditorPlugin } from '@univerjs/sheets-zen-editor';
import { UniverSheetsHyperLinkPlugin } from '@univerjs/sheets-hyper-link';
import { UniverSheetsHyperLinkUIPlugin } from '@univerjs/sheets-hyper-link-ui';
import { UniverUniscriptPlugin } from '@univerjs/uniscript';
import React, { forwardRef, useEffect, useRef, useImperativeHandle } from 'react';
import DesignZhCN from '@univerjs/design/locale/zh-CN';
import UIZhCN from '@univerjs/ui/locale/zh-CN';
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN';
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN';
import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN';
import SheetsFormulaZhCN from '@univerjs/sheets-formula/locale/zh-CN';
import SheetsNumfmtZhCN from '@univerjs/sheets-numfmt/locale/zh-CN';
import DrawingUIZhCN from '@univerjs/drawing-ui/locale/zh-CN';
import SheetsDrawingUIZhCN from '@univerjs/sheets-drawing-ui/locale/zh-CN';
import FindReplaceZhCN from '@univerjs/find-replace/locale/zh-CN';
import SheetsFindReplaceZhCN from '@univerjs/sheets-find-replace/locale/zh-CN';
import SheetsFilterUIZhCN from '@univerjs/sheets-filter-ui/locale/zh-CN';
import SheetsSortUIZhCN from '@univerjs/sheets-sort-ui/locale/zh-CN';
import SheetsDataValidationZhCN from '@univerjs/sheets-data-validation/locale/zh-CN';
import SheetsConditionalFormattingUIZhCN from '@univerjs/sheets-conditional-formatting-ui/locale/zh-CN';
import ThreadCommentUIZhCN from '@univerjs/thread-comment-ui/locale/zh-CN';
import SheetsThreadCommentZhCN from '@univerjs/sheets-thread-comment/locale/zh-CN';
import SheetsZenEditorZhCN from '@univerjs/sheets-zen-editor/locale/zh-CN';
import SheetsHyperLinkUIZhCN from '@univerjs/sheets-hyper-link-ui/locale/zh-CN';
import UniscriptZhCN from '@univerjs/uniscript/locale/zh-CN';
import './index.sass'

import '@univerjs/design/lib/index.css';
import "@univerjs/ui/lib/index.css";
import "@univerjs/docs-ui/lib/index.css";
import "@univerjs/sheets-ui/lib/index.css";
import '@univerjs/sheets-numfmt/lib/index.css';
import '@univerjs/drawing-ui/lib/index.css';
import '@univerjs/sheets-drawing-ui/lib/index.css';
import '@univerjs/find-replace/lib/index.css';
import '@univerjs/sheets-filter-ui/lib/index.css';
import '@univerjs/sheets-sort-ui/lib/index.css';
import '@univerjs/sheets-data-validation/lib/index.css';
import '@univerjs/sheets-conditional-formatting-ui/lib/index.css';
import '@univerjs/thread-comment-ui/lib/index.css';
import '@univerjs/sheets-zen-editor/lib/index.css';
import "@univerjs/sheets-formula/lib/index.css";
import '@univerjs/sheets-hyper-link-ui/lib/index.css';
import '@univerjs/uniscript/lib/index.css';

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
                    SheetsFormulaZhCN,
                    SheetsNumfmtZhCN,
                    DrawingUIZhCN,
                    SheetsDrawingUIZhCN,
                    FindReplaceZhCN,
                    SheetsFindReplaceZhCN,
                    SheetsFilterUIZhCN,
                    SheetsSortUIZhCN,
                    SheetsDataValidationZhCN,
                    SheetsConditionalFormattingUIZhCN,
                    ThreadCommentUIZhCN,
                    SheetsThreadCommentZhCN,
                    SheetsZenEditorZhCN,
                    SheetsHyperLinkUIZhCN,
                    UniscriptZhCN,
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

        univer.registerPlugin(UniverSheetsNumfmtPlugin);
        univer.registerPlugin(UniverDrawingPlugin);
        univer.registerPlugin(UniverDrawingUIPlugin);
        univer.registerPlugin(UniverSheetsDrawingPlugin);
        univer.registerPlugin(UniverSheetsDrawingUIPlugin);
        univer.registerPlugin(UniverFindReplacePlugin);
        univer.registerPlugin(UniverSheetsFindReplacePlugin);
        univer.registerPlugin(UniverSheetsFilterPlugin);
        univer.registerPlugin(UniverSheetsFilterUIPlugin);
        univer.registerPlugin(UniverSheetsSortPlugin);
        univer.registerPlugin(UniverSheetsSortUIPlugin);
        univer.registerPlugin(UniverDataValidationPlugin);
        univer.registerPlugin(UniverSheetsDataValidationPlugin);
        univer.registerPlugin(UniverSheetsConditionalFormattingPlugin);
        univer.registerPlugin(UniverSheetsConditionalFormattingUIPlugin);
        const mockUser = {
            userID: 'mockId',
            name: 'MockUser',
            avatar: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAInSURBVHgBtZU9TxtBEIbfWRzFSIdkikhBSqRQkJqkCKTCFkqVInSUSaT0wC8w/gXxD4gU2nRJkXQWhAZowDUUWKIwEgWWbEEB3mVmx3dn4DA2nB/ppNuPeWd29mMIPXDr+RxwtgRHeW6+guNPRxogqnL7Dwz9psJ27S4NShaeZTH3kwXy6I81dlRKcmRui88swdq9AcSFL7Buz1Vmlns64MiLsCjzwnIYHLH57tbfFbs7KRaXyEU8FVZofqccOfA5l7Q8LPIkGrwnb2RPNEXWFVMUF3L+kDCk0btDDAMzOm5YfAHDwp4tG74wnzAsiOYMnJ3GoDybA7IT98/jm5+JNnfiIzAS6LlqHQBN/i6b2t/cV1Hh6BfwYlHnHP4AXi5q/8kmMMpOs8+BixZw/Fd6xUEHEbnkgclvQP2fGp7uShRKnQ3G32rkjV1th8JhIGG7tR/JyjGteSOZELwGMmNqIIigRCLRh2OZIE6BjItdd7pCW6Uhm1zzkUtungSxwEUzNpQ+GQumtH1ej1MqgmNT6vwmhCq5yuwq56EYTbgeQUz3yvrpV1b4ok3nYJ+eYhgYmjRUqErx2EDq0Fr8FhG++iqVGqxlUJI/70Ar0UgJaWHj6hYVHJrfKssAHot1JfqwE9WVWzXZVd5z2Ws/4PnmtEjkXeKJDvxUecLbWOXH/DP6QQ4J72NS0adedp1aseBfXP8odlZFfPvBF7SN/8hky1TYuPOAXAEipMx15u5ToAAAAABJRU5ErkJggg==',
            anonymous: false,
            canBindAnonymous: false,
        };

        class CustomMentionDataService implements IThreadCommentMentionDataService {
            dataSource: any
            trigger: string = '@';

            async getMentions(search: string) {
                return [
                    {
                        id: mockUser.userID,
                        label: mockUser.name,
                        type: 'user',
                        icon: mockUser.avatar,
                    },
                    {
                        id: '2',
                        label: 'User2',
                        type: 'user',
                        icon: mockUser.avatar,
                    },
                ];
            }
        }
        univer.registerPlugin(UniverThreadCommentPlugin);
        univer.registerPlugin(UniverThreadCommentUIPlugin, {
            overrides: [[IThreadCommentMentionDataService, { useClass: CustomMentionDataService }]],
        });
        univer.registerPlugin(UniverSheetsThreadCommentBasePlugin);
        univer.registerPlugin(UniverSheetsThreadCommentPlugin);
        univer.registerPlugin(UniverSheetsZenEditorPlugin);
        univer.registerPlugin(UniverSheetsHyperLinkPlugin);
        univer.registerPlugin(UniverSheetsHyperLinkUIPlugin);
        univer.registerPlugin(UniverUniscriptPlugin);

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

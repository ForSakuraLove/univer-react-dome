import { useRef, useEffect, useState, } from 'react';
import { NzButton } from 'nuzar-design'
import UniverSheet from './components/UniverSheet/index';
import { importExcel } from './components/utils/importExcelUtils';
import { exportExcel } from './components/utils/exportExcelUtils';
import { useIntl, useLocation, } from '@umijs/max';
import './index.sass'

const UniverComponent: React.FC = () => {
    const { formatMessage } = useIntl();
    const location = useLocation();
    const univerRef = useRef();
    const [univerAPI, setUniverAPI] = useState<any>(null)
    const workbookData = location.state?.workbookData || {}
    const [data, setData] = useState<any>();

    useEffect(() => {
        setUniverAPI(univerRef.current.univerAPI.current)
    }, [])

    useEffect(() => {
        setData(workbookData)
    }, [workbookData,])

    // useEffect(() => {
    //     if (univerAPI) {
    //         const univerWorkbook = univerAPI.getActiveWorkbook()
    //         if (!univerWorkbook) {
    //             console.log('univerWorkbook is null')
    //             return
    //         }
    //         console.log(workbookData)
    //         const unid = univerWorkbook.getId()
    //         univerAPI.disposeUnit(unid)
    //         setTimeout(() => {
    //             univerAPI.createUniverSheet(workbookData)
    //         }, 500)
    //     }
    // }, [workbookData,])

    return (
        <div className='nz_univer'>
            <div className='header'>
                <NzButton onClick={async () => {
                    try {
                        const workbookData = await importExcel(univerAPI);
                        setData(workbookData);
                    } catch (error) {
                        console.error('Error importing Excel:', error);
                    }
                }}>{formatMessage({ id: 'univer.button.import.excel' })}</NzButton>
                <NzButton onClick={() => exportExcel(univerAPI)}>{formatMessage({ id: 'univer.button.export.excel' })}</NzButton>
            </div>
            <div className='body'>
                {/* <UniverSheet ref={univerRef} /> */}
                <UniverSheet ref={univerRef} data={data} />
            </div>
        </div>
    );
}

export default UniverComponent;

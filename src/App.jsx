import { useRef, useEffect, useState, } from 'react';
import UniverSheet from './components/UniverSheet/index'
import { importExcel } from './components/utils/importExcelUtils';
import { exportExcel } from './components/utils/exportExcelUtils';
// import { useIntl, } from '@umijs/max';
import './index.sass'

const UniverComponent = () => {
    // const { formatMessage } = useIntl();
    const univerRef = useRef();
    const [univerAPI, setUniverAPI] = useState(null)

    useEffect(() => {
        setUniverAPI(univerRef.current.univerAPI.current)
    }, [])

    return (
        <div className='nz_univer'>
            <div className='header'>
                <button onClick={() => importExcel(univerAPI)}>import</button>
                <button onClick={() => exportExcel(univerAPI)}>export</button>
            </div>
            <div className='body'>
                <UniverSheet ref={univerRef} />
            </div>
        </div>
    );
}

export default UniverComponent;

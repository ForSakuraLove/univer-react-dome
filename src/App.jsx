import { useRef, useEffect, useState, } from 'react';
import UniverSheet from './index.tsx';
import { importExcel } from './utils/importExcelUtils.ts';
import { exportExcel } from './utils/exportExcelUtils.ts';

const UniverComponent = () => {
    const univerRef = useRef();
    const [univerAPI, setUniverAPI] = useState(null)

    useEffect(() => {
        setUniverAPI(univerRef.current.univerAPI.current)
    }, [])

    return (
        <div id="root">
            <button onClick={() => importExcel(univerAPI)}>import</button>
            <button onClick={() => exportExcel(univerAPI)}>export</button>
            <UniverSheet ref={univerRef} />
        </div>
    );
}

export default UniverComponent;

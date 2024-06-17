import { useRef, useState } from 'react';
import UniverSheet from './components/UniverSheet/index.tsx';
import { DEFAULT_WORKBOOK_DATA } from './assets/default-workbook-data';

function App() {
    const [data] = useState(DEFAULT_WORKBOOK_DATA);
    const univerRef = useRef();

    return (
        <div id="root">
            <UniverSheet style={{ flex: 1 }} ref={univerRef} data={data} />
        </div>
    );
}

export default App;

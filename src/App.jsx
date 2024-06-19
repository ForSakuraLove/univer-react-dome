import { useRef, } from 'react';
import UniverSheet from './index.tsx';

function App() {
    const univerRef = useRef();

    return (
        <div id="root">
            <UniverSheet style={{ flex: 1 }} ref={univerRef} />
        </div>
    );
}

export default App;

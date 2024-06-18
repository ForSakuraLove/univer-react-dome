import { useRef, } from 'react';
import UniverSheet from './components/UniverSheet/index.tsx';

function App() {
    const univerRef = useRef();

    return (
        <div id="root">
            <UniverSheet style={{ flex: 1 }} ref={univerRef} />
        </div>
    );
}

export default App;

// import React, {useState} from 'react';
// import 'bootstrap/dist/css/bootstrap.min.css';
// import { Button, Form, Container } from 'react-bootstrap';

// interface ICheckedFilterDropDownProps {
//     options: Array<{value: string, label: string}>,
//     onChange?: (selected: Array<{value: string, label: string}>) => void
// }

// const CheckedFilterDropDown: React.FC<ICheckedFilterDropDownProps> = ({options, onChange}) => {
//     const [selected, setSelected] = useState<Array<{value: string, label: string}>>([]);
//     const [showList, setShowList] = useState(false);

//     const handleSelect = (option: {value: string, label: string}) => {
//         let selt=[...selected];
//         let newsel;
//         const isSelected = selt.some(item => item.value === option.value);
//         if (isSelected) {
//             newsel=selt.filter(item => item.value !== option.value);
//             setSelected(newsel);
//             if(onChange) onChange(newsel);

//         } else {
//             newsel=[...selt, option];
//             setSelected(newsel);
//             if(onChange) onChange(newsel);
//         }

//     }

//     const handleShowList = () => {
//         setShowList(!showList);
//     }

//     const handleClear = () => {
//         setSelected([]);
//         if(onChange) onChange([]);

//     }

//     const handleApply = () => {
//         if (onChange) {
//             onChange(selected);
//         }
//         setShowList(false);
//     }

//     return (
//         <>
//             <>
//                 <div className="mb-0">
//                     {/* <Form.Label>Filter By</Form.Label> */}
//                     <div className="d-flex flex-wrap">
//                         {selected.map((item, index) => (
//                             <Button key={index} variant="outline-primary" className="me-1 mb-1 p-0" size="sm">
//                                 {item.label}
//                                 <span className="ms-1" onClick={() => handleSelect(item)}>&times;</span>
//                             </Button>
//                         ))}
//                     </div>
//                     <Button variant="link" style={{textAlign:'left'}} onClick={handleShowList} className='p-0 mb-2'>Show List</Button>
//                     {showList && (
//                         <div className="position-absolute bg-white p-2 border shadow-sm" style={{zIndex:1000}}>
//                             <div className="d-flex justify-content-between mb-2">
//                                 <span>Selected: {selected.length}</span>
//                                 <Button variant="link" size="sm" onClick={handleClear}>Clear</Button>
//                             </div>
//                             <ul className="list-group">
//                                 {options.map((option, index) => (
//                                     <li key={index} className="list-group-item d-flex justify-content-between">
//                                         <span>{option.label}</span>
//                                         <input type="checkbox" checked={selected.some(item => item.value === option.value)} onChange={() => handleSelect(option)}/>
//                                     </li>
//                                 ))}
//                             </ul>
//                             {/* <div className="d-flex justify-content-end mt-2">
//                                 <Button variant="primary" onClick={handleApply}>Apply</Button>
//                             </div> */}
//                         </div>
//                     )}
//                 </div>
//             </>
//         </>
//     );
// }

// export default CheckedFilterDropDown;
import React, {useState} from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import { Button, Form, Container } from 'react-bootstrap';

interface ICheckedFilterDropDownProps {
    options: Array<{value: string, label: string}>,
    onChange?: (selected: Array<{value: string, label: string}>) => void ,
    Index?: number
}

const CheckedFilterDropDown: React.FC<ICheckedFilterDropDownProps> = ({options, onChange , Index}) => {
    console.log("CheckedFilterDropDown Index",Index);
    const [selected, setSelected] = useState<Array<{value: string, label: string}>>([]);
    const [showList, setShowList] = useState(false);

    const handleSelect = (option: {value: string, label: string}) => {
        let selt=[...selected];
        let newsel;
        const isSelected = selt.some(item => item.value === option.value);
        if (isSelected) {
            newsel=selt.filter(item => item.value !== option.value);
            setSelected(newsel);
            if(onChange) onChange(newsel);

        } else {
            newsel=[...selt, option];
            setSelected(newsel);
            if(onChange) onChange(newsel);
        }

    }

    const handleShowList = () => {
        setShowList(!showList);
    }

    const handleClear = () => {
        setSelected([]);
        if(onChange) onChange([]);

    }

    const handleApply = () => {
        if (onChange) {
            onChange(selected);
        }
        setShowList(false);
    }

    return (
        <>
            <>
                <div className="mb-0">
                    {/* <Form.Label>Filter By</Form.Label> */}
                    <div className="d-flex flex-wrap">
                        {selected.map((item, index) => (
                            <Button key={index} variant="outline-primary" className="me-1 mb-1 p-0" size="sm">
                                {item.label}
                                <span className="ms-1" onClick={() => handleSelect(item)}>&times;</span>
                            </Button>
                        ))}
                    </div>
                    <Button variant="link" style={{textAlign:'left'}} onClick={handleShowList} className='p-0 mb-2'>Show List</Button>
                    {showList && (
                        <div className="position-absolute bg-white p-2 border shadow-sm" style={{ zIndex: 1000 - (Index * 100) }}>
                            <div className="d-flex justify-content-between mb-2">
                                <span>Selected: {selected.length}</span>
                                <Button variant="link" size="sm" onClick={handleClear}>Clear</Button>
                            </div>
                            <ul className="list-group">
                                {options.map((option, index) => (
                                    <li key={index} className="list-group-item d-flex justify-content-between">
                                        <span>{option.label}</span>
                                        <input type="checkbox" checked={selected.some(item => item.value === option.value)} onChange={() => handleSelect(option)}/>
                                    </li>
                                ))}
                            </ul>
                            {/* <div className="d-flex justify-content-end mt-2">
                                <Button variant="primary" onClick={handleApply}>Apply</Button>
                            </div> */}
                        </div>
                    )}
                </div>
            </>
        </>
    );
}

export default CheckedFilterDropDown;

import * as React from 'react';
import { Routes, Route } from 'react-router-dom';

import Dashboard from '../Forms/Dashboard';
import CategoryMaster from '../Forms/CategoryMaster';
import NavigationBar from '../Navigation/NavigationBar';
import DeparmentMaster from '../Forms/DepartmentMaster';
import ItemMaster from '../Forms/ItemMaster';
import SupplyRequestList from '../Forms/SupplyRequestList';
import SupplyRequestForm from '../Forms/SupplyRequestForm';

interface IRoutesProps {
  context: any;
}

const RoutesItems: React.FC<IRoutesProps> = ({ context }) => {
  return (
    <>
      <NavigationBar />

      <Routes>
        <Route path="/" element={<Dashboard />} />

        <Route
          path="/category-master"
          element={<CategoryMaster context={context} />}
        />
        <Route path="/department-master" element={<DeparmentMaster context={context}/>}/>
        
        <Route path="/item-master" element={<ItemMaster context={context} />}/>
        
        <Route path="/SupplyRequestList" element={<SupplyRequestList context={context}/>}/>
        <Route path="/SupplyRequestForm" element={<SupplyRequestForm context={context} />} />  

 
    

      </Routes>
    </>
  );
};
export default RoutesItems;
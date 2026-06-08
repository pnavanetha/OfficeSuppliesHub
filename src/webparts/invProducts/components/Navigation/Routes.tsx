import * as React from 'react';
import { Routes, Route } from 'react-router-dom';

import Dashboard from '../Forms/Dashboard';
import CategoryMaster from '../Forms/CategoryMaster';
// import NavigationBar from './NavigationBar';
import NavigationBar from '../Navigation/NavigationBar';

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
      </Routes>
    </>
  );
};
export default RoutesItems;
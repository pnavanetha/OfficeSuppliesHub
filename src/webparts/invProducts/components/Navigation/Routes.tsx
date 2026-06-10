// import * as React from 'react';
// import { Routes, Route, Navigate } from 'react-router-dom';

// import AdminDashboard from '../Forms/AdminDashboard';
// import StaffDashboard from '../Forms/StaffDashboard';
// import CategoryMaster from '../Forms/CategoryMaster';
// import NavigationBar from '../Navigation/NavigationBar';
// import DeparmentMaster from '../Forms/DepartmentMaster';
// import ItemMaster from '../Forms/ItemMaster';
// import SupplyRequestList from '../Forms/SupplyRequestList';
// import SupplyRequestForm from '../Forms/SupplyRequestForm';
// import AllRequest from '../Forms/AllRequests';
// import PendingRequest from '../Forms/PendingRequest';
// import MyRequest from '../Forms/MyRequests';

// interface IRoutesProps {
//   context: any;
//   role: 'Admin' | 'Staff';
// }

// const RoutesItems: React.FC<IRoutesProps> = ({ role, context }) => {
//   return (
//     <>
//       <NavigationBar role={role} />

//       <Routes>

//         <Route path="/" element={role === 'Admin' ? <AdminDashboard /> : <StaffDashboard />} />

//         <Route path="/supply-request-list" element={<SupplyRequestList context={context} role={role} />} />
//         <Route
//           path="/SupplyRequestForm"
//           element={<SupplyRequestForm context={context} role={role} />}
//         />

//         <Route
//           path="/SupplyRequestForm/:id"
//           element={<SupplyRequestForm context={context} role={role} />}
//         />

//         <Route path="/myrequest" element={<MyRequest context={context} role={role}  />}/>


//         {role === 'Admin' && (
//           <>
//             <Route path="/category-master" element={<CategoryMaster context={context} />} />
//             <Route path="/department-master" element={<DeparmentMaster context={context} />} />
//             <Route path="/item-master" element={<ItemMaster context={context} />} />
//             <Route path="/pendingrequest" element={<PendingRequest context={context} />} />
//             <Route path="/allrequest" element={<AllRequest context={context} />} />
//           </>
//         )}
//         <Route path="*" element={<Navigate to="/" />} />
//       </Routes>
//     </>
//   );
// };

// export default RoutesItems;
import * as React from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';

import AdminDashboard from '../Forms/AdminDashboard';
import StaffDashboard from '../Forms/StaffDashboard';
import CategoryMaster from '../Forms/CategoryMaster';
import NavigationBar from '../Navigation/NavigationBar';
import DeparmentMaster from '../Forms/DepartmentMaster';
import ItemMaster from '../Forms/ItemMaster';
import SupplyRequestList from '../Forms/SupplyRequestList';
import SupplyRequestForm from '../Forms/SupplyRequestForm';
import AllRequest from '../Forms/AllRequests';
import PendingRequest from '../Forms/PendingRequest';
import MyRequest from '../Forms/MyRequests';

interface IRoutesProps {
  context: any;
  role: 'Admin' | 'Staff';
}

const RoutesItems: React.FC<IRoutesProps> = ({ role, context }) => {
  return (
    <>
      <NavigationBar role={role} />

      <Routes>

        {/* ✅ Dashboard */}
        <Route
          path="/"
          element={role === 'Admin' ? <AdminDashboard /> : <StaffDashboard />}
        />

        {/* ✅ Common Routes */}
        <Route
          path="/supply-request-list"
          element={<SupplyRequestList context={context} role={role} />}
        />

        <Route
          path="/SupplyRequestForm"
          element={<SupplyRequestForm context={context} role={role} />}
        />

        <Route
          path="/SupplyRequestForm/:id"
          element={<SupplyRequestForm context={context} role={role} />}
        />

        <Route
          path="/my-requests"
          element={<MyRequest context={context} role={role} />}
        />

        {/* ✅ Admin Dashboard Navigation */}
        {role === 'Admin' && (
          <>
            <Route
              path="/all-requests"
              element={<AllRequest context={context} />}
            />

            <Route
              path="/pending-requests"
              element={<PendingRequest context={context} />}
            />
          </>
        )}

        {/* ✅ Masters (Admin Only) */}
        {role === 'Admin' && (
          <>
            <Route
              path="/category-master"
              element={<CategoryMaster context={context} />}
            />

            <Route
              path="/department-master"
              element={<DeparmentMaster context={context} />}
            />

            <Route
              path="/item-master"
              element={<ItemMaster context={context} />}
            />
          </>
        )}

        {/* ✅ Fallback */}
        <Route path="*" element={<Navigate to="/" />} />

      </Routes>
    </>
  );
};

export default RoutesItems;
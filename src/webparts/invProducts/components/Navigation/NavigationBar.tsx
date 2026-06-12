// import * as React from "react";
// import { Link } from "react-router-dom";
// import {
//   FaTachometerAlt,
//   FaClipboardList,
//   FaFolder,
//   FaBuilding,
//   FaBox
// } from "react-icons/fa";

// import "../CSS/navigationBar.css";

// interface INavigationBarProps {
//   role: "Admin" | "Staff";
// }

// const NavigationBar: React.FC<INavigationBarProps> = ({ role }) => {
//   return (
//     <div className="sidebar">
//       <div className="logo">
//         <h2>Office Hub</h2>
//       </div>

//       <ul className="menu">
//         <li>
//           <Link to="/">
//             <FaTachometerAlt className="icon" />
//             Dashboard
//           </Link>
//         </li>

//         <li>
//           <Link to="/supply-request-list">
//             <FaClipboardList className="icon" />
//             Supply Requests
//           </Link>
//         </li>

//         {role === "Admin" && (
//           <>
//             <li>
//               <Link to="/category-master">
//                 <FaFolder className="icon" />
//                 Category Master
//               </Link>
//             </li>

//             <li>
//               <Link to="/department-master">
//                 <FaBuilding className="icon" />
//                 Department Master
//               </Link>
//             </li>

//             <li>
//               <Link to="/item-master">
//                 <FaBox className="icon" />
//                 Item Master
//               </Link>
//             </li>
//           </>
//         )}
//       </ul>
//     </div>
//   );
// };

// export default NavigationBar;

import * as React from "react";
import { useState } from "react";
import { Link } from "react-router-dom";

import {
  FiChevronLeft,
  FiChevronRight,
  FiHome,
  FiClipboard,
  FiFolder,
  FiBox,
  FiGrid
} from "react-icons/fi";

import { FaBuilding } from "react-icons/fa";

import "../CSS/navigationBar.css";

interface INavigationBarProps {
  role: "Admin" | "Staff";
}

const NavigationBar: React.FC<INavigationBarProps> = ({ role }) => {
  const [collapsed, setCollapsed] = useState(false);

  return (
    <div className={`sidebar ${collapsed ? "collapsed" : ""}`}>
      
      <button
        className="toggle-btn"
        onClick={() => setCollapsed(!collapsed)}
      >
        {collapsed ? <FiChevronRight /> : <FiChevronLeft />}
      </button>

      <div className="logo">
        <FiGrid className="logo-icon" />

        {!collapsed && (
          <div>
            <h2>Office Hub</h2>
          </div>
        )}
      </div>

      <ul className="menu">

        <li>
          <Link to="/">
            <FiHome className="icon" />
            <span className="menu-text">Dashboard</span>
          </Link>
        </li>

        <li>
          <Link to="/supply-request-list">
            <FiClipboard className="icon" />
            <span className="menu-text">Supply Requests</span>
          </Link>
        </li>

        {role === "Admin" && (
          <>
            <li>
              <Link to="/category-master">
                <FiFolder className="icon" />
                <span className="menu-text">Category Master</span>
              </Link>
            </li>

            <li>
              <Link to="/department-master">
                <FaBuilding className="icon" />
                <span className="menu-text">Department Master</span>
              </Link>
            </li>

            <li>
              <Link to="/item-master">
                <FiBox className="icon" />
                <span className="menu-text">Item Master</span>
              </Link>
            </li>
          </>
        )}

      </ul>
    </div>
  );
};

export default NavigationBar;
// import * as React from "react";
// import { Link } from "react-router-dom";

// interface INavigationBarProps {
//   role: "Admin" | "Staff";
// }

// const NavigationBar: React.FC<INavigationBarProps> = ({ role }) => {
//   return (
//     <div
//       style={{
//         padding: "10px",
//         backgroundColor: "#f3f3f3",
//         marginBottom: "20px"
//       }}
//     >
//       <Link to="/" style={{ marginRight: "20px" }}>
//         Dashboard
//       </Link>
//       <Link to="/supply-request-list">Supply Request List</Link>

//       {role === "Admin" && (
//         <>
//           <Link to="/category-master">Category Master</Link>
//           <Link to="/department-master">Department Master</Link>
//           <Link to="/item-master">Item Master</Link>
//         </>
//       )}
//     </div>

//   );
// };

// export default NavigationBar;
import * as React from "react";
import { Link } from "react-router-dom";
import "../CSS/navigationBar.css";

interface INavigationBarProps {
  role: "Admin" | "Staff";
}

const NavigationBar: React.FC<INavigationBarProps> = ({ role }) => {
  return (
    <div className="sidebar">
      <div className="logo">
        <h2>Office Hub</h2>
      </div>

      <ul className="menu">
        <li>
          <Link to="/">🏠 Dashboard</Link>
        </li>

        <li>
          <Link to="/supply-request-list">📋 Supply Requests</Link>
        </li>

        {role === "Admin" && (
          <>
            <li>
              <Link to="/category-master">📂 Category Master</Link>
            </li>

            <li>
              <Link to="/department-master">🏢 Department Master</Link>
            </li>

            <li>
              <Link to="/item-master">📦 Item Master</Link>
            </li>
          </>
        )}
      </ul>
    </div>
  );
};

export default NavigationBar;
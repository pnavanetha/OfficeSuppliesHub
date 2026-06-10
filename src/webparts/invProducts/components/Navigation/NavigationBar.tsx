import * as React from "react";
import { Link } from "react-router-dom";

interface INavigationBarProps {
  role: "Admin" | "Staff";
}

const NavigationBar: React.FC<INavigationBarProps> = ({ role }) => {
  return (
    <div
      style={{
        padding: "10px",
        backgroundColor: "#f3f3f3",
        marginBottom: "20px"
      }}
    >
      <Link to="/" style={{ marginRight: "20px" }}>
        Dashboard
      </Link>
      <Link to="/SupplyRequestList">Supply Request List</Link>

      {role === "Admin" && (
        <>
          <Link to="/category-master">Category Master</Link>
          <Link to="/department-master">Department Master</Link>
          <Link to="/item-master">Item Master</Link>
        </>
      )}
    </div>

  );
};

export default NavigationBar;
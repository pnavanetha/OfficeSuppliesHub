import * as React from "react";
import { useNavigate } from "react-router-dom";

const AdminDashboard = () => {

  const navigate = useNavigate();

  return (
    <div style={{ padding: "20px" }}>
      <h2>Admin Dashboard</h2>

      <div style={{ display: "flex", gap: "20px", marginTop: "20px" }}>

        <div
          style={cardStyle}
          onClick={() => navigate("/all-requests")}
        >
          All Requests
        </div>

        <div
          style={cardStyle}
          onClick={() => navigate("/pending-requests")}
        >
          Pending Requests
        </div>

        <div
          style={cardStyle}
          onClick={() => navigate("/my-requests")}
        >
          My Requests
        </div>

      </div>
    </div>
  );
};

const cardStyle: React.CSSProperties = {
  padding: "20px",
  border: "1px solid #ccc",
  cursor: "pointer",
  backgroundColor: "#f9f9f9",
  minWidth: "150px",
  textAlign: "center"
};

export default AdminDashboard;
import * as React from "react";
import { useNavigate } from "react-router-dom";

const StaffDashboard = () => {

  const navigate = useNavigate();

  return (
    <div style={{ padding: "20px" }}>
      <h2>Employee Dashboard</h2>

      <div style={{ marginTop: "20px" }}>
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
  width: "200px",
  textAlign: "center"
};

export default StaffDashboard;
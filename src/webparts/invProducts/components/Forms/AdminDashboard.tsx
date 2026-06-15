import { SPFx, spfi } from "@pnp/sp";
import * as React from "react";
import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import "../CSS/Dashboard.css";

const AdminDashboard = (props: any) => {

  const [tableHeaders, setTableHeaders] = useState<string[]>([]);
  const [tableData, setTableData] = useState<string[]>([]);


  const sp = spfi().using(SPFx(props.context));
  const navigate = useNavigate();

  useEffect(() => {
    loadData("All");
  }, []);

  const loadData = async (action: string) => {
    try {
      const headers = ["ID", "Employee", "Department", "Category", "Item", "Request", "Comments", "Status", "Edit"];
      setTableHeaders(headers);

      let filterQuery = '';

      if (action === "Pending") {
        filterQuery = "Status eq 'Submitted'";
      } else if (action === "My") {
        filterQuery = `Author/Id eq ${props.context.pageContext.legacyPageContext.userId}`;
      }

      let items = sp.web.lists.getByTitle("OfficeSupplyRequestList").items;

      if (filterQuery) {
        items = items.filter(filterQuery);
      }

      const res: any = await items
        .select(
          "Id", "RequestDate", "Comments", "Status", "EmployeeName/Title", "EmployeeName/EMail", "Department/Name", "CategoryName/CategoryName", "ItemName/ItemName"
        )
        .expand("EmployeeName", "Department", "CategoryName", "ItemName")();

      console.log(res);

      const data = res.map((item: any) => ({
        ID: item.Id,
        Employee: item.EmployeeName?.Title,
        Department: item.Department?.Name,
        Category: item.CategoryName?.CategoryName,
        Item: item.ItemName?.ItemName,
        Request: item.RequestDate,
        Comments: item.Comments,
        Status: item.Status,
        // Edit: "Edit"
      }));

      setTableData(data);


    } catch (error) {
      console.error("Error loading data:", error);
    }
  };
  
  const editRequest = (id: number): void => {navigate(`/SupplyRequestForm/${id}`);
  };
  return (
    <div className="dashboard-content">
      <div className="header-card">
        <div className="page-title">Admin Dashboard</div>
      </div>
      {/* 
      <h2>Admin Dashboard</h2> */}

      <div className="dashboard-cards">

        <div className="dashboard-card" onClick={() => loadData("All")} >All Requests </div>
        <div className="dashboard-card" onClick={() => loadData("Pending")}>Pending Requests</div>
        <div className="dashboard-card" onClick={() => loadData("My")}> My Requests</div>

      </div>

      <table className="dashboard-table">
        <thead>
          <tr>
            {
              tableHeaders.map((header) => {
                return (
                  <th>{header}</th>
                )
              })
            }
          </tr>
        </thead>

        <tbody>
          {tableData.length > 0 ? (
            tableData.map((item: any) => {
              return (
                <tr>
                  {Object.keys(item).map((field: any) => {
                    return (
                      <td>{item[field]}</td>
                    )
                  })}
                  <td>
                    <button className="edit-btn" onClick={() => editRequest(item.ID)}>
                      Edit
                    </button>
                  </td>
                </tr>
              )
            })
          ) : (
            <tr>
              <td colSpan={9} style={{ textAlign: "center" }}>
                No Records Found
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
};

// const cardStyle: React.CSSProperties = {
//   padding: "20px",
//   border: "1px solid #ccc",
//   cursor: "pointer",
//   backgroundColor: "#f9f9f9",
//   minWidth: "150px",
//   textAlign: "center"
// };

export default AdminDashboard;
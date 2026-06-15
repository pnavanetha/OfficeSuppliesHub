import { SPFx, spfi } from "@pnp/sp";
import * as React from "react";
import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import "../CSS/Dashboard.css";

const StaffDashboard = (props: any) => {

  const [tableHeaders, setTableHeaders] = useState<string[]>([]);
  const [tableData, setTableData] = useState<string[]>([]);

  const sp = spfi().using(SPFx(props.context));
  const navigate = useNavigate();

  useEffect(() => {
    loadData();
  }, []);

const loadData = async () => {
  try {
    const headers = ["ID", "Employee", "Department", "Category", "Item", "Request", "Comments", "Status", "Edit"];
    setTableHeaders(headers);

    const userId = props.context.pageContext.legacyPageContext.userId; 

    console.log("User ID:", userId);

   const res: any = await sp.web.lists
  .getByTitle("OfficeSupplyRequestList").items.select("Id", "RequestDate", "Comments", "Status", "EmployeeName/Id", "EmployeeName/Title", "Department/Name", "CategoryName/CategoryName", "ItemName/ItemName")
  .expand("EmployeeName", "Department", "CategoryName", "ItemName")
  .filter(`EmployeeName/Id eq ${userId}`)();

    console.log("Filtered Data:", res);

    const data = res.map((item: any) => ({
      ID: item.Id,
      Employee: item.EmployeeName?.Title,
      Department: item.Department?.Name,
      Category: item.CategoryName?.CategoryName,
      Item: item.ItemName?.ItemName,
      Request: item.RequestDate,
      Comments: item.Comments,
      Status: item.Status,
    }));

    setTableData(data);

  } catch (error) {
    console.error("Error loading data:", error);
  }
};

  const editRequest = (id: number): void => {
    navigate(`/SupplyRequestForm/${id}`);
  };
  return (
    <div className="dashboard-content">
      <div className="header-card">
        <div className="page-title">Staff Dashboard</div>
      </div>


      <div className="dashboard-cards">

        {/* <div className="dashboard-card" onClick={() => loadData("All")} >All Requests </div>
        <div className="dashboard-card" onClick={() => loadData("Pending")}>Pending Requests</div> */}
        <div className="dashboard-card" onClick={loadData}>My Requests</div>

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



export default StaffDashboard;
import { SPFx, spfi } from "@pnp/sp";
import * as React from "react";
import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import "../CSS/Dashboard.css";
import { formatDateIN } from "../Common/dateHelpers";
import CommonGrid from "../Common/CommonGrid";
import { IColumn } from "@fluentui/react";
import { FiEdit } from "react-icons/fi";

const AdminDashboard = (props: any) => {

  // const [tableHeaders, setTableHeaders] = useState<string[]>([]);
  const [tableData, setTableData] = useState<any[]>([]);


  const sp = spfi().using(SPFx(props.context));
  const navigate = useNavigate();

  useEffect(() => {
    loadData("All");
  }, []);

  const loadData = async (action: string) => {
    try {
      // const headers = ["Edit", "ID", "Employee", "Department", "Category", "Item", "Request", "Comments", "Status", ];
      // setTableHeaders(headers);
      // const headers = ["Edit", "ID", "Employee", "Department", "Category", "Item", "Request", "Comments", "Status",];
      // (headers);

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
        Request: formatDateIN(item.RequestDate),
        Comments: item.Comments,
        Status: item.Status,
        // Edit: "Edit"
      }));

      setTableData(data);


    } catch (error) {
      console.error("Error loading data:", error);
    }
  };

  const editRequest = (id: number): void => {
    navigate(`/SupplyRequestForm/${id}`);
  };

const columns: IColumn[] = [
  {
    key: "Edit",
    name: "Edit",
    minWidth: 70,
    onRender: (item: any) => (
      <FiEdit className="edit-btn" onClick={() => editRequest(item.ID)} />
      // >
      //   Edit
      // </button>
    )
  },
  { key: "ID", name: "ID", fieldName: "ID", minWidth: 60, isResizable: true},
  { key: "Employee", name: "Employee Name", fieldName: "Employee", minWidth: 150, isResizable: true},
  { key: "Department", name: "Department",  fieldName: "Department",minWidth: 120,isResizable: true},
  { key: "Category", name: "Category", fieldName: "Category", minWidth: 120, isResizable: true},
  { key: "Item", name: "Item", fieldName: "Item", minWidth: 120, isResizable: true},
  { key: "Request", name: "Request Date", fieldName: "Request", minWidth: 120, isResizable: true},
  { key: "Comments",name: "Comments", fieldName: "Comments", minWidth: 150, isResizable: true},
  { key: "Status", name: "Status", fieldName: "Status", minWidth: 100, isResizable: true}
];
return (
  <div className="dashboard-content">

    <div className="dashboard-header">
      <h2>Admin Dashboard</h2>
    </div>

    <div className="dashboard-cards">
      <div className="dashboard-card" onClick={() => loadData("All")}>
        {/* <span className="card-count">{tableData.length}</span> */}
        <span className="card-count">All Request</span>
        {/* <span className="card-title">All Requests</span> */}
      </div>

      <div className="dashboard-card" onClick={() => loadData("Pending")}>
        <span className="card-count">Pending Requests</span>
        {/* <span className="card-title">Pending Requests</span> */}
      </div>

      <div className="dashboard-card" onClick={() => loadData("My")}>
        <span className="card-count">My Requests</span>
        {/* <span className="card-title">My Requests</span> */}
      </div>
    </div>

    <div className="grid-container">
      <CommonGrid
        items={tableData}
        columns={columns}
        pageSize={10}
        searchFields={["ID", "Employee", "Department","Category", "Item", "Request", "Comments", "Status"]}
      />
    </div>

  </div>
);
};


export default AdminDashboard;
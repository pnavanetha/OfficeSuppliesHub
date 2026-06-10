import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { useNavigate } from "react-router-dom";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const MyRequest = (props: any) => {

  const sp = spfi().using(SPFx(props.context));
  const navigate = useNavigate();

  const [data, setData] = useState<any[]>([]);

  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    try {

      const userEmail = props.context.pageContext.user.email;

      const res = await sp.web.lists
        .getByTitle("OfficeSupplyRequestList")
        .items
        .select(
          "Id",
          "RequestDate",
          "Comments",
          "Status",
          "EmployeeName/Title",
          "EmployeeName/EMail",
          "Department/Name",
          "CategoryName/CategoryName",
          "ItemName/ItemName"
        )
        .expand("EmployeeName", "Department", "CategoryName", "ItemName")
        .filter(`EmployeeName/EMail eq '${userEmail}'`) // ✅ KEY FILTER
        ();

      setData(res);

    } catch (error) {
      console.log(error);
    }
  };

  // ✅ Navigate to Edit Form
  const editRequest = (id: number): void => {navigate(`/SupplyRequestForm/${id}`);
  };

  return (
    <div style={{ padding: "20px" }}>

      <h2>My Requests</h2>

      <table
        style={{
          width: "100%",
          borderCollapse: "collapse"
        }}
      >

        <thead>
          <tr>
            <th>ID</th>
            <th>Employee Name</th>
            <th>Department</th>
            <th>Category</th>
            <th>Item Name</th>
            <th>Request Date</th>
            <th>Comments</th>
            <th>Status</th>
            <th>Edit</th> {/* ✅ Only Edit */}
          </tr>
        </thead>

        <tbody>
          {data.length > 0 ? (
            data.map((item: any) => (
              <tr key={item.Id}>
                <td>{item.Id}</td>
                <td>{item.EmployeeName?.Title}</td>
                <td>{item.Department?.Name}</td>
                <td>{item.CategoryName?.CategoryName}</td>
                <td>{item.ItemName?.ItemName}</td>
                <td>{item.RequestDate}</td>
                <td>{item.Comments}</td>
                <td>{item.Status}</td>

                {/* ✅ ONLY EDIT BUTTON */}
                <td>
                  <button onClick={() => editRequest(item.Id)}>
                    Edit
                  </button>
                </td>

              </tr>
            ))
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

export default MyRequest;

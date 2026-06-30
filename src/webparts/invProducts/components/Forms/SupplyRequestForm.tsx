import * as React from "react";
import "../CSS/App.css";
import { useEffect, useState, useRef } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { useParams, useNavigate } from "react-router-dom";
import { showSuccess, showError } from "../Common/Toast";
// import { formatDateIN } from "../Common/dateHelpers";
import "../CSS/RequestForm.css";

interface IRequestData {
  DepartmentId: number;
  CategoryNameId: number;
  ItemNameId: number;
  RequestDate: string;
  Comments: string;
}

const SupplyRequestForm = (props: any) => {
  const { role } = props;

  const sp = spfi().using(SPFx(props.context));
  const { id } = useParams();
  const navigate = useNavigate();

  const [currentUser, setCurrentUser] = useState("");
  const [currentUserId, setCurrentUserId] = useState(0);

  const [departmentData, setDepartmentData] = useState<any[]>([]);
  const [categoryData, setCategoryData] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);

  const [status, setStatus] = useState<string>("Draft");


  const [formData, setFormData] = useState<IRequestData>({
    DepartmentId: 0,
    CategoryNameId: 0,
    ItemNameId: 0,
    RequestDate: new Date().toISOString().split("T")[0],
    // RequestDate: formatDateIN(new Date().toISOString()),
    Comments: ""
  });

  useEffect(() => {
    loadData();
  }, []);

  useEffect(() => {
    if (id && itemData.length > 0) {
      loadEditData(Number(id));
    }
  }, [id, itemData]);

  //   useEffect(() => {
  //   console.log("ROLE:", role, "STATUS:", status);
  // }, [role, status]);


  const loadData = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUser(user.Title);
      setCurrentUserId(user.Id);

      const [departments, categories, items] = await Promise.all([
        sp.web.lists.getByTitle("OfficeDepartments").items(),
        sp.web.lists.getByTitle("OfficeCategoryMaster").items(),
        sp.web.lists.getByTitle("OfficeItemMaster")
          .items.select("Id", "ItemName", "CategoryNameId")()
      ]);

      setDepartmentData(departments);
      setCategoryData(categories);
      setItemData(items);

    } catch (error) {
      console.log(error);
    }
  };

  const loadEditData = async (itemId: number) => {
    try {
      const item = await sp.web.lists
        .getByTitle("OfficeSupplyRequestList")
        .items.getById(itemId)
        .select("DepartmentId", "CategoryNameId", "ItemNameId", "RequestDate", "Comments", "Status")();

      setFormData({
        DepartmentId: item.DepartmentId,
        CategoryNameId: item.CategoryNameId,
        ItemNameId: item.ItemNameId,
        RequestDate: item.RequestDate.split("T")[0],
        Comments: item.Comments
      });

      setStatus(item.Status || "Draft");

      const filtered = itemData.filter(
        (i: any) => i.CategoryNameId === item.CategoryNameId
      );
      setFilteredItems(filtered);

    } catch (error) {
      console.log(error);
    }
  };


const handleChange = (e: any) => {
  const { name, value } = e.target;

  setFormData({
    ...formData,
    [name]:
      name === "DepartmentId" ||
      name === "CategoryNameId" ||
      name === "ItemNameId"
        ? Number(value)
        : value
  });
};


  const handleCategoryChange = (e: any) => {
    const categoryId = Number(e.target.value);

    const filtered = itemData.filter(
      (item: any) => item.CategoryNameId === categoryId
    );

    setFilteredItems(filtered);

    setFormData({
      ...formData,
      CategoryNameId: categoryId,
      ItemNameId: 0
    });
  };

  const focusWithError = (element: any) => {
    setTimeout(() => {
      if (element) {
        element.focus();
        element.classList.add("input-error");
      }
    }, 0);
  };
  const deptRef = useRef<HTMLSelectElement>(null);
  const categoryRef = useRef<HTMLSelectElement>(null);
  const itemRef = useRef<HTMLSelectElement>(null);

  const handleSave = async (newStatus: string) => {

    if (formData.DepartmentId === 0) {
      showError("Select Department");
      focusWithError(deptRef.current);
      return;
    }
    if (formData.CategoryNameId === 0) {
      showError("Select Category Name");
      focusWithError(categoryRef.current);
      return;
    }

    if (formData.ItemNameId === 0) {
      showError("Select Item");
      focusWithError(itemRef.current);
      return;
    }


    try {

      if (id) {
        await sp.web.lists
          .getByTitle("OfficeSupplyRequestList")
          .items.getById(Number(id))
          .update({
            DepartmentId: Number(formData.DepartmentId),
            CategoryNameId: Number(formData.CategoryNameId),
            ItemNameId: Number(formData.ItemNameId),
            Comments: formData.Comments,
            Status: newStatus
          });

        showSuccess("Updated Successfully");

      } else {
        await sp.web.lists
          .getByTitle("OfficeSupplyRequestList")
          .items.add({ EmployeeNameId: currentUserId, DepartmentId: Number(formData.DepartmentId), CategoryNameId: Number(formData.CategoryNameId), ItemNameId: Number(formData.ItemNameId), RequestDate: formData.RequestDate, Comments: formData.Comments, Status: newStatus });

        showSuccess("Saved Successfully");
      }

      navigate("/SupplyRequestList");

    } catch (error) {
      console.log(error);
    }
  };

  const handleApproveReject = async (newStatus: string) => {
    try {
      await sp.web.lists.getByTitle("OfficeSupplyRequestList").items.getById(Number(id)).update({ Status: newStatus });
      if (newStatus === "Approved") {
        showSuccess("Request is Approved.");        
      } else if(newStatus === "Rejected"){
        showError("Request is Rejected.");
        
      }

      navigate("/supply-request-list");

    } catch (error) {
      console.log(error);
      showError("Something went wrong");
    }
  };

  const handleCancel = () => {
    navigate("/supply-request-list");
  };

  return (
    // <div style={{ padding: "20px" }}>

    //   <h2>{id ? "Edit Request" : "New Request"}</h2>

    //   <label>Employee Name</label><br />
    //   <input type="text" value={currentUser} readOnly /><br /><br />

    //   <label>Department</label><br />
    //   <select name="DepartmentId" value={formData.DepartmentId} onChange={handleChange} ref={deptRef}>
    //     <option value={0}>Select Department</option>
    //     {departmentData.map((item) => (
    //       <option key={item.Id} value={item.Id}>{item.Name}</option>
    //     ))}
    //   </select><br /><br />

    //   <label>Category</label><br />
    //   <select name="CategoryNameId" value={formData.CategoryNameId} onChange={handleCategoryChange} ref={categoryRef}>
    //     <option value={0}>Select Category</option>
    //     {categoryData.map((item) => (
    //       <option key={item.Id} value={item.Id}>{item.CategoryName}</option>
    //     ))}
    //   </select><br /><br />

    //   <label>Item Name</label><br />
    //   <select name="ItemNameId" value={formData.ItemNameId} onChange={handleChange} ref={itemRef}>
    //     <option value={0}>Select Item</option>
    //     {filteredItems.map((item) => (
    //       <option key={item.Id} value={item.Id}>{item.ItemName}</option>
    //     ))}
    //   </select><br /><br />

    //   <label>Request Date</label><br />
    //   <input type="date" value={formData.RequestDate} readOnly /><br /><br />

    //   <label>Comments</label><br />
    //   <textarea
    //     rows={4}
    //     cols={40}
    //     name="Comments"
    //     value={formData.Comments}
    //     onChange={handleChange}
    //   /><br /><br />


    //   {status === "Draft" && (
    //     <>
    //       <button onClick={() => handleSave("Draft")}>Save</button>

    //       <button
    //         onClick={() => handleSave("Submitted")}
    //         style={{ marginLeft: "10px" }}
    //       >
    //         Submit
    //       </button>
    //     </>
    //   )}


    //   {status === "Submitted" && role === "Staff" && (
    //     <button onClick={() => handleSave("Submitted")}>
    //       Update
    //     </button>
    //   )}

    //   {status === "Submitted" && role === "Admin" && (
    //     <>
    //       <button onClick={() => handleApproveReject("Approved")}>
    //         Approve
    //       </button>

    //       <button
    //         onClick={() => handleApproveReject("Rejected")}
    //         style={{ marginLeft: "10px" }}
    //       >
    //         Reject
    //       </button>
    //     </>
    //   )}


    //   {status === "Rejected" && (
    //     <button onClick={() => handleSave("Submitted")}>
    //       Resubmit
    //     </button>
    //   )}


    //   <button onClick={handleCancel} style={{ marginLeft: "10px" }}>
    //     Cancel
    //   </button>


    // </div>
    <div className="request-container">

      <div className="request-header">
        <h2>{id ? "Edit Supply Request" : "New Supply Request"}</h2>
      </div>

      <div className="request-body">

        {/* Employee */}
        <div className="employee-box">
          <label>Employee Name</label>
          <input type="text" value={currentUser} readOnly />
        </div>
        <div className="form-row">
          <div>
            <label>Department <span className="required">*</span></label>
            <select name="DepartmentId" value={formData.DepartmentId} onChange={handleChange} ref={deptRef}>
              <option value={0}>Select Department</option>
              {departmentData
              .sort((a,b) => a.Name.localeCompare(b.Name))
              .map((item) => (
                <option key={item.Id} value={item.Id}>{item.Name}</option>
              ))}
            </select>
          </div>
          <div>
            <label>Category <span className="required">*</span></label>

            <select name="CategoryNameId" value={formData.CategoryNameId} onChange={handleCategoryChange} ref={categoryRef}>
              <option value={0}>Select Category</option>
              {categoryData
                .map((item) => (
                <option key={item.Id} value={item.Id}>{item.CategoryName}</option>
              ))}
            </select>
          </div>

          <div>
            <label>Item Name <span className="required">*</span></label>

            <select name="ItemNameId" value={formData.ItemNameId} onChange={handleChange} ref={itemRef}>
              <option value={0}>Select Item</option>
              {filteredItems
              .map((item) => (
                <option key={item.Id} value={item.Id}>{item.ItemName}</option>
              ))}
            </select>
          </div>

        </div>

        <div className="form-row-two">
          <div>
            <label>Request Date</label>

            <input type="date" value={formData.RequestDate} readOnly />

          </div>

          <div>
            <label>Comments</label>

            <textarea
              rows={4}
              cols={40}
              name="Comments"
              value={formData.Comments}
              onChange={handleChange}
              placeholder="Add any notes or remarks..."
            />
          </div>

        </div>

        <div className="button-section">

          {status === "Draft" && (
            <>
              <button className="btn" onClick={handleCancel}>Cancel</button>

              <button className="btn btn-primary" onClick={() => handleSave("Draft")}>Save Draft</button>

              <button className="btn btn-success" onClick={() => handleSave("Submitted")}>Submit</button>
            </>
          )}
          {status === "Submitted" && role === "Staff" && (
            <>
              <button className="btn" onClick={handleCancel}>Cancel</button>

              <button className="btn btn-primary" onClick={() => handleSave("Submitted")}>Update</button>
            </>
          )}

          {status === "Submitted" && role === "Admin" && (
            <>
              <button className="btn" onClick={handleCancel}>Cancel</button>
              <button className="btn btn-success" onClick={() => handleApproveReject("Approved")}>Approve</button>
              <button className="btn btn-danger" onClick={() => handleApproveReject("Rejected")} >Reject</button>
            </>
          )}
          {status === "Rejected" && (
            <>
              <button className="btn" onClick={handleCancel}>Cancel</button>
              <button className="btn btn-warning" onClick={() => handleSave("Submitted")}>Resubmit</button>
            </>
          )}

          {/* Approved */}
          {status === "Approved" && (
            <button className="btn" onClick={handleCancel}>Close</button>
          )}

        </div>

      </div>

    </div>
  );
};

export default SupplyRequestForm;
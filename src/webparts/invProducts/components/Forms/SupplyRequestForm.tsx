import * as React from "react";
import { useEffect, useState, useRed } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { useParams, useNavigate } from "react-router-dom";
import {showSuccess, showError } from "../Common/Toast";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

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
      [name]: value
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

  const handleSave = async (newStatus: string) => {

    if (formData.DepartmentId === 0) return showError("Select Department");
    if (formData.CategoryNameId === 0) return showError("Select Category");
    if (formData.ItemNameId === 0) return showError("Select Item");

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
          .items.add({EmployeeNameId: currentUserId, DepartmentId: Number(formData.DepartmentId), CategoryNameId: Number(formData.CategoryNameId),ItemNameId: Number(formData.ItemNameId),RequestDate: formData.RequestDate, Comments: formData.Comments, Status: newStatus});

        showSuccess("Saved Successfully");
      }

      navigate("/SupplyRequestList");

    } catch (error) {
      console.log(error);
    }
  };

  const handleApproveReject = async (newStatus: string) => {
    try {
      await sp.web.lists .getByTitle("OfficeSupplyRequestList") .items.getById(Number(id)).update({ Status: newStatus });      
      alert(`Request ${newStatus}`);
      navigate("/supply-request-list");

    } catch (error) {
      console.log(error);
    }
  };

  const handleCancel = () => {
    navigate("/supply-request-list");
  };

  return (
    <div style={{ padding: "20px" }}>

      <h2>{id ? "Edit Request" : "New Request"}</h2>

      <label>Employee Name</label><br />
      <input type="text" value={currentUser} readOnly /><br /><br />

      <label>Department</label><br />
      <select name="DepartmentId" value={formData.DepartmentId} onChange={handleChange}>
        <option value={0}>Select Department</option>
        {departmentData.map((item) => (
          <option key={item.Id} value={item.Id}>{item.Name}</option>
        ))}
      </select><br /><br />

      <label>Category</label><br />
      <select name="CategoryNameId" value={formData.CategoryNameId} onChange={handleCategoryChange}>
        <option value={0}>Select Category</option>
        {categoryData.map((item) => (
          <option key={item.Id} value={item.Id}>{item.CategoryName}</option>
        ))}
      </select><br /><br />

      <label>Item Name</label><br />
      <select name="ItemNameId" value={formData.ItemNameId} onChange={handleChange}>
        <option value={0}>Select Item</option>
        {filteredItems.map((item) => (
          <option key={item.Id} value={item.Id}>{item.ItemName}</option>
        ))}
      </select><br /><br />

      <label>Request Date</label><br />
      <input type="date" value={formData.RequestDate} readOnly /><br /><br />

      <label>Comments</label><br />
      <textarea
        rows={4}
        cols={40}
        name="Comments"
        value={formData.Comments}
        onChange={handleChange}
      /><br /><br />

  
      {status === "Draft" && (
        <>
          <button onClick={() => handleSave("Draft")}>Save</button>

          <button
            onClick={() => handleSave("Submitted")}
            style={{ marginLeft: "10px" }}
          >
            Submit
          </button>
        </>
      )}

   
      {status === "Submitted" && role === "Staff" && (
        <button onClick={() => handleSave("Submitted")}>
          Update
        </button>
      )}

      {status === "Submitted" && role === "Admin" && (
        <>
          <button onClick={() => handleApproveReject("Approved")}>
            Approve
          </button>

          <button
            onClick={() => handleApproveReject("Rejected")}
            style={{ marginLeft: "10px" }}
          >
            Reject
          </button>
        </>
      )}


      {status === "Rejected" && (
        <button onClick={() => handleSave("Submitted")}>
          Resubmit
        </button>
      )}

   
      <button onClick={handleCancel} style={{ marginLeft: "10px" }}>
        Cancel
      </button>


    </div>
  );
};

export default SupplyRequestForm;
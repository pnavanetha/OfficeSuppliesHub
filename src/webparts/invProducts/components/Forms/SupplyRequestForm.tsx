import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";

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

  const sp = spfi().using(SPFx(props.context));

  const [currentUser, setCurrentUser] = useState("");
  const [currentUserId, setCurrentUserId] = useState(0);

  const [departmentData, setDepartmentData] = useState<any[]>([]);
  const [categoryData, setCategoryData] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);

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

  const loadData = async () => {
    try {

      const user = await sp.web.currentUser();

      setCurrentUser(user.Title);
      setCurrentUserId(user.Id);

      const [departments, categories, items] =
        await Promise.all([

          sp.web.lists
            .getByTitle("OfficeDepartments")
            .items(),

          sp.web.lists
            .getByTitle("OfficeCategoryMaster")
            .items(),

          sp.web.lists
            .getByTitle("OfficeItemMaster")
            .items
            .select(
              "Id",
              "ItemName",
              "CategoryNameId"
            )()
        ]);

      setDepartmentData(departments);
      setCategoryData(categories);
      setItemData(items);

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

  const handleSubmit = async () => {

    if (formData.DepartmentId === 0) {
      alert("Select Department");
      return;
    }

    if (formData.CategoryNameId === 0) {
      alert("Select Category");
      return;
    }

    if (formData.ItemNameId === 0) {
      alert("Select Item");
      return;
    }

    try {

      await sp.web.lists
        .getByTitle("OfficeSupplyRequestList")
        .items
        .add({

          EmployeeNameId: currentUserId,

          DepartmentId: Number(formData.DepartmentId),

          CategoryNameId: Number(formData.CategoryNameId),

          ItemNameId: Number(formData.ItemNameId),

          RequestDate: formData.RequestDate,

          Comments: formData.Comments,

          Status: "Pending"
        });

      alert("Request Submitted Successfully");

      resetForm();

    } catch (error) {
      console.log(error);
    }
  };

  const resetForm = () => {

    setFormData({
      DepartmentId: 0,
      CategoryNameId: 0,
      ItemNameId: 0,
      RequestDate: new Date().toISOString().split("T")[0],
      Comments: ""
    });

    setFilteredItems([]);
  };

  return (
    <div style={{ padding: "20px" }}>

      <h2>Office Supply Request</h2>

      <div>

        <label>Employee Name</label>
        <br />

        <input
          type="text"
          value={currentUser}
          readOnly
        />

        <br /><br />

        <label>Department</label>
        <br />

        <select
          name="DepartmentId"
          value={formData.DepartmentId}
          onChange={handleChange}
        >
          <option value={0}>Select Department</option>

          {
            departmentData.map((item) => (
              <option
                key={item.Id}
                value={item.Id}
              >
                {item.Name}
              </option>
            ))
          }
        </select>

        <br /><br />

        <label>Category</label>
        <br />

        <select
          name="CategoryNameId"
          value={formData.CategoryNameId}
          onChange={handleCategoryChange}
        >
          <option value={0}>Select Category</option>

          {
            categoryData.map((item) => (
              <option
                key={item.Id}
                value={item.Id}
              >
                {item.CategoryName}
              </option>
            ))
          }
        </select>

        <br /><br />

        <label>Item Name</label>
        <br />

        <select
          name="ItemNameId"
          value={formData.ItemNameId}
          onChange={handleChange}
        >
          <option value={0}>Select Item</option>

          {
            filteredItems.map((item) => (
              <option
                key={item.Id}
                value={item.Id}
              >
                {item.ItemName}
              </option>
            ))
          }
        </select>

        <br /><br />

        <label>Request Date</label>
        <br />

        <input
          type="date"
          value={formData.RequestDate}
          readOnly
        />

        <br /><br />

        <label>Comments</label>
        <br />

        <textarea
          rows={4}
          cols={40}
          name="Comments"
          value={formData.Comments}
          onChange={handleChange}
        />

        <br /><br />

        <button onClick={handleSubmit}>
          Submit
        </button>

        <button
          onClick={resetForm}
          style={{ marginLeft: "10px" }}
        >
          Clear
        </button>

      </div>

    </div>
  );
};

export default SupplyRequestForm;
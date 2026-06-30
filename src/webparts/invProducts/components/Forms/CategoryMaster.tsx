import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { FiEdit } from "react-icons/fi";
import { IColumn } from "@fluentui/react";
import CommonGrid from "../Common/CommonGrid";
import { showSuccess, showError } from "../Common/Toast";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "../CSS/CategoryMaster.css";


const CategoryMaster = (props: any) => {

  const listName = "OfficeCategoryMaster";

  const sp = spfi().using(SPFx(props.context));

  const [showForm, setShowForm] = useState(false);
  const [itemId, setItemId] = useState<number>(0);
  const [data, setData] = useState<any[]>([]);
  const [formData, setFormData] = useState({
    CategoryName: "", IsActive: true
  });

  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    try {
      const result = await sp.web.lists.getByTitle(listName).items();
      setData(result);
    } catch (error) {
      console.log(error);
    }
  };

  const handleChange = (e: any) => {
    const { name, value, type, checked } = e.target;

    setFormData({
      ...formData,
      [name]: type === "checkbox" ? checked : value
    });
  };

  const resetForm = () => {
    setItemId(0);
    setFormData({ CategoryName: "", IsActive: true });
  }; 

  const openAddForm = () => {
    resetForm();
    setShowForm(true);

    setTimeout(() => {
      inputRef.current?.focus();
    }, 100);

  };

  const editItem = (item: any) => {
    setItemId(item.Id);
    setFormData({
      CategoryName: item.CategoryName, IsActive: item.IsActive
    });
    setShowForm(true);
  };
  const inputRef = useRef<HTMLInputElement>(null);
  

  const handleSubmit = async () => {
    //setIsSubmitted(true);

    if (!formData.CategoryName.trim()) {
      showError("Category Name is required.");

      setTimeout(() => {
        inputRef.current?.focus();
        inputRef.current?.classList.add('input-error')
      }, 0);

      return;
    }


    try {

      if (itemId > 0) {
        await sp.web.lists.getByTitle(listName).items.getById(itemId).update(formData);
        showSuccess("Category Updated Successfully.");

      } else {
        await sp.web.lists.getByTitle(listName).items.add(formData);
        showSuccess("Category Submitted Successfully.");
      }

      resetForm();
      setShowForm(false);
      loadData();
   

    } catch (error) {
      console.log(Error);
      showError("Something went wrong");
    }
  };

  //  Columns only 
  const columns: IColumn[] = [
    {
      key: "edit", name: "Edit", minWidth: 60, maxWidth: 60,
      onRender: (item: any) => (
        <FiEdit className="edit-icon" onClick={() => editItem(item)} />
      )
    },
    {
      key: "category", name: "Category Name", fieldName: "CategoryName", minWidth: 250
    },
    {
      key: "active", name: "IsActive", fieldName: "IsActive", minWidth: 100,
      onRender: (item: any) =>
        item.IsActive ? "Yes" : "No"
    }
  ];


  return (
    <div className="master-container">

      <div className="header-card">


        <div className="header-top">
          <div className="page-title">Category Master</div>

          <div className="toolbar">
            {!showForm && (
              <button className="add-btn" onClick={openAddForm}>
                Add
              </button>
            )}
          </div>
        </div>


        {showForm && (
          <div className="form-section">
            <div className="form-group">
              <label>Category Name <span className="asterisk">*</span> </label>
              <input type="text" name="CategoryName" value={formData.CategoryName} onChange={handleChange} ref={inputRef}/>
            </div>

            <div className="checkbox-section">
              <label>
                <input type="checkbox" name="IsActive" checked={formData.IsActive} onChange={handleChange} />
                Is Active
              </label>
            </div>

            <div className="button-section">
              <button className="submit-btn" onClick={handleSubmit}>
                {itemId > 0 ? "Update" : "Submit"}
              </button>

              <button className="cancel-btn"
                onClick={() => {
                  resetForm();
                  setShowForm(false);
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>

      <div className="grid-card">
        <CommonGrid
          items={data}
          columns={columns}
          pageSize={10}
          searchFields={["CategoryName", "Active"]}
        />
      </div>

    </div>
  );
};

export default CategoryMaster;
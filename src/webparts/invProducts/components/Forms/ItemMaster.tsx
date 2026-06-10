import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";


import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface ItemsData {
    ItemName: string;
    CategoryNameId: number;
    IsActive: boolean;
}

export const ItemMaster = (props: any) => {
    const listName = "OfficeItemMaster";

    const [formData, setFormData] = useState<ItemsData>({
        ItemName: "",
        CategoryNameId: 0,
        IsActive: true,
    });

    const [data, setData] = useState<any[]>([]);
    const [categoryData, setcategoryData] = useState<any[]>([]);
    const [itemId, setItemId] = useState<number>(0);

    const sp = spfi().using(SPFx(props.context));

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        try {
            // const res = await sp.web.lists.getByTitle(listName).items.select("Id","ItemName","IsActive","CategoryName/Title").expand("CategoryName")();

            const [res, categoryData] = await Promise.all([
                sp.web.lists.getByTitle(listName).items.select("Id", "ItemName", "IsActive", "CategoryNameId","CategoryName/CategoryName").expand("CategoryName")(),
                sp.web.lists.getByTitle("OfficeCategoryMaster").items.select("*")(),
            ]);

            setcategoryData(categoryData);
            console.log(data);

            setData(res);
        } catch (error) {
            console.log(error);
        }
    };

    const handleChange = (e: any) => {
        const { name, value, type, checked } = e.target;

        setFormData({
            ...formData,
            [name]: type === "checkbox" ? checked : name === "CategoryNameId" ? Number(value) : value,
        });
    };

    const handleSubmit = async () => {
        if (!formData.ItemName.trim()) {
            alert("Item Name is required");
            return;
        }
        if (formData.CategoryNameId === 0) {
            alert("Please selct Category");
            return;
        }

        const payload = {
            ItemName: formData.ItemName,
            CategoryNameId: Number(formData.CategoryNameId),
            IsActive: formData.IsActive,
        };

        try {
            if (itemId > 0) {
                await sp.web.lists
                    .getByTitle(listName)
                    .items.getById(itemId)
                    .update(payload);

                alert("Updated Successfully");
            } else {
                await sp.web.lists
                    .getByTitle(listName)
                    .items.add(payload);

                alert("Saved Successfully");
            }

            resetForm();
            loadData();
            
        } catch (error) {
            console.log(error);
        }
    };

    const editItem = (item: any) => {
        setItemId(item.Id);

        setFormData({
            ItemName: item.ItemName,
            CategoryNameId: item.CategoryNameId,
            IsActive: item.IsActive,
        });
    };

    const resetForm = () => {
        setItemId(0);

        setFormData({
            ItemName: "",
            CategoryNameId: 0,
            IsActive: true,
        });
    };

    return (
        <div style={{ padding: "20px" }}>
            <h2>Item Master</h2>

            <div>
                <label>Item Name *</label>
                <br />

                <input type="text" name="ItemName" value={formData.ItemName} onChange={handleChange} />
                <br />
                <label>Category Name</label>
                <select name="CategoryNameId" value={formData.CategoryNameId} onChange={handleChange}
                >
                    <option value={0}>Select Category</option>
                    {
                        categoryData.map((item) => {
                            return (
                                // <option value={0}>{item.CategoryName}</option>
                                <option key={item.Id} value={item.Id}>
                                    {item.CategoryName}
                                </option>
                            )
                        })
                    }
                </select>
                <br />
                <label>
                    <input type="checkbox" name="IsActive" checked={formData.IsActive} onChange={handleChange} />
                    Is Active
                </label>

                <br />
                <br />

                <button onClick={handleSubmit}>
                    {itemId > 0 ? "Update" : "Submit"}
                </button>

                <button
                    onClick={resetForm}
                    style={{ marginLeft: "10px" }}
                >
                    Clear
                </button>
            </div>

            <hr />
            <table
                style={{
                    borderCollapse: "collapse",
                    width: "100%"
                }}
            >
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Category Name</th>
                        <th>Item Name</th>
                        <th>Is Active</th>
                        <th>Edit</th>
                    </tr>
                </thead>

                <tbody>
                    {data.map((item: any) => (
                        <tr key={item.Id}>
                            <td>{item.Id}</td>
                            <td>{item.CategoryName?.CategoryName}</td>
                            {/* <td>{item.CategoryNameId?.CategoryNameId}</td>     */}
                            <td>{item.ItemName}</td>


                            <td>{item.IsActive ? "Yes" : "No"}</td>
                            <td>
                                <button onClick={() => editItem(item)}>
                                    Edit
                                </button>
                            </td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};
export default ItemMaster;
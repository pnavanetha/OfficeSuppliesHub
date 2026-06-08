import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";


import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface IFormData {
    CategoryName: string;
    IsActive: boolean;
}

export const CategoryMaster = (props: any) => {
    const listName = "OfficeCategoryMaster";

    const [formData, setFormData] = useState<IFormData>({
        CategoryName: "",
        IsActive: true,
    });

    const [data, setData] = useState<any[]>([]);
    const [itemId, setItemId] = useState<number>(0);

    const sp = spfi().using(SPFx(props.context));

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        try {
            const res = await sp.web.lists
                .getByTitle(listName)
                .items();

            setData(res);
        } catch (error) {
            console.log(error);
        }
    };

    const handleChange = (e: any) => {
        const { name, value, type, checked } = e.target;

        setFormData({
            ...formData,
            [name]: type === "checkbox" ? checked : value,
        });
    };

    const handleSubmit = async () => {
        if (!formData.CategoryName.trim()) {
            alert("Category Name is required");
            return;
        }

        try {
            if (itemId > 0) {
                await sp.web.lists
                    .getByTitle(listName)
                    .items.getById(itemId)
                    .update(formData);

                alert("Updated Successfully");
            } else {
                await sp.web.lists
                    .getByTitle(listName)
                    .items.add(formData);

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
            CategoryName: item.CategoryName,
            IsActive: item.IsActive,
        });
    };

    const resetForm = () => {
        setItemId(0);

        setFormData({
            CategoryName: "",
            IsActive: true,
        });
    };

    return (
        <div style={{ padding: "20px" }}>
            <h2>Category Master</h2>

            <div>
                <label>Category Name *</label>
                <br />

                <input
                    type="text"
                    name="CategoryName"
                    value={formData.CategoryName}
                    onChange={handleChange}
                />

                <br />
                <br />

                <label>
                    <input
                        type="checkbox"
                        name="IsActive"
                        checked={formData.IsActive}
                        onChange={handleChange}
                    />
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
                        <th>Is Active</th>
                        <th>Edit</th>
                    </tr>
                </thead>

                <tbody>
                    {data.map((item: any) => (
                        <tr key={item.Id}>
                            <td>{item.Id}</td>
                            <td>{item.CategoryName}</td>
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
export default CategoryMaster;
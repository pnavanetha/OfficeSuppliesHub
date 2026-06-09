import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";


import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface DepartmentData {
    Name: string;
    IsActive: boolean;
}

export const DeparmentMaster = (props: any) => {
    const listName = "OfficeDepartments";

    const [formData, setFormData] = useState<DepartmentData>({
        Name: "",
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
        if (!formData.Name.trim()) {
            alert("Department Name is required");
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
            Name: item.Name,
            IsActive: item.IsActive,
        });
    };

    const resetForm = () => {
        setItemId(0);

        setFormData({
            Name: "",
            IsActive: true,
        });
    };

    return (
        <div style={{ padding: "20px" }}>
            <h2>Deparment Master</h2>

            <div>
                <label>Deparment Name *</label>
                <br />

                <input
                    type="text"
                    name="Name"
                    value={formData.Name}
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
                        <th>Deparment Name</th>
                        <th>Is Active</th>
                        <th>Edit</th>
                    </tr>
                </thead>

                <tbody>
                    {data.map((item: any) => (
                        <tr key={item.Id}>
                            <td>{item.Id}</td>
                            <td>{item.Name}</td>
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
export default DeparmentMaster;
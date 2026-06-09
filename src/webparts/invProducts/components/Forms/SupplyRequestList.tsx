import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { useNavigate } from "react-router-dom";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const SupplyRequestList = (props: any) => {

    const sp = spfi().using(SPFx(props.context));
    const navigate = useNavigate();

    const [data, setData] = useState<any[]>([]);

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {

        try {

            const res = await sp.web.lists
                .getByTitle("OfficeSupplyRequestList")
                .items
                .select(
                    "Id",
                    "RequestDate",
                    "Comments",
                    "Status",
                    "EmployeeName/Title",
                    "Department/Title",
                    "CategoryName/Title",
                    "ItemName/Title"
                )
                .expand(
                    "EmployeeName",
                    "Department",
                    "CategoryName",
                    "ItemName"
                )();

            setData(res);

        } catch (error) {
            console.log(error);
        }
    };

    const addRequest = () => {

        navigate("/SupplyRequestForm");
    };

    const editRequest = (id: number) => {

        navigate(`/SupplyRequestForm/${id}`);
    };

    const deleteRequest = async (id: number) => {

        const confirmDelete = window.confirm(
            "Are you sure you want to delete this request?"
        );

        if (!confirmDelete) {
            return;
        }

        try {

            await sp.web.lists
                .getByTitle("OfficeSupplyRequestList")
                .items
                .getById(id)
                .delete();

            alert("Deleted Successfully");

            loadData();

        } catch (error) {
            console.log(error);
        }
    };

    return (
        <div style={{ padding: "20px" }}>

            <div
                style={{
                    display: "flex",
                    justifyContent: "space-between",
                    marginBottom: "20px"
                }}
            >
                <h2>Office Supply Requests</h2>

                <button onClick={addRequest}>
                    + Add Request
                </button>
            </div>

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
                        <th>Edit</th>
                        <th>Delete</th>
                    </tr>
                </thead>

                <tbody>

                    {
                        data.length > 0 ?

                            data.map((item: any) => (

                                <tr key={item.Id}>

                                    <td>{item.Id}</td>

                                    <td>
                                        {item.EmployeeName?.Title}
                                    </td>

                                    <td>
                                        {item.Department?.Title}
                                    </td>

                                    <td>
                                        {item.CategoryName?.Title}
                                    </td>

                                    <td>
                                        {item.ItemName?.Title}
                                    </td>

                                    <td>
                                        {item.RequestDate}
                                    </td>

                                    <td>
                                        {item.Comments}
                                    </td>

                                    <td>
                                        {item.Status}
                                    </td>

                                    <td>
                                        <button
                                            onClick={() =>
                                                editRequest(item.Id)
                                            }
                                        >
                                            Edit
                                        </button>
                                    </td>

                                    <td>
                                        <button
                                            onClick={() =>
                                                deleteRequest(item.Id)
                                            }
                                        >
                                            Delete
                                        </button>
                                    </td>

                                </tr>

                            ))

                            :

                            <tr>
                                <td
                                    colSpan={10}
                                    style={{
                                        textAlign: "center"
                                    }}
                                >
                                    No Records Found
                                </td>
                            </tr>
                    }

                </tbody>

            </table>

        </div>
    );
};

export default SupplyRequestList;
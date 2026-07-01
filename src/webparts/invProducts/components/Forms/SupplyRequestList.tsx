import * as React from "react";
import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { useNavigate } from "react-router-dom";
import { formatDateIN } from "../Common/dateHelpers";
import "../CSS/RequestList.css";
import { IColumn } from "@fluentui/react";
import CommonGrid from "../Common/CommonGrid";



const SupplyRequestList = (props: any) => {


    const sp = spfi().using(SPFx(props.context));
    const navigate = useNavigate();

    const [data, setData] = useState<any[]>([]);

    useEffect(() => {
        loadData();
    }, []);


    // const loadData = async () => {

    //     try {
    //         const res = await sp.web.lists.getByTitle("OfficeSupplyRequestList").items.select("Id", "RequestDate", "Comments", "Status", "EmployeeName/Title", "Department/Name", "CategoryName/CategoryName", "ItemName/ItemName")
    //             .expand("EmployeeName", "Department", "CategoryName", "ItemName")();
    //         setData(res);

    //     } catch (error) {
    //         console.log(error);
    //     }
    // };

    const loadData = async () => {
        try {
            const currentUser = await sp.web.currentUser();
            let items;
            if (props.role === "Admin") {
                items = await sp.web.lists.getByTitle("OfficeSupplyRequestList").items.select("Id", "RequestDate", "Comments", "Status", "EmployeeName/Title", "Department/Name", "CategoryName/CategoryName", "ItemName/ItemName").expand("EmployeeName", "Department", "CategoryName", "ItemName")();
            } else {
                items = await sp.web.lists.getByTitle("OfficeSupplyRequestList").items.select("Id", "RequestDate", "Comments", "Status", "EmployeeName/Title", "Department/Name", "CategoryName/CategoryName", "ItemName/ItemName").expand("EmployeeName", "Department", "CategoryName", "ItemName").filter(`EmployeeNameId eq ${currentUser.Id}`)();
            }
            setData(items);
            console.log("User ID:", currentUser.Id);

        } catch (error) {
            console.log(error);

        }
    }

    const addRequest = () => {
        navigate("/SupplyRequestForm");
    };
    const editRequest = (id: number): void => {
        navigate(`/SupplyRequestForm/${id}`);
    };
    // const deleteRequest = async (id: number) => {
    //     const confirmDelete = window.confirm(
    //         "Are you sure you want to delete this request?"
    //     );

    //     if (!confirmDelete) {
    //         return;
    //     }

    //     try {

    //         await sp.web.lists.getByTitle("OfficeSupplyRequestList").items.getById(id).delete();
    //         alert("Deleted Successfully");
    //         loadData();

    //     } catch (error) {
    //         console.log(error);
    //     }
    // };

    const columns: IColumn[] = [
        {
            key: "Edit", name: "Edit", minWidth: 70, onRender: (item: any) => (
                <button onClick={() => editRequest(item.ID)}>Edit</button>
            )
        },
        // {
        //     key: "Delete", name: "Delete", minWidth: 70, onRender: (item: any) => (
        //         <button onClick={() => deleteRequest(item.ID)}>Delete</button>
        //     )
        // },
        // { key: "ID", name: "ID", fieldName: "ID", minWidth: 60, isResizable: true },
        { key: "Employee", name: "Employee Name", fieldName: "Employee", minWidth: 150, isResizable: true },
        { key: "Department", name: "Department", fieldName: "Department", minWidth: 120, isResizable: true },
        { key: "Category", name: "Category", fieldName: "Category", minWidth: 120, isResizable: true },
        { key: "Item", name: "Item", fieldName: "Item", minWidth: 120, isResizable: true },
        { key: "Request", name: "Request Date", fieldName: "Request", minWidth: 120, isResizable: true },
        { key: "Status", name: "Status", fieldName: "Status", minWidth: 100, isResizable: true },
        { key: "Comments", name: "Comments", fieldName: "Comments", minWidth: 150, isResizable: true },

    ]

    return (
        <div className="list-content">

            <div className="list-header">
                <h2>Office Supply Requests</h2> <br />
            </div>

            <div className="action-bar">
                <button className="add-request-btn" onClick={addRequest}>
                    + Add Request
                </button>
            </div>

            <div className="grid-container">
                <CommonGrid
                    items={data.map((item) => ({
                        ID: item.Id,
                        Employee: item.EmployeeName?.Title || "",
                        Department: item.Department?.Name || "",
                        Category: item.CategoryName?.CategoryName || "",
                        Item: item.ItemName?.ItemName || "",
                        Request: formatDateIN(item.RequestDate),
                        Comments: item.Comments || "",
                        Status: item.Status || ""
                    }))}
                    columns={columns}
                    pageSize={10}
                    searchFields={[
                        "ID",
                        "Employee",
                        "Department",
                        "Category",
                        "Item",
                        "Request",
                        "Comments",
                        "Status"
                    ]}
                />
            </div>

        </div>
    );
};

export default SupplyRequestList;
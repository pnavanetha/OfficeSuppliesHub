import * as React from 'react';
import { useEffect, useState, useRef, useCallback, useMemo } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { NavLink } from 'react-router-dom';
import { hideLoader, showLoader } from '../Shared/Loader';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faEdit } from '@fortawesome/free-solid-svg-icons';
import InputCheckBox from '../Shared/InputCheckBox';
import AGGridDataTable from '../Shared/AGGridDataTable';
import { ActionStatus, ControlType } from '../Constants/Contants';
import formValidation from '../Utilities/FormValidator';
import { showToast } from '../Shared/Toaster';

interface IFormData {
    Title: string;
    IsActive: boolean;
}

interface ITableData {
    Id: number;
    CategoryName: string;
    IsActive: string;
}

export const CategoryMaster = (props: any) => {

    const listName = "Category Master";

    const [formData, setFormData] = useState<IFormData>({
        Title: '',
        IsActive: true
    });

    const [tableData, setTableData] = useState<ITableData[]>([]);
    const [isFormOpen, setIsFormOpen] = useState(false);
    const [ItemId, setItemId] = useState(0);

    const txtCategory = useRef<HTMLInputElement>(null);

    const sp = React.useMemo(
        () => spfi().using(SPFx(props.context)),
        [props.context]
    );

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        try {
            showLoader();

            const res = await sp.web.lists
                .getByTitle(listName)
                .items.orderBy("Modified", false)();

            const tData = res.map(item => ({
                Id: item.Id,
                CategoryName: item.Title,
                IsActive: item.IsActive ? "Yes" : "No"
            }));

            setTableData(tData);
            hideLoader();
        }
        catch (e) {
            console.error(e);
            onError();
        }
    };

    const editItem = async (Id: number) => {
        try {
            showLoader();

            setIsFormOpen(true);
            setItemId(Id);

            const item = await sp.web.lists
                .getByTitle(listName)
                .items.getById(Id)();

            setFormData({
                Title: item.Title || "",
                IsActive: item.IsActive ?? true
            });

            hideLoader();
        }
        catch (e) {
            console.error(e);
            onError();
        }
    };

    const handleChange = (e: any) => {
        const name = e.target.name;
        const value = e.target.type === "checkbox"
            ? e.target.checked
            : e.target.value;

        setFormData(prev => ({
            ...prev,
            [name]: value
        }));
    };

    const handleSubmit = async () => {
        try {
            showLoader();

            const data = {
                category: {
                    val: formData.Title,
                    required: true,
                    Name: "Category Name",
                    Type: ControlType.string,
                    Focusid: txtCategory
                }
            };

            const isValid = formValidation.FormValidation(data);

            if (!isValid.status) {
                showToast("error", isValid.message);
                hideLoader();
                return;
            }

            await InsertOrUpdate();
        }
        catch (e) {
            console.error(e);
            onError();
        }
    };

    const InsertOrUpdate = async () => {
        try {
            if (ItemId > 0) {
                await sp.web.lists
                    .getByTitle(listName)
                    .items.getById(ItemId)
                    .update(formData);

                onSuccess("Category updated successfully");
            } else {
                await sp.web.lists
                    .getByTitle(listName)
                    .items.add(formData);

                onSuccess("Category added successfully");
            }
        }
        catch (e) {
            console.error(e);
            onError();
        }
    };

    const closeForm = () => {
        setFormData({ Title: '', IsActive: true });
        setItemId(0);
        setIsFormOpen(false);
    };

    const onSuccess = async (msg: string) => {
        await loadData();
        showToast("success", msg);
        closeForm();
        hideLoader();
    };

    const onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    };

    const addNew = () => {
        setFormData({ Title: '', IsActive: true });
        setItemId(0);
        setIsFormOpen(true);
    };

    const columns = useMemo(() => [
        {
            field: "Id",
            headerName: "Edit",
            width: 80,
            cellRenderer: (params: any) => (
                <NavLink to={""}
                    onClick={(e) => {
                        e.preventDefault();
                        editItem(params.data.Id);
                    }}>
                    <FontAwesomeIcon icon={faEdit} />
                </NavLink>
            ),
        },
        {
            field: "CategoryName",
            headerName: "Category Name",
            flex: 1
        },
        {
            field: "IsActive",
            headerName: "Is Active",
            width: 130
        }
    ], []);

    return (
        <div className="container-fluid">
            <div className="light-box">

                <div className="form-title">Category Master</div>

                {isFormOpen &&
                    <div className="form-border-box p-2">

                        <div className="row">

                            <div className="col-md-4">
                                <label>Category Name *</label>
                                <input
                                    type="text"
                                    name="Title"
                                    className="form-control"
                                    value={formData.Title}
                                    onChange={handleChange}
                                    ref={txtCategory}
                                />
                            </div>

                            <div className="col-md-4">
                                <InputCheckBox
                                    label="Is Active"
                                    name="IsActive"
                                    checked={formData.IsActive}
                                    onChange={handleChange}
                                />
                            </div>

                            <div className="col-md-12 text-center mt-3">
                                <button onClick={handleSubmit}>
                                    {ItemId ? "Update" : "Submit"}
                                </button>

                                <button onClick={closeForm}>
                                    Cancel
                                </button>
                            </div>

                        </div>

                    </div>
                }

                <AGGridDataTable
                    data={tableData}
                    columns={columns}
                    addNew={addNew}
                    showAddButton={!isFormOpen}
                />

            </div>
        </div>
    );
};

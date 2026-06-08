import * as React from "react"; 
import { useEffect,useState } from "react";
import { spfi, SPFx } from "@pnp/sp";

interface IFormData {
    CategoryName: string;
    IsActive: boolean;
}

export const CategoryMaster = (props: any) => {
    const listName = "OfficeCategoryMaster";
    const [formData, setFormData] = React.useState<IFormData>({
        CategoryName: '',
        IsActive: true
    });
    const [data, setdata] = React.useState<any[]>([]);
    const [itemId, setItemId] = useState<number>(0);

    const sp = spfi().using(SPFx(props.context));

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        const res = await sp.web.lists.getByTitle(listName).items();
        setdata(res);
    };

    const handleChange = (e: any) => {
        const {name, value, type, checked} = e.target;

        setFormData({
            ...formData,
            [name]: type === "checkbox" ? checked : value
        });
    };




}
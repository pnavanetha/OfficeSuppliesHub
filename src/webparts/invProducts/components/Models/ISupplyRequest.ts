export interface ISupplyRequest {
    Id?: number;
    EmployeeNameId: number;
    DepartmentId: number;
    CategoryNameId: number;
    ItemNameId: number;
    RequestDate: string;
    Comments: string;
    Status: string;
}
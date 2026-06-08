export interface OfficeSupplyRequest {
    Id?: number;
    EmployeeName: 'PeoplePicker';
    Department: string;
    CategoryName: string;
    ItemName: string;
    RequestedDate: Date;
    Comments: string;
    Status: string;
}

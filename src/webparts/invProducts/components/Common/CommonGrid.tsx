import * as React from "react";
import { useState } from "react";
import {
  DetailsList,
  IColumn,
  SelectionMode,
  TextField,
  DetailsListLayoutMode,
  ConstrainMode,
  IconButton
} from "@fluentui/react";
import "./CommonGrid.css";


interface CommonGridProps {
  items: any[];
  columns: IColumn[];
  pageSize?: number;
  searchFields?: string[];
}

const CommonGrid: React.FC<CommonGridProps> = ({
  items,
  columns,
  pageSize = 5,
  searchFields = []
}) => {


  const [searchText, setSearchText] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [sortColumn, setSortColumn] = useState("");
  const [isSortedDescending, setIsSortedDescending] = useState(false);

  // Filter
  const filteredData = items.filter((item) =>
    searchFields.length === 0
      ? true
      : searchFields.some((field) =>
        String(item[field] ?? "")
          .toLowerCase()
          .indexOf(searchText.toLowerCase()) !== -1
      )
  );

  // Sort
  const sortedData = [...filteredData];

  if (sortColumn) {
    sortedData.sort((a, b) => {
      const first = String(a[sortColumn] ?? "").toLowerCase();
      const second = String(b[sortColumn] ?? "").toLowerCase();

      if (first < second) return isSortedDescending ? 1 : -1;
      if (first > second) return isSortedDescending ? -1 : 1;
      return 0;
    });
  }

  const onColumnClick = (
    ev?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ) => {

    if (!column?.fieldName) return;

    setSortColumn(column.fieldName);

    setIsSortedDescending(
      sortColumn === column.fieldName
        ? !isSortedDescending
        : false
    );
  };

  const updatedColumns = columns.map((col) => ({
    ...col,
    isSorted: sortColumn === col.fieldName,
    isSortedDescending,
    onColumnClick
  }));

  const totalPages = Math.ceil(sortedData.length / pageSize);

  const pagedData = sortedData.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );

  return (
    <div>

      <div className="search-container">
        <div className="search-box">
          <TextField
            placeholder="Search..."
            value={searchText}
            onChange={(_, value) => {
              setSearchText(value || "");
              setCurrentPage(1);
            }}
          />

          {searchText && (
            <IconButton
              className="clear-btn"
              iconProps={{ iconName: "Cancel" }}
              onClick={() => {
                setSearchText("");
                setCurrentPage(1);
              }}
            />
          )}
        </div>
      </div>

      <div className="grid-section">

        {filteredData.length > 0 ? (

          <DetailsList
            items={pagedData}
            columns={updatedColumns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.horizontalConstrained}
          />

        ) : (

          <>
            <div className="empty-header">
              {updatedColumns.map((col) => (
                <div
                  key={col.key}
                  className="empty-header-cell"
                  style={{ minWidth: col.minWidth }}
                >
                  {col.name}
                </div>
              ))}
            </div>

            <div className="empty-state">
              <div className="empty-icon">📭</div>
              <div className="empty-title">
                No records found
              </div>
            </div>

          </>

        )}

      </div>

      <div className="pagination">

        <button
          disabled={currentPage === 1}
          onClick={() => setCurrentPage(currentPage - 1)}
        >
          Previous
        </button>

        <span>
          Page {totalPages === 0 ? 0 : currentPage} of {totalPages}
        </span>

        <button
          disabled={currentPage === totalPages || totalPages === 0}
          onClick={() => setCurrentPage(currentPage + 1)}
        >
          Next
        </button>

      </div>

    </div>
  );
};

export default CommonGrid;
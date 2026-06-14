import * as React from "react";
import { useState } from "react";
import { DetailsList, IColumn, SelectionMode, TextField } from "@fluentui/react";
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
  const [sortColumn, setSortColumn] = useState<string>("");
  const [isSortedDescending, setIsSortedDescending] = useState<boolean>(true);

  // ✅ Filtering
  const filteredData = items.filter((item: any) =>
    searchFields.length === 0
      ? true
      : searchFields.some((field: string) =>
        item[field]
          ?.toString()
          .toLowerCase()
          .includes(searchText.toLowerCase())
      )
  );

  // ✅ Sorting (AFTER filtering)
  const sortedData = [...filteredData];

  if (sortColumn) {
    sortedData.sort((a: any, b: any) => {

      const first = a[sortColumn]?.toString().toLowerCase() || "";
      const second = b[sortColumn]?.toString().toLowerCase() || "";

      if (first < second) return isSortedDescending ? 1 : -1;
      if (first > second) return isSortedDescending ? -1 : 1;
      return 0;
    });
  }

  // ✅ Column Click (ONLY set state)
  const onColumnClick = (
    ev?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {

    if (!column?.fieldName) return;

    const desc =
      sortColumn === column.fieldName
        ? !isSortedDescending
        : false;

    setSortColumn(column.fieldName);
    setIsSortedDescending(desc);
  };

  // ✅ Attach sorting to columns
  const updatedColumns = columns.map((col) => ({
    ...col,
    isSorted: sortColumn === col.fieldName,
    isSortedDescending: isSortedDescending,
    onColumnClick
  }));

  // ✅ Pagination (AFTER sorting)
  const totalPages = Math.ceil(sortedData.length / pageSize);

  const pagedData = sortedData.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );

  return (
    <div>

      {/* 🔎 Search */}
      <div className="search-container">
        <TextField
          placeholder="Search..."
          value={searchText}
          onChange={(_, value) => {
            setSearchText(value || "");
            setCurrentPage(1);
          }}
        />
      </div>

      {/* 📊 Grid */}
      <div className="grid-section">
        <DetailsList
          items={pagedData}
          columns={updatedColumns}
          selectionMode={SelectionMode.none}
        />
      </div>

      {/* 📄 Pagination */}
      <div className="pagination">
        <button
          disabled={currentPage === 1}
          onClick={() => setCurrentPage(currentPage - 1)}
        >
          Previous
        </button>

        <span>
          Page {currentPage} of {totalPages}
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

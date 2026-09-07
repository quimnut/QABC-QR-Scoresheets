import { useState } from "react";

interface Props {
  headers: string[];
  rows: string[][];
}

export default function CsvPreview({ headers, rows }: Props) {
  const [filter, setFilter] = useState("");

  const filteredRows = filter
    ? rows.filter((row) =>
        row.some((cell) => cell.toLowerCase().includes(filter.toLowerCase()))
      )
    : rows;

  return (
    <div className="preview-container">
      <div className="search-box">
        <input
          type="text"
          placeholder="Search entries…"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
          style={{ width: "100%", maxWidth: 400 }}
        />
      </div>
      <p style={{ marginBottom: 8, fontSize: "0.875rem", color: "var(--text-muted)" }}>
        Showing {filteredRows.length} of {rows.length} entries
      </p>
      <div className="table-container" style={{ maxHeight: 400, overflowY: "auto" }}>
        <table>
          <thead>
            <tr>
              {headers.map((h) => (
                <th key={h}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => (
                  <td key={j}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

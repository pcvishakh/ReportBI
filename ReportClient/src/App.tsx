import { useEffect, useMemo, useState } from "react";
import { Adaptable } from "@adaptabletools/adaptable-react-aggrid";
import type {
  AdaptableApi,
  AdaptableReadyInfo,
  AdaptableState,
  AdaptableStateFunctionConfig,
} from "@adaptabletools/adaptable";

import {
  AllCommunityModule,
  type GridApi,
  ModuleRegistry,
  themeAlpine,
} from "ag-grid-community";
import { AllEnterpriseModule } from "ag-grid-enterprise";

import "@adaptabletools/adaptable-react-aggrid/base.css";
import "@adaptabletools/adaptable-react-aggrid/themes/light.css";
import "@adaptabletools/adaptable-react-aggrid/themes/dark.css";

import "./index.css";
import React from "react";

ModuleRegistry.registerModules([AllCommunityModule, AllEnterpriseModule]);

interface Report {
  reportID: number;
  reportName: string;
  columnDefinitions: string | null;
}

interface ReportDataResponse {
  data: Record<string, any>[];
  gridState: string | null;
  columnDefinitions: string | null;
}

const API_BASE_URL = "http://localhost:5108/api/reports";

function App() {
  const [reports, setReports] = useState<Report[]>([]);
  const [selectedReport, setSelectedReport] = useState<Report | null>(null);
  const [reportData, setReportData] = useState<any[] | null>(null);
  const [gridState, setGridState] = useState<string | null>(null);
  const [columnDefinitions, setColumnDefinitions] = useState<string | null>(null);
  const [loadingList, setLoadingList] = useState(true);
  const [loadingData, setLoadingData] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [saveConfKey, setSaveConfKey] = useState<string>("");
  const adaptableApiRef = React.useRef<AdaptableApi>(null);
  const aggridApiRef = React.useRef<GridApi>(null);

  useEffect(() => {
    fetchReports();
  }, []);

  useEffect(() => {
    adaptableApiRef.current?.stateApi.setAdaptableStateKey(saveConfKey);
  }, [saveConfKey]);

  const fetchReports = async () => {
    try {
      setLoadingList(true);
      const res = await fetch(API_BASE_URL);
      if (!res.ok) throw new Error("Failed to fetch reports");
      const data = await res.json();
      setReports(data);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoadingList(false);
    }
  };

  const handleSelectReport = async (report: Report) => {
    setSelectedReport(report);
    setReportData(null);
    setGridState(null);
    setColumnDefinitions(null);
    setError(null);

    try {
      setLoadingData(true);
      const res = await fetch(`${API_BASE_URL}/${report.reportID}/execute`);
      if (!res.ok) throw new Error("Failed to execute report");
      const result: ReportDataResponse = await res.json();
      setReportData(result.data);
      setGridState(result.gridState);
      setColumnDefinitions(result.columnDefinitions);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoadingData(false);
    }
  };

  const handleExportExcel = async () => {
    if (!selectedReport) return;
    try {
      const res = await fetch(
        `${API_BASE_URL}/${selectedReport.reportID}/export`,
        {
          method: "GET",
        },
      );
      if (!res.ok) throw new Error("Failed to export report");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${selectedReport.reportName}_export.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      setError(err.message);
    }
  };

  const agGridProps = useMemo(() => {
    if (!reportData || reportData.length === 0) return {};

    let columnDefs: any[] = [];
    if (columnDefinitions) {
      try {
        const parsedDefs = JSON.parse(columnDefinitions);
        columnDefs = parsedDefs.map((def: any) => ({
          field: def.field,
          headerName: def.headerName || def.field,
          filter: true,
          sortable: true,
          resizable: true,
          enableRowGroup: true,
          enablePivot: true,
          enableValue: true,
          // Map backend dataType to AG Grid cellDataType
          cellDataType: def.dataType === "number" ? "number" : 
                        def.dataType === "date" ? "date" : 
                        def.dataType === "boolean" ? "boolean" : 
                        "text",
          valueFormatter: (params: any) => {
            if (def.dataType === "date" && params.value) {
              const d = new Date(params.value);
              if (!isNaN(d.getTime())) {
                return (d.getMonth() + 1).toString().padStart(2, "0") + "/" +
                       d.getDate().toString().padStart(2, "0") + "/" +
                       d.getFullYear();
              }
            }
            return params.value;
          }
        }));
      } catch (e) {
        console.error("Failed to parse backend column definitions", e);
      }
    }

    if (columnDefs.length === 0 && reportData && reportData.length > 0) {
      columnDefs = Object.keys(reportData[0]).map((key) => ({
        field: key,
        headerName: key,
        filter: true,
        sortable: true,
        resizable: true,
        enableRowGroup: true,
        enablePivot: true,
        enableValue: true,
        valueFormatter: (params: any) => {
          if (params.value && typeof params.value === "string" && /^\d{4}-\d{2}-\d{2}/.test(params.value)) {
            const d = new Date(params.value);
            if (!isNaN(d.getTime())) {
              return (d.getMonth() + 1).toString().padStart(2, "0") + "/" +
                     d.getDate().toString().padStart(2, "0") + "/" +
                     d.getFullYear();
            }
          }
          return params.value;
        },
      }));
    }

    return {
      columnDefs,
      rowData: reportData,
      theme: themeAlpine,
      defaultColDef: {
        filter: true,
        sortable: true,
        resizable: true,
        enableRowGroup: true,
        enablePivot: true,
        enableValue: true,
      },
      sideBar: {
        toolPanels: [
          {
            id: "columns",
            labelDefault: "Columns",
            labelKey: "columns",
            iconKey: "columns",
            toolPanel: "agColumnsToolPanel",
            toolPanelParams: {
              suppressRowGroups: false,
              suppressValues: false,
              suppressPivots: false,
              suppressPivotMode: false,
            },
          },
          {
            id: "filters",
            labelDefault: "Filters",
            labelKey: "filters",
            iconKey: "filter",
            toolPanel: "agFiltersToolPanel",
          },
        ],
        defaultToolPanel: "columns",
      },
      rowGroupPanelShow: "always" as const,
      pivotPanelShow: "always" as const,
      statusBar: {
        statusPanels: [
          {
            statusPanel: "agTotalAndFilteredRowCountComponent",
            align: "left",
          },
        ],
      },
    };
  }, [reportData, columnDefinitions]);

  const persistanceService = {
    // If there is no saved state we pass the default state.
    // If saved report we return that state and any additional changes.
    loadAdaptableState: async (key: string) => {
      // Expect keys like "Report_1"
      const parts = key.split("_");
      const idPart = parts[1];
      const id = Number(idPart);

      if (!Number.isFinite(id)) {
        console.warn("Unable to derive report id from adaptableStateKey", key);
        return {};
      }

      try {
        // Use the GetReport endpoint so we don't re-execute the heavy query.
        const res = await fetch(`${API_BASE_URL}/${id}`);
        if (!res.ok) {
          console.warn("Failed to load report for state", id, res.status);
          return {};
        }

        const report = await res.json();
        const storedState = report.gridState as string | null | undefined;

        if (!storedState) {
          return {};
        }

        try {
          return JSON.parse(storedState);
        } catch {
          console.warn("Unable to parse stored grid state for report", id);
          return {};
        }
      } catch (err) {
        console.error("Error loading adaptable state from API", err);
        return {};
      }
    },
    persistAdaptableState: async (
      state: Partial<AdaptableState>,
      key: string,
      _user: string,
    ) => {
      const parts = key.split("_");
      const idPart = parts[1];
      const id = Number(idPart);

      if (!Number.isFinite(id)) {
        console.warn("Unable to derive report id from adaptableStateKey", key);
        return;
      }

      const serializedState = JSON.stringify(state);
      setGridState(serializedState);

      try {
        await fetch(`${API_BASE_URL}/${id}/state`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            gridState: serializedState,
          }),
        });
      } catch (err) {
        console.error("Failed to persist adaptable state via API", err);
      }
    },
  };

  const adaptableOptions: any = useMemo(() => {
    if (!selectedReport || !reportData || reportData.length === 0) return {};

    const columns = Object.keys(reportData[0]);

    const options = {
      primaryKey: columns[0] || "id",
      userName: "DefaultUser",
      adaptableId: `Report_${selectedReport.reportID}`,
      adaptableStateKey: `Report_${selectedReport.reportID}`,
      licenseKey: "REPLACE_WITH_LICENSE_KEY_IF_APPLICABLE",
      stateOptions: {
        loadState: async (_config: AdaptableStateFunctionConfig) => {
          return persistanceService.loadAdaptableState(
            _config.adaptableStateKey,
          );
        },
        persistState: async (
          _state: Partial<AdaptableState>,
          config: AdaptableStateFunctionConfig,
        ) => {
          return persistanceService.persistAdaptableState(
            _state,
            config.adaptableStateKey,
            config.userName,
          );
        },
      },
      initialState: {
        Dashboard: {
          Revision: 1,
          Tabs: [
            {
              Name: "Demo",
              Toolbars: ["Layout", "ColumnFilter"],
            },
          ],
        },
        Theme: { CurrentTheme: "light" },
        Layout: {
          CurrentLayout: "Default",
          Layouts: [
            {
              Name: "Default",
              TableColumns: columns,
            },
          ],
        },
      },
    };

    return options;
  }, [selectedReport, reportData, gridState]);

  return (
    <div
      className="container"
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        maxWidth: "none",
        margin: "0",
        padding: "0",
      }}
    >


      {error && (
        <div
          className="error"
          style={{
            margin: "1rem",
            padding: "1rem",
            background: "#fee",
            color: "#c00",
            borderRadius: "4px",
          }}
        >
          {error}
        </div>
      )}

      <div
        className="grid"
        style={{ display: "flex", flex: 1, overflow: "hidden" }}
      >
        <aside
          className="report-list"
          style={{
            width: "250px",
            borderRight: "1px solid #ddd",
            padding: "1rem",
            overflowY: "auto",
          }}
        >
          <h2 style={{ fontSize: "1.25rem" }}>Reports</h2>
          {loadingList
            ? <p className="loading">Loading reports...</p>
            : reports.length === 0
            ? <p className="loading">No reports found.</p>
            : (
              <ul style={{ listStyle: "none", padding: 0 }}>
                {reports.map((report) => (
                  <li
                    key={report.reportID}
                    style={{
                      padding: "0.5rem",
                      cursor: "pointer",
                      marginBottom: "0.25rem",
                      borderRadius: "4px",
                      background: selectedReport?.reportID === report.reportID
                        ? "#e6f7ff"
                        : "transparent",
                      color: selectedReport?.reportID === report.reportID
                        ? "#1890ff"
                        : "inherit",
                      fontWeight: selectedReport?.reportID === report.reportID
                        ? "bold"
                        : "normal",
                    }}
                    onClick={() => handleSelectReport(report)}
                  >
                    {report.reportName}
                  </li>
                ))}
              </ul>
            )}
        </aside>

        <main
          className="report-content"
          style={{
            flex: 1,
            padding: "1rem",
            display: "flex",
            flexDirection: "column",
            overflow: "hidden",
          }}
        >
          {selectedReport
            ? (
              <>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: "1rem",
                  }}
                >
                  <h2 style={{ fontSize: "1.25rem", margin: 0 }}>
                    {selectedReport.reportName}
                  </h2>
                  <button
                    onClick={handleExportExcel}
                    style={{
                      padding: "0.5rem 1rem",
                      background: "#1890ff",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer",
                    }}
                  >
                    Export Excel
                  </button>
                </div>
                {loadingData
                  ? <p className="loading">Executing report...</p>
                  : reportData && reportData.length > 0 &&
                      Object.keys(adaptableOptions).length > 0
                  ? (
                    <div
                      style={{
                        flex: 1,
                        display: "flex",
                        flexDirection: "column",
                        minHeight: 0,
                      }}
                    >
                      <Adaptable.Provider
                        key={selectedReport.reportID}
                        adaptableOptions={adaptableOptions}
                        gridOptions={agGridProps}
                        modules={[AllCommunityModule, AllEnterpriseModule]}
                        onAdaptableReady={(
                          { adaptableApi, agGridApi }: AdaptableReadyInfo,
                        ) => {
                          // save a reference to adaptable api
                          adaptableApiRef.current = adaptableApi;
                          aggridApiRef.current = agGridApi;
                          console.log(
                            "Adaptable grid is initialized and ready.",
                          );
                          //This is responsible for loading the grid.
                          setSaveConfKey(`Report_${selectedReport.reportID}`);
                        }}
                      >
                        <div
                          className="adaptable-container"
                          style={{
                            flex: 1,
                            display: "flex",
                            flexDirection: "column",
                          }}
                        >
                          <Adaptable.UI />
                          <div style={{ flex: 1 }}>
                            <Adaptable.AgGridReact />
                          </div>
                        </div>
                      </Adaptable.Provider>
                    </div>
                  )
                  : reportData && reportData.length === 0
                  ? <p>No data returned for this report.</p>
                  : null}
              </>
            )
            : (
              <p className="loading" style={{ margin: "auto", color: "#999" }}>
                Select a report from the list to view its data.
              </p>
            )}
        </main>
      </div>
    </div>
  );
}

export default App;

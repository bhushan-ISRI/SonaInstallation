import * as React from "react";
import "./userDashboardsc.scss";
import NewAdvanceform from "./NewAdvanceform";
import ViewAdvanceForm from "./ViewAdvanceForm";
import { useState } from "react";
import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import View from "../assets/Eye.png";
import User from "../assets/Userlogo.png";

import ApproverAdvanceForm from "./ApproverAdvanceForm";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface IProps {
  context: any;
  itemId: number;
  formData: any;
}

interface UserDashboardProps {
  context: any;
}

type TabType = "My Request" | "Approved" | "Rejected" | "Paid";

interface IWorkflowHistoryEntry {
  CurrentApprover?: string;
  ActionTaken?: string;
  Comment?: string;
  Date?: string;
}

const parseWorkflowHistory = (raw?: string | null): IWorkflowHistoryEntry[] => {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const userTookAction = (
  workflowHistory: string | null | undefined,
  loggedInUserName: string,
  action: "Approved" | "Reject"
): boolean => {
  const history = parseWorkflowHistory(workflowHistory);
  return history.some(
    (entry) =>
      entry.CurrentApprover?.trim().toLowerCase() === loggedInUserName.toLowerCase() &&
      entry.ActionTaken?.trim().toLowerCase() === action.toLowerCase()
  );
};

const ApproverDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));

  const [formType, setFormType] = useState<"new" | "view" | "approve" | null>(null);
  const [currentUserId, setCurrentUserId] = React.useState<number>(0);
  const [activeMenu, setActiveMenu] = React.useState<TabType>("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  const getLoggedInUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserName(user.Title);
      setCurrentUserId(user.Id);
    } catch (error) {
      console.error("User error:", error);
    }
  };

  const handleApproveClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(item.ID)
        .select("*")();

      setSelectedItem(fullItem);
      setFormType("approve");
      setShowForm(true);
    } catch (error) {
      console.error("Approve error:", error);
    }
  };

  const handleViewClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(item.ID)
        .select("*")();

      setSelectedItem(fullItem);
      setFormType("view");
      setShowForm(true);
    } catch (error) {
      console.error("View error:", error);
    }
  };

  const getCapexData = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle("Installation")
        .items.select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "VendorName",
          "VendorCode",
          "PONumber",
          "TotalamounttobeCapitalized",
          "Status",
          "WorkFlowHistory",
          "CurrentApproverId",
          "CurrentApprover/Title",
          "CurrentApprover/EMail"
        )
        .expand("CurrentApprover")
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        PaymentId: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName || "",
        VendorName: item.VendorName || "",
        VendorCode: item.VendorCode || "",
        PONumber: item.PONumber || "",
        TotalamounttobeCapitalized: item.TotalamounttobeCapitalized || "",
        status: item.Status || "",
        WorkFlowHistory: item.WorkFlowHistory || null,
        CurrentApproverId: item.CurrentApproverId,
        CurrentApprover: item.CurrentApprover?.Title || "",
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  const filteredData = React.useMemo(() => {
    const text = searchText.toLowerCase();
    const status = statusFilter.toLowerCase();
    const loggedInUser = currentUserName.trim().toLowerCase();

    return data.filter((item) => {
      let menuFilter = false;

      if (activeMenu === "My Request") {
        menuFilter =
          item.status?.toLowerCase() === "pending for approval" &&
          item.CurrentApproverId === currentUserId;
      } else if (activeMenu === "Approved") {
        menuFilter = userTookAction(item.WorkFlowHistory, loggedInUser, "Approved");
      } else if (activeMenu === "Rejected") {
        menuFilter = userTookAction(item.WorkFlowHistory, loggedInUser, "Reject");
      } else if (activeMenu === "Paid") {
        menuFilter = item.status?.toLowerCase() === "paid";
      }

      const statusMatch = !status || item.status?.toLowerCase().includes(status);

      const searchMatch =
        !text ||
        item.PaymentId?.toLowerCase().includes(text) ||
        item.EmployeeName?.toLowerCase().includes(text) ||
        item.VendorName?.toLowerCase().includes(text) ||
        item.VendorCode?.toLowerCase().includes(text) ||
        item.PONumber?.toLowerCase().includes(text);

      return menuFilter && statusMatch && searchMatch;
    });
  }, [data, activeMenu, searchText, statusFilter, currentUserName, currentUserId]);

  React.useEffect(() => {
    if (!context) return;
    void getLoggedInUser();
  }, [context]);

  React.useEffect(() => {
    if (currentUserId > 0) {
      void getCapexData();
    }
  }, [currentUserId]);

  if (showForm) {
    if (formType === "approve") {
      return (
        <ApproverAdvanceForm
          context={context}
          formData={selectedItem}
          itemId={selectedItem?.ID}
          onClose={() => {
            setShowForm(false);
            setSelectedItem(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "view") {
      return (
        <ViewAdvanceForm
          context={context}
          formData={selectedItem}
          itemId={selectedItem?.ID}
          onClose={() => {
            setShowForm(false);
            setSelectedItem(null);
          }}
        />
      );
    }
  }

  return (
    <>
      <div style={{ display: "flex", width: "100%" }}>
        <div className="sidebar">
          <div className="sidehead">
            <div className="logo">
              <img src={logo} width="25px" height="25px" />
            </div>
            <div className="sidehead-right">SONA COMSTAR</div>
          </div>

          <div className="sidehead-user">
            <img
              src={User}
              style={{ margin: "10px 20px" }}
              width={20}
              height={20}
            />
            {currentUserName}
          </div>

          <ul className="nav">
            {(["My Request", "Approved", "Rejected", "Paid"] as TabType[]).map((tab) => (
              <li className="nav-item" key={tab}>
                <a
                  className={activeMenu === tab ? "nav-link active" : "nav-link"}
                  onClick={() => setActiveMenu(tab)}
                  style={{ cursor: "pointer" }}
                >
                  {tab}
                </a>
              </li>
            ))}
          </ul>
        </div>

        <div
          className="main"
          style={{ width: "calc(100% - 250px)", transition: "width 0.3s" }}
        >
          <div className="header">
            <div className="left-banner">
              <div className="logo-text">
                <h2>Installation Approver Dashboard</h2>
              </div>
            </div>
          </div>

          <div className="col-md-12 mainsecond">
            <div>
              <input
                placeholder="Search"
                value={searchText}
                className="form-control"
                style={{ width: "250px" }}
                onChange={(e) => setSearchText(e.target.value)}
              />
            </div>
            {activeMenu === "My Request" && (
              <div>
                <select
                  value={statusFilter}
                  className="formtext-control"
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="">All</option>
                  <option value="Pending for Approval">Pending for Approval</option>
                  <option value="Send Back">Send Back</option>
                  <option value="Paid">Paid</option>
                  <option value="Reject">Reject</option>
                  <option value="Save as Draft">Save as Draft</option>
                  <option value="Pending for Vouching Update">Pending for Vouching Update</option>
                  <option value="Pending for UTR Update">Pending for UTR Update</option>
                </select>
              </div>
            )}
          </div>

          <main className="Main-Dash mx-2">
            <div style={{ overflowX: "auto" }}>
              <div className="table-vert-scroll">
                <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                  <thead
                    className="text-white"
                    style={{ backgroundColor: "rgb(60, 62, 69)" }}
                  >
                    <tr>
                      <th className="px-4 py-2">Action</th>
                      <th className="px-4 py-2">Payment ID</th>
                      <th className="px-4 py-2">Requestor Name</th>
                      <th className="px-4 py-2">Requestor Date</th>
                      <th className="px-4 py-2">Requestor Type</th>
                      <th className="px-4 py-2">Vendor Code</th>
                      <th className="px-4 py-2">Vendor Name</th>
                      <th className="px-4 py-2">PO Number</th>
                      <th className="px-4 py-2">Capitalised Amount</th>
                      <th className="px-4 py-2">Pending With</th>
                      <th className="px-4 py-2">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.length === 0 ? (
                      <tr>
                        <td colSpan={11} style={{ textAlign: "center" }}>
                          No Data
                        </td>
                      </tr>
                    ) : (
                      filteredData.map((item, i) => (
                        <tr key={i}>
                          <td className="px-4 py-2">
                            {activeMenu === "My Request" ? (
                              <span
                                onClick={() => handleApproveClick(item)}
                                style={{ cursor: "pointer" }}
                              >
                                <img src={Edit} width={15} alt="Approve" />
                              </span>
                            ) : (
                              <span
                                onClick={() => handleViewClick(item)}
                                style={{ cursor: "pointer" }}
                              >
                                <img src={View} width={15} alt="View" />
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-2">{item.PaymentId}</td>
                          <td className="px-4 py-2">{item.EmployeeName}</td>
                          <td className="px-4 py-2">{item.date}</td>
                          <td className="px-4 py-2">Install commission</td>
                          <td className="px-4 py-2">{item.VendorCode}</td>
                          <td className="px-4 py-2">{item.VendorName}</td>
                          <td className="px-4 py-2">{item.PONumber}</td>
                          <td className="px-4 py-2">
                            ₹ {item.TotalamounttobeCapitalized}
                          </td>
                          <td className="px-4 py-2">
                            {item.CurrentApprover || "-"}
                          </td>
                          <td className="px-4 py-2">{item.status}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </main>
        </div>
      </div>
    </>
  );
};

export default ApproverDashboard;
import * as React from "react";
import "./userDashboardsc.scss";
import NewAdvanceform from "./NewAdvanceform";
import ViewAdvanceForm from "./ViewAdvanceForm";

import { useState } from "react";

import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import View from "../assets/Eye.png";
import User from "../assets/Userlogo.png";

import APperformerAdvanceform from "./APperformerAdvanceform";
import APperformerAdvanceFormForUTR from "./APperformerAdvanceFormForUTR";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface UserDashboardProps {
  context: any;
}

type TabType = "My Request" | "Paid" | "Rejected";

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
  action: "Reject"
): boolean => {
  const history = parseWorkflowHistory(workflowHistory);
  return history.some(
    (entry) =>
      entry.CurrentApprover?.trim().toLowerCase() === loggedInUserName.toLowerCase() &&
      entry.ActionTaken?.trim().toLowerCase() === action.toLowerCase()
  );
};

const APperformerDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));

  const [formType, setFormType] = useState<"approve" | "approveUTR" | "view" | null>(null);
  const [currentUserEmail, setCurrentUserEmail] = React.useState("");
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [activeMenu, setActiveMenu] = React.useState<TabType>("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  const getLoggedInUser = async () => {
    const user = await sp.web.currentUser();
    setCurrentUserName(user.Title);
    setCurrentUserEmail(user.Email.toLowerCase());
    return { email: user.Email.toLowerCase(), name: user.Title };
  };

  const handleApproveClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(item.ID)();

      setSelectedItem(fullItem);

      const status = fullItem.Status?.toLowerCase().trim();

      if (status === "pending for vouching update") {
        setFormType("approve");
      } else if (status === "pending for utr update") {
        setFormType("approveUTR");
      }

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

  const getCapexData = async (userEmail?: string, userName?: string) => {
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
        CurrentApprover: item.CurrentApprover?.Title || "",
        CurrentApproverEmail: item.CurrentApprover?.EMail?.toLowerCase() || "",
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
    const loggedInEmail = currentUserEmail.trim().toLowerCase();

    return data.filter((item) => {
      let menuFilter = false;

      if (activeMenu === "My Request") {
        const isPendingVouching = item.status?.toLowerCase() === "pending for vouching update";
        const isPendingUTR = item.status?.toLowerCase() === "pending for utr update";
        menuFilter =
          (isPendingVouching || isPendingUTR) &&
          item.CurrentApproverEmail === loggedInEmail;
      } else if (activeMenu === "Paid") {
        menuFilter = item.status?.toLowerCase() === "paid";
      } else if (activeMenu === "Rejected") {
        menuFilter = userTookAction(item.WorkFlowHistory, loggedInUser, "Reject");
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
  }, [data, activeMenu, searchText, statusFilter, currentUserName, currentUserEmail]);

  React.useEffect(() => {
    const loadData = async () => {
      if (!context) return;
      const { email, name } = await getLoggedInUser();
      await getCapexData(email, name);
    };
    void loadData();
  }, [context]);

  if (showForm && selectedItem) {
    if (formType === "approve") {
      return (
        <APperformerAdvanceform
          context={context}
          formData={selectedItem}
          itemId={selectedItem.ID}
          onClose={() => {
            setShowForm(false);
            setSelectedItem(null);
            void getCapexData(currentUserEmail, currentUserName);
          }}
        />
      );
    }

    if (formType === "approveUTR") {
      return (
        <APperformerAdvanceFormForUTR
          context={context}
          formData={selectedItem}
          itemId={selectedItem.ID}
          onClose={() => {
            setShowForm(false);
            setSelectedItem(null);
            void getCapexData(currentUserEmail, currentUserName);
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
            {(["My Request", "Paid", "Rejected"] as TabType[]).map((tab) => (
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
                <h2>Installation AP Dashboard</h2>
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
                                <img src={Edit} width={15} alt="Edit" />
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

export default APperformerDashboard;
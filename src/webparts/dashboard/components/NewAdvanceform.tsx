import * as React from "react";
import "./advanced.scss";
import Swal from "sweetalert2";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useState, useRef } from "react";
import logo from "../assets/sona-comstarlogo.png";
import { SPHttpClient } from "@microsoft/sp-http";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

interface IPOData {
  Id: number;
  PONumber: string;
  PODate: string;
  POPaymentTerms: string;
  POAmount: string;
  POBasicAmount: string;
  POGSTAmount: string;
  POOtherAmount: string;
  MRNBasicAmount: string;
  MRNGSTAmount: string;
  MRNOtherAmount: string;
  MRNNumber: string;
  MRNDtae: string;
  MRNAmountwithGST: string;
  RequestedAmountforPayment: string;
  VoucherDate: string;
  VoucherNumber: string;
}

const NewAdvanceform = ({ context, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  const [employee, setEmployee] = React.useState<any>({});
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [advanceHistory, setAdvanceHistory] = useState<any[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [vendorSearch, setVendorSearch] = useState("");
  const [vendorDropdownOpen, setVendorDropdownOpen] = useState(false);
  const vendorDropdownRef = useRef<HTMLDivElement>(null);
  const vendorSearchRef = useRef<HTMLInputElement>(null);

  const filteredVendors = vendors.filter(
    (v) =>
      v.VendorName.toLowerCase().includes(vendorSearch.toLowerCase()) ||
      v.VendorCode.toLowerCase().includes(vendorSearch.toLowerCase()),
  );

  const [poList, setPoList] = useState<IPOData[]>([]);
  const [poLoading, setPoLoading] = useState(false);

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");

  const [poBasicAmount, setPoBasicAmount] = useState("");
  const [poGSTAmount, setPoGSTAmount] = useState("");
  const [poOtherAmount, setPoOtherAmount] = useState("");

  const [mrnBasicAmount, setMrnBasicAmount] = useState("");
  const [mrnGSTAmount, setMrnGSTAmount] = useState("");
  const [mrnOtherAmount, setMrnOtherAmount] = useState("");

  const [assetCodes, setAssetCodes] = useState<string[]>([""]);

  const [gstToBeCapitalized, setGstToBeCapitalized] = useState(false);
  const [showGstTooltip, setShowGstTooltip] = useState(false);
  const [remarks, setRemarks] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [approverDetails, setApproverDetails] = useState<any[]>([]);
  const [approvers, setApprovers] = useState<number[]>([]);
  const [pageLoading, setPageLoading] = useState(true);

  const poAmountTotal = (
    Number(poBasicAmount || 0) +
    Number(poGSTAmount || 0) +
    Number(poOtherAmount || 0)
  ).toFixed(2);

  const mrnAmountTotal = (
    Number(mrnBasicAmount || 0) +
    Number(mrnGSTAmount || 0) +
    Number(mrnOtherAmount || 0)
  ).toFixed(2);

  const paidAmount = Number(advanceAmount || 0).toFixed(2);

  const handleNumberChange = (value: string, setter: any) => {
    if (/^\d*\.?\d*$/.test(value)) setter(value);
  };

  const handleAssetCodeChange = (index: number, value: string) => {
    const updated = [...assetCodes];
    updated[index] = value;
    setAssetCodes(updated);
  };

  const addAssetCode = () => setAssetCodes([...assetCodes, ""]);

  const removeAssetCode = (index: number) => {
    if (assetCodes.length === 1) return;
    setAssetCodes(assetCodes.filter((_, i) => i !== index));
  };

  const getAssetCodesForSave = (): string => {
    return assetCodes
      .map((c) => c.trim())
      .filter((c) => c !== "")
      .join(", ");
  };

  const clearPOAndMRNFields = () => {
    setPoDate("");
    setPoTerms("");
    setPoAmount("");
    setPoBasicAmount("");
    setPoGSTAmount("");
    setPoOtherAmount("");
    setMrnBasicAmount("");
    setMrnGSTAmount("");
    setMrnOtherAmount("");
  };

  const getPaidPOs = async () => {
    setPoLoading(true);
    try {
      const result = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "Id",
          "PONumber",
          "PODate",
          "POPaymentTerms",
          "POAmount",
          "POBasicAmount",
          "POGSTAmount",
          "POOtherAmount",
          "MRNBasicAmount",
          "MRNGSTAmount",
          "MRNOtherAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "RequestedAmountforPayment",
          "VoucherDate",
          "VoucherNumber",
          "Status",
        )
        .filter(`Status eq 'Paid'`)
        .orderBy("Created", false)
        .top(500)();

      const seen = new Set<string>();
      const uniquePOs: IPOData[] = [];

      for (const item of result) {
        if (!item.PONumber || seen.has(item.PONumber)) continue;
        seen.add(item.PONumber);
        uniquePOs.push({
          Id: item.Id,
          PONumber: item.PONumber,
          PODate: item.PODate || "",
          POPaymentTerms: item.POPaymentTerms || "",
          POAmount: item.POAmount || "",
          POBasicAmount:
            item.POBasicAmount != null ? String(item.POBasicAmount) : "0",
          POGSTAmount:
            item.POGSTAmount != null ? String(item.POGSTAmount) : "0",
          POOtherAmount:
            item.POOtherAmount != null ? String(item.POOtherAmount) : "0",
          MRNBasicAmount:
            item.MRNBasicAmount != null ? String(item.MRNBasicAmount) : "0",
          MRNGSTAmount:
            item.MRNGSTAmount != null ? String(item.MRNGSTAmount) : "0",
          MRNOtherAmount:
            item.MRNOtherAmount != null ? String(item.MRNOtherAmount) : "0",
          MRNNumber: item.MRNNumber || "",
          MRNDtae: item.MRNDtae || "",
          MRNAmountwithGST: item.MRNAmountwithGST || "",
          RequestedAmountforPayment: item.RequestedAmountforPayment || "",
          VoucherDate: item.VoucherDate || "",
          VoucherNumber: item.VoucherNumber || "",
        });
      }

      setPoList(uniquePOs);
    } catch (error) {
      console.error("Error fetching PO list from CapexPayment:", error);
      setPoList([]);
    } finally {
      setPoLoading(false);
    }
  };

  const getPastMRNDetails = async (selectedPONumber: string) => {
    if (!selectedPONumber) {
      setPreviousAdvances([]);
      return;
    }
    try {
      const result = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "PONumber",
          "PODate",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "RequestedAmountforPayment",
          "VoucherDate",
          "VoucherNumber",
          "Status",
        )
        .filter(`PONumber eq '${selectedPONumber}' and Status eq 'Paid'`)
        .orderBy("Created", false)();
      setPreviousAdvances(result);
    } catch (error) {
      console.error("Error fetching Past MRN Details:", error);
      setPreviousAdvances([]);
    }
  };

  const getAdvanceHistory = async (selectedPONumber: string) => {
    if (!selectedPONumber) {
      setAdvanceHistory([]);
      return;
    }
    try {
      const result = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",
          "PaidAmount",
          "VouchingNumber",
          "Status",
        )
        .filter(`PONumber eq '${selectedPONumber}'`)
        .orderBy("Created", false)();
      setAdvanceHistory(result);
    } catch (error) {
      console.error("Error fetching Advance History:", error);
      setAdvanceHistory([]);
    }
  };

  const ensureUser = async (email: string): Promise<number> => {
    if (!email) return 0;
    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const response = await context.spHttpClient.post(
        `${webUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ logonName: email }),
        },
      );
      if (!response.ok) {
        console.log("ensureUser failed for:", email);
        return 0;
      }
      const data = await response.json();
      return data.Id || 0;
    } catch (error) {
      console.log("ensureUser error:", email, error);
      return 0;
    }
  };

  const getLoggedInUser = async () => {
    try {
      const toTitleCase = (str: string): string => {
        if (!str) return "";
        return str
          .toLowerCase()
          .split(" ")
          .filter(Boolean)
          .map((w: string) => w.charAt(0).toUpperCase() + w.slice(1))
          .join(" ");
      };

      const cleanLocationForDisplay = (location: string): string => {
        if (!location) return "";
        return location.replace(/^re\s+/i, "").trim();
      };

      const FLOW_URL =
        "https://defaultcb1edbfe8080457d9cae51528f3643.3f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e2bb522aa41443179a72b701b9613471/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=q8b8ADCtK2eKr2f6p3MX7gxmJymPeJbm0mq2M69Rk8E";

      const fetchPage = async (pageNumber: number) => {
        const response = await fetch(FLOW_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ PageSize: 500, PageNumber: pageNumber }),
        });
        if (!response.ok) throw new Error("Failed to fetch employee data");
        return response.json();
      };

      const currentUserEmail = context.pageContext.user.email.toLowerCase();
      let employee: any = null;
      let page = 1;

      while (true) {
        const res = await fetchPage(page);
        const employees = res?.data?.employees || [];
        employee = employees.find(
          (x: any) => x.email?.toLowerCase() === currentUserEmail,
        );
        if (employee) break;
        if (employees.length < 500) break;
        page++;
      }

      if (!employee) {
        console.log("Employee not found.");
        return;
      }

      const attributes = employee.attributes || [];

      const locationAttr = attributes.find(
        (x: any) => x.attributeTypeDescription?.toLowerCase() === "location",
      );
      const departmentAttr = attributes.find(
        (x: any) => x.attributeTypeDescription?.toLowerCase() === "department",
      );
      const hodEmailAttr = attributes.find(
        (x: any) => x.attributeTypeDescription?.toLowerCase() === "hod_email",
      );
      const hodNameAttr = attributes.find(
        (x: any) => x.attributeTypeDescription?.toLowerCase() === "hod name",
      );

      let rmUserId = 0;
      let hodUserId = 0;

      try {
        if (employee.reportingManagerEmail) {
          rmUserId = await ensureUser(employee.reportingManagerEmail);
        }
        if (hodEmailAttr?.attributeTypeUnitDescription) {
          hodUserId = await ensureUser(
            hodEmailAttr.attributeTypeUnitDescription,
          );
        }
      } catch (err) {
        console.log("ensureUser error:", err);
      }

      setEmployee({
        EmployeeCode: employee.employeeCode || "",
        EmployeeName: toTitleCase(employee.employeeName || ""),
        Division: departmentAttr?.attributeTypeUnitDescription || "",
        Location: cleanLocationForDisplay(
          locationAttr?.attributeTypeUnitDescription || "",
        ),
        EmployeeEmail: employee.email || "",
        ContactNo: employee.mobileNo || "",
        EmployeeStatus: employee.employeeStatus || "",
        CostCenter: employee.costCenter || "",
        ReportingManager: {
          Id: rmUserId,
          Title: employee.reportingManagerName || "",
        },
        HOD: {
          Id: hodUserId,
          Title: hodNameAttr?.attributeTypeUnitDescription || "",
        },
      });
    } catch (error) {
      console.error("Error fetching user:", error);
    }
  };

  const buildApprovalFlow = async () => {
    try {
      const baseApprovers: any[] = [];
      if (employee.ReportingManager?.Id) {
        baseApprovers.push({
          Id: employee.ReportingManager.Id,
          Name: employee.ReportingManager.Title,
          Role: "RM",
          Level: 1,
          status: "Pending",
        });
      }
      if (employee.HOD?.Id) {
        baseApprovers.push({
          Id: employee.HOD.Id,
          Name: employee.HOD.Title,
          Role: "HOD",
          Level: 2,
          status: "",
        });
      }
      const matrixData = await sp.web.lists
        .getByTitle("InstallationCommisionApprovalMatrix")
        .items.select(
          "Role/RoleName",
          "Approver/Id",
          "Approver/Title",
          "Level/Level",
        )
        .expand("Role", "Approver", "Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();
      const matrixApprovers = matrixData.map((item: any, index: number) => ({
        Id: item.Approver?.Id,
        Name: item.Approver?.Title,
        Role: item.Role?.RoleName,
        Level: baseApprovers.length + index + 1,
        status: "",
      }));
      const fullFlow = [...baseApprovers, ...matrixApprovers];
      if (fullFlow.length > 0) fullFlow[0].status = "Pending";
      setApproverDetails(fullFlow);
      setApprovers(fullFlow.map((a) => a.Id));
      return fullFlow;
    } catch (error) {
      console.error("Approval Flow Error:", error);
      return [];
    }
  };

  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();
      setVendors(data);
    } catch (error) {
      console.error("Error fetching vendors:", error);
    }
  };

  const getFinancialYear = () => {
    const today = new Date();
    const month = today.getMonth() + 1;
    const year = today.getFullYear();
    return (month >= 4 ? year + 1 : year).toString().slice(-2);
  };

  const generatePaymentId = async (): Promise<string> => {
    try {
      const fy = getFinancialYear();
      const incrementalData = await sp.web.lists
        .getByTitle("InstallationIncrementalID")
        .items.select("Id", "IncrementalID")
        .orderBy("Id", false)
        .top(1)();
      const nextNumber =
        incrementalData.length > 0 && incrementalData[0].IncrementalID
          ? Number(incrementalData[0].IncrementalID) + 1
          : 1;

      await sp.web.lists.getByTitle("InstallationIncrementalID").items.add({
        Title: `INT-${nextNumber}`,
        IncrementalID: nextNumber.toString(),
      });
      return `INT/${fy}/${nextNumber.toString().padStart(5, "0")}`;
    } catch (error) {
      console.error("Generate Payment ID Error:", error);
      return `INT/${getFinancialYear()}/00001`;
    }
  };

  const uploadAttachments = async (PaymentId: string) => {
    if (!attachments || attachments.length === 0) return;
    try {
      const safePaymentId = PaymentId.replace(/\//g, "_");
      const libraryName = "InstallationCommision";
      const webUrl = context.pageContext.web.serverRelativeUrl;
      const folderPath = `${webUrl}/${libraryName}/${safePaymentId}`;
      await sp.web.folders.addUsingPath(`${libraryName}/${safePaymentId}`);
      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }
    } catch (error) {
      console.error("Upload error:", error);
    }
  };

  const validateForm = (): string[] => {
    const errors: string[] = [];
    if (!selectedVendorId) errors.push("Vendor Name is required");
    if (!poNumber) errors.push("PO Number is required");
    if (!poDate) errors.push("PO Date is required");
    if (!advanceAmount || Number(advanceAmount) <= 0)
      errors.push("Total Amount to be Capitalized must be greater than zero");
    if (getAssetCodesForSave() === "")
      errors.push("At least one Asset Code is required");
    if (!attachments || attachments.length === 0)
      errors.push("Please upload at least one attachment");
    return errors;
  };    

  const buildItemPayload = (
    PaymentId: string,
    flow: any[],
    status: string,
    actionTaken: string,
  ) => {
    const history = [
      {
        CurrentApprover: employee.EmployeeName,
        ActionTaken: actionTaken,
        Comment: remarks || actionTaken,
        Date: new Date().toISOString(),
      },
    ];
    const currentApproverId = flow.length > 0 ? flow[0].Id : null;
    const selectedVendor = vendors.find((v) => v.Id === selectedVendorId);
    return {
      Title: PaymentId,
      PaymentId: PaymentId,
      EmployeeCode: employee.EmployeeCode || "",
      EmployeeName: employee.EmployeeName || "",
      Division: employee.Division || "",
      Location: employee.Location || "",
      Email: employee.EmployeeEmail || "",
      ReportingManager: employee.ReportingManager?.Title || "",
      HOD: employee.HOD?.Title || "",
      ContactNo: employee.ContactNo || "",
      EmployeeStatus: employee.EmployeeStatus || "",
      VendorCode: selectedVendor?.VendorCode || "",
      VendorName: selectedVendor?.VendorName || "",
      PONumber: poNumber || "",
      POdate: poDate ? new Date(poDate) : null,
      POPaymentTerms: poTerms || "",
      POAmount: poAmount || "0",
      POAmountBasic: poBasicAmount || "0",
      POAmountGST: poGSTAmount || "0",
      POAmountOther: poOtherAmount || "0",
      MRNAmountGST: mrnGSTAmount || "0",
      MRNAmountOther: mrnOtherAmount || "0",
      MRNAmountTotal: mrnAmountTotal || "0",
      POAmountTotal: poAmountTotal || "0",
      MRNAmountBasic: mrnBasicAmount || "0",
      AssetCodes: getAssetCodesForSave(),
      GSTToBeCapitalized: gstToBeCapitalized ? "Yes" : "No",
      TotalPaymentofProject: advanceAmount || "0",
      GSTAdjustmentifAny: "0",
      OtherAdjustmentifany: "0",
      TotalamounttobeCapitalized: paidAmount || "0",
      Status: status,
      ApprovalMatrix: JSON.stringify(flow),
      CurrentApproverId: currentApproverId,
      WorkFlowHistory: actionTaken === "Saved as draft" ? "" : JSON.stringify(history),
    };
  };

  const handleSubmit = async () => {
    try {
      const errors = validateForm();
      if (errors.length > 0) {
        await Swal.fire({
          title: "Validation",
          html: errors.map((e) => `• ${e}`).join("<br/>"),
          icon: "error",
        });
        return;
      }
      const confirmSubmit = await Swal.fire({
        title: "Submit Request?",
        text: "Are you sure you want to submit this Installation Commissioning Request?",
        icon: "question",
        showCancelButton: true,
        confirmButtonText: "Submit",
        cancelButtonText: "Cancel",
      });
      if (!confirmSubmit.isConfirmed) return;
      Swal.fire({
        title: "Submitting...",
        text: "Please wait",
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading(),
      });
      const PaymentId = await generatePaymentId();
      const flow = await buildApprovalFlow();
      const payload = buildItemPayload(
        PaymentId,
        flow,
        "Pending for Approval",
        "Submitted",
      );
      await sp.web.lists.getByTitle("Installation").items.add(payload);
      await uploadAttachments(PaymentId);
      await Swal.fire({
        title: "Success",
        text: "Request submitted successfully.",
        icon: "success",
        confirmButtonText: "OK",
      });
      onClose();
    } catch (error: any) {
      console.error(error);
      await Swal.fire({
        title: "Submission Failed",
        text: error?.message || "Something went wrong",
        icon: "error",
      });
    }
  };

  const handledraft = async () => {
    try {
      const confirmDraft = await Swal.fire({
        title: "Save as Draft?",
        text: "Do you want to save this request as Draft?",
        icon: "question",
        showCancelButton: true,
        confirmButtonText: "Save Draft",
        cancelButtonText: "Cancel",
      });
      if (!confirmDraft.isConfirmed) return;
      Swal.fire({
        title: "Saving Draft...",
        text: "Please wait",
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading(),
      });
      const PaymentId = await generatePaymentId();
      const flow = await buildApprovalFlow();
      const payload = buildItemPayload(
        PaymentId,
        flow,
        "Save as Draft",
        "Saved as draft",
      );
      await sp.web.lists.getByTitle("Installation").items.add(payload);
      if (attachments.length > 0) await uploadAttachments(PaymentId);
      await Swal.fire({
        title: "Draft Saved",
        text: "Request saved successfully as Draft.",
        icon: "success",
        confirmButtonText: "OK",
      });
      onClose();
    } catch (error: any) {
      console.error(error);
      await Swal.fire({
        title: "Draft Save Failed",
        text: error?.message || "Something went wrong",
        icon: "error",
      });
    }
  };

  const handleExit = async () => {
    const result = await Swal.fire({
      title: "Exit Form?",
      text: "Any unsaved changes will be lost.",
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Yes, Exit",
      cancelButtonText: "Cancel",
    });
    if (result.isConfirmed) onClose();
  };

  React.useEffect(() => {
    if (!context) return;
    void loadPage();
  }, [context]);

  const loadPage = async () => {
    try {
      setPageLoading(true);
      await Promise.all([getLoggedInUser(), getVendors(), getPaidPOs()]);
    } catch (err) {
      console.error(err);
    } finally {
      setPageLoading(false);
    }
  };

  React.useEffect(() => {
    const loadApproval = async () => {
      if (!employee?.EmployeeCode) return;
      setPageLoading(true);
      try {
        await buildApprovalFlow();
      } finally {
        setPageLoading(false);
      }
    };
    void loadApproval();
  }, [employee]);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        vendorDropdownRef.current &&
        !vendorDropdownRef.current.contains(event.target as Node)
      ) {
        setVendorDropdownOpen(false);
        setVendorSearch("");
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  React.useEffect(() => {
    if (vendorDropdownOpen && vendorSearchRef.current) {
      vendorSearchRef.current.focus();
    }
  }, [vendorDropdownOpen]);

  if (pageLoading) {
    return (
      <div className="skeletonWrapper">
        <div className="skeletonLine" style={{ width: "40%" }} />
        <div className="skeletonLine" style={{ width: "80%" }} />
        <div className="skeletonCard"></div>
        <div className="keletonLine" style={{ width: "60%" }} />
        <div className="skeletonCard"></div>
        <div className="skeletonLine" style={{ width: "50%" }} />
        <div className="skeletonCard"></div>
      </div>
    );
  }

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1>Installation Commissioning Request</h1>
            </div>

            <div className="approval-ribbon">
              <div className="ribbon-step current">
                {employee.EmployeeName || "Initiator"}
              </div>
              {approverDetails.map((approver, index) => (
                <div key={index} className="ribbon-step pending">
                  {approver.Name}
                </div>
              ))}
            </div>

            <div className="borderedbox">
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Employee Code</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Name</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Email</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Contact No</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Status</label>{" "}
                    :&nbsp;&nbsp;
                    <label className="fonttext">
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Division</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Location</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">RM</label> :&nbsp;&nbsp;
                    <label className="fonttext">
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">HOD</label> :&nbsp;&nbsp;
                    <label className="fonttext">{employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor &amp; PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <div
                      ref={vendorDropdownRef}
                      style={{ position: "relative" }}
                    >
                      <div
                        className="formtext-control"
                        onClick={() => setVendorDropdownOpen((prev) => !prev)}
                        style={{
                          cursor: "pointer",
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          userSelect: "none",
                          minHeight: "38px",
                          padding: "6px 10px",
                        }}
                      >
                        <span
                          style={{
                            color: selectedVendorName ? "inherit" : "#999",
                          }}
                        >
                          {selectedVendorName || "Select Vendor"}
                        </span>
                        <span style={{ fontSize: "10px", marginLeft: "8px" }}>
                          {vendorDropdownOpen ? "▲" : "▼"}
                        </span>
                      </div>

                      {vendorDropdownOpen && (
                        <div
                          style={{
                            position: "absolute",
                            top: "100%",
                            left: 0,
                            right: 0,
                            zIndex: 1000,
                            backgroundColor: "#fff",
                            border: "1px solid #ccc",
                            borderRadius: "4px",
                            boxShadow: "0 4px 8px rgba(0,0,0,0.15)",
                          }}
                        >
                          <div style={{ padding: "6px" }}>
                            <input
                              ref={vendorSearchRef}
                              type="text"
                              value={vendorSearch}
                              onChange={(e) => setVendorSearch(e.target.value)}
                              placeholder="Search vendor..."
                              className="form-control"
                              style={{ fontSize: "13px" }}
                              onClick={(e) => e.stopPropagation()}
                            />
                          </div>
                          <ul
                            style={{
                              listStyle: "none",
                              margin: 0,
                              padding: 0,
                              maxHeight: "200px",
                              overflowY: "auto",
                            }}
                          >
                            <li
                              onClick={() => {
                                setSelectedVendorId(null);
                                setSelectedVendorName("");
                                setSelectedVendorCode("");
                                setVendorDropdownOpen(false);
                                setVendorSearch("");
                              }}
                              style={{
                                padding: "8px 12px",
                                cursor: "pointer",
                                color: "#999",
                                borderBottom: "1px solid #f0f0f0",
                              }}
                              onMouseEnter={(e) =>
                                ((
                                  e.currentTarget as HTMLLIElement
                                ).style.backgroundColor = "#f5f5f5")
                              }
                              onMouseLeave={(e) =>
                                ((
                                  e.currentTarget as HTMLLIElement
                                ).style.backgroundColor = "transparent")
                              }
                            >
                              Select Vendor
                            </li>
                            {filteredVendors.length === 0 ? (
                              <li
                                style={{
                                  padding: "8px 12px",
                                  color: "#999",
                                  fontSize: "13px",
                                }}
                              >
                                No vendors found
                              </li>
                            ) : (
                              filteredVendors.map((v) => (
                                <li
                                  key={v.Id}
                                  onClick={() => {
                                    setSelectedVendorId(v.Id);
                                    setSelectedVendorName(v.VendorName);
                                    setSelectedVendorCode(v.VendorCode);
                                    setVendorDropdownOpen(false);
                                    setVendorSearch("");
                                  }}
                                  style={{
                                    padding: "8px 12px",
                                    cursor: "pointer",
                                    fontSize: "13px",
                                    backgroundColor:
                                      selectedVendorId === v.Id
                                        ? "#e8f0fe"
                                        : "transparent",
                                    borderBottom: "1px solid #f0f0f0",
                                  }}
                                  onMouseEnter={(e) => {
                                    if (selectedVendorId !== v.Id)
                                      (
                                        e.currentTarget as HTMLLIElement
                                      ).style.backgroundColor = "#f5f5f5";
                                  }}
                                  onMouseLeave={(e) => {
                                    if (selectedVendorId !== v.Id)
                                      (
                                        e.currentTarget as HTMLLIElement
                                      ).style.backgroundColor = "transparent";
                                  }}
                                >
                                  {v.VendorName}
                                </li>
                              ))
                            )}
                          </ul>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <input
                      value={selectedVendorCode}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>

                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <select
                      value={poNumber}
                      className="formtext-control"
                      disabled={poLoading}
                      onChange={(e) => {
                        const val = e.target.value;
                        const selectedPO = poList.find(
                          (item) => item.PONumber === val,
                        );
                        setPoNumber(val);
                        if (selectedPO) {
                          setPoDate(
                            selectedPO.PODate
                              ? new Date(selectedPO.PODate)
                                  .toISOString()
                                  .split("T")[0]
                              : "",
                          );
                          setPoTerms(selectedPO.POPaymentTerms || "");
                          setPoAmount(selectedPO.POAmount || "");
                          setPoBasicAmount(selectedPO.POBasicAmount || "0");
                          setPoGSTAmount(selectedPO.POGSTAmount || "0");
                          setPoOtherAmount(selectedPO.POOtherAmount || "0");
                          setMrnBasicAmount(selectedPO.MRNBasicAmount || "0");
                          setMrnGSTAmount(selectedPO.MRNGSTAmount || "0");
                          setMrnOtherAmount(selectedPO.MRNOtherAmount || "0");
                        } else {
                          clearPOAndMRNFields();
                        }
                        void getPastMRNDetails(val);
                        void getAdvanceHistory(val);
                      }}
                    >
                      <option value="">
                        {poLoading
                          ? "Loading PO Numbers..."
                          : poList.length === 0
                            ? "No Paid POs found"
                            : "Select PO Number"}
                      </option>
                      {poList.map((item) => (
                        <option key={item.Id} value={item.PONumber}>
                          {item.PONumber}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      type="date"
                      value={poDate}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input
                      value={poTerms}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (Incl. GST)</label>
                    <input
                      value={poAmount}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Basic Amount</label>
                    <input
                      value={poBasicAmount}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO GST Amount</label>
                    <input
                      value={poGSTAmount}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Other Amount</label>
                    <input
                      value={poOtherAmount}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Basic Amount</label>
                    <input
                      value={mrnBasicAmount}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN GST Amount</label>
                    <input
                      value={mrnGSTAmount}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Other Amount</label>
                    <input
                      value={mrnOtherAmount}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Amount Total</label>
                    <input
                      value={poAmountTotal}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount Total</label>
                    <input
                      value={mrnAmountTotal}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font" style={{ color: "#000000" }}>
                      Total Amount to be Capitalized
                    </label>
                    <input
                      value={advanceAmount}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setAdvanceAmount)
                      }
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div
                    className="col-md-4"
                    style={{
                      display: "flex",
                      alignItems: "flex-end",
                      paddingBottom: "4px",
                    }}
                  >
                    <div className="gst-capitalized-row">
                      <input
                        type="checkbox"
                        id="gstCapitalized"
                        checked={gstToBeCapitalized}
                        onChange={(e) =>
                          setGstToBeCapitalized(e.target.checked)
                        }
                        className="gst-checkbox"
                      />
                      <label
                        htmlFor="gstCapitalized"
                        className="gst-checkbox-label font"
                      >
                        Whether GST to be Capitalized
                      </label>
                      <span
                        className="info-icon"
                        onMouseEnter={() => setShowGstTooltip(true)}
                        onMouseLeave={() => setShowGstTooltip(false)}
                      >
                        &#9432;
                        {showGstTooltip && (
                          <span className="info-tooltip">
                            Info not added yet!
                          </span>
                        )}
                      </span>
                    </div>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">
                      Asset Code(s) <span style={{ color: "red" }}>*</span>
                    </label>
                    <div className="asset-codes-container">
                      {assetCodes.map((code, index) => (
                        <div key={index} className="asset-code-row">
                          <input
                            value={code}
                            className={`form-control asset-code-input${code.trim() === "" ? " input-error" : ""}`}
                            placeholder={`Asset Code ${index + 1}`}
                            onChange={(e) =>
                              handleAssetCodeChange(index, e.target.value)
                            }
                          />
                          <button
                            type="button"
                            className="asset-code-remove-btn"
                            onClick={() => removeAssetCode(index)}
                            disabled={assetCodes.length === 1}
                            title="Remove"
                          >
                            &times;
                          </button>
                          {index === assetCodes.length - 1 && (
                            <button
                              type="button"
                              className="asset-code-add-btn"
                              onClick={addAssetCode}
                              title="Add another asset code"
                            >
                              + Add
                            </button>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Past MRN Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th className="px-4 py-2">PO Number</th>
                            <th className="px-4 py-2">PO Date</th>
                            <th className="px-4 py-2">PO Amount</th>
                            <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">MRN Date</th>
                            <th className="px-4 py-2">MRN Amount</th>
                            <th className="px-4 py-2">Advance Adjustment</th>
                            <th className="px-4 py-2">Paid Amount</th>
                            <th className="px-4 py-2">Remarks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previousAdvances.length === 0 ? (
                            <tr>
                              <td colSpan={9} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => (
                              <tr key={index}>
                                <td className="px-4 py-2">{item.PONumber}</td>
                                <td className="px-4 py-2">
                                  {item.PODate
                                    ? new Date(item.PODate).toLocaleDateString()
                                    : ""}
                                </td>
                                <td className="px-4 py-2">{item.POAmount}</td>
                                <td className="px-4 py-2">{item.MRNNumber}</td>
                                <td className="px-4 py-2">
                                  {item.MRNDtae
                                    ? new Date(
                                        item.MRNDtae,
                                      ).toLocaleDateString()
                                    : ""}
                                </td>
                                <td className="px-4 py-2">
                                  {item.MRNAmountwithGST}
                                </td>
                                <td className="px-4 py-2"></td>
                                <td className="px-4 py-2">
                                  {item.RequestedAmountforPayment}
                                </td>
                                <td className="px-4 py-2"></td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Advance History (to be PO Specific)</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th className="px-4 py-2">PO Number</th>
                            <th className="px-4 py-2">Previous Advance</th>
                            <th className="px-4 py-2">Amount Requested Date</th>
                            <th className="px-4 py-2">Amount Paid Date</th>
                            <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">Settled Amount</th>
                            <th className="px-4 py-2">Pending Advance</th>
                          </tr>
                        </thead>
                        <tbody>
                          {advanceHistory.length === 0 ? (
                            <tr>
                              <td colSpan={7} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            advanceHistory.map((item: any, index: number) => {
                              const pending = Math.max(
                                0,
                                Number(item.RequestAdvanceAmount || 0) -
                                  Number(item.PaidAmount || 0),
                              );
                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">
                                    {item.RequestAdvanceAmount}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(
                                          item.Created,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(
                                          item.VoucherDate,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>
                                  <td className="px-4 py-2"></td>
                                  <td className="px-4 py-2">
                                    {item.PaidAmount}
                                  </td>
                                  <td className="px-4 py-2">{pending}</td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Document</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachment</label>
                    <input
                      type="file"
                      className="form-control"
                      multiple
                      onChange={(e) => {
                        if (e.target.files)
                          setAttachments(Array.from(e.target.files));
                      }}
                    />
                  </div>
                </div>
              </div>

              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "5px",
                  marginBottom: "1rem",
                  marginTop: "1rem",
                }}
              >
                <button
                  type="button"
                  onClick={() => void handleSubmit()}
                  className="submit-btn"
                >
                  Submit
                </button>
                <button
                  type="button"
                  onClick={() => void handledraft()}
                  className="Rework-btn"
                >
                  Save as Draft
                </button>
                <button
                  type="button"
                  onClick={() => void handleExit()}
                  className="reset-btn"
                >
                  Exit
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default NewAdvanceform;

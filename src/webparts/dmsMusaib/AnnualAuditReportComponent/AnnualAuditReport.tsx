import * as React from 'react';
import type { IAnnualAuditReportProps } from './IAnnualAuditReportProps';
import { escape } from '@microsoft/sp-lodash-subset';

import Provider from '../../../GlobalContext/provider';

import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';

import HorizontalNavbar from '../../horizontalNavBar/components/HorizontalNavBar';

import { getSP } from '../../dmsMusaib/loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import UserContext from "../../../GlobalContext/context";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import "./annualauditReport.scss";
import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import Select from "react-select";
import Swal from 'sweetalert2';
import { FormSubmissionMode } from '../../../Shared/Interfaces';
import { decryptId } from '../../../APISearvice/CryptoService';
import { WorkflowAction } from '../../../CustomJSComponents/WorkflowAction/WorkflowAction';
import { WorkflowAuditHistory } from '../../../CustomJSComponents/WorkflowAuditHistory/WorkflowAuditHistory';
import { CONTENTTYPE_AuditReport, CONTENTTYPE_AuditReportTemp, LIST_TITLE_AuditReport, SITE_URL, Tenant_URL } from '../../../Shared/Constants';
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Modal } from 'react-bootstrap';
import { faDownload, faEye } from '@fortawesome/free-solid-svg-icons';
import CustomBreadcrumb from './CustomBreadcrumb/CustomBreadcrumb';
import { addAllProcessItem, addItem, addItem2, getAllApprovedAuditplan, getAllAuditType, getAllDepartment, getAllDepartment1, getAllProcessData, getApprovalByID, getApprovalByID2, getDataRoles, getDocumentLinkByID, getDocumentLinkByIDPlan, getDraftApprovalByID, getFormNameID, getGeneratedTemplateDocCR, getItemByID, getItemByID2, getItemfromChecklistMaster, getLatestChangeRequestTemplateType, getListNameID, UpdateAllProcessItem, updateApprovalItem, updateItem, updateItem2, uploadAllFiles } from './AuditReportService';
import { TextField } from '@fluentui/react';
import { isMac } from 'office-ui-fabric-react';
import moment from 'moment';

// let myloader = '../../'
let newfileupload: any
let newfilepreview: any;
let filechanged: boolean = false;
interface ForwardTo {
    id: number;
    role: number;
    level: number;
    approvers: any[]; // Or a more specific type like `string[]` or `SPUser[]`
    approvalType: string;
}

const AnnualAuditReportContext = ({ props }: any) => {
    const sp: SPFI = getSP();
    const elementRef = React.useRef<HTMLDivElement>(null);
    const siteUrl = props.siteUrl;
    const { useHide }: any = React.useContext(UserContext);
    const [InputDisabled, setInputDisabled] = React.useState(false);
    const selectedTextDiv = document.getElementById('selectedText');

    selectedTextDiv.style.display = 'none';


    const [FilesArr, setFilesArr] = React.useState<any>([]);
    const [FilesArr1, setFilesArr1] = React.useState<any>([]);
    const [FilesArrDoclink, setFilesArrDoclink] = React.useState<any>([]);
    const [Loading, setLoading] = React.useState(false);
    const [FormLoading, setFormLoading] = React.useState(false);
    const [showForwardapproval, setshowForwardapproval] = React.useState(true);
    const [FormItemId, setFormItemId] = React.useState(null);
    const [editID, setEditID] = React.useState(null);
    const [editItemID, setEditItemID] = React.useState(null);
    const [MainEditItem, setMainEditItem] = React.useState(null);
    const [rows, setRows] = React.useState<any>([]);
    const [UserRoles, setUserRoles] = React.useState<any>([]);
    const [rows1, setRows1] = React.useState<any>([]);
    const [currentUser, setCurrentUser] = React.useState(null);
    const [selectedOption, setSelectedOption] = React.useState(null);
    const [selectedPeople, setSelectedPeople] = React.useState(null);
    const [selectedRole, setSelectedRole] = React.useState(null);
    const [ValidDRecomm, setValidDRecomm] = React.useState(true);
    const [ValidDraft, setValidDraft] = React.useState(true);
    const [ValidAudit, setValidAudit] = React.useState(true);
    const [ValidSubmit, setValidSubmit] = React.useState(true);
    const [ValidCancelReason, setValidCancelReason] = React.useState(true);
    const [ValidForwardTo, setValidForwardTo] = React.useState(true);
    const [RequesterRoleId, setRequesterRoleId] = React.useState(null);
    const [RequestTypeId, setRequestTypeId] = React.useState(null);
    const [FormNameId, setFormNameId] = React.useState(null);
    const [AuditPlanType, setAuditPlanType] = React.useState([]);
    const [editForm, setEditForm] = React.useState(false);
    const [modeValue, setmode] = React.useState("");
    const [showviewdownload, setshowviewdownload] = React.useState(true);
    const [currentUserDept, setcurrentUserDept] = React.useState("");
    const [selectUserDept, setselectUserDept] = React.useState(null);
    const [selectAuditplan, setselectAuditplan] = React.useState(null);
    const [doccode, setdoccode] = React.useState("");
    const [AllDept, setAllDept] = React.useState([]);
    const [RowErrors, setRowErrors] = React.useState([]);
    const [DocumentLink, setDocumentLink] = React.useState(null);
    const [DraftApprovalItem, setDraftApprovalItem] = React.useState(null);
    // const [cancellReason, setcancellReason] = React.useState([{ id: 0, description: "", reason: "" }]);
    // const [RecommendRows, setRecommendRows] = React.useState([]);
    const [showModal, setShowModal] = React.useState(false);
    const [ShowModalpre, setShowModalpre] = React.useState(false);
    //error
    const [dateerr, setdateerr] = React.useState(false);
    const [isoreferenceerr, setisoreferenceerr] = React.useState(false);
    const [departmenterr, setdepartmenterr] = React.useState(false);
    const [imsprocedureerr, setimsprocedureerr] = React.useState(false);
    const [inquirieserr, setinquirieserr] = React.useState(false);
    const [timeerr, settimeerr] = React.useState(false);
    const [auditorcommentserr, setauditorcommentserr] = React.useState(false);
    const [attachmenterr, setattachmenterr] = React.useState(false);
    const [sharewitherr, setsharewitherr] = React.useState(false);
    const [approvedauditplanerr, setapprovedauditplanerr] = React.useState(false);
    const [TemplateDoc, setTemplateDoc] = React.useState<any>([]);
    const [sharewithusers, setSharewithusers] = React.useState([]);
    //error end
    const [formData, setFormData] = React.useState({
        approvedauditplanId: 0,
        deptId: 0,
        date: "",
        documentCode: "",
        issueNo: "",
        revisionNo: "",
        Status: "",
        revisionDate: null,
        issueDate: null,
        referenceNo: "",
        documentLink: "",
        attachment: null,
        attachmentIds: null,
        attachmentJson: null

    });
    const [selectCCUsers, setSelectCCUsers] = React.useState([]);
    const [ListNameId, setListNameId] = React.useState(null);

    const [selectToUsers, setSelectToUsers] = React.useState([]);

    const handleDepartmentChange = (selectedOption: any) => {
        setselectUserDept(selectedOption);
        setFormData({ ...formData, deptId: selectedOption.value });
    };


    // ////// Recommendation
    const [recommendationRows, setRecommendationRows] = React.useState([
        { id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }
    ]);

    const [recommendationRowsEdit, setRecommendationRowsEdit] = React.useState([]);

    const handleAddRecommendationRow = () => {
        setRecommendationRows([...recommendationRows, { id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }]);
    };

    const handleRecommendationChange = (index: number, field: string, value: any) => {
        debugger
        let updatedRows;
        if (field == "sharewith") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = recommendationRows.map((row, i) =>
                i === index ? { ...row, [field]: value, sharewithIds: value.value } : row
            );
        } else {
            updatedRows = recommendationRows.map((row, i) =>
                i === index ? { ...row, [field]: value } : row
            );
        }

        setRecommendationRows(updatedRows);
    };

    const handleDeleteRecommendationRow = (index: number) => {
        const updatedRows = recommendationRows.filter((_, i) => i !== index);
        setRecommendationRows(updatedRows);
    };
    // //////


    // Handle change event
    const handleChange = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        const updatedArr = forwardToArr.map(row =>
            row.level === lvl ? { ...row, approvalType: event.target.value } : row
        );
        //   setApprovalType(event.target.value);
        setForwardToArr(updatedArr);
    };


    const [forwardToArr, setForwardToArr] = React.useState<ForwardTo[]>([
        { id: 0, role: 0, level: 1, approvers: [], approvalType: "One" } // Default row
    ]);

    const Breadcrumb = [
        {
            MainComponent: "My Request",
            MainComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx`,
        },
        {
            ChildComponent: "IMS Audit Report and Checklist",
            ChildComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx#/AnnualAuditReport`,
        },
    ];
    const [forwardToArrEdit, setForwardToArrEdit] = React.useState<ForwardTo[]>([]);
    const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);

    // const [remark, setRemark] = React.useState("");

    // Function to handle People Picker selection
    const onPeoplePickerChange = (items: any[]) => {
        setSelectedUsers(items);
    };

    const ApiCallFunc = async () => {
        setAuditPlanType(await getAllAuditType(sp));
        var DepartmentArr = await getAllDepartment(sp);
        DepartmentArr.sort((a, b) => a.Department.localeCompare(b.Department));
        const optionsDepartment = DepartmentArr.map((item: any) => ({
            value: item.ID,
            label: item.Department,
            itemId: item.ID,
            department: item.Department,
            departmentcode: item.DepartmentCode
        }));
        setAllDept(optionsDepartment);
        var setAllDept1 = await getAllDepartment1(sp);

        var DocCodeArr = await getAllApprovedAuditplan(sp);
        console.log("DocCodeArrDocCodeArr", DocCodeArr);
        const options = DocCodeArr.map((item: any) => ({
            value: item.ID,
            label: item.MemoNumber,
            ApprovedAuditPlanId: item.ID,
            //IssueNumber: item.IssueNumber,
            ReferenceNumber: item.ReferenceNumber,
            //RevisionNumber: item.RevisionNumber,
            SubmiitedDate: item.SubmiitedDate,
            SubmitStatus: item.SubmitStatus,
            MemoNumber: item.MemoNumber,
            DepartmentId: item.DepartmentId,
            AttachmentId: item.AttachmentId,
            AttachmentJson: item.AttachmentJson,
            ID: item.ID,
            Status: item.Status,

        }));
        setRows(options);

        const path1 = window.location.href;

        if (path1.includes("/view/") || path1.includes("/approve/")) {
            setFormLoading(true); ////
            setInputDisabled(true);
            setshowForwardapproval(false)
        }
        else {
            setInputDisabled(false);
            setshowForwardapproval(true)
        }
        if (path1.includes("/edit/")) {
            setFormLoading(true); ////
            setshowForwardapproval(true)
        }
        const rowData: any[] = await getItemfromChecklistMaster(sp) //baseUrl
        if (rowData.length > 0) {
            const initialRows = rowData.map((item: any) => ({
                id: item.Id,
                isoreference: item.ISOReference,
                imsprocedure: item.IMSProcedure,
                inquiries: item.Inquiries,
                auditorcomments: item.AuditorsComments,
                time: item.Time,
                sharewithIds: item.SharewithId,
                //endTime: "",
                sharewith: item.Sharewith ? { label: item.Sharewith.Title, value: item.Sharewith.ID } : null // Convert single object
            }));
            setRecommendationRows(initialRows);
            setRecommendationRowsEdit(initialRows);
        } else {
            setRecommendationRowsEdit([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
            setRecommendationRows([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
        }

        const Currusers: any = await getCurrentUser(sp, siteUrl);
        setCurrentUser(await getCurrentUser(sp, siteUrl));
        const userProfile = await sp.profiles.myProperties();
        setcurrentUserDept(userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "")
        const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
        setselectUserDept(setAllDept1.filter(user => user.label === UserDept));
        setFormData({ ...formData, deptId: setAllDept1.filter(user => user.label === UserDept)[0]?.value });
        const AllUserRoles = await getDataRoles(sp);
        const setRolesValue = AllUserRoles.map((item: any) => ({
            value: item.Id,
            label: item.Role,

        }));

        setUserRoles(setRolesValue)

        const users = await sp.web.siteUsers();
        const people = users.filter(user => user.PrincipalType === PrincipalType.User);

        const Selectedoptions = people.map(item => ({
            value: item.Id,
            label: item.Title,
            UserName: item.Title,
            UserEmail: item.Email
        }));

        setRows1(Selectedoptions);

        setFormData(prevData => ({
            ...prevData,
            from: Currusers?.Id || 0,
            fromEmail: Currusers?.Email,

        }));
        let ChangeRequestTemplateType = await getLatestChangeRequestTemplateType(sp, CONTENTTYPE_AuditReportTemp);
        debugger
        if (ChangeRequestTemplateType.length > 0) {
            const template = ChangeRequestTemplateType[0];
            console.log("template", template);
            setFormData(prevData => ({
                ...prevData,
                documentCode: template.DocumentCode || "",
                revisionNo: template.RevisionNumber,
                issueNo: template.IssueNumber,
                referenceNo: template.ReferenceNumber || "",
                revisionDate: template.RevisionDate == null ? null : new Date(template.RevisionDate).toLocaleDateString("en-CA"),
                issueDate: template.IssueDate == null ? null : new Date(template.IssueDate).toLocaleDateString("en-CA"),
            }));
        }


        let formitemid;
        //#region getdataByID
        if (sessionStorage.getItem("AuditReportId") != undefined) {
            const iD = sessionStorage.getItem("AuditReportId")
            let iDs = decryptId(iD)
            formitemid = Number(iDs);
            setFormItemId(Number(iDs))
        }
        else {


            const path = window.location.href;
            const segments = path.split('/').filter(Boolean); // Remove empty elements

            // Check if "edit" or "view" exists in the URL
            const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");


            if (paramIndex !== -1 && segments[paramIndex + 1]) {
                setmode(segments[paramIndex])
                // mode = segments[paramIndex]; // Will be "edit" or "view"
                formitemid = segments[paramIndex + 1]; // Get the ID
                if (segments[paramIndex + 2] !== undefined) {
                    // var ProcessListItem={
                    //     Status:"",
                    //     Level:0,
                    //     CurrentUserRole:"",

                    // }

                    //  ProcessListItem =await getApprovalByID(sp, Number(segments[paramIndex + 2]),CONTENTTYPE_DocumentCancel);
                    // setInputDisabled((ProcessListItem.Status == "Pending" || ProcessListItem?.Status === "Save as draft") && ProcessListItem.Level === 0 && ProcessListItem.CurrentUserRole !=="OES")
                    setEditID(await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditReport));
                    // var ProcessItemId: any = await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditReport);
                    setInputDisabled(await getApprovalByID2(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditReport));
                }
                // else {

                //     setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditReport))

                // }
            }

            setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditReport))

        }
        // formitemid =20;
        if (formitemid) {

            setEditItemID(Number(formitemid));

            const setBannerById = await getItemByID(sp, Number(formitemid))

            if (setBannerById.length > 0) {
                debugger
                setEditForm(true);
                setMainEditItem(setBannerById[0]);


                // const valuesOnly = selectedOptions.map((option: any) => option.value);
                // setFormData({ ...formData, [fieldName]: valuesOnly });

                setFormData(prevData => ({
                    ...prevData,
                    // memoNo: "",
                    deptId: setBannerById[0].DepartmentId,
                    issueNo: setBannerById[0].IssueNumber,
                    revisionNo: setBannerById[0].RevisionNumber,
                    revisionDate: setBannerById[0].RevisionDate == null ? null : new Date(setBannerById[0].RevisionDate).toLocaleDateString("en-CA") || null,
                    issueDate: setBannerById[0].IssueDate == null ? null : new Date(setBannerById[0].IssueDate).toLocaleDateString("en-CA") || null,
                    date: new Date(setBannerById[0].Date).toLocaleDateString("en-CA"),
                    approvedauditplanId: setBannerById[0].ApprovedAuditPlanId,
                    documentcode: setBannerById[0].DocumentCode,
                    referenceNo: setBannerById[0].ReferenceNumber,
                    AnnualAuditPlanDocumentLinkId: setBannerById[0].AnnualAuditPlanDocumentLinkId,
                    SubmiitedDate: setBannerById[0].SubmiitedDate,
                    submitstatus: setBannerById[0].SubmitStatus,
                    Status: setBannerById[0].Status,
                    documentname: setBannerById[0].DocumentName,
                    isrework: setBannerById[0].IsRework,
                    attachmentIds: setBannerById[0].AttachmentId || null,
                    attachmentJson: setBannerById[0].AttachmentJson || null
                }));
                let sharewithuser = setBannerById[0].Sharewith?.map((approver: any) => ({
                    value: approver.ID,
                    label: approver.Title,
                    UserName: approver.Title,
                    //UserEmail: approver.EMail
                })) || [];
                setSharewithusers(sharewithuser);
                setdoccode(setBannerById[0].Title);
                debugger
                setselectUserDept(setAllDept1.filter(user => user.value === setBannerById[0].DepartmentId));
                const selectedauditplan = options.filter((cust: { value: any; }) => cust.value === setBannerById[0].ApprovedAuditPlanId) || null;
                setselectAuditplan(selectedauditplan);
                setSelectedOption(selectedauditplan);
                console.log(" setBannerById[0].AttachmentId if", setBannerById[0], selectedauditplan)
                if (setBannerById[0].AttachmentId) {
                    // setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].AttachmentId));
                    let arrn = await getDocumentLinkByID(sp, setBannerById[0].AttachmentId);
                    setFilesArr([...FilesArr, ...arrn]);
                    setFilesArr1([...FilesArr1, ...arrn]);

                }
                setTemplateDoc(await getGeneratedTemplateDocCR(sp, Number(formitemid)));
                if (setBannerById[0].AnnualAuditPlanDocumentLinkId.length > 0) {
                    let arrn = await getDocumentLinkByIDPlan(sp, setBannerById[0].AnnualAuditPlanDocumentLinkId);
                    //let arraynew: any[];
                    //arraynew.push(arrn)
                    console.log("arrrrrrn5ghgh6", arrn);
                    setFilesArrDoclink([...FilesArrDoclink, ...arrn]);
                    console.log("arrrrrrn5ghjghj6 after", arrn);
                    //setAttachmentarr(arrn);
                    // setDocumentLink(await getDocumentLinkByIDPlan(sp, selectedList.AttachmentId[0]))
                }
                const ApprowData: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_AuditReport, setBannerById[0].ReferenceNumber)

                if (ApprowData.length > 0) {

                    const EditApprowData = ApprowData.map((item: any) => ({
                        id: item.ID,
                        role: item.ApproverRole?.Id || 0, // Assuming role comes from ApproverRole
                        level: item.Level || 1, // Default to 1 if missing
                        approvalType: item.LevelType,
                        approvers: item.Approvers?.map((approver: any) => ({
                            value: approver.Id,
                            label: approver.Title,
                        })) || []

                    }));
                    setForwardToArr(EditApprowData);
                    setForwardToArrEdit(EditApprowData);

                }
                const rowData1: any[] = await getItemByID2(sp, Number(formitemid));
                const rowData: any[] = await getItemfromChecklistMaster(sp) //baseUrl
                if (rowData1.length > 0) {
                    const initialRows1 = rowData1.map((item: any) => ({
                        id: item.Id,
                        isoreference: item.ISOReference,
                        imsprocedure: item.IMSProcedure,
                        inquiries: item.Inquiries,
                        auditorcomments: item.AuditorsComments,
                        time: item.Time,
                        sharewithIds: item.SharewithId,
                        //endTime: "",
                        sharewith: item.Sharewith ? { label: item.Sharewith.Title, value: item.Sharewith.ID } : null // Convert single object
                    }));
                    setRecommendationRows(initialRows1);
                    setRecommendationRowsEdit(initialRows1);
                } else {
                    if (rowData.length > 0) {
                        const initialRows = rowData.map((item: any) => ({
                            id: item.Id,
                            isoreference: item.ISOReference,
                            imsprocedure: item.IMSProcedure,
                            inquiries: item.Inquiries,
                            auditorcomments: item.AuditorsComments,
                            time: item.Time,
                            sharewithIds: item.SharewithId,
                            endTime: "",
                            sharewith: item.Sharewith ? { label: item.Sharewith.Title, value: item.Sharewith.ID } : null // Convert single object
                        }));
                        setRecommendationRows(initialRows);
                        setRecommendationRowsEdit(initialRows);
                    } else {
                        setRecommendationRowsEdit([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
                        setRecommendationRows([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
                    }
                }



            }





        }
        setFormLoading(false);

        // setRequesterRoleId(await getRequesterID(sp))
        setFormNameId(await getFormNameID(sp, CONTENTTYPE_AuditReport))
        setListNameId(await getListNameID(sp, LIST_TITLE_AuditReport))

        //}
        //#endregion


    };

    const onSelectsharewith = (selectedOptions: any) => {
        const newSelections = selectedOptions || [];
        const allOptions = [...sharewithusers, ...newSelections];

        const uniqueOptions = allOptions.filter(
            (option, index, self) =>
                index === self.findIndex((o) => o.value === option.value)
        );

        setSharewithusers(uniqueOptions);

    };
    const onSelectApprovers = (selectedOptions: any, lvl: number) => {
        setForwardToArr((prev) =>
            prev.map((row) =>
                row.level === lvl ? { ...row, approvers: selectedOptions || [] } : row
            )
        );
    };



    const onSelectRole = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        const updatedArr = forwardToArr.map(row =>
            row.level === lvl ? { ...row, role: Number(event.target.value) } : row
        );
        setForwardToArr(updatedArr);
        setSelectedRole(forwardToArr.map(r => r.role).filter(role => role))
    };


    const handleAddRow = () => {
        setForwardToArr((prev) => [
            ...prev,
            { id: 0, role: 0, level: prev.length + 1, approvers: [], approvalType: "One" }
        ]);
    };

    const handleDeleteRow = (index: number) => {
        const updatedRows = forwardToArr
            .filter((_, i) => i !== index) // Remove selected row
            .map((row, newIndex) => ({ ...row, level: newIndex + 1 })); // Reassign levels

        setForwardToArr([...updatedRows]); // Ensure a new array reference
    };



    React.useEffect(() => {

        ApiCallFunc();

    }, [useHide]);

    const handleCancel = () => {
        // window.location.reload();
        window.history.back();
        // window.location.reload();
        setTimeout(() => {
            location.reload();
        }, 100);

    }



    const OpenFile = (obj: any, sts: string) => {
        debugger
        console.log("obbbj", obj)
        const fileUrl = `${Tenant_URL}${obj.FileRef}`;

        if (sts == "Open") {
            if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {

                window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=default`, "_blank");
            } else {
                window.open(fileUrl, "_blank"); // Open PDF and other files normally
            }

        } else if (sts == "Download") {
            const link = document.createElement("a");
            link.href = fileUrl;
            link.setAttribute("download", obj.FileLeafRef); // Suggests a filename for download
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        }



    }

    const onSelectDocCode = async (selectedList: any) => {
        debugger
        console.log(selectedList, "selectedList");
        if (selectedList != null) {
            setLoading(true);
            setFormData(prevData => ({
                ...prevData,
                //issueNo: selectedList.IssueNumber != "" || selectedList.IssueNumber != null ? selectedList.IssueNumber : 0,
                //referenceNo: selectedList.ReferenceNumber != "" || selectedList.ReferenceNumber != null ? selectedList.ReferenceNumber : 0,
                //revisionNo: selectedList.RevisionNumber != "" || selectedList.RevisionNumber != null ? selectedList.RevisionNumber : 0,
                deptId: selectedList.DepartmentId,
                attachmentIds: selectedList.AttachmentId,
                attachmentJson: selectedList.AttachmentJson,
                ApprovedAuditPlanId: selectedList.ID,
                //documentCode: selectedList.MemoNumber
                // Format as YYYY-MM-DD
            }));
            setSelectedOption(selectedList);
            setdoccode(selectedList.MemoNumber);
            //const rowData: any[] = await getItemByID2(sp, Number(selectedList.ID)) //baseUrl
            // const rowData: any[] = await getItemfromChecklistMaster(sp) //baseUrl
            // if (rowData.length > 0) {
            //     const initialRows = rowData.map((item: any) => ({
            //         id: item.Id,
            //         isoreference: item.ISOReference,
            //         imsprocedure: item.IMSProcedure,
            //         inquiries: item.Inquiries,
            //         auditorcomments: item.auditorcomments,
            //         time: item.Time,
            //         sharewithIds: item.sharewithId,
            //         endTime: "",
            //         sharewith: item.sharewith ? { label: item.sharewith.Title, value: item.sharewith.ID } : null // Convert single object
            //     }));
            //     setRecommendationRows(initialRows);
            //     setRecommendationRowsEdit(initialRows);
            // } else {
            //     setRecommendationRowsEdit([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
            //     setRecommendationRows([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
            // }

            console.log("alllldept", AllDept);
            const selecteddepart = AllDept.filter((cust: { value: any; }) => cust.value === selectedList.DepartmentId)[0] || null;
            setselectUserDept(selecteddepart);
            const selectedauditplan = rows.filter((cust: { value: any; }) => cust.value === selectedList.ID)[0] || null;
            setselectAuditplan(selectedauditplan);
            console.log("selectedList.AttachmentId[0]", selecteddepart, selectedList.AttachmentId, selectedauditplan)
            if (selectedList.AttachmentId.length > 0) {
                let arrn = await getDocumentLinkByIDPlan(sp, selectedList.AttachmentId);
                //let arraynew: any[];
                //arraynew.push(arrn)
                console.log("arrrrrrn56", arrn);
                setFilesArrDoclink([...FilesArrDoclink, ...arrn]);
                console.log("arrrrrrn56 after", arrn);
                //setAttachmentarr(arrn);
                // setDocumentLink(await getDocumentLinkByIDPlan(sp, selectedList.AttachmentId[0]))
            }
            else {
                setDocumentLink(null);
            }  // Set the selected users
            setLoading(false);
        };
    };

    const validateForm = async (fmode: FormSubmissionMode) => {
        const {
            approvedauditplanId,
            deptId,
            date,
            documentCode,
            issueNo,
            revisionNo,
            referenceNo,
            documentLink,
            attachment,
            attachmentIds,
            attachmentJson } = formData;
        // const { description } = richTextValues;
        let valid = true;
        let validraft = true;
        let valid1 = true;
        let validRec = true;
        let validAudit = true;
        // let validateOverview:boolean = false;
        // let validatetitlelength = false;
        // let validateTitle = false;
        // setValidDraft(true);
        setdepartmenterr(false);
        setdateerr(false);
        setattachmenterr(false);
        setisoreferenceerr(false);
        setimsprocedureerr(false);
        setinquirieserr(false);
        settimeerr(false);
        setauditorcommentserr(false);
        setsharewitherr(false);
        setapprovedauditplanerr(false);
        setValidSubmit(true);
        setValidCancelReason(true);
        setValidForwardTo(true);
        setValidAudit(true);
        setValidDraft(true);
        let errormsg = "";

        if (fmode == FormSubmissionMode.SUBMIT) {
            if (!selectedOption) {
                setapprovedauditplanerr(true);
                valid = false;
            }
            if (!deptId) {
                setdepartmenterr(true);
                //Swal.fire('Error', 'Title is required!', 'error');
                valid = false;
            }
            if (!FilesArr.length) {
                setattachmenterr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                valid = false;
            }
            if (!date) {
                setdateerr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                valid = false;
            }
            if (date == "Invalid Date") {
                setdateerr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                valid = false;
            }
            //  else if (!auditPlanTypeId.length) {
            //     //Swal.fire('Error', 'Entity is required!', 'error');
            //     valid = false;
            // }
            if (!sharewithusers || sharewithusers.length === 0) {
                setsharewitherr(true);
            }
            if (!recommendationRows.length) {
                validRec = false;
            }

            // if (recommendationRows.length > 0 && recommendationRows.every((row: any) =>
            //     row.isoreference.trim() !== "" && row.imsprocedure.trim() !== "" && row.inquiries.trim() !== "" && row.time.trim() !== "" && row.auditorcomments.trim() !== ""
            //     && row.sharewith != null && row.sharewith.length != 0) == false) {
            //     validRec = false;
            // }
            // if (recommendationRows.length > 0) {
            //     let isoreferenceError = false;
            //     let imsprocedureError = false;
            //     let inquiriesError = false;
            //     let timeError = false;
            //     let auditorcommentsError = false;
            //     let sharewithError = false;

            //     recommendationRows.forEach((row: any) => {
            //         // Check each field individually
            //         if (row.isoreference === null || row.isoreference.trim() === "") {
            //             isoreferenceError = true;
            //         }
            //         if (row.imsprocedure === null || row.imsprocedure.trim() === "") {
            //             imsprocedureError = true;
            //         }
            //         if (row.inquiries === null || row.inquiries.trim() === "") {
            //             inquiriesError = true;
            //         }
            //         if (row.time === null || row.time.trim() === "") {
            //             timeError = true;
            //         }
            //         if (row.auditorcomments === null || row.auditorcomments.trim() === "") {
            //             auditorcommentsError = true;
            //         }
            //         if (row.sharewith === null || row.sharewith.length === 0) {
            //             sharewithError = true;
            //         }
            //     });

            //     // Set error flags for fields that are invalid
            //     if (isoreferenceError) {
            //         setisoreferenceerr(true);
            //     }
            //     if (imsprocedureError) {
            //         setimsprocedureerr(true);
            //     }
            //     if (inquiriesError) {
            //         setinquirieserr(true);
            //     }
            //     if (timeError) {
            //         settimeerr(true);
            //     }
            //     if (auditorcommentsError) {
            //         setauditorcommentserr(true);
            //     }
            //     if (sharewithError) {
            //         setsharewitherr(true);
            //     }

            //     // If any field has an error, set the validation flag to false
            //     if (isoreferenceError || imsprocedureError || inquiriesError || timeError || auditorcommentsError || sharewithError) {
            //         validRec = false;
            //     }
            // }
            if (recommendationRows.length > 0) {
                debugger
                let validRec = true; // Assume valid initially
                let rowErrors: any[] = []; // Store errors for each row

                recommendationRows.forEach((row: any, index: number) => {
                    let rowError: any = {}; // Store errors for this row

                    if (!row.isoreference || row.isoreference.trim() === "") {
                        rowError.isoreference = true;
                    }
                    if (!row.imsprocedure || row.imsprocedure.trim() === "") {
                        rowError.imsprocedure = true;
                    }
                    if (!row.inquiries || row.inquiries.trim() === "") {
                        rowError.inquiries = true;
                    }
                    if (!row.time || row.time.trim() === "") {
                        rowError.time = true;
                    }
                    if (!row.auditorcomments || row.auditorcomments.trim() === "") {
                        rowError.auditorcomments = true;
                    }
                    // if (!row.sharewith || row.sharewith.length === 0) {
                    //     rowError.sharewith = true;
                    // }

                    // If there are errors in this row, store them
                    if (Object.keys(rowError).length > 0) {
                        rowErrors[index] = rowError; // Assign the errors for this row
                        validRec = false; // Mark as invalid
                    }
                });

                // Update the state or UI with the errors
                setRowErrors(rowErrors);

                // Set validation flag
                if (!validRec) {
                    console.log("Validation failed. Highlight errors accordingly.");
                }
            }

            if (!forwardToArr) {
                valid1 = false;
            }
            if (forwardToArr.length > 0 && forwardToArr.every((row: any) => row.role !== 0 && row.approvalType.trim() !== "" && row.approvers.length != 0) == false) {
                // const isValid = cancellReason.every((row:any) => row.description.trim() !== "" && row.reason.trim() !== "");
                valid1 = false;
            }

            setValidSubmit(valid);
            setValidDRecomm(validRec);
            setValidAudit(validAudit);
            setValidForwardTo(valid1);

        }
        else {
            if (!date) {
                setdateerr(true);
                //Swal.fire('Error', 'Title is required!', 'error');
                validraft = false;
            }
            if (date == "Invalid Date") {
                setdateerr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                validraft = false;
            }
            if (!selectedOption) {
                setapprovedauditplanerr(true);
                validraft = false;
            }
            if (!deptId) {
                setdepartmenterr(true);
                //Swal.fire('Error', 'Type is required!', 'error');
                validraft = false;
            }
            // else if (selectedOption == null || !selectedOption.value) {
            //     //Swal.fire('Error', 'Entity is required!', 'error');
            //     valid = false;
            // }


            setValidDraft(validraft);
            // setValidCancelReason(valid1);

        }

        // console.log("validateTitle", validateTitle,"errormsg", errormsg,"valid,", valid, ImagepostArr.length);
        // if (!valid && fmode == FormSubmissionMode.SUBMIT){
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');

        // }
        // else if (!valid1 && fmode == FormSubmissionMode.SUBMIT)
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields in description section.');
        // else if (!valid && fmode == FormSubmissionMode.DRAFT) {
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');
        // }
        // else if (!valid1 && fmode == FormSubmissionMode.DRAFT) {
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields in description section..');
        // }
        if (valid == false || valid1 == false || validRec == false || validraft == false) {
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');

            return false
        }
        else {
            return true
        }
        // return valid;
    };
    // #region  Submit Form
    const handleFormSubmit = async () => {
        if (await validateForm(FormSubmissionMode.SUBMIT)) {
            if (editForm) {
                Swal.fire({
                    title: 'Do you want to submit this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    console.log(result)
                    if (result.isConfirmed) {
                        setLoading(true);
                        debugger

                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditReportDocs');



                        if (FilesArr.length > 0) {
                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    DocumentName = file.name;
                                    const fileAddResult = await folder.files.addChunked(file.name, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult rep)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        debugger
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {

                            MemoNumber: doccode,
                            // MemoSerialNumber:,
                            // IssueNumber:,
                            // RevisionNumber:,
                            Title: doccode,
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.deptId,
                            Date: formData.date,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            DocumentName: DocumentName,
                            IsRework: "No",
                            SharewithId: sharewithIds,
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || ""


                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        debugger
                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditReportIDId: editItemID, // Assuming "Title" column exists
                                ISOReference: row.isoreference,
                                IMSProcedure: row.imsprocedure,
                                Inquiries: row.inquiries,
                                AuditorsComments: row.auditorcomments,
                                Time: row.time,
                                SharewithId: row.sharewithIds ? row.sharewithIds : 0
                                //SharewithId: row.sharewithIds,
                                // sharewith: row.sharewith ? { label: row.sharewith.Title, value: row.sharewith.ID } : null
                            }

                            if (row.id) {

                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {
                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;

                            }

                            // debugger
                            // if (!postId2) {
                            //     console.error("Post creation failed.");
                            //     return;
                            // }
                        }

                        let isValid = true;

                        if (forwardToArr.length) {
                            isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                                row.approvalType.trim() !== "");

                            // if (!isValid) {
                            //     // alert("Each row must have a role selected and at least one approver.");
                            //     valid = false;
                            // }



                            // if (!valid) {
                            //     Swal.fire('Please fill all the mandatory fields.');
                            //     // setValidSubmit(false);
                            //     setValidForwardTo(false);
                            //     return;
                            // }

                        }

                        if (isValid) {
                            for (const item of forwardToArr) {

                                const approversIds: any[] = [];
                                item.approvers.forEach((user: any) => {
                                    if (user?.value) {
                                        approversIds.push(user.value);
                                    }
                                });

                                let arr2 = {
                                    Title: currentUser.Title,


                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: item.approvers?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(editItemID),
                                    ContentTitle: formData.referenceNo,
                                    RequestId: doccode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Annual Audit Report",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    // IsApprovalGenerated: "No"
                                    // RedirectionLink:,



                                }
                                if (item.id) {
                                    const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }
                        // *******************************???????????




                        // /////////*************************** */

                        const toDelete = forwardToArrEdit.filter(
                            (itemEdit) => !forwardToArr.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        const toDelete1 = recommendationRowsEdit.filter(
                            (itemEdit) => !recommendationRows.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete1) {
                            try {
                                await sp.web.lists.getByTitle("AnnualAuditReportCheckList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }


                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            let arr2 = {
                                ActionTakenById: currentUser.Id,
                                ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                                // ActionTakenRoleId: formData.RequesterDesignation,
                                Status: "Approved",
                                // Remark: remark,

                            }
                            const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                            const postId = postResult?.data?.ID;

                        }


                        let boolval = false;

                        // if (boolval == true) {
                        setLoading(false);

                        sessionStorage.removeItem("DocumentCancelId");
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://officeindia.sharepoint.com/sites/edcspfx/SitePages/EDCMAIN.aspx`;
                            }
                        });
                        // setTimeout(() => {

                        //     window.history.back();
                        //     // window.location.reload();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        //     // let url = window.location.href;
                        //     // let baseUrl = url.split("#")[0];
                        // }, 500);
                        // }
                    }

                })
            }
            else {
                Swal.fire({
                    title: 'Do you want to submit this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    //console.log("Form Submitted:", formValues, bannerImages, galleryImages, documents);
                    if (result.isConfirmed) {
                        setLoading(true);

                        debugger
                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditReportDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    DocumentName = file.name;
                                    const fileAddResult = await folder.files.addChunked(file.name, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        debugger
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        let arr = {
                            MemoNumber: doccode,
                            Title: doccode,
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.deptId,
                            Date: formData.date,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            DocumentName: DocumentName,
                            IsRework: "No",
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            SharewithId: sharewithIds


                        }

                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }
                        debugger
                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditReportIDId: postId, // Assuming "Title" column exists
                                ISOReference: row.isoreference,
                                IMSProcedure: row.imsprocedure,
                                Inquiries: row.inquiries,
                                AuditorsComments: row.auditorcomments,
                                Time: row.time,
                                SharewithId: row.sharewithIds ? row.sharewithIds : 0
                            }

                            const postResult2 = await addItem2(postPayload2, sp);
                            const postId2 = postResult2?.data?.ID;
                            // debugger
                            if (!postId2) {
                                console.error("Post creation failed.");
                                return;
                            }
                        }

                        let isValid = true;

                        if (forwardToArr.length) {
                            isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                                row.approvalType.trim() !== "");

                            // if (!isValid) {
                            //     // alert("Each row must have a role selected and at least one approver.");
                            //     valid = false;
                            // }



                            // if (!valid) {
                            //     Swal.fire('Please fill all the mandatory fields.');
                            //     // setValidSubmit(false);
                            //     setValidForwardTo(false);
                            //     return;
                            // }

                        }

                        if (isValid) {
                            for (const item of forwardToArr) {

                                const approversIds: any[] = [];
                                item.approvers.forEach((user: any) => {
                                    if (user?.value) {
                                        approversIds.push(user.value);
                                    }
                                });

                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.referenceNo,
                                    RequestId: doccode,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: item.approvers?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(postId),
                                    // RequestId: selectedOption.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Annual Audit Report",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No"
                                    // RedirectionLink:,



                                }
                                if (item.id) {
                                    const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }




                        let boolval;

                        // if (boolval == true) {
                        setLoading(false);
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://officeindia.sharepoint.com/sites/edcspfx/SitePages/EDCMAIN.aspx`;
                            }
                        });
                        //Swal.fire('Submitted successfully.', '', 'success');
                        // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 500);
                        // }

                    }
                })

            }
        }

    }

    const handleSaveAsDraft = async () => {
        if (await validateForm(FormSubmissionMode.DRAFT)) {
            if (editForm) {
                Swal.fire({
                    title: 'Do you want to save this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    debugger
                    console.log(result)
                    if (result.isConfirmed) {
                        setLoading(true);
                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditReportDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    DocumentName = file.name;
                                    const fileAddResult = await folder.files.addChunked(file.name, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.deptId,
                            Date: formData.date,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            MemoNumber: doccode,
                            Title: doccode,
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "No",
                            Status: "Save as draft",
                            DocumentName: DocumentName,
                            IsRework: "No",
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            SharewithId: sharewithIds
                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        //  ////////////////////////////
                        debugger
                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditReportIDId: editItemID, // Assuming "Title" column exists
                                ISOReference: row.isoreference || "",
                                IMSProcedure: row.imsprocedure || "",
                                Inquiries: row.inquiries || "",
                                AuditorsComments: row.auditorcomments || "",
                                Time: row.time || "",
                                SharewithId: row.sharewithIds ? row.sharewithIds : 0,
                            }

                            if (row.id) {

                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {

                                if ((row.isoreference.trim() == "" && row.imsprocedure.trim() == "" && row.inquiries.trim() == "" && row.auditorcomments.trim() == "" && row.time.trim() == "" && (row.sharewith == null || row.sharewith.length == 0)) == false) {


                                    const postResult2 = await addItem2(postPayload2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                    if (!postId2) {
                                        console.error("Post creation failed.");
                                        return;
                                    }

                                }



                            }
                            // }


                            // debugger

                        }

                        let isValid = true;

                        // if (forwardToArr.length) {
                        //     isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                        //         row.approvalType.trim() !== "");


                        // }

                        // if (isValid) {
                        for (const item of forwardToArr) {

                            const approversIds: any[] = [];
                            item.approvers.forEach((user: any) => {
                                if (user?.value) {
                                    approversIds.push(user.value);
                                }
                            });

                            let arr2 = {
                                Title: currentUser.Title,
                                // ContentTitle: selectedOption.ReferenceNumber,
                                ContentTitle: formData.referenceNo,
                                RequestId: doccode,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role || 0,
                                Level: Number(item.level),
                                ApproversId: approversIds || [],
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "No",
                                Maxlevel: item.approvers?.length,

                                // MainListID: String(editItemID),
                                MainListID: String(editItemID),
                                // RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Annual Audit Report",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                // IsApprovalGenerated: "No"
                                // RedirectionLink:,



                            }
                            if (item.id) {
                                const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {
                                if (forwardToArr.every(row => row.role == 0 && row.approvers.length == 0 &&
                                    row.approvalType.trim() == "") == false) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }

                        // }



                        // /#############//////////////////

                        const toDelete = forwardToArrEdit.filter(
                            (itemEdit) => !forwardToArr.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        const toDelete1 = recommendationRowsEdit.filter(
                            (itemEdit) => !recommendationRows.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete1) {
                            try {
                                await sp.web.lists.getByTitle("AnnualAuditReportCheckList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            let arr2 = {
                                ActionTakenById: currentUser.Id,
                                ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                                // ActionTakenRoleId: formData.RequesterDesignation,
                                Status: "Save as draft",
                                // Remark: remark,

                            }
                            const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                            const postId = postResult?.data?.ID;

                        }


                        let boolval = false;

                        // if (boolval == true) {
                        setLoading(false);
                        sessionStorage.removeItem("DocumentCancelId")
                        Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://officeindia.sharepoint.com/sites/edcspfx/SitePages/EDCMAIN.aspx`;
                            }
                        });

                        // setTimeout(() => {

                        //     window.history.back();
                        //     // window.location.reload();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        // }, 1000);
                        // }
                    }

                })
            }
            else {
                Swal.fire({
                    title: 'Do you want to save this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    //console.log("Form Submitted:", formValues, bannerImages, galleryImages, documents);
                    if (result.isConfirmed) {
                        setLoading(true);
                        debugger

                        // let galleryIds: any[] = [];

                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditReportDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    DocumentName = file.name;
                                    const fileAddResult = await folder.files.addChunked(file.name, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        let arr = {

                            // MemoNumber:,
                            // MemoSerialNumber:,
                            // IssueNumber:,
                            // RevisionNumber:,
                            MemoNumber: doccode,
                            Title: doccode,
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.deptId,
                            Date: formData.date,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "No",
                            Status: "Save as draft",
                            DocumentName: DocumentName,
                            IsRework: "No",
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            SharewithId: sharewithIds

                        }
                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }

                        for (const row of recommendationRows) {
                            if ((row.isoreference.trim() == "" && row.imsprocedure.trim() == "" && row.inquiries.trim() == "" && row.auditorcomments.trim() == "" &&
                                row.time.trim() == "" && (row.sharewith == null || row.sharewith.length == 0)) == false) {
                                // if ((row.section.trim() == "" && row.date.trim() == "" && row.startTime.trim() == "" && 
                                // (row.auditor == null || row.auditor.length == 0)) == false) {
                                debugger
                                const postPayload2 = {
                                    AnnualAuditReportIDId: postId, // Assuming "Title" column exists
                                    ISOReference: row.isoreference || "",
                                    IMSProcedure: row.imsprocedure || "",
                                    Inquiries: row.inquiries || "",
                                    AuditorsComments: row.auditorcomments || "",
                                    Time: row.time || "",
                                    SharewithId: row.sharewithIds ? row.sharewithIds : 0,

                                }

                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }
                            }


                        }

                        let isValid = true;

                        if (forwardToArr.length) {
                            isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                                row.approvalType.trim() !== "");
                        }

                        // if (isValid) {
                        for (const item of forwardToArr) {
                            if ((item.role == 0 && item.approvers.length == 0 &&
                                item.approvalType.trim() == "") == false) {



                                const approversIds: any[] = [];
                                item.approvers.forEach((user: any) => {
                                    if (user?.value) {
                                        approversIds.push(user.value);
                                    }
                                });

                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.referenceNo,
                                    RequestId: doccode,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role ? item.role : 0,
                                    Level: Number(item.level),
                                    ApproversId: approversIds || [],
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: item.approvers?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(postId),
                                    // RequestId: selectedOption.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Annual Audit Report",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No"
                                    // RedirectionLink:,



                                }
                                if (item.id) {
                                    const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }

                        // }




                        setLoading(false);
                        Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://officeindia.sharepoint.com/sites/edcspfx/SitePages/EDCMAIN.aspx`;
                            }
                        });
                        // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 1000);
                    }
                })

            }
        }

    }

    const formatToDDMMYYYY = (dateInput: string | number | Date) => {
        const date = new Date(dateInput);
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0'); // months are zero-indexed
        const year = date.getFullYear();

        return `${day}-${month}-${year}`;
    }

    const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>, libraryName: string, docLib: string) => {
        event.preventDefault();
        setshowviewdownload(false);
        const allowedTypes = [
            "image/jpeg",
            "image/png",
            "image/gif",
            "image/bmp",
            "image/svg+xml",
            "application/pdf",
            "application/msword",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-powerpoint",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ];

        filechanged = true;
        newfileupload = true;
        let uloadBannerImageFiles: any[] = [];
        let uloadImageFiles: any[] = [];
        let uloadImageFiles1: any[] = [];


        if (event.target.files && event.target.files.length > 0) {
            const files = Array.from(event.target.files);
            (event.target as HTMLInputElement).value = '';

            if (files.length > 0) {

                for (const fn of files) {
                    // const file = files[0];
                    if (!allowedTypes.includes(fn.type)) {
                        Swal.fire({
                            icon: "error",
                            title: "Invalid File Type",
                            text: "Only images and document files are allowed.",
                        });
                        return;
                    }
                    const fileType = fn.type.split("/")[0]; // Extract file type (image, pdf, etc.)
                    const preview = URL.createObjectURL(fn);
                }
                setFilesArr([...FilesArr, ...files]);
            } else {
                Swal.fire("upload a document")
            }
        }
    };

    const handleDelete = (index: number) => {
        setFilesArr((prevFiles: any[]) => prevFiles.filter((_file: any, i: number) => i !== index));
    };


    return (
        <div id="wrapper" ref={elementRef}>
            {/* <div
                className="app-menu"
                id="myHeader">
                <VerticalSideBar _context={sp} />
            </div> */}
            <div className="content-page">
                {/* <HorizontalNavbar _context={sp} siteUrl={siteUrl} /> */}
                {/* <div className="content" style={{ marginLeft: `${!useHide ? '0px' : '80px'}`, marginTop: '2.3rem' }}> */}
                <div>
                    <div className="">
                        <div className="row">
                            <div className="col-lg-4">
                                <CustomBreadcrumb Breadcrumb={Breadcrumb} />
                            </div>

                        </div>
                        <div className="row mt-0">

                            {/* <!-- Right Sidebar --> */}
                            <div className="col-12">
                                <div >
                                    <div>
                                        {/* <!-- Left sidebar --> */}


                                        {Loading ?

                                            <div className="loadernewadd mt-10">
                                                <div>
                                                    <img
                                                        src={require("../assets/edc-gif.gif")}
                                                        className="alignrightl"
                                                        alt="Loading..."
                                                    />
                                                </div>
                                                <span>Loading </span>{" "}
                                                <span>
                                                    <img
                                                        src={require("../assets/edcnew.gif")}
                                                        className="alignrightl"
                                                        alt="Loading..."
                                                    />
                                                </span>
                                            </div>
                                            :
                                            <div style={{ width: '100%' }} className="inbox-rightbar">

                                                <div className="card">
                                                    <div className="card-body">
                                                        {/* <h4 className="text-dark font-16 fw-bold mb-3">Memo Details</h4>
 */}
                                                        <div className="previewIcon">
                                                            <h4 style={{ textAlign: 'left' }} className="text-dark font-16 fw-bold mb-3">Requested By</h4>

                                                            {TemplateDoc && TemplateDoc.length > 0 && (
                                                                <span
                                                                    onClick={() => OpenFile(TemplateDoc[0], "Open")}
                                                                    style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                >
                                                                    <div className="btn btn-primary">
                                                                        <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../assets/noun-download-5006210.png")} ></img></div>
                                                                </span>
                                                            )}
                                                        </div>

                                                        <div style={{ clear: "both" }}></div>

                                                        <form className="form-horizontal">
                                                            <div className="row">
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="DocumentCode" className="col-form-label">Approved Memo No./Audit Plan:<span className="text-danger1">*</span>
                                                                        </label>
                                                                        <div
                                                                            title={selectedOption?.label || "Select a audit plan"}
                                                                            style={{ width: "100%" }}
                                                                        >
                                                                            <Select
                                                                                options={rows}
                                                                                value={selectedOption}
                                                                                name="AuditPlan"
                                                                                isClearable={true}
                                                                                isSearchable={true}
                                                                                className={`newse  ${(!ValidSubmit && approvedauditplanerr) ? "border-on-error" : ""} ${(!ValidDraft && approvedauditplanerr) ? "border-on-error" : ""}`}
                                                                                onChange={(selectedOption: any) => onSelectDocCode(selectedOption)}
                                                                                placeholder={"Search Audit plan"}
                                                                                isDisabled={InputDisabled}
                                                                            />
                                                                        </div>

                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="Department" className="col-form-label">Department<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <div
                                                                                title={selectUserDept?.label || "Select a department"}
                                                                                style={{ width: "100%" }}
                                                                            >
                                                                                <Select
                                                                                    options={AllDept}
                                                                                    isDisabled={InputDisabled}
                                                                                    value={selectUserDept}
                                                                                    name="deptId"
                                                                                    className={`newse  ${(!ValidSubmit && departmenterr) ? "border-on-error" : ""} ${(!ValidDraft && departmenterr) ? "border-on-error" : ""}`}
                                                                                    // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                    // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                    // onChange={handleDepartmentChange}
                                                                                    onChange={(selectedOptions: any) => handleDepartmentChange(selectedOptions)}
                                                                                    placeholder="Select Department"
                                                                                />
                                                                            </div>

                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="date" className="col-form-label">Date<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <input
                                                                                type="date"
                                                                                className={`form-control ${(!ValidSubmit && dateerr) ? "border-on-error" : ""}${(!ValidDraft && dateerr) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="date"
                                                                                value={formData.date}
                                                                                onChange={(e) => setFormData({ ...formData, date: new Date(e.target.value).toLocaleDateString("en-CA") })}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                            {/* {(modeValue === "" || modeValue === "edit" || InputDisabled != true) ||
                                                                             (modeValue == "approve" && formData?.Status == "Rework") ?
                                                                                <input
                                                                                    type="date"
                                                                                    className={`form-control ${(!ValidSubmit && dateerr) ? "border-on-error" : ""}${(!ValidDraft && dateerr) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    id="date"
                                                                                    value={formData.date}
                                                                                    onChange={(e) => setFormData({ ...formData, date: new Date(e.target.value).toLocaleDateString("en-CA") })}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                                :
                                                                                <input
                                                                                    type="text"
                                                                                    className={`form-control ${(!ValidSubmit && dateerr) ? "border-on-error" : ""}${(!ValidDraft && dateerr) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    id="date"
                                                                                    value={new Date(formData.date).toLocaleDateString("en-GB")  }
                                                                                    //onChange={(e) => setFormData({ ...formData, date: new Date(e.target.value).toLocaleDateString("en-CA") })}
                                                                                    disabled={true}
                                                                                />
                                                                            } */}
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                {console.log("doccodedoccodedoccode", formData, doccode)}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Memo Code<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="documentCode"
                                                                                value={doccode}
                                                                            //onChange={(e) => setFormData({ ...formData, documentCode: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Document Code<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="docCode"
                                                                                value={formData.documentCode}

                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="issueNo" className="col-form-label">Issue No<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <input

                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issueNo"
                                                                                value={formData.issueNo}
                                                                                //onChange={(e) => setFormData({ ...formData, issueNo: e.target.value })}
                                                                                disabled={true}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {/* <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="issuedate" className="col-form-label">Issue Date<span className="text-danger1"> *</span></label>
                                                                        <div>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issuedate"
                                                                                value={formData.issueDate == null ? "" : moment(formData.issueDate).format("MM/DD/YYYY")}
                                                                            //onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>*/}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisionNo" className="col-form-label">Revision No<span className="text-danger1"> *</span></label>
                                                                        <div>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="revisionNo"
                                                                                value={formData.revisionNo}
                                                                            //onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {/*}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisiondate" className="col-form-label">Revision Date<span className="text-danger1"> *</span></label>
                                                                        <div>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="revisionNo"
                                                                                value={formData.revisionDate == null ? "" : moment(formData.revisionDate).format("MM/DD/YYYY")}
                                                                            //onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div> 
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisionNo" className="col-form-label">Reference No<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="referenceNo"
                                                                                value={formData.referenceNo}
                                                                                onChange={(e) => setFormData({ ...formData, referenceNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>*/}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisionNo" className="col-form-label">Share With<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <Select
                                                                                options={rows1}
                                                                                isMulti
                                                                                value={sharewithusers}
                                                                                name="share with"
                                                                                className={`newse ${(!ValidSubmit && sharewitherr) ? "border-on-error" : ""}`}
                                                                                // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                onChange={(selectedOptions: any) => onSelectsharewith(selectedOptions)}
                                                                                placeholder="Enter Share with"
                                                                                isDisabled={InputDisabled}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {console.log("documentlinkkkkkkk", DocumentLink, FilesArrDoclink)}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="DocumentCode" className="form-label">Previous Document:</label>
                                                                        {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}

                                                                        {/* <div className="text-dark mt-0"> <span >
                                                                            <a onClick={() => setShowModal(true)} ><FontAwesomeIcon icon={faPaperclip} />{DocumentLink && "1 file Attached"}</a>

                                                                        </span>
                                                                        </div> */}
                                                                        {FilesArrDoclink.length > 0 ?
                                                                            (<a style={{ fontSize: '0.875rem' }} onClick={() => setShowModalpre(true)}>
                                                                                <FontAwesomeIcon icon={faPaperclip} />{FilesArrDoclink.length} {FilesArrDoclink.length > 0 ? "files" : "file"} Attached
                                                                            </a>) : ""

                                                                        }
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="attachment" className="col-form-label">Attachment<span className="text-danger1"> *</span></label>
                                                                        <div>

                                                                            <div>
                                                                                <input
                                                                                    type="file"
                                                                                    className={`form-control ${(!ValidSubmit && attachmenterr) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    id="attachment"
                                                                                    accept=".jpg,.jpeg,.png,.gif,.bmp,.svg,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx"
                                                                                    onChange={(e) => onFileChange(e, "Gallery", "AnnualAuditReportDocs")}
                                                                                    // onChange={(e) => setFormData({ ...formData, attachment: e.target.files[0] })}
                                                                                    disabled={InputDisabled}
                                                                                    multiple
                                                                                />

                                                                            </div>

                                                                            <div>
                                                                                {FilesArr.length > 0 ?
                                                                                    (<a style={{ fontSize: '0.875rem' }} onClick={() => setShowModal(true)}>
                                                                                        <FontAwesomeIcon icon={faPaperclip} />{FilesArr.length} {FilesArr.length > 0 ? "files" : "file"} Attached
                                                                                    </a>) : ""

                                                                                }
                                                                            </div>
                                                                        </div>

                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </form>
                                                    </div>
                                                </div>

                                                <section className='card card-body mt-2'>
                                                    <fieldset>
                                                        <div className='row'>
                                                            <div className='col-sm-6'>
                                                                <h3 className='text-dark font-16 fw-bold mb-3'>Checklist</h3>
                                                            </div>
                                                            <div style={{ textAlign: 'right' }} className='col-sm-6'>
                                                                {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")} onClick={handleAddRecommendationRow}></img>}

                                                            </div>
                                                        </div>

                                                        <div style={{ display: 'grid', overflow: 'auto' }}>
                                                            <table id="tabRec" className='mtbalenew overhi'>
                                                                <thead>
                                                                    <tr><th>ISO reference<span className="text-danger1"> *</span></th>
                                                                        <th >IMS procedure<span className="text-danger1"> *</span></th>
                                                                        <th colSpan={2}>Inquiries<span className="text-danger1"> *</span></th>
                                                                        <th >Auditor’s
                                                                            Comments<span className="text-danger1"> *</span></th>
                                                                        <th>Time<span className="text-danger1"> *</span></th>
                                                                        {/* <th>Share with<span className="text-danger1"> *</span></th> */}

                                                                        {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <th>Action</th>}
                                                                    </tr>
                                                                </thead>

                                                                <tbody>
                                                                    {console.log("ValidDRecommValidDRecomm", ValidDRecomm)}
                                                                    {recommendationRows.map((row, index) => (
                                                                        <tr key={index}>
                                                                            <td title={row?.isoreference ? row?.isoreference : row?.isoreference}>
                                                                                <input
                                                                                    type="text"
                                                                                    className={`form-control ${(RowErrors[index]?.isoreference) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.isoreference}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'isoreference', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            {/* <td>
                                                                            <input
                                                                                type="text"
                                                                                className={`form-control ${(!ValidDRecomm && imsprocedureerr) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                value={row.imsprocedure}
                                                                                onChange={(e) => handleRecommendationChange(index, 'imsprocedure', e.target.value)}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                        </td> */}
                                                                            <td title={row?.imsprocedure ? row?.imsprocedure : row?.imsprocedure}>
                                                                                <input
                                                                                    type="text"
                                                                                    className={`form-control ${RowErrors[index]?.imsprocedure ? "border-on-error" : ""}`}
                                                                                    value={row.imsprocedure}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'imsprocedure', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            <td title={row?.inquiries ? row?.inquiries : row?.inquiries}>
                                                                                <textarea
                                                                                    id="simpleinput"
                                                                                    className={`form-control ${(RowErrors[index]?.inquiries) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.inquiries}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'inquiries', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                                {/* <input
                                                                                    type="text"
                                                                                    className={`form-control ${(RowErrors[index]?.inquiries) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.inquiries}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'inquiries', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                /> */}
                                                                            </td>
                                                                            <td title={row?.auditorcomments ? row?.auditorcomments : row?.auditorcomments}>
                                                                                <textarea
                                                                                    id="simpleinput"
                                                                                    className={`form-control ${(RowErrors[index]?.auditorcomments) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.auditorcomments}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'auditorcomments', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            <td title={row?.time ? row?.time : row?.time}>
                                                                                <input
                                                                                    type="time"
                                                                                    className={`form-control ${(RowErrors[index]?.time) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.time}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'time', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            {/* <td title={row?.sharewith ? row?.sharewith : row?.sharewith}>
                                                                                <Select
                                                                                    options={rows1}
                                                                                    // isMulti
                                                                                    className={`${(RowErrors[index]?.sharewith) ? "border-on-error" : ""}`}
                                                                                    value={row.sharewith}
                                                                                    onChange={(selectedOptions: any) => handleRecommendationChange(index, 'sharewith', selectedOptions)}
                                                                                    placeholder="Select"
                                                                                    isDisabled={InputDisabled}
                                                                                />
                                                                            </td> */}
                                                                            {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td>
                                                                                <img src={require("../assets/del.png")} onClick={() => handleDeleteRecommendationRow(index)} />

                                                                            </td>
                                                                            }
                                                                        </tr>
                                                                    ))}
                                                                </tbody>

                                                            </table>
                                                        </div>
                                                    </fieldset>
                                                </section>

                                                {/* {modeValue === "approve" && editID != null && editID.Status === "Pending" && editID.CurrentUserRole !== "Initiator" && */}

                                                <div className="card mt-2" style={{ marginBottom: '17px' }}>
                                                    <div className="card-body">
                                                        <div className='row'>
                                                            <div className='col-sm-8'>
                                                                <h4 className="text-dark font-16 fw-bold mb-3 ">Approval Hierarchy</h4>
                                                                <label>Define the approval hierarchy to ensure requests are routed to the appropriate approvers.
                                                                </label>
                                                            </div>
                                                            <div className='col-sm-4'>
                                                                <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right", paddingRight: "22px" }}>
                                                                    {!InputDisabled &&
                                                                        <img style={{ width: '34px' }} src={require("../assets/plus.png")} onClick={handleAddRow} className='' />
                                                                    }

                                                                </div>
                                                            </div>

                                                        </div>

                                                        <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">
                                                            <table style={{ overflow: 'inherit' }} className="mtbalenew  table-centered table-nowrap table-borderless mb-0 overhi" id="myTabl">
                                                                <thead >
                                                                    <tr>
                                                                        <th style={{ minWidth: "35px", maxWidth: "35px" }}>S.No</th>
                                                                        <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Role<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '70px', maxWidth: '70px' }} >Level</th>
                                                                        <th>Approver name<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '70px', maxWidth: '70px' }} >Approval criteria<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '70px', maxWidth: '70px' }}>Action</th>
                                                                    </tr>
                                                                </thead>
                                                                <tbody style={{ maxHeight: "8007px", overflow: 'inherit' }}>
                                                                    {forwardToArr.map((row, index) => (
                                                                        <tr>
                                                                            <td style={{ minWidth: "35px", maxWidth: "35px", overflow: 'inherit' }}> <div
                                                                                style={{ marginLeft: "5px" }}
                                                                                className="indexdesign"
                                                                            >
                                                                                {index + 1}</div>
                                                                            </td>
                                                                            <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px', }} className="ng-binding">
                                                                                <select
                                                                                    // className="form-select"
                                                                                    className={`form-select newse ${(!ValidForwardTo) ? "border-on-error" : ""} `}

                                                                                    onChange={(e) => onSelectRole(e, row.level)} value={row.role} disabled={InputDisabled}>

                                                                                    <option value="" selected>Select Role</option>
                                                                                    {/* {UserRoles.map((role: any, index: number) => (
                                                                                    <option key={index} value={role.value}>{role.label}</option>
                                                                                ))} */}
                                                                                    {UserRoles.filter((role: any) =>
                                                                                        !forwardToArr.some((r) => r.role === role.value && r.level !== row.level) || role.value === row.role // Allow the current row's role
                                                                                    ).map((role: any, idx: number) => (
                                                                                        <option key={idx} value={role.value}>{role.label}</option>
                                                                                    ))}



                                                                                </select>

                                                                            </td>
                                                                            <td style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>Level {index + 1}</td>
                                                                            <td style={{ overflow: 'inherit' }}>

                                                                                <Select
                                                                                    options={rows1}
                                                                                    isMulti
                                                                                    value={row.approvers}
                                                                                    name="Approvers"
                                                                                    className={`newse ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                    // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                    onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                                                                    placeholder="Enter Approver Name"
                                                                                    isDisabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            <td style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                                                {/* <label htmlFor="approvalType">Approval Type: </label> */}
                                                                                <select id="approvalType" value={row.approvalType} onChange={(e) => handleChange(e, row.level)} className={`newse form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`} disabled={InputDisabled} >
                                                                                    <option value="">Select </option>
                                                                                    <option value="One">Anyone</option>
                                                                                    <option value="All">Everyone</option>
                                                                                </select>
                                                                            </td>
                                                                            <td style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>

                                                                                {/* {editID.CurrentUserRole === "OES" ?  */}

                                                                                {!InputDisabled ? <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} /> : <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' />}
                                                                                {/* : <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' /> */}
                                                                                {/* } */}

                                                                            </td>
                                                                        </tr>

                                                                    ))}


                                                                </tbody>
                                                            </table>
                                                        </div>



                                                        {/* {editID.CurrentUserRole === "OES" &&  */}
                                                        {/* <div className="row mt-3">
                                                            <div className="col-12 text-center">
                                                                <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardApproval("Forward")} >
                                                                    <i className="fe-check-circle me-1"></i> Forward
                                                                </button>

                                                                <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={() => ForwardApproval("Rework")} >
                                                                    <i className="fe-corner-up-left me-1"></i> Rework
                                                                </button>

                                                                <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={() => ForwardApproval("Rejected")} >
                                                                    <i className="fe-x me-1"></i> Reject
                                                                </button>

                                                                <button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}>
                                                                    <i className="fe-x me-1"></i> Cancel
                                                                </button>

                                                            </div>
                                                        </div> */}
                                                        {/* } */}
                                                    </div>
                                                </div>
                                                {/* // } */}



                                                {/* ////////////Approval card */}

                                                {
                                                    (InputDisabled && editID != null && modeValue === "approve" && editID.ApprovalType === "Approval" && editID.Status === "Pending") ? (
                                                        <WorkflowAction currentItem={editID} ctx={props.context} ContentType={CONTENTTYPE_AuditReport}
                                                            DisableApproval={false} DisableCancel={false}
                                                        />
                                                    ) : (<div></div>)
                                                }

                                                {/* ////////////Audit History card */}
                                                {/* {editID !== null && editID.length != 0 && modeValue === "approve" && */}
                                                {MainEditItem !== null && MainEditItem.length != 0 && MainEditItem?.Status !== "Save as draft" &&
                                                    <WorkflowAuditHistory ContentItemId={MainEditItem} ContentType={CONTENTTYPE_AuditReport} ctx={props.context} />
                                                }
                                                {/* ////////////Audit History card */}


                                                {/* /////////////////%%%%%%%%%%%%%%%%%%%%%%%% */}

                                                <div className="row mt-3">
                                                    <div className="col-12 text-center">


                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") || (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Save As Draft</button>
                                                        }

                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") || (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Submit</button>
                                                        }

                                                        {/* {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</button>}


                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</button>}
 */}


                                                        {((modeValue === "" || modeValue === "edit" || modeValue === "view") || (editID !== null && editID.ApprovalType !== "Approval")) &&
                                                            <button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</button>
                                                        }

                                                    </div>
                                                </div>


                                                {/* ////////////Approval card */}


                                                {/* </div> */}

                                                {/* /////////// */}

                                                <Modal show={showModal} onHide={() => setShowModal(false)} size='lg' className='filemodal'>
                                                    <Modal.Header closeButton>
                                                        <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-0'>Attachment Details</h4>
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for Annual Audit Report
                                                            </p>

                                                        </Modal.Title>


                                                    </Modal.Header>
                                                    <Modal.Body className="" id="style-5">

                                                        {/* {DocumentLink &&
                                                            (
                                                                <> */}
                                                        <table className="mtbalenew">
                                                            <thead style={{ background: '#eef6f7' }}>
                                                                <tr>
                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                                                    <th>File Name</th>
                                                                    {((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                                                        || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                                                        <th > File Link </th>
                                                                    }
                                                                    {/* <th>File Link</th> */}
                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Upload date</th>
                                                                    {(modeValue == "edit" || modeValue == null || modeValue == ""
                                                                        || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                                                        <th style={{ textAlign: 'center' }}> Action </th>
                                                                    }
                                                                </tr>
                                                            </thead>
                                                            <tbody>
                                                                {FilesArr.length > 0 && (
                                                                    FilesArr.map((row: any, index: number) => (
                                                                        <tr>
                                                                            <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>
                                                                            <td title={row.name || row.FileLeafRef}>{row.name || row.FileLeafRef}</td>
                                                                            {((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                                                                || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                                                                <td style={{ textAlign: 'center' }}>
                                                                                    <span onClick={() => OpenFile(FilesArr && FilesArr[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                        <FontAwesomeIcon icon={faEye} /></span>
                                                                                    <span onClick={() => OpenFile(FilesArr && FilesArr[0], "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                        <FontAwesomeIcon icon={faDownload} /></span>
                                                                                </td>
                                                                            }
                                                                            <td style={{ minWidth: '50px', maxWidth: '50px' }}>{row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                day: "2-digit",
                                                                                month: "short",
                                                                                year: "numeric"
                                                                            }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                day: "2-digit",
                                                                                month: "short",
                                                                                year: "numeric"
                                                                            }).replace(/ /g, "/")}</td>

                                                                            {(modeValue == "edit" || modeValue == null || modeValue == ""
                                                                                || (modeValue == "approve" && formData?.Status == "Rework")) && <td style={{ textAlign: 'center' }}> <img src={require("../assets/del.png")} style={{ cursor: "pointer" }} onClick={() => handleDelete(index)} /></td>}


                                                                        </tr>
                                                                    ))
                                                                )}
                                                            </tbody>
                                                        </table>
                                                        {/* </>
                                                            )
                                                        } */}

                                                    </Modal.Body>

                                                </Modal>
                                                <Modal show={ShowModalpre} onHide={() => setShowModalpre(false)} size='lg' className='filemodal'>
                                                    <Modal.Header closeButton>
                                                        <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-0'>Attachment Details</h4>
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for Annual Audit Report
                                                            </p>

                                                        </Modal.Title>


                                                    </Modal.Header>
                                                    <Modal.Body className="" id="style-5">

                                                        {/* {DocumentLink &&
                                                            (
                                                                <> */}
                                                        <table className="mtbalenew">
                                                            <thead style={{ background: '#eef6f7' }}>
                                                                <tr>
                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                                                    <th>File Name</th>
                                                                    <th>File Link</th>
                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Upload date</th>
                                                                    {/* {(InputDisabled || !InputDisabled) && <th className='text-center'>Action</th>} */}
                                                                </tr>
                                                            </thead>
                                                            <tbody>
                                                                {FilesArrDoclink.length > 0 && (
                                                                    FilesArrDoclink.map((row: any, index: number) => (
                                                                        <tr>
                                                                            <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>
                                                                            <td title={row.name || row.FileLeafRef}>{row.name || row.FileLeafRef}</td>
                                                                            {/* <td style={{ textAlign: 'center' }} >
                                                                                       
                                                                                        <span onClick={() => OpenFile(DocumentLink, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}><FontAwesomeIcon icon={faEye} /></span>
                                                                                         </td> */}
                                                                            {/* <td>{DocumentLink.Created
                                                                                        ? new Intl.DateTimeFormat('en-GB', {
                                                                                            day: '2-digit',
                                                                                            month: 'short',
                                                                                            year: 'numeric'
                                                                                        }).format(new Date(DocumentLink.Created)).replace(/ /g, "/")
                                                                                        : ""}</td> */}

                                                                            <td style={{ textAlign: 'center' }}>
                                                                                <span onClick={() => OpenFile(FilesArrDoclink && FilesArrDoclink[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                    <FontAwesomeIcon icon={faEye} /></span>
                                                                                <span onClick={() => OpenFile(FilesArrDoclink && FilesArrDoclink[0], "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                    <FontAwesomeIcon icon={faDownload} /></span>
                                                                            </td>

                                                                            <td style={{ minWidth: '50px', maxWidth: '50px' }}>{row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                day: "2-digit",
                                                                                month: "short",
                                                                                year: "numeric"
                                                                            }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                day: "2-digit",
                                                                                month: "short",
                                                                                year: "numeric"
                                                                            }).replace(/ /g, "/")}</td>

                                                                            {/* {!InputDisabled && <td style={{ textAlign: 'center' }}> <img src={require("../assets/del.png")} style={{ cursor: "pointer" }} onClick={() => handleDelete(index)} /></td>} */}



                                                                        </tr>
                                                                    ))
                                                                )}
                                                            </tbody>
                                                        </table>
                                                        {/* </>
                                                            )
                                                        } */}

                                                    </Modal.Body>

                                                </Modal>
                                                {/* ///////////////// */}

                                            </div>



                                        }


                                    </div>
                                </div>

                            </div>
                        </div>


                    </div>
                </div>
            </div>
        </div>
    );
}

const AnnualAuditReport: React.FC<IAnnualAuditReportProps> = (props) => (
    <Provider>
        <AnnualAuditReportContext props={props} />
    </Provider>
);



export default AnnualAuditReport;
import * as React from 'react';
import type { IAnnualAuditPlanProps } from './IAnnualAuditPlanProps';
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
import "./annualaudit.scss";
import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import Select from "react-select";
import Swal from 'sweetalert2';
import { FormSubmissionMode } from '../../../Shared/Interfaces';
import { decryptId } from '../../../APISearvice/CryptoService';
import { WorkflowAction } from '../../../CustomJSComponents/WorkflowAction/WorkflowAction';
import { WorkflowAuditHistory } from '../../../CustomJSComponents/WorkflowAuditHistory/WorkflowAuditHistory';
import { CONTENTTYPE_AuditPlan, LIST_TITLE_AuditPlan, SITE_URL, Tenant_URL } from '../../../Shared/Constants';
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Modal } from 'react-bootstrap';
import { faDownload, faEye } from '@fortawesome/free-solid-svg-icons';
import CustomBreadcrumb from './CustomBreadcrumb/CustomBreadcrumb';
import { addAllProcessItem, addItem, addItem2, getAllAuditType, getAllDepartment, getAllProcessData, getApprovalByID, getApprovalByID2, getDataRoles, getDocumentLinkByID, getDraftApprovalByID, getFormNameID, getItemByID, getItemByID2, getListNameID, UpdateAllProcessItem, updateApprovalItem, updateItem, updateItem2, uploadAllFiles } from './AuditPlanService';
import { TextField } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Tooltip } from 'react-tooltip';
import 'react-tooltip/dist/react-tooltip.css';

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

const AnnualAuditPlanContext = ({ props }: any) => {
    const sp: SPFI = getSP();
    const elementRef = React.useRef<HTMLDivElement>(null);
    const siteUrl = props.siteUrl;
    const { useHide }: any = React.useContext(UserContext);
    const [InputDisabled, setInputDisabled] = React.useState(false);
    const selectedTextDiv = document.getElementById('selectedText');

    selectedTextDiv.style.display = 'none';


    const [FilesArr, setFilesArr] = React.useState<any>([]);
    const [FilesArr1, setFilesArr1] = React.useState<any>([]);
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
    const [currentUserDept, setcurrentUserDept] = React.useState("");
    const [selectUserDept, setselectUserDept] = React.useState([]);
    const [selectUserDeptTo, setselectUserDeptTo] = React.useState([]);
    const [selectUserDeptCC, setselectUserDeptCC] = React.useState(null);
    const [AllDept, setAllDept] = React.useState([]);
    const [DocumentLink, setDocumentLink] = React.useState(null);
    const [DraftApprovalItem, setDraftApprovalItem] = React.useState(null);
    // const [cancellReason, setcancellReason] = React.useState([{ id: 0, description: "", reason: "" }]);
    // const [RecommendRows, setRecommendRows] = React.useState([]);
    const [tooltipText, settooltipText] = React.useState("");
    const [tooltipText1, settooltipText1] = React.useState("");
    const [showModal, setShowModal] = React.useState(false);
    const [formData, setFormData] = React.useState({
        // infoCheck: false,
        // signCheck: false,
        // approvalCheck: false,
        memoNo: "",
        memoSerialNo: 0,
        deptId: 0,
        issueNo: "",
        revisionNo: "",
        from: 0,
        fromEmail: "",
        to: [],
        subject: "",
        attachment: null,
        date: "",
        background: "",
        issues: "",
        recommendationforApproval: "",
        approval: "",
        CC: [],
        auditPlanTypeId: [],
        exclusions: "",
        boundary: "",
        objective: "",
        criteria: "",
        scope: "",
        assignedTo: "",
        ToDepartments: [],
        CCDepartments: [],
        attachmentIds: null,
        attachmentJson: null

    });
    const [selectCCUsers, setSelectCCUsers] = React.useState([]);
    const [ListNameId, setListNameId] = React.useState(null);

    const handleCCChange = (selectedOptions: any, fieldName: string) => {
        setSelectCCUsers(selectedOptions);
        const valuesOnly = selectedOptions.map((option: any) => option.value);

        setFormData({ ...formData, [fieldName]: valuesOnly });

        // setFormData({ ...formData, [fieldName]: selectedOptions });

    };

    const [selectToUsers, setSelectToUsers] = React.useState([]);

    const handleToChange = (selectedOptions: any, fieldName: string) => {
        setSelectToUsers(selectedOptions);
        const valuesOnly = selectedOptions.map((option: any) => option.value);
        setFormData({ ...formData, [fieldName]: valuesOnly });

        // setFormData({ ...formData, [fieldName]: selectedOptions.value });
    };

    const handleDepartmentChange = (selectedOption: any) => {
        setselectUserDept(selectedOption);
        const formattedMemoSerialNo = formData.memoSerialNo < 10
            ? `00${formData.memoSerialNo}`
            : formData.memoSerialNo < 100
                ? `0${formData.memoSerialNo}`
                : formData.memoSerialNo;
        setFormData({
            ...formData,
            deptId: selectedOption.value,
            memoNo: `${selectedOption.DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
        });

        if (selectedOption) {
            document.getElementById("DeptID")?.classList.remove("border-on-error");
        }
    };


    const handleDepartmentChangeTo = (selectedOption: any) => {
        setselectUserDeptTo(selectedOption);
        const valuesOnly = selectedOption.map((option: any) => option.value);
        const consolidatedToUsers = Array.from(new Set(selectedOption
            .flatMap((option: any) => option.ToUsers || []))); // Flatten, handle undefined ToUsers, and remove duplicates

        setFormData({
            ...formData,
            // deptId: selectedOption.value,
            ToDepartments: valuesOnly, // Assuming it's an array of IDs
            to: consolidatedToUsers  // Assuming it's an array of IDs
        });
        if (selectedOption) {
            document.getElementById("ToDept")?.classList.remove("border-on-error");
        }
        createTooltipContentTo(selectedOption);
    }

    const handleDepartmentChangeCC = (selectedOption: any) => {
        setselectUserDeptCC(selectedOption);
        const valuesOnly = selectedOption.map((option: any) => option.value);
        const consolidatedCCUsers = Array.from(new Set(selectedOption
            .flatMap((option: any) => option.CCUsers || []))); // Flatten and handle undefined CCUsers

        setFormData({
            ...formData,
            // deptId: selectedOption.value,
            // ToDepartments: [selectedOption.value], // Assuming it's an array of IDs
            CCDepartments: valuesOnly,
            CC: consolidatedCCUsers  // Assuming it's an array of IDs
        });
        if (selectedOption) {
            document.getElementById("CCDept")?.classList.remove("border-on-error");
        }
        createTooltipContent(selectedOption);
    }


    // ////// Recommendation
    const [recommendationRows, setRecommendationRows] = React.useState([
        { id: 0, section: "", date: "", startTime: "", endTime: "", auditor: null, auditorIds: null }
    ]);

    const [recommendationRowsEdit, setRecommendationRowsEdit] = React.useState([]);

    const handleAddRecommendationRow = () => {
        setRecommendationRows([...recommendationRows, { id: 0, section: "", date: "", startTime: "", endTime: "", auditor: null, auditorIds: null }]);
    };

    const handleRecommendationChange = (index: number, field: string, value: any) => {
        let updatedRows;
        if (field == "auditor") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = recommendationRows.map((row, i) =>
                i === index ? { ...row, [field]: value, auditorIds: value.value } : row
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
            ChildComponent: "Audit Plan",
            ChildComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx#/AnnualAuditPlan`,
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

        setAllDept(await getAllDepartment(sp));
        var setAllDept1 = await getAllDepartment(sp);


        const path1 = window.location.href;
        const path = window.location.href;
        const segments = path.split('/').filter(Boolean); // Remove empty elements
        let formMode = "";
        // Check if "edit" or "view" exists in the URL
        const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");
        if (paramIndex !== -1) {
            setmode(segments[paramIndex])
            formMode = segments[paramIndex]; // Will be "edit" or "view"
        }
        else {

            setmode("");
            formMode = "";
        }


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
        let memo: number = 0;
        let filteredDeptArrayTo: any[] = [];
        let filteredDeptArrayCC: any[] = [];

        const Currusers: any = await getCurrentUser(sp, siteUrl);
        setCurrentUser(await getCurrentUser(sp, siteUrl));
        const userProfile = await sp.profiles.myProperties();
        setcurrentUserDept(userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "")
        const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
        setselectUserDept(setAllDept1.filter(user => user.label === UserDept));

        if (formMode == "") {

            const listItems = await sp.web.lists.getByTitle("AnnualAuditPlanList").items.orderBy("MemoSerialNumber", false).top(1)();
            if (listItems.length > 0) {
                memo = listItems[0].MemoSerialNumber ? listItems[0].MemoSerialNumber + 1 : 1;
                // if (memo < 999) {
                //     memo = ("0000" + memo).slice(-3);
                // }
                // setFormData((prevData) => ({ ...prevData, memoSerialNo: memo }));


                // setFormData({
                //     ...formData,
                //     memoSerialNo:memo 
                // });
            } else {
                memo = 1;
                // setFormData((prevData) => ({ ...prevData, memoSerialNo: memo }));
                // setFormData({
                //     ...formData,
                //     memoSerialNo: memo 
                // });
            }
            const formattedMemoSerialNo = memo < 10
                ? `00${memo}`
                : memo < 100
                    ? `0${memo}`
                    : memo;

            setFormData({
                ...formData,
                memoSerialNo: memo,
                deptId: setAllDept1.filter(user => user.label === UserDept)[0].value,
                memoNo: `${setAllDept1.filter(user => user.label === UserDept)[0].DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
            });

        }

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



        let formitemid;
        //#region getdataByID
        if (sessionStorage.getItem("AuditPlanId") != undefined) {
            const iD = sessionStorage.getItem("AuditPlanId")
            let iDs = decryptId(iD)
            formitemid = Number(iDs);
            setFormItemId(Number(iDs))
        }
        else {
            // const path = window.location.href;
            // const segments = path.split('/').filter(Boolean); // Remove empty elements

            // // Check if "edit" or "view" exists in the URL
            // const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");


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
                    setEditID(await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditPlan));
                    // var ProcessItemId: any = await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditPlan);
                    setInputDisabled(await getApprovalByID2(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditPlan));
                }
                // else {

                //     setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditPlan))

                // }
            }

            setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditPlan))

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
                    memoNo: setBannerById[0].MemoNumber,
                    memoSerialNo: setBannerById[0].MemoSerialNumber,
                    deptId: setBannerById[0].DepartmentId,
                    // issueNo: "",
                    // revisionNo: "",
                    from: setBannerById[0].FromId,
                    fromEmail: setBannerById[0].From?.EMail,
                    to: setBannerById[0].ToId || [],
                    subject: setBannerById[0].Subject,
                    // attachment: null,
                    date: new Date(setBannerById[0].Date).toLocaleDateString("en-CA"),
                    background: setBannerById[0].Background,
                    issues: setBannerById[0].Issues,
                    recommendationforApproval: setBannerById[0].RecommendedforApproval,
                    // approval: setBannerById[0].,
                    CC: setBannerById[0].CcId || [],
                    auditPlanTypeId: setBannerById[0].AuditPlanTypeId || [],
                    exclusions: setBannerById[0].Exclusions,
                    boundary: setBannerById[0].Boundary,
                    objective: setBannerById[0].AimObjective,
                    criteria: setBannerById[0].Criteria,
                    scope: setBannerById[0].Scope,
                    // assignedTo: "",
                    ToDepartments: setBannerById[0].ToDepartmentsId || [],
                    CCDepartments: setBannerById[0].CCDepartmentsId || [],
                    attachmentIds: setBannerById[0].AttachmentId || null,
                    attachmentJson: setBannerById[0].AttachmentJson || null
                }));

                setselectUserDept(setAllDept1.filter(user => user.value === setBannerById[0].DepartmentId));

                setselectUserDeptCC(setBannerById[0].CCDepartments?.map((obj: any) => {
                    const filteredDept = setAllDept1.find((dept: any) => dept.value === obj.ID);
                    if (filteredDept) {
                        filteredDeptArrayCC.push(filteredDept);
                    }
                    return {
                        value: obj.ID,
                        label: obj.Department,
                        Department: obj.Department,
                        DepartmentCode: obj.DepartmentCode,
                        ToUsers: filteredDept?.ToUsers || [],
                        CCUsers: filteredDept?.CCUsers || [],
                        ToUsersTitle: filteredDept?.ToUsersTitle || [],
                        CCUsersTitle: filteredDept?.CCUsersTitle || [],

                    };

                }) || []);

                setselectUserDeptTo(setBannerById[0].ToDepartments?.map((obj: any) => {
                    const filteredDept = setAllDept1.find((dept: any) => dept.value === obj.ID);
                    if (filteredDept) {
                        filteredDeptArrayTo.push(filteredDept);
                    }
                    return {
                        value: obj.ID,
                        label: obj.Department,
                        Department: obj.Department,
                        DepartmentCode: obj.DepartmentCode,
                        ToUsers: filteredDept?.ToUsers || [],
                        CCUsers: filteredDept?.CCUsers || [],
                        ToUsersTitle: filteredDept?.ToUsersTitle || [],
                        CCUsersTitle: filteredDept?.CCUsersTitle || [],
                    };
                }) || []);
                if (setBannerById[0].AttachmentId) {
                    // setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].AttachmentId[0]));
                    let arrn = await getDocumentLinkByID(sp, setBannerById[0].AttachmentId);
                    setFilesArr([...FilesArr, ...arrn]);
                    setFilesArr1([...FilesArr1, ...arrn]);

                }

                const ApprowData: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_AuditPlan, setBannerById[0].ReferenceNumber)

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

                const rowData: any[] = await getItemByID2(sp, Number(setBannerById[0].ID)) //baseUrl
                const initialRows = rowData.map((item: any) => ({
                    id: item.Id,
                    // AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                    section: item.Section,
                    date: new Date(item.Date).toLocaleDateString("en-CA"),
                    startTime: item.Time,
                    auditorIds: item.AuditorId,
                    endTime: "",
                    auditor: item.Auditor ? { label: item.Auditor.Title, value: item.Auditor.ID } : null // Convert single object
                }));
                setRecommendationRows(initialRows);
                setRecommendationRowsEdit(initialRows);


            }





        }
        setFormLoading(false);

        // setRequesterRoleId(await getRequesterID(sp))
        setFormNameId(await getFormNameID(sp, CONTENTTYPE_AuditPlan))
        setListNameId(await getListNameID(sp, LIST_TITLE_AuditPlan))
        createTooltipContent(filteredDeptArrayCC);
        createTooltipContentTo(filteredDeptArrayTo)

        //}
        //#endregion


    };

    // const onSelect = async (selectedList: any) => {
    //     console.log(selectedList, "selectedList");
    //     setFormData(prevData => ({
    //         ...prevData,
    //         IssueNumber: selectedList.IssueNumber,
    //         ReferenceNumber: selectedList.ReferenceNumber,
    //         RevisionNumber: selectedList.RevisionNumber,
    //         ChangeRequestID: selectedList.ChangeRequestID,
    //         IssueDate: selectedList.IssueDate,
    //         LocationId: selectedList.LocationId,
    //         CustodianId: selectedList.CustodianId,
    //         SerialNumber: selectedList.SerialNumber,
    //         RevisionDate: selectedList.RevisionDate,
    //         AmendmentTypeId: selectedList.AmendmentTypeId,
    //         ClassificationId: selectedList.ClassificationId,
    //         ChangeRequestTypeId: selectedList.ChangeRequestTypeId,
    //         RequestTypeId: RequestTypeId,
    //         SubmiitedDate: selectedList.SubmiitedDate,
    //         SubmitStatus: selectedList.SubmitStatus,
    //         DocumentCode: selectedList.value,
    //         DocumentTypeId: selectedList.DocumentTypeId,
    //         Department: selectedList.Department,
    //         AttachmentId: selectedList.AttachmentId,
    //         AttachmentJson: selectedList.AttachmentJson
    //         // Format as YYYY-MM-DD
    //     }));
    //     setSelectedOption(selectedList);
    //     if (selectedList.AttachmentId) {
    //         // setDocumentLink(await getDocumentLinkByID(sp, selectedList.AttachmentId))
    //     }
    //     else {
    //         setDocumentLink(null);
    //     }  // Set the selected users
    // };

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
        // getMemoNumber();

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


    const validateForm = async (fmode: FormSubmissionMode) => {
        Array.from(document.getElementsByClassName("border-on-error")).forEach((element: Element) => {
            element.classList.remove("border-on-error");
        });
        const {
            memoNo,
            deptId,
            issueNo,
            revisionNo,
            from,
            fromEmail,
            to,
            subject,
            attachment,
            date,
            background,
            issues,
            recommendationforApproval,
            approval,
            CC,
            auditPlanTypeId,
            exclusions,
            boundary,
            objective,
            criteria,
            scope,
            ToDepartments,
            CCDepartments,
            assignedTo, } = formData;
        // const { description } = richTextValues;
        let valid = true;
        let validraft = true;
        let valid1 = true;
        let validRec = true;
        let validAudit = true;

        setValidSubmit(true);
        setValidCancelReason(true);
        setValidForwardTo(true);
        setValidAudit(true);
        setValidDraft(true);
        setValidDRecomm(true);
        let errormsg = "";

        if (fmode == FormSubmissionMode.SUBMIT) {
            if (!deptId) {
                document.getElementById("DeptID")?.classList.add("border-on-error");
                valid = false;
            }
            if (!from) {
                document.getElementById("fromEmail")?.classList.add("border-on-error");
                valid = false;
            }
            if (!ToDepartments.length) {
                document.getElementById("ToDept")?.classList.add("border-on-error");
                valid = false;
            }
            if (!CCDepartments.length) {
                document.getElementById("CCDept")?.classList.add("border-on-error");
                valid = false;
            }
            //  if (!FilesArr.length) {
            //     //Swal.fire('Error', 'Category is required!', 'error');
            //     valid = false;
            // }
            if (!subject) {
                document.getElementById("subject")?.classList.add("border-on-error");
                valid = false;
            }
            if (!date) {
                document.getElementById("date")?.classList.add("border-on-error");
                valid = false;
            }
            if (date == "Invalid Date") {
                document.getElementById("date")?.classList.add("border-on-error");
                valid = false;
            }
            if (!background) {
                document.getElementById("background")?.classList.add("border-on-error");
                valid = false;
            }
            if (!issues) {
                document.getElementById("issues")?.classList.add("border-on-error");
                valid = false;
            }
            if (!auditPlanTypeId.length) {
                // document.getElementById("date")?.classList.add("border-on-error");
                Array.from(document.getElementsByClassName("auditPlanType")).forEach((element: Element) => {
                    element.classList.add("border-on-error");
                });
                valid = false;
            }

            if (!recommendationRows.length) {
                // document.getElementById("date")?.classList.add("border-on-error");
                validRec = false;
            }

            if (recommendationRows.length > 0 && recommendationRows.every((row: any) => row.section.trim() !== "" && row.date.trim() !== "" && row.startTime.trim() !== "" && row.auditor != null && row.auditor.length != 0) == false) {
                // document.getElementById("date")?.classList.add("border-on-error");
                validRec = false;

                Array.from(document.getElementsByClassName("recommendClsErr")).forEach((element: Element) => {
                    if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "")) {
                        element.classList.add("border-on-error");
                    }
                    else if (element.tagName === "INPUT" && (element as HTMLInputElement).value.trim() === "") {
                        element.classList.add("border-on-error");
                    }


                });
            }
            if (!recommendationforApproval) {
                document.getElementById("recApp")?.classList.add("border-on-error");
                validRec = false;
            }
            if (!forwardToArr) {
                valid1 = false;
            }
            if (forwardToArr.length > 0 && forwardToArr.every((row: any) => row.role !== 0 && row.approvalType.trim() !== "" && row.approvers.length != 0) == false) {


                Array.from(document.getElementsByClassName("HierarchyClsErr")).forEach((element: Element) => {
                    if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "Enter Approver Name" || element.textContent?.trim() === "")) {
                        element.classList.add("border-on-error");
                    }
                    else if (element.tagName === "SELECT" && (element as HTMLSelectElement).value.trim() === "") {
                        element.classList.add("border-on-error");
                    }
                    else if (element.tagName === "SELECT" && (element as HTMLInputElement).value.trim() === "") {
                        element.classList.add("border-on-error");
                    }

                });

                // Array.from(document.getElementsByClassName("newse")).forEach((element: Element) => {
                //     if (element.tagName === "SELECT" && (element as HTMLSelectElement).value.trim() === "") {
                //         element.classList.add("border-on-error");
                //     }
                // });
                valid1 = false;
            }
            if (!exclusions) {
                document.getElementById("exclusions")?.classList.add("border-on-error");
                validAudit = false;
            }
            if (!boundary) {
                document.getElementById("boundary")?.classList.add("border-on-error");
                validAudit = false;
            }
            if (!objective) {
                document.getElementById("objective")?.classList.add("border-on-error");
                validAudit = false;
            }
            if (!criteria) {
                document.getElementById("criteria")?.classList.add("border-on-error");
                validAudit = false;
            }
            if (!scope) {
                document.getElementById("scope")?.classList.add("border-on-error");
                validAudit = false;
            }


            // setValidSubmit(valid);
            // setValidDRecomm(validRec);
            // setValidAudit(validAudit);
            // setValidForwardTo(valid1);

        }
        else {
            if (!date) {
                document.getElementById("date")?.classList.add("border-on-error");
                validraft = false;
            }
            if (date == "Invalid Date") {
                document.getElementById("date")?.classList.add("border-on-error");
                validraft = false;
            }
            if (!deptId) {
                document.getElementById("DeptID")?.classList.add("border-on-error");
                validraft = false;
            }
            // else if (selectedOption == null || !selectedOption.value) {
            //     //Swal.fire('Error', 'Entity is required!', 'error');
            //     valid = false;
            // }


            // setValidDraft(validraft);
            // setValidCancelReason(valid1);

        }

        if (valid == false || valid1 == false || validRec == false || validAudit == false || validraft == false) {
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');

            return false
        }
        else {
            return true
        }
        // return valid;
    };

    const getNewFileName = async (originalFileName: string): Promise<string> => {
        const userId = currentUser.Id; // Or however you get the current user ID
        const date = new Date();
        const fileExtension = originalFileName.split('.').pop();

        const components = [
            date.getFullYear(),
            (date.getMonth() + 1).toString().padStart(2, '0'),
            date.getDate().toString().padStart(2, '0'),
            date.getHours().toString().padStart(2, '0'),
            date.getMinutes().toString().padStart(2, '0'),
            date.getSeconds().toString().padStart(2, '0'),
            date.getMilliseconds().toString().padStart(3, '0')
        ];

        return `${userId}_${components.join('')}_${originalFileName}`;
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


                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditPlanDocs');



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

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {

                            MemoNumber: formData.memoNo,
                            // MemoSerialNumber:formData.memoSerialNo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            Boundary: formData.boundary,
                            AimObjective: formData.objective,
                            Criteria: formData.criteria,
                            Scope: formData.scope,
                            ToDepartmentsId: formData.ToDepartments || [],
                            CCDepartmentsId: formData.CCDepartments || [],
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            // DocumentName:"",
                            IsRework: "No",
                            // DigitalSignStatus                


                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || ""


                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;

                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditPlanIDId: editItemID, // Assuming "Title" column exists
                                Section: row.section,
                                Date: row.date,
                                Time: row.startTime,
                                AuditorId: row.auditorIds
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
                                    // ContentTitle: selectedOption.ReferenceNumber,

                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: forwardToArr?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(editItemID),
                                    RequestId: formData.memoNo,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Annual Audit Plan",
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
                                await sp.web.lists.getByTitle("AnnualAuditPlanRecommendationList").items.getById(item.id).delete();
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
                        Swal.fire('Submitted successfully.', '', 'success');
                        sessionStorage.removeItem("DocumentCancelId")
                        setTimeout(() => {

                            window.history.back();
                            // window.location.reload();
                            setTimeout(() => {
                                location.reload();
                            }, 100);
                            // let url = window.location.href;
                            // let baseUrl = url.split("#")[0];
                        }, 500);
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


                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditPlanDocs');


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

                        let arr = {

                            MemoNumber: formData.memoNo,
                            MemoSerialNumber: formData.memoSerialNo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            Boundary: formData.boundary,
                            AimObjective: formData.objective,
                            Criteria: formData.criteria,
                            Scope: formData.scope,
                            ToDepartmentsId: formData.ToDepartments || [],
                            CCDepartmentsId: formData.CCDepartments || [],
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            // DocumentName:"",
                            IsRework: "No",
                            // DigitalSignStatus                


                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || ""


                        }

                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }

                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                                Section: row.section,
                                Date: row.date,
                                Time: row.startTime,
                                AuditorId: row.auditorIds
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

                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: forwardToArr?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(postId),
                                    RequestId: formData.memoNo,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Annual Audit Plan",
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
                        Swal.fire('Submitted successfully.', '', 'success');
                        // sessionStorage.removeItem("bannerId")
                        setTimeout(() => {
                            window.location.reload();
                            // window.history.back();
                        }, 500);
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
                    console.log(result)
                    if (result.isConfirmed) {
                        setLoading(true);
                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditPlanDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    // DocumentName = file.name;
                                    const newFileName = await getNewFileName(file.name);
                                    DocumentName = newFileName;
                                    const fileAddResult = await folder.files.addChunked(newFileName, file);
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

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {

                            MemoNumber: formData.memoNo,
                            // MemoSerialNumber:formData.memoSerialNo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            Boundary: formData.boundary,
                            AimObjective: formData.objective,
                            Criteria: formData.criteria,
                            Scope: formData.scope,
                            ToDepartmentsId: formData.ToDepartments || [],
                            CCDepartmentsId: formData.CCDepartments || [],
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "No",
                            Status: "Save as draft",
                            // DocumentName:"",
                            IsRework: "No",
                            // DigitalSignStatus                


                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || ""


                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        //  ////////////////////////////

                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditPlanIDId: editItemID, // Assuming "Title" column exists
                                Section: row.section || "",
                                Date: row.date ? row.date : null,
                                Time: row.startTime || "",
                                AuditorId: row.auditorIds ? row.auditorIds : 0
                            }

                            if (row.id) {

                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {

                                if ((row.section.trim() == "" && row.date.trim() == "" && row.startTime.trim() == "" && (row.auditor == null || row.auditor.length == 0)) == false) {


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

                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role || 0,
                                Level: Number(item.level),
                                ApproversId: approversIds || [],
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "No",
                                Maxlevel: forwardToArr?.length,

                                // MainListID: String(editItemID),
                                MainListID: String(editItemID),
                                RequestId: formData.memoNo,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Annual Audit Plan",
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
                                await sp.web.lists.getByTitle("AnnualAuditPlanRecommendationList").items.getById(item.id).delete();
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
                        Swal.fire('Saved successfully.', '', 'success');
                        sessionStorage.removeItem("DocumentCancelId")
                        setTimeout(() => {

                            window.history.back();
                            // window.location.reload();
                            setTimeout(() => {
                                location.reload();
                            }, 100);
                        }, 1000);
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


                        // let galleryIds: any[] = [];

                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditPlanDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    // DocumentName = file.name;
                                    const newFileName = await getNewFileName(file.name);
                                    DocumentName = newFileName;
                                    const fileAddResult = await folder.files.addChunked(newFileName, file);
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

                        let arr = {

                            MemoNumber: formData.memoNo,
                            MemoSerialNumber: formData.memoSerialNo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            Boundary: formData.boundary,
                            AimObjective: formData.objective,
                            Criteria: formData.criteria,
                            Scope: formData.scope,
                            ToDepartmentsId: formData.ToDepartments || [],
                            CCDepartmentsId: formData.CCDepartments || [],
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "No",
                            Status: "Save as draft",
                            // DocumentName:"",
                            IsRework: "No",
                            // DigitalSignStatus                


                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || ""


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

                            if ((row.section.trim() == "" && row.date.trim() == "" && row.startTime.trim() == "" && (row.auditor == null || row.auditor.length == 0)) == false) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                                    Section: row.section || "",
                                    Date: row.date ? row.date : null,
                                    Time: row.startTime || "",
                                    AuditorId: row.auditorIds ? row.auditorIds : 0
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

                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role ? item.role : 0,
                                    Level: Number(item.level),
                                    ApproversId: approversIds || [],
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: forwardToArr?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(postId),
                                    RequestId: formData.memoNo,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Annual Audit Plan",
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
                        Swal.fire('Saved successfully.', '', 'success');
                        // sessionStorage.removeItem("bannerId")
                        setTimeout(() => {
                            window.location.reload();
                            // window.history.back();
                        }, 1000);
                    }
                })

            }
        }

    }



    const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>, libraryName: string, docLib: string) => {
        event.preventDefault();
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
                    // const folder = sp.web.getFolderByServerRelativePath('Socialfeedimages');
                    // const uploadResult = await folder.files.addChunked(file.name, file);
                    // console.log("File uploaded successfully", uploadResult);

                    // Generate the preview URL dynamically
                    // const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);

                    //previewFile(previewUrl);
                    const preview = URL.createObjectURL(fn);

                    // newfilepreview = preview
                    // // setPreviewUrl(preview);
                    // // setFileType(fileType);

                    // var arr = {};
                    // arr = {
                    //     // files: files,
                    //     libraryName: libraryName,
                    //     docLib: docLib,
                    //     name: fn.name,
                    //     fileName: fn.name,
                    //     fileSize: fn.size,
                    //     date: new Date().toLocaleDateString("en-GB", {
                    //         day: "2-digit",
                    //         month: "short",
                    //         year: "numeric"
                    //     }).replace(/ /g, "/"),
                    //     fileUrl: preview,
                    //     fileType: fileType,
                    //     //   previewUrl: previewUrl
                    // };
                    // uloadBannerImageFiles.push(arr);
                    // setFilesArr1(uloadBannerImageFiles);
                }




                // uloadBannerImageFiles.push(arr);
                setFilesArr([...FilesArr, ...files]);




            } else {
                Swal.fire("upload a document")
            }
        }
    };

    //   const generatePreviewUrl = async (serverRelativeUrl: string) => {
    //     // Encode the file name and construct the preview URL
    //     const encodedFilePath = encodeURIComponent(serverRelativeUrl);

    //     // Example:
    //     // serverRelativeUrl = "/sites/AlRostmani/test/DocumentLibraryInsideTest/Book.xlsx"
    //     const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    //     const siteUrl = window.location.origin;

    //     // const previewUrl = `${siteUrl}/sites/AlRostmani/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    //     const previewUrl = `${siteUrl}${locationPath}/ChangeRequestDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    //     // const previewUrl = `${siteUrl}/sites/SPFXDemo/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    //     console.log("Generated Preview URL:", previewUrl);
    //     if (previewUrl) {
    //       console.log("enter herr")
    //       const deletebut = document.getElementById('closeCommand') as HTMLElement
    //       if (deletebut) {
    //         console.log(" here ", deletebut)
    //       }
    //     }
    //     return previewUrl;
    //   };

    const handleDelete = (index: number) => {
        setFilesArr((prevFiles: any[]) => prevFiles.filter((_file: any, i: number) => i !== index));
    };


    const createTooltipContent = (deptArr: any) => {
        // return "This is your tooltip content";

        const tableHeader = `
    <thead>
                    <tr style="background-color: #f2f2f2;">
                        <th style="border: 1px solid #ddd; padding: 8px;">S.No</th>
                        <th style="border: 1px solid #ddd; padding: 8px;">Department</th>
                       
                        <th style="border: 1px solid #ddd; padding: 8px;">Users</th>
                    </tr>
    </thead>`;

        // Generate table rows from AllDept data
        const tableRows = deptArr.map((item: any, index: number) => {
            // Extract and format data
            const department = item.Department || '';

            const ccUsers = item.CCUsersTitle?.map((user: any) => user.Title).join(", ") || '';

            // Return formatted row
            return `
                        <tr>
                            <td style="border: 1px solid #ddd; padding: 8px;">${index + 1}</td>
        <td style="border: 1px solid #ddd; padding: 8px;">${department}</td>
       
        <td style="border: 1px solid #ddd; padding: 8px;">${ccUsers}</td>
      </tr>`;
        }).join('');

        // Combine header and rows into complete table
        const tooltipTable = `
    <table style="border-collapse: collapse; width: 100%;">
      ${tableHeader}
      <tbody>
        ${tableRows}
                </tbody>
    </table>
  `;

        // Set tooltip text
        settooltipText(tooltipTable);
    };

    const createTooltipContentTo = (deptArr: any) => {
        // return "This is your tooltip content";

        const tableHeader = `
    <thead>
                    <tr style="background-color: #f2f2f2;">
                        <th style="border: 1px solid #ddd; padding: 8px;">S.No</th>
                        <th style="border: 1px solid #ddd; padding: 8px;">Department</th>
                        <th style="border: 1px solid #ddd; padding: 8px;">Users</th>
                        
                    </tr>
    </thead>`;

        // Generate table rows from AllDept data
        const tableRows = deptArr.map((item: any, index: number) => {
            // Extract and format data
            const department = item.Department || '';
            const toUsers = item.ToUsersTitle?.map((user: any) => user.Title).join(", ") || '';
            // const ccUsers = item.CCUsersTitle?.map((user: any) => user.Title).join(", ") || '';

            // Return formatted row
            return `
                        <tr>
                            <td style="border: 1px solid #ddd; padding: 8px;">${index + 1}</td>
        <td style="border: 1px solid #ddd; padding: 8px;">${department}</td>
        <td style="border: 1px solid #ddd; padding: 8px;">${toUsers}</td>
       
      </tr>`;
        }).join('');

        // Combine header and rows into complete table
        const tooltipTable = `
    <table style="border-collapse: collapse; width: 100%;">
      ${tableHeader}
      <tbody>
        ${tableRows}
                </tbody>
    </table>
  `;


        settooltipText1(tooltipTable);
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
                                                        <h4 className="text-dark font-16 fw-bold mb-3">Memorandum</h4>
                                                        {/* <p className="sub-header">
                                                            Lorem ipsum dolor sit amet consectetur adipisicing elit. Numquam, autem.
                                                        </p> */}

                                                        <div className="row mb-3">
                                                            {AuditPlanType.map((row, index) => (<div className="col-lg-3">
                                                                <div className="mb-2">
                                                                    <div className="form-check">
                                                                        <input type="checkbox" className={`form-check-input auditPlanType ${(!ValidSubmit) ? "border-on-error" : ""}`} id={`auditPlanType_${row.Id}`} disabled={InputDisabled} checked={formData.auditPlanTypeId.includes(row.Id)}
                                                                            onChange={(e) => {
                                                                                setFormData((prevState) => {
                                                                                    const isChecked = e.target.checked;
                                                                                    const updatedAuditPlanTypeId = isChecked
                                                                                        ? [...prevState.auditPlanTypeId, row.Id] // Add ID if checked
                                                                                        : prevState.auditPlanTypeId.filter(id => id !== row.Id); // Remove ID if unchecked

                                                                                    if (isChecked) {
                                                                                        Array.from(document.getElementsByClassName("auditPlanType")).forEach((element: Element) => {
                                                                                            element.classList.remove("border-on-error");
                                                                                        });
                                                                                    }

                                                                                    return {
                                                                                        ...prevState,
                                                                                        auditPlanTypeId: updatedAuditPlanTypeId
                                                                                    };
                                                                                });
                                                                            }}
                                                                        />
                                                                        <label className="form-check-label" htmlFor="infoCheck">{row.AuditPlanType}</label>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                            ))}


                                                            {/* onChange={(e) => {
                                                                setFormData((prevState) => {
                                                                    const isChecked = e.target.checked;
                                                                    return {
                                                                        ...prevState,
                                                                        auditPlanTypeId: isChecked
                                                                            ? [...prevState.auditPlanTypeId, row.Id] // Add ID if checked
                                                                            : prevState.auditPlanTypeId.filter(id => id !== row.Id) // Remove ID if unchecked
                                                                    };
                                                                });
                                                            }} */}

                                                            {/* <div className="col-lg-3">
                                                                <div className="mb-2">
                                                                    <div className="form-check">
                                                                        <input type="checkbox" className="form-check-input" id="signCheck" checked={formData.signCheck} onChange={(e) => setFormData(prevState => ({
                                                                            ...prevState,
                                                                            signCheck: e.target.checked
                                                                        }))} />
                                                                        <label className="form-check-label" htmlFor="signCheck">Request for Signing</label>
                                                                    </div>
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-3">
                                                                <div className="mb-2">
                                                                    <div className="form-check">
                                                                        <input type="checkbox" className="form-check-input" id="approvalCheck" checked={formData.approvalCheck} onChange={(e) => setFormData(prevState => ({
                                                                            ...prevState,
                                                                            approvalCheck: e.target.checked
                                                                        }))} />
                                                                        <label className="form-check-label" htmlFor="approvalCheck">For Approval</label>
                                                                    </div>
                                                                </div>
                                                            </div> */}
                                                        </div>

                                                        <div style={{ clear: "both" }}></div>

                                                        <form className="form-horizontal">
                                                            <div className="row">
                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="Department" className="col-4 col-xl-3 col-form-label">Department<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <Select
                                                                                options={AllDept}
                                                                                isDisabled={InputDisabled}
                                                                                value={selectUserDept}
                                                                                name="deptId"
                                                                                id="DeptID"
                                                                                className={`newse  ${(!ValidSubmit) ? "border-on-error" : ""} ${(!ValidDraft) ? "border-on-error" : ""}`}
                                                                                // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                // onChange={handleDepartmentChange}
                                                                                onChange={(selectedOptions: any) => handleDepartmentChange(selectedOptions)}
                                                                                placeholder="Select Department"
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="memoNo" className="col-4 col-xl-3 col-form-label">Memo No<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="memoNo"
                                                                                value={formData.memoNo}
                                                                                onChange={(e) => setFormData({ ...formData, memoNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>


                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="fromEmail" className="col-4 col-xl-3 col-form-label">From<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <input
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="fromEmail"
                                                                                value={formData.fromEmail}
                                                                                disabled={true}
                                                                            // onChange={(e) => setFormData({ ...formData, from: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="to" className="col-4 col-xl-3 col-form-label">To

                                                                            <Icon
                                                                                iconName="Info"
                                                                                className="ms-1"
                                                                                // title={tooltipText}
                                                                                // data-html={true}
                                                                                // data-tip={tooltipText}
                                                                                data-tooltip-id="my-tooltip"
                                                                                style={{ fontSize: '14px', cursor: 'pointer' }}
                                                                            />
                                                                            <Tooltip
                                                                                id="my-tooltip"
                                                                                content={tooltipText1}
                                                                                className="custom-tooltip"
                                                                                render={({ content }) => (
                                                                                    <div dangerouslySetInnerHTML={{ __html: content }} />
                                                                                )}
                                                                                style={{
                                                                                    backgroundColor: 'white',
                                                                                    color: 'black',
                                                                                    zIndex: 999
                                                                                }}
                                                                            />
                                                                            <span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <Select
                                                                                options={AllDept}
                                                                                isDisabled={InputDisabled}
                                                                                value={selectUserDeptTo}
                                                                                isMulti
                                                                                name="to"
                                                                                id="ToDept"
                                                                                className={`newse  ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                // onChange={handleDepartmentChange}
                                                                                onChange={(selectedOptions: any) => handleDepartmentChangeTo(selectedOptions)}
                                                                                placeholder="Select"
                                                                            />
                                                                            {/* <Select
                                                                                options={rows1}
                                                                                isMulti
                                                                                // value={formData.to}
                                                                                value={selectToUsers}
                                                                                name="CC"

                                                                                className={`newse ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // onChange={(selectedOptions:any) => handleToChange(selectedOptions, 'CC')}
                                                                                onChange={(selectedOptions: any) => handleToChange(selectedOptions, 'to')}
                                                                                placeholder="Select"
                                                                                isDisabled={InputDisabled}
                                                                            /> */}
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="recommendation" className="col-4 col-xl-3 col-form-label">CC
                                                                            <Icon
                                                                                iconName="Info"
                                                                                className="ms-1"
                                                                                // title={tooltipText}
                                                                                // data-html={true}
                                                                                // data-tip={tooltipText}
                                                                                data-tooltip-id="my-tooltip2"
                                                                                style={{ fontSize: '14px', cursor: 'pointer' }}
                                                                            />

                                                                            <Tooltip
                                                                                id="my-tooltip2"
                                                                                content={tooltipText}
                                                                                className="custom-tooltip"

                                                                                render={({ content }) => (
                                                                                    <div dangerouslySetInnerHTML={{ __html: content }} />
                                                                                )}
                                                                                style={{
                                                                                    backgroundColor: 'white',
                                                                                    color: 'black',
                                                                                    zIndex: 999
                                                                                }}
                                                                            />
                                                                            <span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <Select
                                                                                options={AllDept}
                                                                                isDisabled={InputDisabled}
                                                                                isMulti
                                                                                value={selectUserDeptCC}
                                                                                name="CC"
                                                                                id="CCDept"
                                                                                className={`newse  ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                // onChange={handleDepartmentChange}
                                                                                onChange={(selectedOptions: any) => handleDepartmentChangeCC(selectedOptions)}
                                                                                // onChange={(selectedOptions: any) => setFormData({ ...formData, CC: selectedOptions })}
                                                                                placeholder="Select"

                                                                            />
                                                                            {/* <Select
                                                                            options={rows1}
                                                                            isMulti
                                                                            // value={formData.CC}
                                                                            value={selectCCUsers}
                                                                            name="CC"
                                                                            className={`newse ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                            onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                            // onChange={handleCCChange}
                                                                            // onChange={(selectedOptions) => setFormData({ ...formData, CC: selectedOptions })}
                                                                            placeholder="Select"
                                                                            isDisabled={InputDisabled}
                                                                        /> */}
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="issueNo" className="col-4 col-xl-3 col-form-label">Issue No<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <input

                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issueNo"
                                                                                value={formData.issueNo}
                                                                                onChange={(e) => setFormData({ ...formData, issueNo: e.target.value })}
                                                                                disabled={true}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="revisionNo" className="col-4 col-xl-3 col-form-label">Revision No<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="revisionNo"
                                                                                value={formData.revisionNo}
                                                                                onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="subject" className="col-4 col-xl-3 col-form-label">Subject<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <input
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                id="subject"
                                                                                value={formData.subject}
                                                                                onChange={(e) => {
                                                                                    setFormData({ ...formData, subject: e.target.value });
                                                                                    if (e.target.value) {
                                                                                        document.getElementById("subject")?.classList.remove("border-on-error");
                                                                                    }
                                                                                }}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="attachment" className="col-4 col-xl-3 col-form-label">Attachment</label>
                                                                        <div className="col-8 col-xl-9">

                                                                            <div>
                                                                                <input
                                                                                    type="file"
                                                                                    // className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                    className="form-control"
                                                                                    id="attachment"
                                                                                    accept=".jpg,.jpeg,.png,.gif,.bmp,.svg,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx"
                                                                                    onChange={(e) => onFileChange(e, "Gallery", "AnnualAuditPlanDocs")}
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

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="date" className="col-4 col-xl-3 col-form-label">Date<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">

                                                                            <input
                                                                                type="date"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}${(!ValidDraft) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="date"
                                                                                value={formData.date}
                                                                                onChange={(e) => {
                                                                                    setFormData({ ...formData, date: new Date(e.target.value).toLocaleDateString("en-CA") });
                                                                                    if (e.target.value) {
                                                                                        document.getElementById("date")?.classList.remove("border-on-error");
                                                                                    }
                                                                                }}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6 mb-3">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="background" className="col-4 col-xl-3 col-form-label">Background<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <textarea
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                id="background"
                                                                                value={formData.background}
                                                                                onChange={(e) => {
                                                                                    setFormData({ ...formData, background: e.target.value });
                                                                                    if (e.target.value) {
                                                                                        document.getElementById("background")?.classList.remove("border-on-error");
                                                                                    }
                                                                                }}
                                                                                disabled={InputDisabled}
                                                                            ></textarea>
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="issues" className="col-4 col-xl-3 col-form-label">Description<span className="text-danger1"> *</span></label>
                                                                        <div className="col-8 col-xl-9">
                                                                            <textarea
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issues"
                                                                                value={formData.issues}
                                                                                onChange={(e) => {
                                                                                    setFormData({ ...formData, issues: e.target.value });
                                                                                    if (e.target.value) {
                                                                                        document.getElementById("issues")?.classList.remove("border-on-error");
                                                                                    }

                                                                                }}
                                                                                disabled={InputDisabled}
                                                                            ></textarea>
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
                                                                <h3 className='text-dark font-16 fw-bold mb-3'>Recommendation</h3>
                                                            </div>
                                                            <div style={{ textAlign: 'right' }} className='col-sm-6'>
                                                                {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")} onClick={handleAddRecommendationRow}></img>}

                                                            </div>
                                                        </div>


                                                        <table id="tabRec" className='mtbalenew overhi'>
                                                            <thead>
                                                                <tr><th>Section<span className="text-danger1"> *</span></th>
                                                                    <th>Date<span className="text-danger1"> *</span></th>
                                                                    <th colSpan={2}>Time<span className="text-danger1"> *</span></th>
                                                                    <th>Auditor<span className="text-danger1"> *</span></th>
                                                                    {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <th>Action</th>}
                                                                </tr>
                                                            </thead>

                                                            <tbody>

                                                                {recommendationRows.map((row, index) => (
                                                                    <tr key={index}>
                                                                        <td>
                                                                            <input
                                                                                type="text"
                                                                                className={`form-control recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                value={row.section}
                                                                                onChange={(e) => handleRecommendationChange(index, 'section', e.target.value)}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                        </td>
                                                                        <td>
                                                                            <input
                                                                                type="date"
                                                                                className={`form-control recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                value={row.date}
                                                                                onChange={(e) => handleRecommendationChange(index, 'date', e.target.value)}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                        </td>
                                                                        <td>
                                                                            <input
                                                                                type="time"
                                                                                className={`form-control recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                value={row.startTime}
                                                                                onChange={(e) => handleRecommendationChange(index, 'startTime', e.target.value)}
                                                                                disabled={InputDisabled}
                                                                            />

                                                                        </td>

                                                                        <td>
                                                                            <Select
                                                                                options={rows1}
                                                                                // isMulti
                                                                                className={`recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                value={row.auditor}
                                                                                onChange={(selectedOptions: any) => handleRecommendationChange(index, 'auditor', selectedOptions)}
                                                                                placeholder="Select"
                                                                                isDisabled={InputDisabled}
                                                                            />
                                                                        </td>
                                                                        {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td>
                                                                            <img src={require("../assets/del.png")} onClick={() => handleDeleteRecommendationRow(index)} />

                                                                        </td>
                                                                        }
                                                                    </tr>
                                                                ))}
                                                            </tbody>

                                                        </table>



                                                        <TextField id="recApp" className={`form-control ${(!ValidDRecomm) ? "border-on-error" : ""}`} onChange={(e, newValue) => { setFormData(prevState => ({ ...prevState, recommendationforApproval: newValue })); if (newValue) { document.getElementById("recApp")?.classList.remove("border-on-error") } }} errorMessage={""} multiline autoAdjustHeight value={formData.recommendationforApproval} validateOnFocusOut={true} required={true} label="Recommendation for Approval" disabled={InputDisabled} />


                                                    </fieldset>
                                                </section>

                                                <div className="card mt-2">
                                                    <div className="card-body">
                                                        <h4 className="text-dark font-16 fw-bold mb-3">Audit Plan Detail</h4>

                                                        <div className="row">


                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="exclusions" className="form-label">Exclusions<span className="text-danger1"> *</span></label>
                                                                    <textarea

                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="exclusions"
                                                                        placeholder=""
                                                                        value={formData.exclusions}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, exclusions: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("exclusions")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="boundary" className="form-label">Boundary<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="boundary"
                                                                        placeholder=""
                                                                        value={formData.boundary}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, boundary: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("boundary")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="objective" className="form-label">Aim / Objective<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="objective"
                                                                        placeholder=""
                                                                        value={formData.objective}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, objective: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("objective")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="criteria" className="form-label">Criteria<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="criteria"
                                                                        placeholder=""
                                                                        value={formData.criteria}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, criteria: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("scope")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="scope" className="form-label">Scope<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="scope"
                                                                        placeholder=""
                                                                        value={formData.scope}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, scope: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("scope")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div>


                                                        </div>
                                                    </div>
                                                </div>



                                                {/* /////////////////%%%%%%%%%%%%%%%%%%%%%%%% */}

                                                {/* {modeValue === "approve" && editID != null && editID.Status === "Pending" && editID.CurrentUserRole !== "Initiator" && */}

                                                <div className="card mt-2" style={{ marginBottom: '17px' }}>
                                                    <div className="card-body">
                                                        <div className='row'>
                                                            <div className='col-sm-8'>
                                                                <h4 className="text-dark font-16 fw-bold mb-3 ">Approval Hierarchy</h4>
                                                                <label>Define approval hierarchy for the documents submitted by Team members in this folder.
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
                                                                                    className={`form-select HierarchyClsErr newse ${(!ValidForwardTo) ? "border-on-error" : ""} `}

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
                                                                                    className={`newse HierarchyClsErr ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                    // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                    onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                                                                    placeholder="Enter Approver Name"
                                                                                    isDisabled={InputDisabled}
                                                                                />



                                                                            </td>
                                                                            <td style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                                                {/* <label htmlFor="approvalType">Approval Type: </label> */}
                                                                                <select id="approvalType" value={row.approvalType} onChange={(e) => handleChange(e, row.level)} className={`newse HierarchyClsErr form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`} disabled={InputDisabled} >
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
                                                        <WorkflowAction currentItem={editID} ctx={props.context} ContentType={CONTENTTYPE_AuditPlan}
                                                            DisableApproval={false} DisableCancel={false}
                                                        />
                                                    ) : (<div></div>)
                                                }

                                                {/* ////////////Audit History card */}
                                                {/* {editID !== null && editID.length != 0 && modeValue === "approve" && */}
                                                {MainEditItem !== null && MainEditItem.length != 0 && MainEditItem?.Status !== "Save as draft" &&
                                                    <WorkflowAuditHistory ContentItemId={MainEditItem} ContentType={CONTENTTYPE_AuditPlan} ctx={props.context} />
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
                                                        <Modal.Title > <h4 className='font-16 text-dark fw-bold'>Attachment Details</h4>  <br></br>
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for Annual Audit Plan
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
                                                                    {/* {editForm && <th>File Link</th>} */}
                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Upload date</th>
                                                                    {/* {!InputDisabled && <th className='text-center'>Action</th>} */}
                                                                    <th className='text-center'>Action</th>
                                                                </tr>
                                                            </thead>
                                                            <tbody>
                                                                {FilesArr.length > 0 && (
                                                                    FilesArr.map((row: any, index: number) => (
                                                                        <tr>
                                                                            <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>
                                                                            <td title={row.name || row.FileLeafRef}>{row.name || row.FileLeafRef}</td>
                                                                            {/* {row.Id && <td style={{ textAlign: 'center' }} >
                                                                                <span onClick={() => OpenFile(row, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                    <FontAwesomeIcon icon={faDownload} /></span>
                                                                               {row.Id && <span onClick={() => OpenFile(row, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                    <FontAwesomeIcon icon={faEye} /></span>}
                                                                            </td>} */}
                                                                            {/* <td>{DocumentLink.Created
                                                                                        ? new Intl.DateTimeFormat('en-GB', {
                                                                                            day: '2-digit',
                                                                                            month: 'short',
                                                                                            year: 'numeric'
                                                                                        }).format(new Date(DocumentLink.Created)).replace(/ /g, "/")
                                                                                        : ""}</td> */}
                                                                            <td style={{ minWidth: '50px', maxWidth: '50px' }}>{row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                day: "2-digit",
                                                                                month: "short",
                                                                                year: "numeric"
                                                                            }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                day: "2-digit",
                                                                                month: "short",
                                                                                year: "numeric"
                                                                            }).replace(/ /g, "/")}</td>

                                                                            <td>
                                                                                {row.Id && (
                                                                                    <>
                                                                                        <span
                                                                                            onClick={() => OpenFile(row, "Download")}
                                                                                            style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                                        >
                                                                                            <FontAwesomeIcon icon={faDownload} />
                                                                                        </span>
                                                                                        <span
                                                                                            onClick={() => OpenFile(row, "Open")}
                                                                                            style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                                        >
                                                                                            <FontAwesomeIcon icon={faEye} />
                                                                                        </span>
                                                                                    </>
                                                                                )}

                                                                                {!InputDisabled && <img src={require("../assets/del.png")} style={{ cursor: "pointer" }} onClick={() => handleDelete(index)} />}
                                                                            </td>


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

const AnnualAuditPlan: React.FC<IAnnualAuditPlanProps> = (props) => (
    <Provider>
        <AnnualAuditPlanContext props={props} />
    </Provider>
);



export default AnnualAuditPlan;
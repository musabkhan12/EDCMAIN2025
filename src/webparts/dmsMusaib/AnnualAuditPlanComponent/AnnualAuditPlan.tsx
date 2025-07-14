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
import { CONTENTTYPE_AuditPlan, CONTENTTYPE_AuditPlanForm, CONTENTTYPE_AuditPlanTemp, CONTENTTYPE_Memo, LIST_AuditPlan, LIST_TITLE_AuditPlan, SITE_URL, Tenant_URL } from '../../../Shared/Constants';
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Modal } from 'react-bootstrap';
import { faDownload, faEye } from '@fortawesome/free-solid-svg-icons';
import CustomBreadcrumb from './CustomBreadcrumb/CustomBreadcrumb';
import { addAllProcessItem, addItem, addItem2, addItem3, addMemoNumber, getAllAuditType, getAllClassificationMaster, getAllDepartment, getAllMemoNumberList, getAllProcessData, getApprovalByID, getApprovalByID2, getDataRoles, getdigitalsignaturerequestbyID, getDocumentLinkByID, getDraftApprovalByID, getFormNameID, getGeneratedTemplateDoc, getItemByID, getItemByID2, getItemByID3, getLatestChangeRequestTemplateType, getListNameID, getRecommendationTypes, UpdateAllProcessItem, updateApprovalItem, updateDigitalsign, updateItem, updateItem2, updateItem3, updateMemoNumber, uploadAllFiles } from './AuditPlanService';
import { TextField } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Tooltip } from 'react-tooltip';
import 'react-tooltip/dist/react-tooltip.css';
import { Attachment } from '@pnp/sp/attachments';
import { DatePicker } from 'office-ui-fabric-react';
import moment from 'moment';
import { LocalizationProvider, TimePicker } from '@mui/x-date-pickers';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';
import FileViewer from '../components/fileviewer';

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

interface Location {
    locationId: number; // ID for Location lookup
    locationName: string;
    // locationCode: string;
}

const AnnualAuditPlanContext = ({ props }: any) => {
    const sp: SPFI = getSP();
    const elementRef = React.useRef<HTMLDivElement>(null);
    const siteUrl = props.siteUrl;
    const { useHide }: any = React.useContext(UserContext);
    const [InputDisabled, setInputDisabled] = React.useState(false);
    const selectedTextDiv = document.getElementById('selectedText');
    const [disableDepartment, setdisableDepartment] = React.useState(false);
    selectedTextDiv.style.display = 'none';


    const [FilesArr, setFilesArr] = React.useState<any>([]);
    const [TemplateDoc, setTemplateDoc] = React.useState<any>([]);
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
    const [MemoNumDrpdown, setMemoNumDrpdown] = React.useState([]);
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
    const [FormNameId, setFormNameVal] = React.useState(null);
    const [AuditPlanType, setAuditPlanType] = React.useState([]);
    const [editForm, setEditForm] = React.useState(false);
    const [modeValue, setmode] = React.useState("");
    const [currentUserDept, setcurrentUserDept] = React.useState("");
    const [selectUserDept, setselectUserDept] = React.useState(null);
    const [selectUserDeptTo, setselectUserDeptTo] = React.useState([]);
    const [selectUserDeptCC, setselectUserDeptCC] = React.useState(null);
    const [AllDept, setAllDept] = React.useState([]);
    const [DocumentLink, setDocumentLink] = React.useState(null);
    const [LocationOpt, setLocationOpt] = React.useState<any>([]);
    const [DraftApprovalItem, setDraftApprovalItem] = React.useState(null);
    const [RecommType, setRecommType] = React.useState([]);
    const [Classificationopt, setClassificationopt] = React.useState<any>([]);
    const [tooltipText, settooltipText] = React.useState("");
    const [tooltipText1, settooltipText1] = React.useState("");
    const [showModal, setShowModal] = React.useState(false);

    const [DigitalsignID, setDigitalsignID] = React.useState(null);
    const [hidedigisign, sethidedigisign] = React.useState(false);

    const [Showfile, setShowfile] = React.useState(false);
    const [redirecturl, setredirecturl] = React.useState(null);
    const [ShowModalTemplateDoc, setShowModalTemplateDoc] = React.useState(false);

    const [formData, setFormData] = React.useState({
        // infoCheck: false,
        // signCheck: false,
        // approvalCheck: false,
        MemoListId: 0,
        memoNo: null,
        MemoId: 0,
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
        // boundary: "",
        objective: "",
        criteria: "",
        scope: "",
        assignedTo: "",
        ToDepartments: [],
        CCDepartments: [],
        attachmentIds: null,
        attachmentJson: null,
        recommendationTypeId: 0,
        recommendationDetails: "",
        RecommendationTypeValue: "",
        DocCode: "",
        RevisionNo: "",
        IssueNo: "",
        RevisionDate: null,
        IssueDate: null,
        changeReqListID: 0,
        classificationValue: null,
        classificationId: 0,

        MDocumentCode: "",
        MIssueNumber: null,
        MRevisionNumber: null,
        MRevisionDate: "",
        MIssueDate: "",

        memoFileName: "",
        Status: "",


        RequesterNameId: null,
        RequesterDesignation: "",
        RequesterName: "",
        RequestDate: null,

    });

    const cancelModalAction = (refresh?: boolean,) => {
        debugger
        setredirecturl(window.location.href);
        //setShowfileNew(false);
        setShowModalTemplateDoc(false);
        setShowfile(false);
    }

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

    const handleDepartmentChange = async (selectedOption: any) => {
        setselectUserDept(selectedOption);
        let memo: number = 0;
        let memoId: number = 0;
        let listItems = [];
        let onloadDeptId: any;
        if (selectedOption == null) {
            onloadDeptId = AllDept.filter((user: any) => user.ADDepartmentName === currentUserDept)[0]?.value || 0;

            listItems = await sp.web.lists.getByTitle("MemoNumberLogic").items.filter(`Department/ID eq ${onloadDeptId}`).orderBy("SerialNumber", false).top(1)();


        }
        else {
            listItems = await sp.web.lists.getByTitle("MemoNumberLogic").items.filter(`Department/ID eq ${selectedOption.value}`).orderBy("SerialNumber", false).top(1)();

        }

        if (listItems.length > 0) {
            // if (modeValue == "") {
            memo = listItems[0].SerialNumber ? listItems[0].SerialNumber + 1 : 1;
            memoId = listItems[0].Id;
            // }
            // else {
            //     memo = listItems[0].SerialNumber;
            //     memoId = listItems[0].Id;

            // }

        } else {
            memo = 1;
            memoId = 0;

        }
        const formattedMemoSerialNo = memo < 10
            ? `00${memo}`
            : memo < 100
                ? `0${memo}`
                : memo;

        if (selectedOption == null) {
            setFormData((prevFormData) => ({
                ...prevFormData,
                MemoListId: memoId,
                memoSerialNo: memo,
                deptId: onloadDeptId,
                // memoNo: setAllDept1.filter(user => user.label === UserDept)[0]
                //     ? `${setAllDept1.filter(user => user.label === UserDept)[0].DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
                //     : `0/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
                memoNo: AllDept.filter(user => user.ADDepartmentName === currentUserDept)[0]
                    ? `${AllDept.filter(user => user.ADDepartmentName === currentUserDept)[0].DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
                    : `0/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`,
                memoFileName: AllDept.filter((user: any) => user.ADDepartmentName === currentUserDept)[0]
                    ? `${AllDept.filter((user: any) => user.ADDepartmentName === currentUserDept)[0].DepartmentCode}_${String(new Date().getMonth() + 1).padStart(2, '0')}_${formattedMemoSerialNo}`
                    : `0_${String(new Date().getMonth() + 1).padStart(2, '0')}_${formattedMemoSerialNo}`,
            }));


        }
        else {

            setFormData({
                ...formData,
                MemoListId: memoId,
                memoSerialNo: memo,
                deptId: selectedOption?.value,
                memoNo: `${selectedOption.DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`,
                memoFileName: `${selectedOption.DepartmentCode}_${String(new Date().getMonth() + 1).padStart(2, '0')}_${formattedMemoSerialNo}`,

            });

        }





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
        { id: 0, section: "", date: "", startTime: "", endTime: "", auditor: null, auditorIds: null, validtime: true }
    ]);

    const [coverageAuditCriteria, setcoverageAuditCriteria] = React.useState([
        { id: 0, ProcessActivity: "", date: "", dept: null, deptId: null, startTime: "", auditor: null, auditorIds: null, LocationId: null, Location: null, validtime: true }
    ]);

    const [coverageAuditCriteriaEdit, setcoverageAuditCriteriaEdit] = React.useState([]);

    const [recommendationRowsEdit, setRecommendationRowsEdit] = React.useState([]);

    const handleAddRecommendationRow = () => {
        setRecommendationRows([...recommendationRows, { id: 0, section: "", date: "", startTime: "", endTime: "", auditor: null, auditorIds: null, validtime: true }]);
    };

    const handleAddCoverageRow = () => {
        setcoverageAuditCriteria([...coverageAuditCriteria, { id: 0, ProcessActivity: "", date: "", dept: null, deptId: null, startTime: "", auditor: null, auditorIds: null, LocationId: null, Location: null, validtime: true }]);
    };

    const handleRecommendationChange = (index: number, field: string, value: any) => {
        let updatedRows;
        if (field == "auditor") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = recommendationRows.map((row, i) =>
                i === index ? { ...row, [field]: value ? value : null, auditorIds: value ? value.value : null } : row
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
    const handleDeleteCoverageRow = (index: number) => {
        const updatedRows = coverageAuditCriteria.filter((_, i) => i !== index);
        setcoverageAuditCriteria(updatedRows);
    };

    const handleCoverageRow = (index: number, field: string, value: any) => {
        let updatedRows;
        if (field == "auditor") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = coverageAuditCriteria.map((row, i) =>
                i === index ? { ...row, [field]: value ? value : null, auditorIds: value ? value.value : null } : row
            );

        }
        else if (field == "Location") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = coverageAuditCriteria.map((row, i) =>
                i === index ? { ...row, [field]: value ? value : null, LocationId: value ? value.value : null } : row
            );

        }
        else if (field == "dept") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = coverageAuditCriteria.map((row, i) =>
                i === index ? { ...row, [field]: value ? value : null, deptId: value ? value.value : null } : row
            );

        }
        else {
            updatedRows = coverageAuditCriteria.map((row, i) =>
                i === index ? { ...row, [field]: value } : row
            );

        }


        setcoverageAuditCriteria(updatedRows);
    };

    const fetchLocations = async () => {
        try {
            // Fetch the items
            const items = await sp.web.lists
                // .getByTitle("LocationMaster") // Your list name
                .getByTitle("AuditPlanLocationMaster") // Your list name
                .items
                // .filter("IsActive eq 'Yes'") // Filter active items
                // .select("ID", "Location", "LocationCode") // Select required fields
                .top(5000) // Limit number of records
                (); // Call get() to fetch data
            items.sort((a, b) => a.Location.localeCompare(b.Location));
            // Use map on the result to create the desired array structure
            const locations: Location[] = items.map((item: any) => ({
                locationId: item.ID, // Store the ID for lookup
                locationName: item.Location, // Name of the location
                // locationCode: item.LocationCode, // Code of the location
                label: item.Location,
                value: item.ID
            }));

            // Update state with the fetched locations
            setLocationOpt(locations);
            return locations;
        } catch (error) {
            console.error("Error fetching locations: ", error);
        }
    };


    // Handle change event
    const handleChange = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        event.preventDefault();
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
            ChildComponent: "IMS Audit Plan",
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


        var setAllDept1 = await getAllDepartment(sp);
        setAllDept(setAllDept1);


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
        if (formMode == "") {
            setFormData(prevFormData => ({
                ...prevFormData,
                date: new Date().toLocaleDateString("en-CA"),

            }));
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
        let memoId: number = 0;
        let filteredDeptArrayTo: any[] = [];
        let filteredDeptArrayCC: any[] = [];

        const Currusers: any = await getCurrentUser(sp, siteUrl);
        setCurrentUser(await getCurrentUser(sp, siteUrl));
        const userProfile = await sp.profiles.myProperties();
        setcurrentUserDept(userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "")
        const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
        // setselectUserDept(setAllDept1.filter(user => user.label === UserDept));
        setselectUserDept(setAllDept1.filter(user => user.ADDepartmentName === UserDept));
        const recommendationTypes = await getRecommendationTypes(sp);
        setRecommType(recommendationTypes);
        var ClassificationArr = await getAllClassificationMaster(sp);
        ClassificationArr.sort((a, b) => a.Classification.localeCompare(b.Classification));
        const optionsclassification = ClassificationArr.map((item: any) => ({
            value: item.ID,
            label: item.Classification,
            itemId: item.ID
        }));
        setClassificationopt(optionsclassification);
        // const AllMemoNumber = await getAllMemoNumberList(sp);
        // const formattedMemoNumbers = AllMemoNumber.map((memo: any) => ({
        //     label: memo.MemoNumber,
        //     value: memo.Id,
        //     ID: memo.ID,
        //     // DocumentCode: memo.DocumentCode,
        //     // IssueNumber: memo.IssueNumber,
        //     // RevisionNumber: memo.RevisionNumber,
        //     Background: memo.Background,
        //     Subject: memo.Subject,
        //     Issues: memo.Issues,
        //     // RevisionDate: memo.RevisionDate || null,
        //     // IssueDate: memo.IssueDate || null,
        //     AttachmentId: memo.AttachmentId || null,
        //     From: memo.From,
        //     To: memo.To,
        //     CC: memo.CC,
        //     AuditType: memo.AuditType,
        //     AuditTypeId: memo.AuditTypeId,
        //     Date: new Date(memo.Date).toLocaleDateString("en-CA"),
        //     ToId: memo.ToId || [],
        //     CcId: memo.CcId || [],
        //     ToDepartmentsId: memo.ToDepartmentsId || [],
        //     CCDepartmentsId: memo.CCDepartmentsId || [],
        //     ToDepartments: memo.ToDepartments || [],
        //     CCDepartments: memo.CCDepartments || [],
        //     Department: memo.Department,
        //     DepartmentId: memo.DepartmentId,
        //     RecommendedforApproval: memo.RecommendedforApproval,
        //     RecommendationType: memo.RecommendationType,
        //     RecommendationDetails: memo.RecommendationDetails || "",
        //     RecommendationTypeId: memo.RecommendationTypeId || 0,






        // }));
        // setMemoNumDrpdown(formattedMemoNumbers);
        if (formMode == "") {



            if (recommendationTypes.length > 0) {
                const defaultRecommendationType = recommendationTypes.find(type => type.RecommendationTypeValue === "Table");
                if (defaultRecommendationType) {
                    setFormData(prevFormData => ({
                        ...prevFormData,
                        RecommendationTypeValue: "Table",
                        recommendationTypeId: defaultRecommendationType.Id
                    }));
                }
            }

            // const AllMemoNumber = await getAllMemoNumberList(sp);
            // const formattedMemoNumbers = AllMemoNumber.map((memo: any) => ({
            //     label: memo.MemoNumber,
            //     value: memo.Id,
            //     ID: memo.ID,
            //     DocumentCode: memo.DocumentCode,
            //     IssueNumber: memo.IssueNumber,
            //     RevisionNumber: memo.RevisionNumber,
            //     Background: memo.Background,
            //     Subject: memo.Subject,
            //     Issues: memo.Issues,
            //     RevisionDate: memo.RevisionDate || null,
            //     IssueDate: memo.IssueDate || null,
            //     AttachmentId: memo.AttachmentId || null,
            //     From: memo.From,
            //     To: memo.To,
            //     CC: memo.CC,
            //     AuditType: memo.AuditType,
            //     AuditTypeId: memo.AuditTypeId,
            //     Date: new Date(memo.Date).toLocaleDateString("en-CA"),
            //     ToId: memo.ToId || [],
            //     CcId: memo.CcId || [],
            //     ToDepartmentsId: memo.ToDepartmentsId || [],
            //     CCDepartmentsId: memo.CCDepartmentsId || [],
            //     ToDepartments: memo.ToDepartments || [],
            //     CCDepartments: memo.CCDepartments || [],
            //     Department: memo.Department,
            //     DepartmentId: memo.DepartmentId,
            //     RecommendedforApproval: memo.RecommendedforApproval,
            //     RecommendationType: memo.RecommendationType,
            //     RecommendationDetails: memo.RecommendationDetails || "",
            //     RecommendationTypeId: memo.RecommendationTypeId || 0,





            // }));
            // setMemoNumDrpdown(formattedMemoNumbers);

            // const listItems = await sp.web.lists.getByTitle("AnnualAuditPlanList").items.orderBy("MemoSerialNumber", false).top(1)();
            // if (listItems.length > 0) {
            //     memo = listItems[0].MemoSerialNumber ? listItems[0].MemoSerialNumber + 1 : 1;

            // } else {
            //     memo = 1;

            // }

            // const onloadDeptId = setAllDept1.filter((user: any) => user.label === UserDept)[0]?.value || 0;
            const onloadDeptId = setAllDept1.filter((user: any) => user.ADDepartmentName === UserDept)[0]?.value || 0;

            const listItems = await sp.web.lists.getByTitle("MemoNumberLogic").items.filter(`Department/ID eq ${onloadDeptId}`).orderBy("SerialNumber", false).top(1)();
            if (listItems.length > 0) {
                if (modeValue == "") {
                    memo = listItems[0].SerialNumber ? listItems[0].SerialNumber + 1 : 1;
                    memoId = listItems[0].Id;
                }
                else {
                    memo = listItems[0].SerialNumber;
                    memoId = listItems[0].Id;

                }

            } else {
                memo = 1;
                memoId = 0;

            }
            const formattedMemoSerialNo = memo < 10
                ? `00${memo}`
                : memo < 100
                    ? `0${memo}`
                    : memo;

            setFormData((prevFormData) => ({
                ...prevFormData,
                MemoListId: memoId,
                memoSerialNo: memo,
                deptId: onloadDeptId,
                // memoNo: setAllDept1.filter(user => user.label === UserDept)[0]
                //     ? `${setAllDept1.filter(user => user.label === UserDept)[0].DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
                //     : `0/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
                memoNo: setAllDept1.filter(user => user.ADDepartmentName === UserDept)[0]
                    ? `${setAllDept1.filter(user => user.ADDepartmentName === UserDept)[0].DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`
                    : `0/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`,
                memoFileName: setAllDept1.filter((user: any) => user.ADDepartmentName === UserDept)[0]
                    ? `${setAllDept1.filter((user: any) => user.ADDepartmentName === UserDept)[0].DepartmentCode}_${String(new Date().getMonth() + 1).padStart(2, '0')}_${formattedMemoSerialNo}`
                    : `0_${String(new Date().getMonth() + 1).padStart(2, '0')}_${formattedMemoSerialNo}`,
            }));

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


            RequesterNameId: Currusers?.Id || "",
            RequesterDesignation: userProfile?.Title || "",
            RequesterName: userProfile?.DisplayName || "",
            RequestDate: new Date().toLocaleDateString("en-CA"),

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



            setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditPlan));



        }
        // formitemid =20;
        if (formitemid) {

            setEditItemID(Number(formitemid));


            const setBannerById = await getItemByID(sp, Number(formitemid))
            var increaseMemo = true;
            const draftedItem = await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditPlan);
            if (draftedItem != null && draftedItem != undefined && draftedItem.length > 0) {
                increaseMemo = false;
            }

            const newItem1 = await getdigitalsignaturerequestbyID(LIST_TITLE_AuditPlan, sp, Number(formitemid));
            // console.log("newItem1newItem1", newItem1);
            if (newItem1.length > 0) {
                setDigitalsignID(newItem1[0].ID)
            }

            if (setBannerById.length > 0) {
                debugger
                let varmemoNum = "";
                let varmemofilename = "";
                setEditForm(true);
                setMainEditItem(setBannerById[0]);
                setBannerById[0].label = setBannerById[0].MemoNumber;
                setBannerById[0].value = setBannerById[0].MemorandumIDId;
                if (formMode == "edit") {
                    const listItems = await sp.web.lists.getByTitle("MemoNumberLogic").items.filter(`Department/ID eq ${setBannerById[0]?.DepartmentId}`).orderBy("SerialNumber", false).top(1)();
                    // if (listItems.length > 0 && (setBannerById[0].Status == "Rework" || (setBannerById[0].Status == "Save as draft"))) {
                    if (listItems.length > 0 && ((setBannerById[0].Status == "Save as draft" && increaseMemo))) {
                        if (listItems[0].SerialNumber > setBannerById[0].MemoSerialNumber) {
                            memo = listItems[0].SerialNumber + 1;
                        }
                        else {
                            memo = setBannerById[0].MemoSerialNumber;
                        }
                        // memoId = listItems[0].Id;  

                    } else {
                        memo = setBannerById[0].MemoSerialNumber;
                        // memoId= 0;

                    }
                    const formattedMemoSerialNo = memo < 10
                        ? `00${memo}`
                        : memo < 100
                            ? `0${memo}`
                            : memo;
                    varmemoNum = `${setAllDept1.filter((user: any) => user.value === setBannerById[0].DepartmentId)[0].DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`;
                    varmemofilename = `${setAllDept1.filter((user: any) => user.value === setBannerById[0].DepartmentId)[0].DepartmentCode}_${String(new Date().getMonth() + 1).padStart(2, '0')}_${formattedMemoSerialNo}`;

                }
                else {
                    varmemoNum = setBannerById[0].MemoNumber;
                    memo = setBannerById[0].MemoSerialNumber;
                    varmemofilename = setBannerById[0].MemoNumber.replace(/\//g, "_");
                }

                let ClassificationVal = optionsclassification.filter((docType: { value: any; }) => docType.value === setBannerById[0].ClassificationId) || null;

                setFormData(prevData => ({
                    ...prevData,
                    // label: setBannerById[0].MemoNumber,
                    // value: setBannerById[0].MemorandumIDId,
                    // MemoListId: memoId,
                    MemoId: setBannerById[0].MemorandumIDId,
                    // memoNo: setBannerById[0],
                    // memoSerialNo: setBannerById[0].MemoSerialNumber,
                    memoNo: varmemoNum,
                    memoFileName: varmemofilename,
                    memoSerialNo: memo,
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
                    // boundary: setBannerById[0].Boundary,
                    objective: setBannerById[0].AimObjective,
                    criteria: setBannerById[0].Criteria,
                    scope: setBannerById[0].Scope,
                    // assignedTo: "",
                    ToDepartments: setBannerById[0].ToDepartmentsId || [],
                    CCDepartments: setBannerById[0].CCDepartmentsId || [],
                    attachmentIds: setBannerById[0].AttachmentId || null,
                    attachmentJson: setBannerById[0].AttachmentJson || null,
                    recommendationTypeId: setBannerById[0].RecommendationTypeId || 0,
                    RecommendationTypeValue: setBannerById[0].RecommendationType?.RecommendationTypeValue || "",
                    recommendationDetails: setBannerById[0].RecommendationDetails || "",
                    DocCode: setBannerById[0].DocumentCode || "",
                    RevisionNo: setBannerById[0].RevisionNumber,
                    IssueNo: setBannerById[0].IssueNumber,
                    RevisionDate: setBannerById[0].RevisionDate || null,
                    IssueDate: setBannerById[0].IssueDate || null,
                    classificationId: setBannerById[0].ClassificationId,
                    classificationValue: ClassificationVal?.[0] || null,

                    Status: setBannerById[0].Status,
                }));
                setdisableDepartment(setBannerById[0].DepartmentId ? true : false);

                setselectUserDept(setAllDept1.filter(user => user.value === setBannerById[0].DepartmentId)?.[0] || null);

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
                    // setFilesArr1([...FilesArr1, ...arrn]);

                }
                setTemplateDoc(await getGeneratedTemplateDoc(sp, Number(formitemid)));

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

                if (rowData.length > 0) {
                    const initialRows = rowData.map((item: any) => ({
                        id: item.Id,
                        // AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                        section: item.Section || "",
                        date: item.Date ? new Date(item.Date).toLocaleDateString("en-CA") : "",
                        startTime: item.Time || "",
                        auditorIds: item.AuditorId,
                        endTime: "",
                        validtime: true,
                        auditor: item.Auditor ? { label: item.Auditor.Title, value: item.Auditor.ID } : null // Convert single object
                    }));
                    if (setBannerById[0].RecommendationType?.RecommendationTypeValue == "Table") {
                        setRecommendationRows(initialRows);
                    }
                    setRecommendationRowsEdit(initialRows);


                }

                const rowData1: any[] = await getItemByID3(sp, Number(setBannerById[0].ID)) //baseUrl

                if (rowData1.length > 0) {
                    const initialRows1 = rowData1.map((item: any) => ({
                        id: item.Id,
                        // AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                        ProcessActivity: item.ProcessActivity || "",
                        // StandardClauses: item.StandardClauses ||"",
                        date: item.Date ? new Date(item.Date).toLocaleDateString("en-CA") : "",
                        startTime: item.Time || "",
                        LocationId: item.LocationId,
                        dept: item?.Department ? { label: item.Department.Department, value: item.Department.ID } : null,
                        deptId: item.DepartmentId ? item.DepartmentId : null,
                        Location: item.Location ? { label: item.Location.Location, value: item.Location.ID } : null,
                        auditorIds: item.AuditorId,
                        validtime: true,
                        auditor: item.Auditor ? { label: item.Auditor.Title, value: item.Auditor.ID } : null // Convert single object
                    }));
                    // if (setBannerById[0].RecommendationType?.RecommendationTypeValue == "Table") {
                    setcoverageAuditCriteria(initialRows1);
                    // }
                    setcoverageAuditCriteriaEdit(initialRows1);


                }


            }





        }
        setFormLoading(false);


        setFormNameVal(await getFormNameID(sp, CONTENTTYPE_AuditPlanForm))
        setListNameId(await getListNameID(sp, LIST_TITLE_AuditPlan))
        createTooltipContent(filteredDeptArrayCC);
        createTooltipContentTo(filteredDeptArrayTo);
        let locationoptions = await fetchLocations();
        setRecommType(await getRecommendationTypes(sp));
        let ChangeRequestTemplateType = await getLatestChangeRequestTemplateType(sp, CONTENTTYPE_AuditPlanTemp);

        if (ChangeRequestTemplateType.length > 0) {
            const template = ChangeRequestTemplateType[0];
            setFormData((prevFormData) => ({
                ...prevFormData,
                changeReqListID: template.ID,
                DocCode: template.DocumentCode || "",
                RevisionNo: template.RevisionNumber,
                IssueNo: template.IssueNumber,
                RevisionDate: new Date(template.RevisionDate).toLocaleDateString("en-CA") || null,
                IssueDate: new Date(template.IssueDate).toLocaleDateString("en-CA") || null,

            }));
        }


        let ChangeRequestMemoTemplateType = await getLatestChangeRequestTemplateType(sp, CONTENTTYPE_Memo);

        if (ChangeRequestMemoTemplateType.length > 0) {
            const template = ChangeRequestMemoTemplateType[0];
            setFormData((prevFormData) => ({
                ...prevFormData,

                MDocumentCode: template.DocumentCode || "",
                MRevisionNumber: template.RevisionNumber,
                MIssueNumber: template.IssueNumber,
                MRevisionDate: new Date(template.RevisionDate).toLocaleDateString("en-CA") || null,
                MIssueDate: new Date(template.IssueDate).toLocaleDateString("en-CA") || null,

            }));
        }


        //}
        //#endregion


    };



    const onSelectApprovers = (selectedOptions: any, lvl: number) => {
        setForwardToArr((prev) =>
            prev.map((row) =>
                row.level === lvl ? { ...row, approvers: selectedOptions || [] } : row
            )
        );
    };



    const onSelectRole = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        event.preventDefault();
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
        const handleScroll = () => {
            // Close the dropdown on scroll
            document.activeElement && (document.activeElement as HTMLElement).blur();
        };

        const container = document.querySelector('.scroll-container');
        container?.addEventListener('scroll', handleScroll);

        return () => {
            container?.removeEventListener('scroll', handleScroll);
        };

    }, [useHide]);

    const handleCancel = () => {

        window.history.back();

        setTimeout(() => {
            location.reload();
        }, 100);

    }



    const OpenFile = (obj: any, sts: string) => {

        const fileUrl = `${Tenant_URL}${obj.FileRef}`;
        if (sts == "Open") {
            setShowfile(true);
        }

        if (sts == "Open") {
            if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {

                // window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=default`, "_blank");
                const viewerUrlppt = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=embedview`
                setredirecturl(viewerUrlppt);
            } else {
                // window.open(fileUrl, "_blank"); // Open PDF and other files normally
                setredirecturl(fileUrl);
            }

        } else if (sts == "Download") {
            const link = document.createElement("a");
            link.href = fileUrl;
            link.setAttribute("download", obj.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1')); // Suggests a filename for download
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
            // boundary,
            objective,
            criteria,
            scope,
            ToDepartments,
            CCDepartments,
            assignedTo,
            DocCode,
            RevisionNo,
            IssueNo,
            classificationValue,
            classificationId
        } = formData;
        // const { description } = richTextValues;
        let valid = true;
        let validraft = true;
        let valid1 = true;
        let validRec = true;
        let validAudit = true;

        const updatedRows1 = coverageAuditCriteria.map(row => ({
            ...row,
            validtime: true
        }));
        setcoverageAuditCriteria(updatedRows1); // triggers re-render

        const updatedRows = recommendationRows.map(row => ({
            ...row,
            validtime: true
        }));
        setRecommendationRows(updatedRows); // triggers re-render

        setValidSubmit(true);
        setValidCancelReason(true);
        setValidForwardTo(true);
        setValidAudit(true);
        setValidDraft(true);
        setValidDRecomm(true);
        let errormsg = "";

        if (fmode == FormSubmissionMode.SUBMIT) {
            if (!memoNo) {
                document.getElementById("memoNo")?.classList.add("border-on-error");
                valid = false;
            }
            if (!deptId) {
                document.getElementById("DeptID")?.classList.add("border-on-error");
                valid = false;
            }
            if (!classificationId) {
                document.getElementById("Classification")?.classList.add("border-on-error");
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
                // Array.from(document.getElementsByClassName("ms-TextField-fieldGroup")).forEach((element: Element) => {
                //     if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "" || element.textContent?.trim() === "")) {
                //       element.classList.add("border-on-error");
                //     }
                //   });
                Array.from(document.getElementsByClassName("ms-TextField-fieldGroup")).forEach((element: Element) => {
                    // Skip if the current element or any of its ancestors has the class "Exclude-date-picker"
                    // if ((element as HTMLElement).closest(".Exclude-date-picker")) {
                    //     return;
                    // }

                    const text = element.textContent?.trim();
                    if (
                        element.tagName === "DIV" &&
                        (text === "Select" || text === "" || text === "")
                    ) {
                        element.classList.add("border-on-error");
                    }
                });
                valid = false;
            }
            if (date == "Invalid Date") {
                document.getElementById("date")?.classList.add("border-on-error");
                // Array.from(document.getElementsByClassName("ms-TextField-fieldGroup")).forEach((element: Element) => {
                //     if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "" || element.textContent?.trim() === "")) {
                //       element.classList.add("border-on-error");
                //     }
                //   });

                Array.from(document.getElementsByClassName("ms-TextField-fieldGroup")).forEach((element: Element) => {
                    // Skip if the current element or any of its ancestors has the class "Exclude-date-picker"
                    // if ((element as HTMLElement).closest(".Exclude-date-picker")) {
                    //     return;
                    // }

                    const text = element.textContent?.trim();
                    if (
                        element.tagName === "DIV" &&
                        (text === "Select" || text === "" || text === "")
                    ) {
                        element.classList.add("border-on-error");
                    }
                });
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
            if (!formData.DocCode) {
                document.getElementById("docCode")?.classList.add("border-on-error");
                valid = false;
            }

            if (formData.IssueNo == null || formData.IssueNo == undefined || formData.IssueNo === "") {
                document.getElementById("IssueNo")?.classList.add("border-on-error");
                valid = false;
            }

            if (formData.RevisionNo == null || formData.RevisionNo == undefined || formData.RevisionNo === "") {
                document.getElementById("RevNo")?.classList.add("border-on-error");
                valid = false;
            }

            if (!formData.recommendationTypeId) {
                Array.from(document.getElementsByClassName("RecTypeClsErr")).forEach((element: Element) => {
                    element.classList.add("border-on-error");
                });
                valid = false;
            }

            if (formData.RecommendationTypeValue == "Table") {

                if (!recommendationRows.length) {
                    // document.getElementById("date")?.classList.add("border-on-error");
                    validRec = false;
                }

                if (recommendationRows.length > 0 && recommendationRows.every((row: any) => row.section.trim() !== "" && row.date.trim() !== "" && row.startTime !== null && row.startTime.trim() !== "" && row.auditor != null && row.auditor.length != 0) == false) {
                    // document.getElementById("date")?.classList.add("border-on-error");
                    validRec = false;
                    // recommendationRows.forEach(row => {
                    //     row.validtime = row.startTime === "" ? false : true;
                    //   });
                    const updatedRows = recommendationRows.map(row => ({
                        ...row,
                        validtime: (row.startTime === "" || row.startTime === null) ? false : true
                    }));
                    setRecommendationRows(updatedRows); // triggers re-render

                    Array.from(document.getElementsByClassName("recommendClsErr")).forEach((element: Element) => {
                        if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "")) {
                            element.classList.add("border-on-error");
                        }
                        else if (element.tagName === "INPUT" && (element as HTMLInputElement).value.trim() === "") {
                            element.classList.add("border-on-error");
                        }


                    });


                    Array.from(document.getElementsByClassName("ms-TextField-fieldGroup")).forEach((element: Element) => {
                        // Skip if the current element or any of its ancestors has the class "Exclude-date-picker"
                        // if ((element as HTMLElement).closest(".Exclude-date-picker")) {
                        //     return;
                        // }

                        const text = element.textContent?.trim();
                        if (
                            element.tagName === "DIV" &&
                            (text === "Select" || text === "" || text === "")
                        ) {
                            element.classList.add("border-on-error");
                        }
                    });
                }
            }
            else if (formData.RecommendationTypeValue === "TextBox") {
                if (!formData.recommendationDetails.trim()) {
                    document
                        .getElementById("recommendationDetails")
                        ?.classList.add("border-on-error");
                    validRec = false;
                }
            }
            // if (!recommendationforApproval) {
            //     document.getElementById("recApp")?.classList.add("border-on-error");
            //     validRec = false;
            // }
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
            if (!coverageAuditCriteria.length) {
                valid1 = false;
            }
            if (coverageAuditCriteria.length > 0 && coverageAuditCriteria.every((row: any) => row.Location != null && row.LocationId != null
                // && row.StandardClauses.trim() !== ""
                && row.ProcessActivity.trim() !== "" && row.date.trim() !== "" && row.startTime !== null && row.startTime.trim() !== "" && row.auditorIds != null && row.auditor != null && row.auditor.length != 0 && row.deptId != 0 && row.dept != null && row.dept != null) == false) {
                valid1 = false;

                // coverageAuditCriteria.forEach(row => {
                //     row.validtime = row.startTime === "" ? false : true;
                //   });

                const updatedRows1 = coverageAuditCriteria.map(row => ({
                    ...row,
                    validtime: (row.startTime === "" || row.startTime === null) ? false : true
                }));
                setcoverageAuditCriteria(updatedRows1); // triggers re-render

                Array.from(document.getElementsByClassName("coverageClsErr")).forEach((element: Element) => {
                    if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "Select department" || element.textContent?.trim() === "")) {
                        element.classList.add("border-on-error");
                    }
                    else if ((element.tagName === "INPUT" || element.tagName === "TEXTAREA") && (element as HTMLInputElement).value.trim() === "") {
                        element.classList.add("border-on-error");
                    }


                });

                Array.from(document.getElementsByClassName("ms-TextField-fieldGroup")).forEach((element: Element) => {
                    // Skip if the current element or any of its ancestors has the class "Exclude-date-picker"
                    // if ((element as HTMLElement).closest(".Exclude-date-picker")) {
                    //     return;
                    // }

                    const text = element.textContent?.trim();
                    if (
                        element.tagName === "DIV" &&
                        (text === "Select" || text === "" || text === "")
                    ) {
                        element.classList.add("border-on-error");
                    }
                });
            }
            if (!exclusions) {
                document.getElementById("exclusions")?.classList.add("border-on-error");
                validAudit = false;
            }
            // if (!boundary) {
            //     document.getElementById("boundary")?.classList.add("border-on-error");
            //     validAudit = false;
            // }
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
            // if (!memoNo) {
            //     document.getElementById("memoNo")?.classList.add("border-on-error");
            //     validraft = false;
            // }
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

    const getNewFileName = async (originalFileName: string, MemoNum: string): Promise<string> => {
        const userId = currentUser.Id; // Or however you get the current user ID
        const date = new Date();
        // const fileExtension = originalFileName.split('.').pop();

        // const components = [
        //     date.getFullYear(),
        //     (date.getMonth() + 1).toString().padStart(2, '0'),
        //     date.getDate().toString().padStart(2, '0'),
        //     date.getHours().toString().padStart(2, '0'),
        //     date.getMinutes().toString().padStart(2, '0'),
        //     date.getSeconds().toString().padStart(2, '0'),
        //     date.getMilliseconds().toString().padStart(3, '0')
        // ];

        // return `${userId}_${components.join('')}_${originalFileName}`;

        const components = [
            date.getDate().toString().padStart(2, '0'),
            (date.getMonth() + 1).toString().padStart(2, '0'),
            date.getFullYear().toString(),
            date.getHours().toString().padStart(2, '0'),
            date.getMinutes().toString().padStart(2, '0'),
            date.getSeconds().toString().padStart(2, '0'),
            date.getMilliseconds().toString().padStart(3, '0')
        ];
        const fileExtension = originalFileName.split('.').pop();
        const fileNameWithoutExtension = originalFileName.split('.').slice(0, -1).join('.');
        // const NewFileName = `${formData.memoFileName}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
        const NewFileName = MemoNum != "" ? `${MemoNum}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}` : `${formData.memoFileName}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;


        return NewFileName;
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
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditPlanDocs');



                        if (FilesArr.length > 0) {
                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    // DocumentName = file.name;
                                    const newFileName = await getNewFileName(file.name, "");
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

                            MemorandumIDId: formData.MemoId,

                            // MemoNumber: formData.memoNo.label,
                            MemoNumber: formData.memoNo,
                            MemoSerialNumber: formData.memoSerialNo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date ? formData.date : null,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            // Boundary: formData.boundary,
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
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            RecommendationTypeId: formData.recommendationTypeId,
                            RecommendationDetails: formData.RecommendationTypeValue === "TextBox" ? formData.recommendationDetails : "",
                            DocumentCode: formData.DocCode,
                            RevisionDate: formData.RevisionDate,
                            RevisionNumber: formData.RevisionNo,
                            IssueDate: formData.IssueDate,
                            IssueNumber: formData.IssueNo,
                            ChangeRequestIDId: formData.changeReqListID,
                            ClassificationId: formData.classificationId,

                            // MDocumentCode: formData.MDocumentCode,
                            // MRevisionNumber: formData.MRevisionNumber,
                            // MIssueNumber: formData.MIssueNumber,
                            // MRevisionDate: formData.MRevisionDate,
                            // MIssueDate:formData.MIssueDate


                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        if (formData.RecommendationTypeValue == "Table") {
                            for (const row of recommendationRows) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: editItemID, // Assuming "Title" column exists
                                    Section: row.section,
                                    Date: row.date || null,
                                    Time: row.startTime || "",
                                    AuditorId: row.auditorIds || null
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
                                    ContentTitle: formData.subject,

                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds || null,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: forwardToArr?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(editItemID),
                                    RequestId: formData.memoNo,
                                    // RequestId: String(formData.memoNo.label),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "IMS Audit Plan",
                                    FormNameId: FormNameId.Id,
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

                        for (const cov of coverageAuditCriteria) {

                            // if ((cov.ProcessActivity.trim() == "" &&cov.StandardClauses.trim() == "" && cov.date.trim() == "" && cov.startTime.trim() == "" && (cov.auditor == null || cov.auditor.length == 0)&& (cov.Location == null || cov.Location.length == 0)) == false) {

                            const postPayload2 = {
                                AnnualAuditPlanIDId: editItemID, // Assuming "Title" column exists
                                ProcessActivity: cov.ProcessActivity || "",
                                // StandardClauses: cov.StandardClauses || "",
                                DepartmentId: cov.deptId ? cov.deptId : null,
                                Date: cov.date ? cov.date : null,
                                Time: cov.startTime || "",
                                AuditorId: cov.auditorIds ? cov.auditorIds : null,
                                LocationId: cov.LocationId ? cov.LocationId : null,
                            }

                            if (cov.id) {

                                const postResult2 = await updateItem3(postPayload2, sp, cov.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {

                                // if ((cov.ProcessActivity.trim() == "" &&cov.StandardClauses.trim() == "" && cov.date.trim() == "" && cov.startTime.trim() == "" && (cov.auditor == null || cov.auditor.length == 0)&& (cov.Location == null || cov.Location.length == 0)) == false) {


                                const postResult2 = await addItem3(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }

                                // }



                            }
                            // }


                        }

                        const toDelete2 = coverageAuditCriteriaEdit.filter(
                            (itemEdit) => !coverageAuditCriteria.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete2) {
                            try {
                                await sp.web.lists.getByTitle("AnnualAuditPlanAuditCriteriaList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }




                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            let arr2 = {
                                ActionTakenById: currentUser.Id,
                                // ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                                ActionTakenOn: new Date().toISOString(),
                                // ActionTakenRoleId: formData.RequesterDesignation,
                                Status: "Approved",
                                // Remark: remark,

                            }
                            const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                            const postId = postResult?.data?.ID;

                        }
                        // else {
                        if (DraftApprovalItem == null || DraftApprovalItem == undefined || DraftApprovalItem.length == 0) {
                            if (modeValue != "approve") {
                                let arry = {
                                    DepartmentId: formData.deptId,
                                    SerialNumber: formData.memoSerialNo,
                                    ProcessName: FormNameId.FormName

                                }
                                // if (formData.MemoListId) {
                                //     const postResults = await updateMemoNumber(arry, sp, formData.MemoListId);
                                //     const postIds = postResults?.data?.ID;

                                // }
                                // else {
                                const postResults = await addMemoNumber(arry, sp);
                                const postIds = postResults?.data?.ID;

                                // }

                            }

                        }

                        // if (boolval == true) {
                        setLoading(false);
                        // Swal.fire('Submitted successfully.', '', 'success');
                        // sessionStorage.removeItem("DocumentCancelId")
                        // setTimeout(() => {

                        //     window.history.back();

                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        //     // let url = window.location.href;
                        //     // let baseUrl = url.split("#")[0];
                        // }, 500);
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
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


                        // //////*************** */
                        const listItems = await sp.web.lists.getByTitle("MemoNumberLogic").items.filter(`Department/ID eq ${formData.deptId}`).orderBy("SerialNumber", false).top(1)();
                        let memoNum;

                        let memo;
                        if (listItems.length > 0) {
                            // if (modeValue == "") {
                            memo = listItems[0].SerialNumber ? listItems[0].SerialNumber + 1 : 1;
                            // memoId = listItems[0].Id;
                            // }
                            // else {
                            //   memo = listItems[0].SerialNumber;
                            //   // memoId = listItems[0].Id;

                            // }

                        } else {
                            memo = 1;
                            // memoId = 0;

                        }
                        const formattedMemoSerialNo = memo < 10
                            ? `00${memo}`
                            : memo < 100
                                ? `0${memo}`
                                : memo;

                        memoNum = `${selectUserDept?.DepartmentCode}/${String(new Date().getMonth() + 1).padStart(2, '0')}/${formattedMemoSerialNo}`;

                        const memoFileName = memoNum.replace(/\//g, "_");



                        // //////////****************** */


                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditPlanDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    // DocumentName = file.name;
                                    const newFileName = await getNewFileName(file.name, memoFileName);
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

                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            RequestDate: new Date(formData.RequestDate).toISOString(),



                            // MemoNumber: formData.memoNo.label,
                            // MemoNumber: formData.memoNo,
                            MemorandumIDId: formData.MemoId,
                            // MemoSerialNumber: formData.memoSerialNo,
                            MemoNumber: memoNum,
                            MemoSerialNumber: memo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date || null,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            // Boundary: formData.boundary,
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
                            RecommendationTypeId: formData.recommendationTypeId,
                            RecommendationDetails: formData.RecommendationTypeValue === "TextBox" ? formData.recommendationDetails : "",
                            DocumentCode: formData.DocCode,
                            RevisionDate: formData.RevisionDate,
                            RevisionNumber: formData.RevisionNo,
                            IssueDate: formData.IssueDate,
                            IssueNumber: formData.IssueNo,
                            ChangeRequestIDId: formData.changeReqListID,
                            ClassificationId: formData.classificationId,
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",

                            MDocumentCode: formData.MDocumentCode,
                            MRevisionNumber: formData.MRevisionNumber,
                            MIssueNumber: formData.MIssueNumber,
                            MRevisionDate: formData.MRevisionDate,
                            MIssueDate: formData.MIssueDate


                        }

                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }
                        if (formData.RecommendationTypeValue == "Table") {
                            for (const row of recommendationRows) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                                    Section: row.section,
                                    Date: row.date || null,
                                    Time: row.startTime || "",
                                    AuditorId: row.auditorIds || null
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
                                    ContentTitle: formData.subject,

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
                                    // RequestId: formData.memoNo,
                                    RequestId: memoNum,
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "IMS Audit Plan",
                                    FormNameId: FormNameId.Id,
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

                        let arry = {
                            DepartmentId: formData.deptId,
                            // SerialNumber: formData.memoSerialNo,
                            SerialNumber: memo,
                            ProcessName: FormNameId.FormName
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            // Status: "Approved",
                            // Remark: remark,

                        }
                        // if (formData.MemoListId) {
                        //     const postResults = await updateMemoNumber(arry, sp, formData.MemoListId);
                        //     const postIds = postResult?.data?.ID;

                        // }
                        // else {
                        const postResults = await addMemoNumber(arry, sp);
                        const postIds = postResult?.data?.ID;

                        // }

                        for (const cov of coverageAuditCriteria) {

                            if ((cov.ProcessActivity.trim() == "" &&
                                //  cov.StandardClauses.trim() == "" &&
                                cov.date.trim() == "" && cov.startTime.trim() == "" && (cov.auditor == null || cov.auditor.length == 0) && (cov.Location == null || cov.Location.length == 0)) == false) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                                    ProcessActivity: cov.ProcessActivity || "",
                                    // StandardClauses: cov.StandardClauses || "",
                                    DepartmentId: cov.deptId ? cov.deptId : null,
                                    Date: cov.date ? cov.date : null,
                                    Time: cov.startTime || "",
                                    AuditorId: cov.auditorIds ? cov.auditorIds : null,
                                    LocationId: cov.LocationId ? cov.LocationId : null,
                                }

                                const postResult2 = await addItem3(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }
                            }


                        }




                        let boolval;

                        // if (boolval == true) {
                        setLoading(false);
                        // Swal.fire('Submitted successfully.', '', 'success');
                        // // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 500);
                        // }
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });

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
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditPlanDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    // DocumentName = file.name;
                                    const newFileName = await getNewFileName(file.name, "");
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

                        let arr = {};

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            arr = {

                                // MemorandumIDId: formData.MemoId,

                                // MemoNumber: formData.memoNo.label,
                                MemoNumber: formData.memoNo,
                                MemoSerialNumber: formData.memoSerialNo,
                                // IssueNumber:,
                                // RevisionNumber:,
                                AuditPlanTypeId: formData.auditPlanTypeId,
                                FromId: formData.from,
                                ToId: formData.to,
                                CcId: formData.CC,
                                Subject: formData.subject,
                                Date: formData.date || null,
                                Background: formData.background,
                                Issues: formData.issues,
                                RecommendedforApproval: formData.recommendationforApproval,
                                DepartmentId: formData.deptId,
                                Exclusions: formData.exclusions,
                                // Boundary: formData.boundary,
                                AimObjective: formData.objective,
                                Criteria: formData.criteria,
                                Scope: formData.scope,
                                ToDepartmentsId: formData.ToDepartments || [],
                                CCDepartmentsId: formData.CCDepartments || [],
                                SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                                SubmitStatus: "No",
                                // Status: "Save as draft",

                                // IsRework: "No",
                                // IsRework: "No",


                                Status: "Rework",
                                IsRework: "Yes",
                               


                                // //////
                                RecommendationTypeId: formData.recommendationTypeId,
                                RecommendationDetails: formData.RecommendationTypeValue === "TextBox" ? formData.recommendationDetails : "",
                                DocumentCode: formData.DocCode,
                                RevisionDate: formData.RevisionDate,
                                RevisionNumber: formData.RevisionNo,
                                IssueDate: formData.IssueDate,
                                IssueNumber: formData.IssueNo,
                                ChangeRequestIDId: formData.changeReqListID,
                                ClassificationId: formData.classificationId,
                                AttachmentId: attachmentIds || [],
                                AttachmentJson: JSON.stringify(bannerImageArray) || "",

                                // MDocumentCode: formData.MDocumentCode,
                                // MRevisionNumber: formData.MRevisionNumber,
                                // MIssueNumber: formData.MIssueNumber,
                                // MRevisionDate: formData.MRevisionDate,
                                // MIssueDate:formData.MIssueDate


                            }

                        }
                        else {
                            arr = {

                                // MemorandumIDId: formData.MemoId,

                                // MemoNumber: formData.memoNo.label,
                                MemoNumber: formData.memoNo,
                                MemoSerialNumber: formData.memoSerialNo,
                                // IssueNumber:,
                                // RevisionNumber:,
                                AuditPlanTypeId: formData.auditPlanTypeId,
                                FromId: formData.from,
                                ToId: formData.to,
                                CcId: formData.CC,
                                Subject: formData.subject,
                                Date: formData.date || null,
                                Background: formData.background,
                                Issues: formData.issues,
                                RecommendedforApproval: formData.recommendationforApproval,
                                DepartmentId: formData.deptId,
                                Exclusions: formData.exclusions,
                                // Boundary: formData.boundary,
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
                                RecommendationTypeId: formData.recommendationTypeId,
                                RecommendationDetails: formData.RecommendationTypeValue === "TextBox" ? formData.recommendationDetails : "",
                                DocumentCode: formData.DocCode,
                                RevisionDate: formData.RevisionDate,
                                RevisionNumber: formData.RevisionNo,
                                IssueDate: formData.IssueDate,
                                IssueNumber: formData.IssueNo,
                                ChangeRequestIDId: formData.changeReqListID,
                                ClassificationId: formData.classificationId,
                                AttachmentId: attachmentIds || [],
                                AttachmentJson: JSON.stringify(bannerImageArray) || "",

                                // MDocumentCode: formData.MDocumentCode,
                                // MRevisionNumber: formData.MRevisionNumber,
                                // MIssueNumber: formData.MIssueNumber,
                                // MRevisionDate: formData.MRevisionDate,
                                // MIssueDate:formData.MIssueDate


                            }
                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        //  ////////////////////////////
                        if (formData.RecommendationTypeValue == "Table") {
                            for (const row of recommendationRows) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: editItemID, // Assuming "Title" column exists
                                    Section: row.section || "",
                                    Date: row.date ? row.date : null,
                                    Time: row.startTime || "",
                                    AuditorId: row.auditorIds ? row.auditorIds : null
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
                                ContentTitle: formData.subject,

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
                                // RequestId: String(formData.memoNo.label),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "IMS Audit Plan",
                                FormNameId: FormNameId.Id,
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

                        for (const cov of coverageAuditCriteria) {

                            if ((cov.ProcessActivity.trim() == ""
                                //  && cov.StandardClauses.trim() == ""
                                && cov.date.trim() == "" && cov.startTime.trim() == "" && (cov.auditor == null || cov.auditor.length == 0) && (cov.Location == null || cov.Location.length == 0)) == false) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: editItemID, // Assuming "Title" column exists
                                    ProcessActivity: cov.ProcessActivity || "",
                                    // StandardClauses: cov.StandardClauses || "",
                                    DepartmentId: cov.deptId ? cov.deptId : null,
                                    Date: cov.date ? cov.date : null,
                                    Time: cov.startTime || "",
                                    AuditorId: cov.auditorIds ? cov.auditorIds : null,
                                    LocationId: cov.LocationId ? cov.LocationId : null,
                                }

                                if (cov.id) {

                                    const postResult2 = await updateItem3(postPayload2, sp, cov.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {

                                    if ((cov.ProcessActivity.trim() == ""
                                        // && cov.StandardClauses.trim() == "" 
                                        && cov.date.trim() == "" && cov.startTime.trim() == "" && (cov.auditor == null || cov.auditor.length == 0) && (cov.Location == null || cov.Location.length == 0)) == false) {


                                        const postResult2 = await addItem3(postPayload2, sp);
                                        const postId2 = postResult2?.data?.ID;
                                        if (!postId2) {
                                            console.error("Post creation failed.");
                                            return;
                                        }

                                    }



                                }
                            }


                        }

                        const toDelete2 = coverageAuditCriteriaEdit.filter(
                            (itemEdit) => !coverageAuditCriteria.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete2) {
                            try {
                                await sp.web.lists.getByTitle("AnnualAuditPlanAuditCriteriaList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }


                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            let arr2 = {
                                // ActionTakenById: currentUser.Id,
                                // ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                                // // ActionTakenRoleId: formData.RequesterDesignation,
                                // Status: "Save as draft",
                                // // Remark: remark,
                                Title: formData.subject,
                                ContentTitle: formData.subject,

                            }
                            const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                            const postId = postResult?.data?.ID;

                        }


                        let boolval = false;

                        // if (boolval == true) {
                        setLoading(false);
                        // Swal.fire('Saved successfully.', '', 'success');
                        // sessionStorage.removeItem("DocumentCancelId")
                        // setTimeout(() => {

                        //     window.history.back();
                        //     // window.location.reload();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        // }, 1000);
                        // }
                        Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;

                                // window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
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
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditPlanDocs');


                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    // DocumentName = file.name;
                                    const newFileName = await getNewFileName(file.name, "");
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

                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            RequestDate: new Date(formData.RequestDate).toISOString(),

                            MemorandumIDId: formData.MemoId,
                            // MemoNumber: formData.memoNo.label,
                            MemoNumber: formData.memoNo,
                            MemoSerialNumber: formData.memoSerialNo,
                            // IssueNumber:,
                            // RevisionNumber:,
                            AuditPlanTypeId: formData.auditPlanTypeId,
                            FromId: formData.from,
                            ToId: formData.to,
                            CcId: formData.CC,
                            Subject: formData.subject,
                            Date: formData.date || null,
                            Background: formData.background,
                            Issues: formData.issues,
                            RecommendedforApproval: formData.recommendationforApproval,
                            DepartmentId: formData.deptId,
                            Exclusions: formData.exclusions,
                            // Boundary: formData.boundary,
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
                            RecommendationTypeId: formData.recommendationTypeId,
                            RecommendationDetails: formData.RecommendationTypeValue === "TextBox" ? formData.recommendationDetails : "",
                            DocumentCode: formData.DocCode,
                            RevisionDate: formData.RevisionDate,
                            RevisionNumber: formData.RevisionNo,
                            IssueDate: formData.IssueDate,
                            IssueNumber: formData.IssueNo,
                            ChangeRequestIDId: formData.changeReqListID,
                            ClassificationId: formData.classificationId,
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",




                            MDocumentCode: formData.MDocumentCode,
                            MRevisionNumber: formData.MRevisionNumber,
                            MIssueNumber: formData.MIssueNumber,
                            MRevisionDate: formData.MRevisionDate,
                            MIssueDate: formData.MIssueDate


                        }
                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }
                        if (formData.RecommendationTypeValue == "Table") {
                            for (const row of recommendationRows) {

                                if ((row.section.trim() == "" && row.date.trim() == "" && row.startTime.trim() == "" && (row.auditor == null || row.auditor.length == 0)) == false) {

                                    const postPayload2 = {
                                        AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                                        Section: row.section || "",
                                        Date: row.date ? row.date : null,
                                        Time: row.startTime || "",
                                        AuditorId: row.auditorIds ? row.auditorIds : null
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
                                    ContentTitle: formData.subject,

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
                                    // RequestId: String(formData.memoNo.label),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "IMS Audit Plan",
                                    FormNameId: FormNameId.Id,
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

                        for (const cov of coverageAuditCriteria) {

                            if ((cov.ProcessActivity.trim() == ""
                                // && cov.StandardClauses.trim() == "" 
                                && cov.date.trim() == "" && cov.startTime.trim() == "" && (cov.auditor == null || cov.auditor.length == 0) && (cov.Location == null || cov.Location.length == 0)) == false) {

                                const postPayload2 = {
                                    AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                                    ProcessActivity: cov.ProcessActivity || "",
                                    // StandardClauses: cov.StandardClauses || "",
                                    DepartmentId: cov.deptId ? cov.deptId : null,
                                    Date: cov.date ? cov.date : null,
                                    Time: cov.startTime || "",
                                    AuditorId: cov.auditorIds ? cov.auditorIds : null,
                                    LocationId: cov.LocationId ? cov.LocationId : null,
                                }

                                const postResult2 = await addItem3(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }
                            }


                        }

                        setLoading(false);
                        // Swal.fire('Saved successfully.', '', 'success');
                        // // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 1000);
                        Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                                // window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
                    }
                })

            }
        }

    }

    // const handleKeyDown = (e: React.KeyboardEvent<HTMLSelectElement>) => {
    //     // if (e.key === 'Enter') {
    //     //   e.preventDefault(); // 🛑 Prevents page reload
    //     // }
    // };
    const handleKeyDowntext = (e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.key === 'Enter') {
            e.preventDefault();
        }
    };



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
    //     // serverRelativeUrl = "/sites/ededms/test/DocumentLibraryInsideTest/Book.xlsx"
    //     const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    //     const siteUrl = window.location.origin;

    //     // const previewUrl = `${siteUrl}/sites/ededms/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
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

    const handleMemoChange = async (selectedOption: any) => {
        setFormData({
            ...formData,
            background: selectedOption.Background,
            // IssueNo: selectedOption.IssueNumber,
            // RevisionNo: selectedOption.RevisionNumber,
            issues: selectedOption.Issues,
            // RevisionDate: selectedOption.RevisionDate,
            // IssueDate: selectedOption.IssueDate,
            // DocCode: selectedOption.DocumentCode,
            from: selectedOption.From?.ID,
            fromEmail: selectedOption.From?.EMail,
            subject: selectedOption.Subject,
            date: selectedOption.Date,
            deptId: selectedOption.DepartmentId,
            memoNo: selectedOption, // Set the selected memo object
            MemoId: selectedOption.value, // Set MemoId with the value of the selected item
            recommendationforApproval: selectedOption.RecommendedforApproval,
            recommendationDetails: selectedOption.RecommendationDetails,
            recommendationTypeId: selectedOption.RecommendationTypeId,
            RecommendationTypeValue: selectedOption.RecommendationType?.RecommendationTypeValue || "",
            auditPlanTypeId: selectedOption.AuditTypeId,
            ToDepartments: selectedOption.ToDepartmentsId || [],
            CCDepartments: selectedOption.CCDepartmentsId || [],
            to: selectedOption.ToId || [],
            CC: selectedOption.CcId || [],

        });
        let filteredDeptArrayTo: any[] = [];
        let filteredDeptArrayCC: any[] = [];
        setselectUserDept(AllDept.filter((user: any) => user.value === selectedOption.DepartmentId));
        setselectUserDeptCC(selectedOption.CCDepartments?.map((obj: any) => {
            const filteredDept = AllDept.find((dept: any) => dept.value === obj.ID);
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

        setselectUserDeptTo(selectedOption.ToDepartments?.map((obj: any) => {
            const filteredDept = AllDept.find((dept: any) => dept.value === obj.ID);
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
        const rowData: any[] = await getItemByID2(sp, Number(selectedOption.ID)) //baseUrl

        if (rowData.length > 0) {
            const initialRows = rowData.map((item: any) => ({
                id: item.Id,
                // AnnualAuditPlanIDId: postId, // Assuming "Title" column exists
                section: item.Section,
                date: new Date(item.Date).toLocaleDateString("en-CA"),
                startTime: item.Time || "",
                auditorIds: item.AuditorId,
                endTime: "",
                validtime: true,
                auditor: item.Auditor ? { label: item.Auditor.Title, value: item.Auditor.ID } : null // Convert single object
            }));
            // if (selectedOption.RecommendationType?.RecommendationTypeValue == "Table") {
            setRecommendationRows(initialRows);
            // }
            // setRecommendationRowsEdit(initialRows);


        }

        createTooltipContent(filteredDeptArrayCC);
        createTooltipContentTo(filteredDeptArrayTo);

    };





    // const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
    //     // if (dropdownOpen && highlightedIndex !== -1 && event.key === "Enter") {
    //     //   // Let the dropdown handle selection
    //     //   return;
    //     // }
    //     // If dropdown is closed or no option is highlighted, prevent form submit
    //     if (event.key === "Enter") {
    //       event.preventDefault();
    //     }
    //   };
    // const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>, field: string, indx: number) => {
    //     if (event.key === 'Enter') {
    //         // Get currently focused option from react-select
    //         //   const focusedOption = document.querySelector('.react-select__option--is-focused');

    //         //   if (focusedOption && focusedOption.textContent && selectUserDept != null) {
    //         let alreadySelected = false;
    //         // switch (field) {
    //         //     case 'deptId': if (selectUserDept != null && selectUserDept?.length > 0) {
    //         //         alreadySelected = true;
    //         //         if (alreadySelected) {
    //         //             //   console.log('Option already selected — preventing default.');
    //         //             event.preventDefault();
    //         //         }
    //         //         // else {
    //         //         //   console.log('Option not selected yet — allowing default.');
    //         //         // }
    //         //     }
    //         //         break;
    //         //     case 'auditor': if (recommendationRows[indx].auditor != null && recommendationRows[indx].auditor?.length != 0) {
    //         //         alreadySelected = true;
    //         //         if (alreadySelected) {

    //         //             event.preventDefault();
    //         //         }

    //         //     }
    //         //     case 'Covauditor': if (coverageAuditCriteria[indx].auditor != null && coverageAuditCriteria[indx].auditor?.length != 0) {
    //         //         alreadySelected = true;
    //         //         if (alreadySelected) {

    //         //             event.preventDefault();
    //         //         }

    //         //     }

    //         //     case 'dept': if (coverageAuditCriteria[indx].dept != null && coverageAuditCriteria[indx].dept?.length != 0) {
    //         //         alreadySelected = true;
    //         //         if (alreadySelected) {

    //         //             event.preventDefault();
    //         //         }

    //         //     }

    //         //     case 'Location': if (coverageAuditCriteria[indx].Location != null && coverageAuditCriteria[indx].Location?.length != 0) {
    //         //         alreadySelected = true;
    //         //         if (alreadySelected) {

    //         //             event.preventDefault();
    //         //         }

    //         //     }


    //         // }
    //     }
    // };

    // const [selectedOpt, setSelectedOpt] = React.useState(null);
    // const [menuIsOpen, setMenuIsOpen] = React.useState(false);
    // const [inputValue, setInputValue] = React.useState('');
    // const [focusedOption, setFocusedOption] = React.useState(null);

    // const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
    //     if (event.key === 'Enter') {
    //       // ✅ If menu is open and an option is focused, allow default behavior
    //       if (menuIsOpen && focusedOption) {
    //         console.log('Focused option exists — allow Enter');
    //         return;
    //       }

    //       // ❌ No focused option, already selected — prevent
    //       if (selectedOpt) {
    //         console.log('No focused option — prevent default');
    //         event.preventDefault();
    //       }
    //     }
    //   };

    const updatedigisignnew = async () => {
        let items = await updateDigitalsign(LIST_TITLE_AuditPlan, sp, DigitalsignID);
        if (items) {
            sethidedigisign(true);
        }
    }

    const OpenFileTemplate = (obj: any, sts: string) => {
        debugger
        setShowModalTemplateDoc(true);
        const fileUrl = `${Tenant_URL}${obj.FileRef}`;
        if (sts == "Open") {
            setShowfile(true);
        }
        if (sts == "Open") {
            if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
                const viewerUrl = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=embedview`;

                //window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=view`);
                setredirecturl(viewerUrl);
            } else {
                setredirecturl(fileUrl);
                //window.open(fileUrl, "_blank"); // Open PDF and other files normally
            }

        } else if (sts == "Download") {
            const link = document.createElement("a");
            link.href = fileUrl;
            link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        }

    }
    const cleanFileName = (filename: string) => {
        // Match a 14-digit datetime suffix before the file extension
        const datetimePattern = /_\d{14}(?=\.[^.]+$)/;

        if (datetimePattern.test(filename)) {
            return filename.replace(datetimePattern, '');
        }

        return filename;
    }


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
                                                        {/* <h4 style={{ textAlign: 'left' }} className="text-dark font-16 fw-bold mb-3">Memorandum</h4> */}
                                                        <div className="previewIcon">
                                                            <h4 style={{ textAlign: 'left', margin: 'inherit' }} className="text-dark font-16 fw-bold mb-3">Memorandum</h4>

                                                            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>

                                                                {(formData.Status === "Approved" || formData.Status === "Rejected") && !hidedigisign && DigitalsignID != null && (
                                                                    <span
                                                                        onClick={() => updatedigisignnew()}
                                                                        style={{ cursor: "pointer" }}
                                                                    >
                                                                        <div className="" title='Sync digital signed document from Signing Hub'>
                                                                            <img
                                                                                style={{ cursor: 'pointer', height: '40px' }}
                                                                                className='mt-0'
                                                                                src={require("../assets/digisign.png")}
                                                                                alt="Signature Icon"
                                                                            />
                                                                        </div>
                                                                    </span>
                                                                )}
                                                                {TemplateDoc && TemplateDoc.length > 0 && <div className='btn btn-primary'
                                                                    onClick={() => OpenFileTemplate(TemplateDoc[0], "Open")}

                                                                >
                                                                    {(() => {
                                                                        const parts = TemplateDoc[0]?.FileRef?.split('/');
                                                                        const folderName = parts && parts[3] ? parts[3] : null;
                                                                        return folderName === "AnnualAuditPlanDigitalSignedDocs" ? (
                                                                            // <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../../assets/noun-download-5006210.png")} alt="Download Icon" />
                                                                            // <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../assets/digisigndownload.png")} alt="Digital Sign Download Icon" />
                                                                            <img style={{ cursor: 'pointer', height: '24px' }} className='mt-0' src={require("../assets/signicon.png")} alt="Digital Sign Download Icon" />


                                                                        ) : (
                                                                            <img style={{ cursor: 'pointer', height: '24px' }} className='mt-0' src={require("../assets/noun-download-5006210.png")} alt="Download Icon" />
                                                                            // <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../../assets/digisigndownload.png")} alt="Digital Sign Download Icon" />
                                                                        );
                                                                    })()}

                                                                    {/* <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../assets/noun-download-5006210.png")} ></img> */}
                                                                    {/* <FontAwesomeIcon icon={faEye} /> */}
                                                                </div>
                                                                }
                                                            </div>
                                                        </div>

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

                                                        <form className="form-horizontal" noValidate onSubmit={(e) => { e.preventDefault(); return false; }}>
                                                            {/*<div className="form-horizontal"> */}

                                                            <div className="row">
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className=" col-form-label">Memo No<span className="text-danger1"> *</span></label>
                                                                        <div className="">

                                                                            <input style={{ height: '47px' }}
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                title={formData.memoNo}
                                                                                // className="form-control"
                                                                                id="memoNo"
                                                                                value={formData.memoNo}
                                                                                onChange={(e) => setFormData({ ...formData, memoNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Document Code<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <input style={{ height: '47px' }}
                                                                                disabled
                                                                                type="text"
                                                                                title={formData.DocCode}
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="docCode"
                                                                                value={formData.DocCode}

                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Issue No<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <input style={{ height: '47px' }}
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="IssueNo"
                                                                                value={formData.IssueNo}
                                                                                title={formData.IssueNo}

                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Revision No<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <input style={{ height: '47px' }}
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="RevNo"
                                                                                value={formData.RevisionNo}
                                                                                title={formData.RevisionNo}

                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="Department" className="col-form-label">From Department<span className="text-danger1"> *</span></label>
                                                                        <div title={selectUserDept?.label || "Select Department"}>
                                                                            <Select
                                                                                // options={AllDept}
                                                                                options={AllDept.sort((a: any, b: any) => a.label.localeCompare(b.label))}
                                                                                isDisabled={InputDisabled || (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0 ? true : false)}
                                                                                value={selectUserDept}
                                                                                // onKeyDown={(e: any) => handleKeyDown(e, 'deptId', 0)}
                                                                                isClearable
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



                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="fromEmail" className="col-form-label">From<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <input style={{ height: '47px' }}
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="fromEmail"
                                                                                value={formData.fromEmail}
                                                                                disabled={InputDisabled}
                                                                                title={formData.fromEmail}
                                                                            // onChange={(e) => setFormData({ ...formData, from: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label style={{ display: 'flex' }} htmlFor="to" className="col-form-label">To

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
                                                                        <div className="" title={selectUserDeptTo?.map((dept: any) => dept.label).join(", ") || "Select Department"}
                                                                        >
                                                                            <Select
                                                                                // options={AllDept}
                                                                                options={AllDept.sort((a: any, b: any) => a.label.localeCompare(b.label))}

                                                                                isDisabled={InputDisabled}
                                                                                value={selectUserDeptTo}
                                                                                isMulti
                                                                                name="to"
                                                                                id="ToDept"
                                                                                className={`newse  ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                onChange={(selectedOptions: any) => handleDepartmentChangeTo(selectedOptions)}
                                                                                placeholder="Select"

                                                                            />

                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label style={{ display: 'flex' }} htmlFor="recommendation" className="col-form-label">CC
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
                                                                        <div className="" title={selectUserDeptCC?.map((dept: any) => dept.label).join(", ") || "Select Department"}>
                                                                            <Select
                                                                                // options={AllDept}
                                                                                options={AllDept.sort((a: any, b: any) => a.label.localeCompare(b.label))}

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

                                                                        </div>
                                                                    </div>
                                                                </div>





                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="subject" className="col-form-label">Subject<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <input style={{ height: '47px' }}
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                id="subject"
                                                                                value={formData.subject}
                                                                                title={formData.subject}
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



                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="date" className="col-form-label">Date<span className="text-danger1"> *</span></label>
                                                                        <div className="" title={
                                                                            formData?.date
                                                                                ? moment(formData.date).format('DD/MMM/YYYY')
                                                                                : ""
                                                                        }>

                                                                            <DatePicker id="date"
                                                                                value={
                                                                                    formData?.date
                                                                                        ? new Date(moment(formData?.date).format('YYYY-MM-DD'))
                                                                                        : null
                                                                                }
                                                                                onSelectDate={(date: Date | null) => {
                                                                                    setFormData({ ...formData, date: new Date(date).toLocaleDateString("en-CA") });
                                                                                    // setFormData({ ...formData, date: moment(date).format('DD/MMM/YYYY') });
                                                                                }}
                                                                                // maxDate={new Date()}
                                                                                minDate={new Date()}
                                                                                disabled={InputDisabled}
                                                                                formatDate={(date: any) => moment(date).format('DD/MMM/YYYY')}
                                                                            />

                                                                            {/* <input
                                                                                type="date"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}${(!ValidDraft) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="date"
                                                                                value={formData.date}
                                                                                title={formData.date}
                                                                                onChange={(e) => {
                                                                                    setFormData({ ...formData, date: new Date(e.target.value).toLocaleDateString("en-CA") });
                                                                                    if (e.target.value) {
                                                                                        document.getElementById("date")?.classList.remove("border-on-error");
                                                                                    }
                                                                                }}
                                                                                disabled={InputDisabled}
                                                                            /> */}
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-4">

                                                                    <div className="mb-3">
                                                                        <label htmlFor="DocumentCode" className=" col-form-label">Classification<span className="text-danger1"> *</span></label>
                                                                        {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                                                        <div
                                                                            title={formData.classificationValue?.label || "Select a classification"}
                                                                            style={{ width: "100%" }}
                                                                        >
                                                                            <Select
                                                                                options={Classificationopt}
                                                                                value={formData.classificationValue}
                                                                                name="Classification"
                                                                                isClearable
                                                                                id="Classification"
                                                                                className={`${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                onChange={(selectedOption: any) => {
                                                                                    setFormData({
                                                                                        ...formData,
                                                                                        classificationValue: selectedOption ? selectedOption : null,
                                                                                        classificationId: selectedOption ? selectedOption.value : 0,
                                                                                    });
                                                                                }}
                                                                                placeholder="Select Classification"
                                                                                isDisabled={InputDisabled}

                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <div className='d-flex justify-content-between'>
                                                                            <label htmlFor="attachment" className="col-form-label">Attachment</label>
                                                                            <div>

                                                                                <div>
                                                                                    {FilesArr.length > 0 ?
                                                                                        (<a style={{ fontSize: '0.875rem' }} onClick={() => setShowModal(true)}>
                                                                                            <FontAwesomeIcon icon={faPaperclip} />{" "}{FilesArr.length} {FilesArr.length > 0 ? "files" : "file"} Attached
                                                                                        </a>) : ""

                                                                                    }
                                                                                </div>
                                                                            </div>
                                                                        </div>

                                                                        <div>
                                                                            <input style={{ height: '47px', padding: '10px' }}
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





                                                                        {/* </div> */}

                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-12">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="background" className="col-form-label">Background<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <textarea style={{ height: '80px' }}
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                id="background"
                                                                                value={formData.background}
                                                                                title={formData.background}
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

                                                                <div className="col-lg-12">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="issues" className="col-form-label">Description<span className="text-danger1"> *</span></label>
                                                                        <div className="">
                                                                            <textarea style={{ height: '80px' }}
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issues"
                                                                                value={formData.issues}
                                                                                title={formData.issues}
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
                                                            {/* // changes */}
                                                            {/* </div> */}
                                                        </form>

                                                    </div>
                                                </div>

                                                <section className='card card-body mt-2'>
                                                    <fieldset>
                                                        <div className='row'>
                                                            <div className='col-sm-4'>
                                                                <h3 className='text-dark font-16 fw-bold mb-3'>Recommendation</h3>
                                                            </div>
                                                            <div className='col-sm-8'>
                                                                <div style={{ display: 'flex', justifyContent: 'flex-end', paddingTop: '0px' }}>
                                                                    <h5 style={{ textAlign: 'right', margin: '5px 24px 0px 0px' }} className="text-dark font-14 fw-bold mb-2">Select Recommendation type<span className="text-danger1"> *</span></h5>
                                                                    <div className=''>

                                                                        {RecommType.map((type, index) => (
                                                                            <div key={index} className="form-check form-check-inline">
                                                                                <input
                                                                                    className="form-check-input"
                                                                                    type="radio"
                                                                                    name="recommendationType"
                                                                                    id={`recommendationType_${type.Id}`}
                                                                                    value={type.Id}
                                                                                    disabled={InputDisabled}
                                                                                    onChange={(e) => setFormData({ ...formData, recommendationTypeId: Number(e.target.value), RecommendationTypeValue: type.RecommendationTypeValue })}
                                                                                    checked={formData.recommendationTypeId === type.Id}
                                                                                />
                                                                                <label className="form-check-label" htmlFor={`recommendationType_${type.Id}`}>
                                                                                    {type.RecommendationTypeValue}
                                                                                </label>
                                                                            </div>
                                                                        ))}
                                                                    </div>
                                                                </div>

                                                            </div>

                                                        </div>

                                                        <div className="row mb-1">
                                                            <div className="col-sm-12">


                                                            </div>
                                                            <div style={{ textAlign: 'right' }} className='col-sm-12 mt-0'>
                                                                {!InputDisabled && formData.RecommendationTypeValue === "Table" && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")} onClick={handleAddRecommendationRow}></img>}

                                                            </div>
                                                        </div>


                                                        {formData.RecommendationTypeValue === "Table" ? (
                                                            // className='newclasstabls scroll-container'
                                                            <div style={{ display: 'grid' }} >
                                                                <table id="tabRec" className='mtbalenew  overhi mb-3 cont-scroll-mtb'>
                                                                    <thead>
                                                                        <tr><th style={{ minWidth: '190px', maxWidth: '190px' }}>Section<span className="text-danger1"> *</span></th>
                                                                            <th>Date<span className="text-danger1"> *</span></th>
                                                                            <th colSpan={2}>Time<span className="text-danger1"> *</span></th>
                                                                            <th>Auditor<span className="text-danger1"> *</span></th>
                                                                            {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <th style={{ minWidth: '70px', maxWidth: '70px' }}>Action</th>}
                                                                        </tr>
                                                                    </thead>

                                                                    <tbody>

                                                                        {recommendationRows.map((row, index) => (
                                                                            <tr key={index}>
                                                                                <td style={{ minWidth: '190px', maxWidth: '190px' }}>
                                                                                    <input
                                                                                        type="text"
                                                                                        className={`form-control recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                        // className="form-control"
                                                                                        onKeyDown={handleKeyDowntext}
                                                                                        value={row.section}
                                                                                        title={row.section}
                                                                                        onChange={(e) => handleRecommendationChange(index, 'section', e.target.value)}
                                                                                        disabled={InputDisabled}
                                                                                    />
                                                                                </td>
                                                                                <td title={
                                                                                    row?.date
                                                                                        ? moment(row?.date).format('DD/MMM/YYYY')
                                                                                        : "Select a date"
                                                                                }>

                                                                                    <DatePicker
                                                                                        value={
                                                                                            row?.date
                                                                                                ? new Date(moment(row?.date).format('YYYY-MM-DD'))
                                                                                                : null
                                                                                        }
                                                                                        onSelectDate={(date: Date | null) => {
                                                                                            if (date) {
                                                                                                const formattedDate = new Date(date).toLocaleDateString("en-CA"); // Format as "yyyy-MM-dd"
                                                                                                handleRecommendationChange(index, 'date', formattedDate); // Pass formatted date to handleRecommendationChange
                                                                                            }
                                                                                        }}
                                                                                        // maxDate={new Date()}
                                                                                        minDate={new Date()}
                                                                                        disabled={InputDisabled}
                                                                                        formatDate={(date: any) => moment(date).format('DD/MMM/YYYY')}
                                                                                    />
                                                                                    {/* <input
                                                                                    type="date"
                                                                                    className={`form-control recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.date}
                                                                                    title={row.date}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'date', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                /> */}
                                                                                </td>
                                                                                {/* <td>
                                                                                    <input
                                                                                        type="time"
                                                                                        className={`form-control recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                        // className="form-control"
                                                                                        value={row.startTime}
                                                                                        title={row.startTime}
                                                                                        onChange={(e) => handleRecommendationChange(index, 'startTime', e.target.value)}
                                                                                        disabled={InputDisabled}
                                                                                    />

                                                                                </td> */}
                                                                                <td style={{ overflow: "inherit", minWidth: '152px', maxWidth: '152px' }} title={row.startTime ? new Date(`1970-01-01T${row.startTime}`).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true }) : "Select a time"}>
                                                                                    <LocalizationProvider dateAdapter={AdapterDateFns}>
                                                                                        <TimePicker
                                                                                            label="Select Time"
                                                                                            className={`form-control recommendClsErr `}
                                                                                            value={row.startTime ? new Date(`1970-01-01T${row.startTime}`) : null}
                                                                                            onChange={(newValue: any) => {
                                                                                                const formattedTime = newValue
                                                                                                    ? newValue.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false })
                                                                                                    : '';
                                                                                                handleRecommendationChange(index, 'startTime', formattedTime);
                                                                                            }}
                                                                                            disabled={InputDisabled}
                                                                                            slotProps={{
                                                                                                textField: {
                                                                                                    fullWidth: true,
                                                                                                    className: `form-control ${row.validtime == false ? 'ErrBorder-ErrColor' : ''}`
                                                                                                    // className: `form-control ${RowErrors[index]?.time ? 'border-on-error' : ''}`
                                                                                                }
                                                                                            }}
                                                                                        />
                                                                                    </LocalizationProvider></td>

                                                                                <td title={row?.auditor?.label || "Select Auditor"} >
                                                                                    <Select
                                                                                        options={rows1}
                                                                                        // isMulti
                                                                                        className={`recommendClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                        value={row.auditor}
                                                                                        isClearable
                                                                                        // onKeyDown={(e: any) => handleKeyDown(e, 'auditor', index)}
                                                                                        onChange={(selectedOptions: any) => handleRecommendationChange(index, 'auditor', selectedOptions)}
                                                                                        placeholder="Select"
                                                                                        isDisabled={InputDisabled}

                                                                                    // Added title tooltip
                                                                                    />
                                                                                </td>
                                                                                {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td style={{ minWidth: '70px', maxWidth: '70px' }}>
                                                                                    <img src={require("../assets/del.png")} onClick={() => handleDeleteRecommendationRow(index)} />

                                                                                </td>
                                                                                }
                                                                            </tr>
                                                                        ))}
                                                                    </tbody>

                                                                </table>
                                                            </div>
                                                        ) : formData.RecommendationTypeValue === "TextBox" ? (
                                                            <div className="row mb-3">
                                                                <div className="col-lg-12">
                                                                    <label htmlFor="recommendationDetails" className="form-label">
                                                                        Recommendation Details
                                                                    </label>
                                                                    <textarea style={{ height: '80px' }}
                                                                        id="recommendationDetails"
                                                                        className="form-control"
                                                                        value={formData.recommendationDetails || ""}
                                                                        title={formData.recommendationDetails || ""}
                                                                        onChange={(e) =>
                                                                            setFormData({ ...formData, recommendationDetails: e.target.value })
                                                                        }
                                                                        disabled={InputDisabled}
                                                                        placeholder="Enter recommendation details here"
                                                                    ></textarea>
                                                                </div>
                                                            </div>
                                                        ) : null}



                                                        {/* <TextField id="recApp" className={`form-control ${(!ValidDRecomm) ? "border-on-error" : ""}`} onChange={(e, newValue) => { setFormData(prevState => ({ ...prevState, recommendationforApproval: newValue })); if (newValue) { document.getElementById("recApp")?.classList.remove("border-on-error") } }} errorMessage={""} multiline autoAdjustHeight value={formData.recommendationforApproval} validateOnFocusOut={true} required={true} label="Recommendation for Approval" disabled={InputDisabled} /> */}
                                                        {/* ////// */}
                                                        {/* <div className="row mb-3">
                                                            <div className="col-lg-12">
                                                                <label htmlFor="recApp" className="form-label">
                                                                    Recommendation for Approval <span className="text-danger1"> *</span>
                                                                </label>
                                                                <textarea style={{ height: '80px' }}
                                                                    id="recApp"
                                                                    className={`form-control ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                    value={formData.recommendationforApproval || ""}
                                                                    title={formData.recommendationforApproval || ""}
                                                                    onChange={(e) => {
                                                                        setFormData((prevState) => ({
                                                                            ...prevState,
                                                                            recommendationforApproval: e.target.value,
                                                                        }));
                                                                        if (e.target.value) {
                                                                            document.getElementById("recApp")?.classList.remove("border-on-error");
                                                                        }
                                                                    }}
                                                                    disabled={InputDisabled}
                                                                    placeholder="Enter recommendation for approval"
                                                                ></textarea>

                                                            </div>
                                                        </div> */}
                                                        {/* ///// */}

                                                    </fieldset>
                                                </section>

                                                <div className="card mt-2">
                                                    <div className="card-body">
                                                        <h4 className="text-dark font-16 fw-bold mb-3">Audit Plan Detail</h4>

                                                        <div className="row">
                                                            <div className="col-lg-6">
                                                                <div className="mb-3">
                                                                    <label htmlFor="scope" className="form-label">Scope<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="scope"
                                                                        placeholder=""
                                                                        value={formData.scope}
                                                                        title={formData.scope}
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

                                                            <div className="col-lg-6">
                                                                <div className="mb-3">
                                                                    <label htmlFor="exclusions" className="form-label">Exclusions<span className="text-danger1"> *</span></label>
                                                                    <textarea

                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="exclusions"
                                                                        placeholder=""
                                                                        value={formData.exclusions}
                                                                        title={formData.exclusions}
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

                                                            {/* <div className="col-lg-6">
                                                                <div className="mb-3">
                                                                    <label htmlFor="boundary" className="form-label">Boundary<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="boundary"
                                                                        placeholder=""
                                                                        value={formData.boundary}
                                                                        title={formData.boundary}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, boundary: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("boundary")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div> */}

                                                            <div className="col-lg-6">
                                                                <div className="mb-3">
                                                                    <label htmlFor="objective" className="form-label">Aim / Objective<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="objective"
                                                                        placeholder=""
                                                                        value={formData.objective}
                                                                        title={formData.objective}
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

                                                            <div className="col-lg-6">
                                                                <div className="mb-3">
                                                                    <label htmlFor="criteria" className="form-label">Criteria<span className="text-danger1"> *</span></label>
                                                                    <textarea
                                                                        className={`form-control ${(!ValidAudit) ? "border-on-error" : ""}`}
                                                                        id="criteria"
                                                                        placeholder=""
                                                                        value={formData.criteria}
                                                                        title={formData.criteria}
                                                                        onChange={(e) => {
                                                                            setFormData({ ...formData, criteria: e.target.value })
                                                                            if (e.target.value) {
                                                                                document.getElementById("criteria")?.classList.remove("border-on-error");
                                                                            }
                                                                        }}
                                                                        disabled={InputDisabled}
                                                                    ></textarea>
                                                                </div>
                                                            </div>




                                                        </div>
                                                    </div>
                                                </div>

                                                {/* //// new change */}

                                                <section className='card card-body mt-2'>
                                                    <fieldset>
                                                        <div className='row'>
                                                            <div className='col-sm-6'>
                                                                <h3 className='text-dark font-16 fw-bold mb-3'>Coverage Of The Audit Criteria</h3>
                                                            </div>
                                                            <div style={{ textAlign: 'right' }} className='col-sm-6 mt-0'>
                                                                {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")} onClick={handleAddCoverageRow}></img>}

                                                            </div>

                                                        </div>




                                                        {/* {formData.RecommendationTypeValue === "Table" ? ( */}
                                                        <div style={{ display: 'grid' }} className='newclasstabls scroll-container'>
                                                            {/* mtbalenewscrollnew4 */}
                                                            {/* <table id="tabCov" className=' mtbalenew mb-3 cont-scroll-mtb'> */}
                                                            <table id="tabCov" className='mtbalenew  overhi mb-3 cont-scroll-mtb'>
                                                                <thead>
                                                                    <tr><th style={{ minWidth: '150px', maxWidth: '150px' }}>Date
                                                                        <span className="text-danger1"> *</span>
                                                                    </th>
                                                                        <th style={{ minWidth: '140px', maxWidth: '140px' }} colSpan={2}>Time
                                                                            <span className="text-danger1"> *</span>
                                                                        </th>
                                                                        <th style={{ minWidth: '200px', maxWidth: '200px' }}>Department
                                                                            <span className="text-danger1"> *</span>
                                                                        </th>
                                                                        <th style={{ minWidth: '260px', maxWidth: '260px' }}>Process/Activity
                                                                            <span className="text-danger1"> *</span>
                                                                        </th>
                                                                        <th style={{ minWidth: '200px', maxWidth: '200px' }}>Location
                                                                            <span className="text-danger1"> *</span>
                                                                        </th>
                                                                        <th style={{ minWidth: '200px', maxWidth: '200px' }}>Auditor
                                                                            <span className="text-danger1"> *</span>
                                                                        </th>
                                                                        {/* <th style={{ minWidth: '260px', maxWidth: '260px' }}>Standard & Clauses
                                                                            <span className="text-danger1"> *</span>
                                                                            </th> */}
                                                                        {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <th style={{ minWidth: '70px', maxWidth: '70px' }}>Action</th>}
                                                                    </tr>
                                                                </thead>

                                                                <tbody>

                                                                    {coverageAuditCriteria.map((row, index) => (
                                                                        <tr key={index}>
                                                                            <td style={{ overflow: "inherit", minWidth: '150px', maxWidth: '150px' }} title={
                                                                                row?.date
                                                                                    ? moment(row?.date).format('DD/MMM/YYYY')
                                                                                    : "Select a date"
                                                                            }>

                                                                                <DatePicker
                                                                                    // className='Exclude-date-picker'
                                                                                    value={
                                                                                        row?.date
                                                                                            ? new Date(moment(row?.date).format('YYYY-MM-DD'))
                                                                                            : null
                                                                                    }
                                                                                    onSelectDate={(date: Date | null) => {
                                                                                        if (date) {
                                                                                            const formattedDate = new Date(date).toLocaleDateString("en-CA"); // Format as "yyyy-MM-dd"
                                                                                            handleCoverageRow(index, 'date', formattedDate); // Pass formatted date to handleRecommendationChange
                                                                                        }
                                                                                    }}
                                                                                    // maxDate={new Date()}
                                                                                    minDate={new Date()}
                                                                                    disabled={InputDisabled}
                                                                                    formatDate={(date: any) => moment(date).format('DD/MMM/YYYY')}
                                                                                />
                                                                                {/* <input style={{ paddingLeft: '2px', paddingRight: '0px' }}
                                                                                    type="date"
                                                                                    className={`form-control  ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.date}
                                                                                    title={row.date}
                                                                                    onChange={(e) => handleCoverageRow(index, 'date', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                /> */}
                                                                            </td>
                                                                            {/* <td style={{ overflow: "inherit", minWidth: '140px', maxWidth: '140px' }}>
                                                                                <input style={{ paddingLeft: '2px', paddingRight: '0px' }}
                                                                                    type="time"
                                                                                    className={`coverageClsErr form-control  ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.startTime}
                                                                                    title={row.startTime}
                                                                                    onChange={(e) => handleCoverageRow(index, 'startTime', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />

                                                                            </td> */}
                                                                            <td style={{ overflow: "inherit", minWidth: '152px', maxWidth: '152px' }} title={row.startTime ? new Date(`1970-01-01T${row.startTime}`).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true }) : "Select a time"}>
                                                                                <LocalizationProvider dateAdapter={AdapterDateFns}>
                                                                                    <TimePicker
                                                                                        label="Select Time"
                                                                                        className={`form-control recommendClsErr`}
                                                                                        value={row.startTime ? new Date(`1970-01-01T${row.startTime}`) : null}
                                                                                        onChange={(newValue: any) => {
                                                                                            const formattedTime = newValue
                                                                                                ? newValue.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false })
                                                                                                : '';
                                                                                            handleCoverageRow(index, 'startTime', formattedTime);
                                                                                        }}
                                                                                        disabled={InputDisabled}
                                                                                        slotProps={{
                                                                                            textField: {
                                                                                                fullWidth: true,
                                                                                                className: `form-control ${row.validtime == false ? 'ErrBorder-ErrColor' : ''}`
                                                                                                // className: `form-control ${RowErrors[index]?.time ? 'border-on-error' : ''}`
                                                                                            }
                                                                                        }}
                                                                                    />
                                                                                </LocalizationProvider>
                                                                            </td>
                                                                            <td style={{ overflow: "inherit", minWidth: '200px', maxWidth: '200px' }} title={row.dept?.label || "Select department"} className='mtbalenew3'>
                                                                                <Select
                                                                                    options={AllDept.sort((a: any, b: any) => a.label.localeCompare(b.label))}
                                                                                    // isMulti
                                                                                    className={`coverageClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    value={row.dept}
                                                                                    isClearable
                                                                                    // onKeyDown={(e: any) => handleKeyDown(e, 'dept', index)}
                                                                                    title={row.dept?.label || "Select department"} // Added title tooltip
                                                                                    onChange={(selectedOptions: any) => handleCoverageRow(index, 'dept', selectedOptions)}
                                                                                    placeholder="Select department"
                                                                                    isDisabled={InputDisabled}

                                                                                    menuPortalTarget={document.body}
                                                                                    // styles={{ menuPortal: (base:any) => ({ ...base, zIndex: 9,position:'absolute'}) }}
                                                                                    styles={{
                                                                                        menu: (base: any) => ({
                                                                                            ...base,
                                                                                            position: 'absolute',
                                                                                            zIndex: 9,
                                                                                            top: '100%',
                                                                                            left: 0,
                                                                                        }),
                                                                                        container: (base: any) => ({
                                                                                            ...base,
                                                                                            zIndex: 0,
                                                                                            position: 'relative'
                                                                                        }),
                                                                                    }}
                                                                                />
                                                                            </td>
                                                                            <td style={{ minWidth: '260px', maxWidth: '260px', overflow: "inherit" }}>
                                                                                <textarea
                                                                                    className={`form-control coverageClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    value={row.ProcessActivity}
                                                                                    title={row.ProcessActivity}
                                                                                    onChange={(e) => handleCoverageRow(index, 'ProcessActivity', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                ></textarea>
                                                                            </td>



                                                                            <td style={{ overflow: "inherit", minWidth: '200px', maxWidth: '200px' }} title={row.Location?.label || "Select Location"}>
                                                                                <Select
                                                                                    options={LocationOpt}
                                                                                    // isMulti
                                                                                    className={`coverageClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    value={row.Location}
                                                                                    isClearable
                                                                                    // onKeyDown={(e: any) => handleKeyDown(e, 'Location', index)}
                                                                                    title={row.Location?.label || "Select Location"} // Added title tooltip
                                                                                    onChange={(selectedOptions: any) => handleCoverageRow(index, 'Location', selectedOptions)}
                                                                                    placeholder="Select"
                                                                                    isDisabled={InputDisabled}

                                                                                    menuPortalTarget={document.body}
                                                                                    // styles={{ menuPortal: (base:any) => ({ ...base, zIndex: 9,position:'absolute'}) }}
                                                                                    styles={{
                                                                                        menu: (base: any) => ({
                                                                                            ...base,
                                                                                            position: 'absolute',
                                                                                            zIndex: 9,
                                                                                            top: '100%',
                                                                                            left: 0,
                                                                                        }),
                                                                                        container: (base: any) => ({
                                                                                            ...base,
                                                                                            zIndex: 0,
                                                                                            position: 'relative'
                                                                                        }),
                                                                                    }}



                                                                                // 

                                                                                // onKeyDown={handleKeyDown}
                                                                                // onMenuOpen={() => setMenuIsOpen(true)}
                                                                                // onMenuClose={() => {
                                                                                //   setMenuIsOpen(false);
                                                                                //   setFocusedOption(null);
                                                                                // }}
                                                                                // onInputChange={(value: string, { action }: { action: string }) => {
                                                                                //   if (action === 'input-change') {
                                                                                //     setInputValue(value);
                                                                                //   }
                                                                                // }}
                                                                                // onFocus={() => setMenuIsOpen(true)}
                                                                                // menuIsOpen={menuIsOpen}
                                                                                // onMenuScrollToTop={() => setFocusedOption(null)}
                                                                                // onFocusOptionChange={(option:any) => {
                                                                                //   setFocusedOption(option); // Custom prop - not native
                                                                                // }}
                                                                                />
                                                                            </td>


                                                                            <td style={{ overflow: "inherit", minWidth: '200px', maxWidth: '200px' }} title={row?.auditor?.label || "Select Auditor"} >
                                                                                <Select
                                                                                    options={rows1}
                                                                                    // isMulti
                                                                                    className={`coverageClsErr ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    value={row.auditor}
                                                                                    isClearable
                                                                                    // onKeyDown={(e: any) => handleKeyDown(e, 'Covauditor', index)}
                                                                                    title={row.auditor?.label || "Select Auditor"} // Added title tooltip
                                                                                    onChange={(selectedOptions: any) => handleCoverageRow(index, 'auditor', selectedOptions)}
                                                                                    placeholder="Select"
                                                                                    isDisabled={InputDisabled}
                                                                                    menuPortalTarget={document.body}

                                                                                    // styles={{ menuPortal: (base:any) => ({ ...base, zIndex: 9,position:'absolute'}) }}
                                                                                    styles={{
                                                                                        menu: (base: any) => ({
                                                                                            ...base,
                                                                                            position: 'absolute',
                                                                                            zIndex: 9,
                                                                                            top: '100%',
                                                                                            left: 0,
                                                                                        }),
                                                                                        container: (base: any) => ({
                                                                                            ...base,
                                                                                            zIndex: 0,
                                                                                            position: 'relative'
                                                                                        }),
                                                                                    }}
                                                                                />
                                                                            </td>

                                                                            {/* <td style={{ minWidth: '260px', maxWidth: '260px', overflow: "inherit" }}>
                                                                                <textarea
                                                                                    className={`coverageClsErr form-control  ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                                                                    value={row.StandardClauses}
                                                                                    title={row.StandardClauses}
                                                                                    onChange={(e) => handleCoverageRow(index, 'StandardClauses', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                ></textarea>
                                                                            </td> */}
                                                                            {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td style={{ minWidth: '70px', maxWidth: '70px', overflow: "inherit" }}>
                                                                                <img src={require("../assets/del.png")} onClick={() => handleDeleteCoverageRow(index)} />

                                                                            </td>
                                                                            }
                                                                        </tr>
                                                                    ))}
                                                                </tbody>

                                                            </table>
                                                        </div>
                                                        {/* ) : formData.RecommendationTypeValue === "TextBox" ? (
                                                            <div className="row mb-3">
                                                                <div className="col-lg-12">
                                                                    <label htmlFor="recommendationDetails" className="form-label">
                                                                        Recommendation Details
                                                                    </label>
                                                                    <textarea
                                                                        id="recommendationDetails"
                                                                        className="form-control"
                                                                        value={formData.recommendationDetails || ""}
                                                                        onChange={(e) =>
                                                                            setFormData({ ...formData, recommendationDetails: e.target.value })
                                                                        }
                                                                        disabled={InputDisabled}
                                                                        placeholder="Enter recommendation details here"
                                                                    ></textarea>
                                                                </div>
                                                            </div>
                                                        ) : null} */}



                                                    </fieldset>
                                                </section>

                                                {/* ........ */}


                                                {/* /////////////////%%%%%%%%%%%%%%%%%%%%%%%% */}

                                                {/* {modeValue === "approve" && editID != null && editID.Status === "Pending" && editID.CurrentUserRole !== "Initiator" && */}

                                                <div className="card mt-2" style={{ marginBottom: '17px' }}>
                                                    <div className="card-body">
                                                        <div className='row'>
                                                            <div className='col-sm-8'>
                                                                <h4 className="text-dark font-16 fw-bold mb-1 ">Approval Hierarchy</h4>
                                                                <label>Define the approval hierarchy to ensure requests are routed to the appropriate approvers.
                                                                </label>
                                                            </div>
                                                            <div className='col-sm-4'>
                                                                <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right" }}>
                                                                    {!InputDisabled &&
                                                                        <img style={{ width: '30px' }} src={require("../assets/plus.png")} onClick={handleAddRow} className='' />
                                                                    }

                                                                </div>
                                                            </div>

                                                        </div>

                                                        <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">
                                                            <table style={{ overflow: 'inherit' }} className="mtbalenew  table-centered table-nowrap table-borderless mb-0 overhi" id="myTabl">
                                                                <thead >
                                                                    <tr>
                                                                        <th style={{ minWidth: "30px", maxWidth: "30px" }}>S.No</th>
                                                                        <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Role<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '40px', maxWidth: '40px' }} >Level</th>
                                                                        <th>Approver name<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '80px', maxWidth: '80px' }} >Approval criteria<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '40px', maxWidth: '40px' }}>Action</th>
                                                                    </tr>
                                                                </thead>
                                                                <tbody style={{ maxHeight: "8007px", overflow: 'inherit' }}>
                                                                    {forwardToArr.map((row, index) => (
                                                                        <tr>
                                                                            <td style={{ minWidth: "30px", maxWidth: "30px", overflow: 'inherit' }}> <div
                                                                                style={{ marginLeft: "5px" }}
                                                                                className="indexdesign"
                                                                            >
                                                                                {index + 1}</div>
                                                                            </td>
                                                                            <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px', }} className="ng-binding">
                                                                                <select
                                                                                    // className="form-select"
                                                                                    className={`form-select HierarchyClsErr newse ${(!ValidForwardTo) ? "border-on-error" : ""} `}
                                                                                    onChange={(e) => onSelectRole(e, row.level)}

                                                                                    value={row.role}
                                                                                    disabled={InputDisabled}
                                                                                    title={UserRoles.find((role: any) => role.value === row.role)?.label || "Select Role"} // Added title tooltip
                                                                                >
                                                                                    <option value="" selected>Select Role</option>
                                                                                    {UserRoles.filter((role: any) =>
                                                                                        !forwardToArr.some((r) => r.role === role.value && r.level !== row.level) || role.value === row.role // Allow the current row's role
                                                                                    ).map((role: any, idx: number) => (
                                                                                        <option key={idx} value={role.value}>{role.label}</option>
                                                                                    ))}



                                                                                </select>

                                                                            </td>
                                                                            <td style={{ minWidth: '40px', maxWidth: '40px', overflow: 'inherit' }}>Level {index + 1}</td>
                                                                            <td style={{ overflow: 'inherit' }} title={row?.approvers.map((approver: any) => approver.label).join(", ") || "Enter Approver Name"}>

                                                                                <Select
                                                                                    options={rows1}
                                                                                    isMulti
                                                                                    value={row.approvers}
                                                                                    // onKeyDown={(e: any) => handleKeyDown(e, 'Approvers', index)}
                                                                                    name="Approvers"
                                                                                    className={`newse HierarchyClsErr ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                    onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                                                                    placeholder="Enter Approver Name"
                                                                                    isDisabled={InputDisabled}
                                                                                    title={row.approvers.map((approver: any) => approver.label).join(", ") || "Enter Approver Name"} // Added title tooltip
                                                                                />



                                                                            </td>
                                                                            <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px', }}>
                                                                                {/* <label htmlFor="approvalType">Approval Type: </label> */}
                                                                                <select id="approvalType" value={row.approvalType} onChange={(e) => handleChange(e, row.level)} className={`newse HierarchyClsErr form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`} disabled={InputDisabled} title={row.approvalType === "One" ? "Anyone" : row.approvalType === "All" ? "Everyone" : "Select Approval Criteria"} >
                                                                                    <option value="">Select </option>
                                                                                    <option value="One">Anyone</option>
                                                                                    <option value="All">Everyone</option>
                                                                                </select>
                                                                            </td>
                                                                            <td style={{ minWidth: '40px', maxWidth: '40px', overflow: 'inherit' }}>

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


                                                        {/* {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") || (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Save As Draft</button>
                                                        }

                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") || (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Submit</button>
                                                        }

                                                       
                                                        {((modeValue === "" || modeValue === "edit" || modeValue === "view") || (editID !== null && editID.ApprovalType !== "Approval")) &&
                                                            <button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</button>
                                                        } */}
                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") || (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <div className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Save As Draft</div>
                                                        }

                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") || (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <div className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Submit</div>
                                                        }


                                                        {((modeValue === "" || modeValue === "edit" || modeValue === "view") || (editID !== null && editID.ApprovalType !== "Approval")) &&
                                                            <div className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</div>
                                                        }

                                                    </div>
                                                </div>


                                                {/* ////////////Approval card */}


                                                {/* </div> */}

                                                {/* /////////// */}

                                                <Modal show={showModal} onHide={() => setShowModal(false)} size={Showfile ? "xl" : "lg"} className='filemodal'>
                                                    <Modal.Header closeButton>
                                                        <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-1'>Attachment Details</h4>
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for IMS Audit Plan
                                                            </p>

                                                        </Modal.Title>


                                                    </Modal.Header>
                                                    <Modal.Body className="" id="style-5">

                                                        <>

                                                            {Showfile ?

                                                                <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                                                :
                                                                <table className="mtbalenew">
                                                                    <thead style={{ background: '#eef6f7' }}>
                                                                        <tr>
                                                                            <th style={{ minWidth: '30px', maxWidth: '30px' }}>S.No.</th>
                                                                            <th>File Name</th>
                                                                            {/* {editForm && <th>File Link</th>} */}
                                                                            <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Upload date</th>
                                                                            {/* {!InputDisabled && <th className='text-center'>Action</th>} */}
                                                                            <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Action</th>
                                                                        </tr>
                                                                    </thead>
                                                                    <tbody>
                                                                        {FilesArr.length > 0 && (
                                                                            FilesArr.map((row: any, index: number) => {

                                                                                const date = new Date();
                                                                                // const components = [
                                                                                //     date.getDate().toString().padStart(2, '0'),
                                                                                //     (date.getMonth() + 1).toString().padStart(2, '0'),
                                                                                //     date.getFullYear().toString(),
                                                                                //     date.getHours().toString().padStart(2, '0'),
                                                                                //     date.getMinutes().toString().padStart(2, '0'),
                                                                                //     date.getSeconds().toString().padStart(2, '0'),
                                                                                //     date.getMilliseconds().toString().padStart(3, '0')
                                                                                // ];
                                                                                const fileExtension = row.name ? row.name.split('.').pop() : "";
                                                                                const fileNameWithoutExtension = row.name ? row.name.split('.').slice(0, -1).join('.') : "";
                                                                                // const NewFileName = `${formData.memoFileName}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
                                                                                const NewFileName = `${formData.memoFileName}_${fileNameWithoutExtension}.${fileExtension}`;


                                                                                return (<tr>
                                                                                    <td style={{ minWidth: '30px', maxWidth: '30px' }} className='text-center'>{index + 1}</td>
                                                                                    <td title={row.name ? NewFileName : row.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1')}>
                                                                                        {row.name ? NewFileName : row.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1')}
                                                                                    </td>
                                                                                    {/* <td title={row.name || (row.FileLeafRef.includes('_') ? row.FileLeafRef.split('_')[2] : row.FileLeafRef)}>{row.name || (row.FileLeafRef.includes('_') ? row.FileLeafRef.split('_')[2] : row.FileLeafRef)}</td> */}
                                                                                    {/* <td title={row.name || (row.FileLeafRef)?.split('_')[2]}>{row.name || (row.FileLeafRef)?.split('_')[2]}</td> */}
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
                                                                                    <td style={{ minWidth: '50px', maxWidth: '50px' }} title={row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                        day: "2-digit",
                                                                                        month: "short",
                                                                                        year: "numeric"
                                                                                    }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                        day: "2-digit",
                                                                                        month: "short",
                                                                                        year: "numeric"
                                                                                    }).replace(/ /g, "/")}>

                                                                                        {row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                            day: "2-digit",
                                                                                            month: "short",
                                                                                            year: "numeric"
                                                                                        }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                            day: "2-digit",
                                                                                            month: "short",
                                                                                            year: "numeric"
                                                                                        }).replace(/ /g, "/")}</td>

                                                                                    <td style={{ minWidth: '50px', maxWidth: '50px' }}>
                                                                                        {row.Id && (
                                                                                            <>

                                                                                                <span title='preview file'
                                                                                                    onClick={() => OpenFile(row, "Open")}
                                                                                                    style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                                                >
                                                                                                    <FontAwesomeIcon icon={faEye} />
                                                                                                </span>
                                                                                                <span title='download file'
                                                                                                    onClick={() => OpenFile(row, "Download")}
                                                                                                    style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                                                >
                                                                                                    <FontAwesomeIcon icon={faDownload} />
                                                                                                </span>
                                                                                            </>
                                                                                        )}

                                                                                        {!InputDisabled && <img src={require("../assets/del.png")} style={{ cursor: "pointer" }} onClick={() => handleDelete(index)} />}
                                                                                    </td>


                                                                                </tr>
                                                                                )
                                                                            })
                                                                        )}
                                                                    </tbody>
                                                                </table>
                                                            }</>
                                                        {/* </>
                                                            )
                                                        } */}

                                                    </Modal.Body>

                                                </Modal>


                                                <Modal show={ShowModalTemplateDoc} onHide={() => setShowModalTemplateDoc(false)} size={Showfile ? "xl" : "lg"} className='newmobmodal'>

                                                    <Modal.Body className="" id="style-5">
                                                        <>
                                                            {Showfile &&

                                                                <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />

                                                            }
                                                        </>
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
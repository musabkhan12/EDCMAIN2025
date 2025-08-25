import * as React from 'react';

// import type { IDocumentCancellationProcessProps } from './IDocumentCancellationProcessProps';
import type { IDocumentCancellationProcessProps } from './IDocumentCancellationProcessProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import Provider from '../../../GlobalContext/provider';
import Provider from '../../../../GlobalContext/provider';
// import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import VerticalSideBar from '../../../verticalSideBar/components/VerticalSideBar';
// import HorizontalNavbar from '../../horizontalNavBar/components/HorizontalNavBar';
import HorizontalNavbar from '../../../horizontalNavBar/components/HorizontalNavBar';
// import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';
// import { getSP } from '../loc/pnpjsConfig';
import { getSP } from '../../../dmsMusaib/loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
// import UserContext from "../../../GlobalContext/context";
import UserContext from "../../../../GlobalContext/context";
import "bootstrap/dist/css/bootstrap.min.css";
// import "../../../CustomCss/mainCustom.scss";
import "../../../../CustomCss/mainCustom.scss";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
// import "../../verticalSideBar/components/VerticalSidebar.scss";
import "../../../verticalSideBar/components/VerticalSidebar.scss";
// import "../components/documentCancellation.scss";
import "./documentCancellation.scss";
// import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import { allowstringonly, getCurrentUser } from '../../../../APISearvice/CustomService';
// import { addAllProcessItem, addApprovalItem, addItem, addItem2, getAllDocumentCode, getAllProcessData, getApprovalByID, getApprovalByID2, getDataRoles, getDocumentLinkByID, getFormNameID, getItemByID, getItemByID2, getListNameID, getRequesterID, getRequestTypeID, UpdateAllProcessItem, updateApprovalItem, updateItem, updateItem2 } from '../../../APISearvice/DocumentCancellation';
import { addAllProcessItem, addApprovalItem, addItem, addItem2, CheckIfAlreadyactionTaken, getAllDocumentCode, getAllProcessData, getApprovalByID, getApprovalByID2, getDataRoles, getdigitalsignaturerequestbyID, getDocumentLinkByID, getDocumentLinkByIDSigned, getDraftApprovalByID, getFormNameID, getGeneratedTemplateDoc, getGeneratedTemplateDoc2, getItemByID, getItemByID2, getLatestChangeRequestTemplateType, getListNameID, getRequesterID, getRequestTypeID, getUserDepartment, UpdateAllProcessItem, updateApprovalItem, updateDigitalsign, updateItem, updateItem2 } from '../../../../APISearvice/DocumentCancellation';
import Select from "react-select";
import Swal from 'sweetalert2';
// import { FormSubmissionMode } from '../../../Shared/Interfaces';
import { FormSubmissionMode } from '../../../../Shared/Interfaces';
// import { decryptId } from '../../../APISearvice/CryptoService';
import { decryptId } from '../../../../APISearvice/CryptoService';
// import { getUrlParameterValue } from '../../../Shared/Helper';
import { getUrlParameterValue } from '../../../../Shared/Helper';
// import { WorkflowAction } from '../../../CustomJSComponents/WorkflowAction/WorkflowAction';
import { WorkflowAction } from '../../../../CustomJSComponents/WorkflowAction/WorkflowAction';
import { WorkflowAuditHistory } from '../../../../CustomJSComponents/WorkflowAuditHistory/WorkflowAuditHistory';
// import { WorkflowAuditHistory } from '../../ChangerequestComponent/WorkflowAuditHistory/WorkflowAuditHistory';
// import { CONTENTTYPE_DocumentCancel, LIST_TITLE_DocCancel, Tenant_URL } from '../../../Shared/Constants';
import { CONTENTTYPE_ChangeReqTemp, CONTENTTYPE_DocumentCancel, LIST_TITLE_DocCancel, SITE_URL, Tenant_URL } from '../../../../Shared/Constants';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Modal } from 'react-bootstrap';
import { faDownload, faEye } from '@fortawesome/free-solid-svg-icons';
import CustomBreadcrumb from '../../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
import FileViewer from '../../components/fileviewer';
let myloader = '../../'
interface ForwardTo {
    id: number;
    role: number;
    level: number;
    approvers: any[]; // Or a more specific type like `string[]` or `SPUser[]`
    approvalType: string;
    Responsibility: string;
    IsSignatureRequired: boolean;
    IsDisabled: boolean;
}

const ErrorCls = "border-on-error";

const DocumentCancellationProcessContext = ({ props }: any) => {
    const sp: SPFI = getSP();
    const elementRef = React.useRef<HTMLDivElement>(null);
    const siteUrl = props.siteUrl;
    const { useHide }: any = React.useContext(UserContext);
    const [InputDisabled, setInputDisabled] = React.useState(false);
    const selectedTextDiv = document.getElementById('selectedText');

    selectedTextDiv.style.display = 'none';

    const [Loading, setLoading] = React.useState(false);
    const [TemplateDoc, setTemplateDoc] = React.useState<any>([]);
    const [FormLoading, setFormLoading] = React.useState(false);
    const [FormItemId, setFormItemId] = React.useState(null);
    const [editID, setEditID] = React.useState(null);
    const [DigitalsignID, setDigitalsignID] = React.useState(null);
    const [MainEditItem, setMainEditItem] = React.useState(null);
    const [rows, setRows] = React.useState<any>([]);
    const [UserRoles, setUserRoles] = React.useState<any>([]);
    const [rows1, setRows1] = React.useState<any>([]);
    const [currentUser, setCurrentUser] = React.useState(null);
    const [selectedOption, setSelectedOption] = React.useState(null);
    const [selectedPeople, setSelectedPeople] = React.useState(null);
    const [selectedRole, setSelectedRole] = React.useState(null);
    const [ValidDraft, setValidDraft] = React.useState(true);
    const [ValidSubmit, setValidSubmit] = React.useState(true);
    const [ValidCancelReason, setValidCancelReason] = React.useState(true);
    const [ValidForwardTo, setValidForwardTo] = React.useState(true);
    const [RequesterRoleId, setRequesterRoleId] = React.useState(null);
    const [RequestTypeId, setRequestTypeId] = React.useState(null);
    const [FormNameId, setFormNameId] = React.useState(null);
    const [ListNameId, setListNameId] = React.useState(null);
    const [editForm, setEditForm] = React.useState(false);
    const [modeValue, setmode] = React.useState("");
    const [currentUserDept, setcurrentUserDept] = React.useState("");
    const [DraftApprovalItem, setDraftApprovalItem] = React.useState(null);
    const [DocumentLink, setDocumentLink] = React.useState(null);
    const [cancellReason, setcancellReason] = React.useState([{ id: 0, description: "", reason: "" }]);
    const [cancellReasonEdit, setcancellReasonEdit] = React.useState([]);
    const [showModal, setShowModal] = React.useState(false);

    const [DocCancellByEdit, setDocCancellByEdit] = React.useState([]);

    const [editItemID, setEditItemID] = React.useState(null);

    const [hidedigisign, sethidedigisign] = React.useState(false);


    const [Showfile, setShowfile] = React.useState(false);
    const [redirecturl, setredirecturl] = React.useState(null);
    const [ShowModalTemplateDoc, setShowModalTemplateDoc] = React.useState(false);
    const [sharewithusers, setSharewithusers] = React.useState([]);
    const [formData, setFormData] = React.useState({
        RequesterNameId: 0,
        RequesterName: "",
        RequesterDesignation: "",
        Department: null,
        DepartmentName: "",
        RequestDate: "",
        RequestDateNew: "",
        IssueDate: "",
        LocationId: 0,
        CustodianId: 0,
        AmendmentTypeId: 0,
        ClassificationId: 0,
        Location: "",
        Custodian: "",
        AmendmentType: "",
        Classification: "",
        DocumentType: "",
        RequestTypeId: 0,

        SerialNumber: "",
        IssueNumber: "",
        RevisionNumber: "",
        RevisionDate: "",
        DocumentCode: "",
        ReferenceNumber: "",
        ChangeRequestTypeId: [],
        SubmiitedDate: "",
        SubmitStatus: "",
        Status: "",
        DocumentName: "",
        IsRework: false,
        DigitalSignStatus: false,
        ChangeRequestID: 0,
        AttachmentId: [],
        AttachmentDigitalSignatureId: [],
        AttachmentJson: "",
        TemplateTypeId: 0,
        TemplateTypeValue: "",


        CDocumentCode: "",
        CIssueNumber: null,
        CRevisionNumber: null,
        CRevisionDate: "",
        CIssueDate: "",


        filename: "",
        Remark: "",
        PreparedById: [],
        PreparedBy: [],
        DocumentCancelledBy: []



    });


    const cancelModalAction = (refresh?: boolean,) => {
        debugger
        setredirecturl(window.location.href);
        //setShowfileNew(false);
        setShowModalTemplateDoc(false);
        setShowfile(false);
    }


    // Handle change event
    const handleChange = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        event.preventDefault();
        const updatedArr = forwardToArr.map(row =>
            row.level === lvl ? { ...row, approvalType: event.target.value } : row
        );
        //   setApprovalType(event.target.value);
        setForwardToArr(updatedArr);
    };
    const handleChangeResp = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        event.preventDefault();
        const updatedArr = forwardToArr.map(row =>
            // row.level === lvl ? { ...row, Responsibility: event.target.value} : row

            row.level === lvl ? { ...row, Responsibility: event.target.value, IsSignatureRequired: event.target.value === "" ? false : true } : row
        );
        //   setApprovalType(event.target.value);
        setForwardToArr(updatedArr);
    };


    const [forwardToArr, setForwardToArr] = React.useState<ForwardTo[]>([
        // { id: 0, role: 0, level: 1, approvers: [], approvalType: "One" } // Default row
    ]);

    const Breadcrumb = [
        {
            MainComponent: "My Request",
            MainComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx`,
        },
        {
            ChildComponent: "Document Cancellation",
            ChildComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx#/Document%20Cancellation`,
        },
    ];
    const [forwardToArrEdit, setForwardToArrEdit] = React.useState<ForwardTo[]>([]);
    const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);
    // const [remark, setRemark] = React.useState("");
    const getCustomStyles = (valueCount: any) => ({
        control: (provided: any) => ({
            ...provided,
            minHeight: valueCount <= 1 ? '47px' : 'auto',
            //fontSize: '14px',
            flexWrap: 'wrap',
        }),
        valueContainer: (provided: any) => ({
            ...provided,
            padding: '2px 6px',
        }),
        multiValue: (provided: any) => ({
            ...provided,
            //fontSize: '12px',
            margin: '2px',
        }),
        indicatorsContainer: (provided: any) => ({
            ...provided,
            height: '47px',
        }),
    });

    // Function to handle People Picker selection
    const onPeoplePickerChange = (items: any[]) => {
        setSelectedUsers(items);
    };
    const handleKeyDown = (e: React.KeyboardEvent<HTMLSelectElement>) => {
        if (e.key === 'Enter') {
            e.preventDefault(); // 🛑 Prevents page reload
        }
    };
    const handleKeyDowntext = (e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.key === 'Enter') {
            e.preventDefault();
        }
    };

    const ApiCallFunc = async () => {


        var ReqId = await getRequestTypeID(sp);
        setRequestTypeId(ReqId);
        // setRequestTypeId(await getRequestTypeID(sp));

        const path1 = window.location.href;

        if (path1.includes("/view/") || path1.includes("/approve/")) {
            setFormLoading(true); ////
            setInputDisabled(true);
            // if (path1.includes("/approve/")) {
            //     setForwardToArr((prev) => [
            //         ...prev,
            //         { id: 0, role: 0, level: prev.length + 1, approvers: [], approvalType: "One", Responsibility: "Signer", IsSignatureRequired: true }
            //     ]);
            // }
        }
        else {
            // setForwardToArr((prev) => [
            //     ...prev,
            //     { id: 0, role: 0, level: prev.length + 1, approvers: [], approvalType: "One", Responsibility: "Signer", IsSignatureRequired: true }
            // ]);
            setInputDisabled(false);
        }
        if (path1.includes("/edit/")) {
            setFormLoading(true); ////
        }


        const Currusers: any = await getCurrentUser(sp, siteUrl);
        setCurrentUser(await getCurrentUser(sp, siteUrl));
        const userProfile = await sp.profiles.myProperties();
        setcurrentUserDept(userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "")
        const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";

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

        var fetchUserDept = await getUserDepartment(sp, UserDept)

        setFormData(prevData => ({
            ...prevData,
            RequesterNameId: Currusers?.Id || "",
            RequesterDesignation: userProfile?.Title || "",
            RequesterName: userProfile?.DisplayName || "",
            RequestDate: new Date().toLocaleDateString("en-CA"),
            // RequestedDate: new Date().toISOString().split("T")[0] // Format as YYYY-MM-DD
            // RequestDateNew: new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }).replace(",", "")
            RequestDateNew: new Date().toLocaleDateString("en-GB", {
                day: "2-digit",
                month: "short",
                year: "numeric"
            }).replace(/ /g, "/"),

            DepartmentName: fetchUserDept

        }));
        setFormLoading(true);
        var DocCodeArr = await getAllDocumentCode(sp, UserDept);
        // var DocCodeArr = await getAllDocumentCode(sp);
        const options = DocCodeArr.map((item: any) => ({
            value: item.DocumentCode,
            label: item.DocumentCode,
            IssueNumber: item.IssueNumber,
            ReferenceNumber: item.ReferenceNumber,
            RevisionNumber: item.RevisionNumber,
            ChangeRequestID: item.ID,
            IssueDate: item.IssueDate,
            LocationId: item.LocationId,
            CustodianId: item.CustodianId,
            SerialNumber: item.SerialNumber,
            RevisionDate: item.RevisionDate,
            AmendmentTypeId: item.AmendmentTypeId,
            RequestTypeId: ReqId,
            ClassificationId: item.ClassificationId,
            ChangeRequestTypeId: item.ChangeRequestTypeId,
            // ChangeRequestTypeId: RequestTypeId,
            SubmiitedDate: item.SubmiitedDate,
            SubmitStatus: item.SubmitStatus,
            DocumentTypeId: item.DocumentTypeId,
            Department: item.DepartmentId || null,
            DepartmentName: item.Department?.Department || "",
            AttachmentId: item.AttachmentId || [],
            AttachmentDigitalSignatureId: item.AttachmentDigitalSignatureId || [],
            AttachmentJson: item.AttachmentJson,

            Location: item.Location.Location,
            Custodian: item.Custodian.Custodian,
            AmendmentType: item.AmendmentType.AmendmentType,
            Classification: item.Classification.Classification,
            DocumentType: item.DocumentType.DocumentType,
            TemplateTypeId: item.TemplateTypeId,
            TemplateTypeValue: item.TemplateType?.TemplateTypeName || "",

            filename: item?.FileName || "",
            PreparedById: item?.PreparedById || [],

            PreparedBy: item?.PreparedBy || [],


        }));




        setRows(options);


        let formitemid;
        let formMode = "";
        //#region getdataByID
        if (sessionStorage.getItem("DocumentCancelId") != undefined) {
            const iD = sessionStorage.getItem("DocumentCancelId")
            let iDs = decryptId(iD)
            formitemid = Number(iDs);
            setFormItemId(Number(iDs))
        }
        else {
            //   let formitemidparam = getUrlParameterValue('contentid');
            //   if (formitemidparam) {
            //     formitemid = Number(formitemidparam);
            //     setFormItemId(Number(formitemid));
            //   }

            const path = window.location.href;
            const segments = path.split('/').filter(Boolean); // Remove empty elements

            // Check if "edit" or "view" exists in the URL
            const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");

            var EditID = null;
            if (paramIndex !== -1 && segments[paramIndex + 1]) {
                setmode(segments[paramIndex])
                formMode = segments[paramIndex];
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

                    EditID = await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_DocumentCancel);
                    setEditID(EditID);
                    var ProcessItemId: any = await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_DocumentCancel);
                    setInputDisabled(await getApprovalByID2(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_DocumentCancel));
                }
                // else{


                // }
            }

            // setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_DocumentCancel))

            if (formitemid) {
                let ProcessApprovalItem = await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_DocumentCancel);

                setDraftApprovalItem(ProcessApprovalItem);
                setFormData(prevData => ({
                    ...prevData,

                    Remark: ProcessApprovalItem[0]?.Remark || "",
                }));
            }




        }
        // formitemid =20;
        if (formitemid) {

            setEditItemID(Number(formitemid));

            const setBannerById = await getItemByID(sp, Number(formitemid))

            const newItem1 = await getdigitalsignaturerequestbyID(LIST_TITLE_DocCancel, sp, Number(formitemid));
            // console.log("newItem1newItem1", newItem1);
            if (newItem1.length > 0) {
                setDigitalsignID(newItem1[0].ID)
            }

            if (setBannerById.length > 0) {
                debugger
                setEditForm(true);
                setMainEditItem(setBannerById[0]);
                // setCategoryData(await getCategory(sp, Number(setBannerById[0]?.TypeMaster))) // Category
                if (setBannerById[0].Status == "Approved") {
                    setDocumentLink(await getGeneratedTemplateDoc2(sp, Number(formitemid), setBannerById[0].ChangeRequestIDId))
                }
                else {
                    if (setBannerById[0].AttachmentId.length) {
                        setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].AttachmentId))
                    }
                    if (setBannerById[0].AttachmentDigitalSignatureId.length) {
                        setDocumentLink(await getDocumentLinkByIDSigned(sp, setBannerById[0].AttachmentDigitalSignatureId))
                    }
                }

                // if (ProcessItemId && ProcessItemId.Level === 0 && ProcessItemId.CurrentUserRole === "OES" && ProcessItemId.IsInitiator == "No") {
                const ApprowData: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_DocumentCancel, setBannerById[0].DocumentCode)

                if (ApprowData.length > 0) {

                    const EditApprowData = ApprowData.map((item: any) => ({
                        id: item.ID,
                        role: item.ApproverRole?.Id || 0, // Assuming role comes from ApproverRole
                        level: item.Level || 1, // Default to 1 if missing
                        approvalType: item.LevelType,
                        Responsibility: item.Responsibility || "",
                        IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
                        approvers: item.Approvers?.map((approver: any) => ({
                            value: approver.Id,
                            label: approver.Title,
                        })) || [],
                        IsDisabled: item.Responsibility == "Preparer" ? true : false

                    }));



                    setForwardToArr(EditApprowData);
                    setForwardToArrEdit(EditApprowData);

                    //     setDocCancellByEdit(
                    //     (ApprowData || [])
                    //         .filter((user: any) => user.Id !== setBannerById[0].RequesterNameId)
                    //         .map((user: any) => ({
                    //             value: user.Id,
                    //             label: user.Title,
                    //             UserName: user.Title,
                    //             UserEmail: user.EMail
                    //         }))
                    // )



                }
                if (setBannerById[0].Status == "Pending" && formMode == "approve" && ApprowData.length == 0) {
                    handleAddRow();
                }
                // else if (
                //     setBannerById[0].DocumentCancellationBy?.length > 0 && EditID != null && setBannerById[0].Status == "Pending" && (EditID as any)?.CurrentUserRole == "OES"
                // ) {
                //     for (let i = 0; i < setBannerById[0].DocumentCancellationBy.length; i++) {
                //         let user = setBannerById[0].DocumentCancellationBy[i];
                //         if (user.Id != Currusers?.Id) {
                //             setForwardToArr((prev) => [
                //                 ...prev,
                //                 {
                //                     id: 0,
                //                     role: 0,
                //                     level: prev.length + 1,
                //                     approvers: [
                //                         {
                //                             value: user.Id,
                //                             label: user.Title,
                //                             UserName: user.Title,
                //                             UserEmail: user.EMail
                //                         }
                //                     ],
                //                     approvalType: "One",
                //                     Responsibility: "Preparer",
                //                     IsSignatureRequired: true,
                //                     IsDisabled: true
                //                 }
                //             ]);

                //             setForwardToArrEdit((prev) => [
                //                 ...prev,
                //                 {
                //                     id: 0,
                //                     role: 0,
                //                     level: prev.length + 1,
                //                     approvers: [
                //                         {
                //                             value: user.Id,
                //                             label: user.Title,
                //                             UserName: user.Title,
                //                             UserEmail: user.EMail
                //                         }
                //                     ],
                //                     approvalType: "One",
                //                     Responsibility: "Preparer",
                //                     IsSignatureRequired: true,
                //                     IsDisabled: true
                //                 }
                //             ]);

                //         }

                //     }
                // }

                // setDocCancellByEdit(
                //     (setBannerById[0].DocumentCancellationBy || [])
                //         .filter((user: any) => user.Id !== setBannerById[0].RequesterNameId)
                //         .map((user: any) => ({
                //             value: user.Id,
                //             label: user.Title,
                //             UserName: user.Title,
                //             UserEmail: user.EMail
                //         }))
                // )

                let arr = {

                    RequesterName: setBannerById[0].Title,
                    RequesterNameId: setBannerById[0].RequesterNameId,
                    RequesterDesignation: setBannerById[0].RequesterDesignation,
                    // Department: setBannerById[0].Department.Department||"",
                    Department: setBannerById[0].DepartmentId || null,
                    DepartmentName: setBannerById[0].Department?.Department || "",
                    RequestDate: setBannerById[0].RequestDate,
                    RequestDateNew: new Date(setBannerById[0].RequestDate).toLocaleDateString("en-GB", {
                        day: "2-digit",
                        month: "short",
                        year: "numeric"
                    }).replace(/ /g, "/"),

                    // RequestDateNew: new Date(setBannerById[0].RequestDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }).replace(",", ""),

                    IssueDate: setBannerById[0].IssueDate,
                    LocationId: setBannerById[0].LocationId,
                    CustodianId: setBannerById[0].CustodianId,
                    SerialNumber: setBannerById[0].SerialNumber,
                    IssueNumber: setBannerById[0].IssueNumber,
                    RevisionNumber: setBannerById[0].RevisionNumber,
                    RevisionDate: setBannerById[0].RevisionDate,
                    DocumentCode: setBannerById[0].DocumentCode,
                    ReferenceNumber: setBannerById[0].ReferenceNumber,
                    AmendmentTypeId: setBannerById[0].AmendmentTypeId,
                    ClassificationId: setBannerById[0].ClassificationId,
                    ChangeRequestTypeId: setBannerById[0].ChangeRequestTypeId,
                    RequestTypeId: ReqId,
                    SubmiitedDate: setBannerById[0].SubmiitedDate,
                    SubmitStatus: setBannerById[0].SubmitStatus,
                    value: setBannerById[0].DocumentCode,
                    label: setBannerById[0].DocumentCode,
                    //
                    Status: setBannerById[0].Status,
                    // DocumentName: "",
                    // IsRework: false,
                    // DigitalSignStatus: false,
                    ChangeRequestIDId: setBannerById[0].ChangeRequestID,
                    DocumentTypeId: setBannerById[0].DocumentTypeId,
                    AttachmentId: setBannerById[0].AttachmentId || [],
                    AttachmentDigitalSignatureId: setBannerById[0].AttachmentDigitalSignatureId || [],
                    AttachmentJson: setBannerById[0].AttachmentJson,

                    Location: setBannerById[0].Location.Location,
                    Custodian: setBannerById[0].Custodian.Custodian,
                    AmendmentType: setBannerById[0].AmendmentType.AmendmentType,
                    Classification: setBannerById[0].Classification.Classification,
                    DocumentType: setBannerById[0].DocumentType.DocumentType,
                    TemplateTypeId: setBannerById[0].TemplateTypeId,
                    TemplateTypeValue: setBannerById[0].TemplateType?.TemplateTypeName || "",

                    filename: setBannerById[0]?.FileName || "",

                    PreparedById: setBannerById[0].PreparedById || [],
                    PreparedBy: setBannerById[0].PreparedBy || [],
                    DocumentCancelledBy: setBannerById[0].DocumentCancellationBy?.map((user: any) => ({
                        value: user.Id,
                        label: user.Title,
                        UserName: user.Title,
                        UserEmail: user.EMail
                    })) || []

                }


                setFormData(prevData => ({
                    ...prevData,
                    RequesterName: setBannerById[0].Title,
                    RequesterNameId: setBannerById[0].RequesterNameId,
                    RequesterDesignation: setBannerById[0].RequesterDesignation,
                    RequestDate: setBannerById[0].RequestDate,
                    RequestDateNew: new Date(setBannerById[0].RequestDate).toLocaleDateString("en-GB", {
                        day: "2-digit",
                        month: "short",
                        year: "numeric"
                    }).replace(/ /g, "/"),
                    IssueNumber: setBannerById[0].IssueNumber,
                    ReferenceNumber: setBannerById[0].ReferenceNumber,
                    RevisionNumber: setBannerById[0].RevisionNumber,
                    ChangeRequestID: setBannerById[0].ChangeRequestIDId,
                    IssueDate: setBannerById[0].IssueDate,
                    LocationId: setBannerById[0].LocationId,
                    CustodianId: setBannerById[0].CustodianId,
                    SerialNumber: setBannerById[0].SerialNumber,
                    RevisionDate: setBannerById[0].RevisionDate,
                    AmendmentTypeId: setBannerById[0].AmendmentTypeId,
                    ClassificationId: setBannerById[0].ClassificationId,
                    ChangeRequestTypeId: setBannerById[0].ChangeRequestTypeId,
                    RequestTypeId: ReqId,
                    SubmiitedDate: setBannerById[0].SubmiitedDate,
                    SubmitStatus: setBannerById[0].SubmitStatus,
                    DocumentCode: setBannerById[0].DocumentCode,
                    DocumentTypeId: setBannerById[0].DocumentTypeId,
                    DepartmentName: setBannerById[0].Department?.Department || "",
                    Department: setBannerById[0].DepartmentId || null,
                    AttachmentId: setBannerById[0].AttachmentId,
                    AttachmentDigitalSignatureId: setBannerById[0].AttachmentDigitalSignatureId || [],
                    AttachmentJson: setBannerById[0].AttachmentJson,

                    Location: setBannerById[0].Location.Location,
                    Custodian: setBannerById[0].Custodian.Custodian,
                    AmendmentType: setBannerById[0].AmendmentType.AmendmentType,
                    Classification: setBannerById[0].Classification.Classification,
                    DocumentType: setBannerById[0].DocumentType.DocumentType,
                    TemplateTypeId: setBannerById[0].TemplateTypeId,
                    TemplateTypeValue: setBannerById[0].TemplateType?.TemplateTypeName || "",

                    filename: setBannerById[0]?.FileName || "",
                    Status: setBannerById[0].Status,
                    PreparedById: setBannerById[0].PreparedById || [],
                    PreparedBy: setBannerById[0].PreparedBy || [],
                    DocumentCancelledBy: setBannerById[0].DocumentCancellationBy?.map((user: any) => ({
                        value: user.Id,
                        label: user.Title,
                        UserName: user.Title,
                        UserEmail: user.EMail
                    })) || []

                    // Format as YYYY-MM-DD
                }));

                // setFormData(arr2);


                setSelectedOption(arr);
                setSharewithusers(setBannerById[0]?.PreparedBy?.map((obj: any) => {
                    // const filteredUser = Selectedoptions.find((dept: any) => dept.value === obj.ID);

                    return {
                        value: obj.ID,
                        label: obj.Title,
                        UserName: obj.Title,
                        UserEmail: obj.Email


                    };

                }) || []);

                const rowData: any[] = await getItemByID2(sp, Number(setBannerById[0].ID)) //baseUrl
                if (rowData.length > 0) {
                    const initialRows = rowData.map((item: any) => ({
                        id: item.Id,
                        description: item.ChangeDescription,
                        reason: item.ReasonforChange,
                    }));
                    setcancellReason(initialRows);
                    setcancellReasonEdit(initialRows);

                }

                setTemplateDoc(await getGeneratedTemplateDoc(sp, Number(formitemid)));

            }

            // setRequesterRoleId(await getRequesterID(sp))
            // setFormNameId(await getFormNameID(sp, CONTENTTYPE_DocumentCancel))
            // setListNameId(await getListNameID(sp, LIST_TITLE_DocCancel))



        }
        setRequesterRoleId(await getRequesterID(sp))
        setFormNameId(await getFormNameID(sp, CONTENTTYPE_DocumentCancel))
        setListNameId(await getListNameID(sp, LIST_TITLE_DocCancel))
        setFormLoading(false);


        let ChangeRequestTemplateType = await getLatestChangeRequestTemplateType(sp, CONTENTTYPE_ChangeReqTemp);

        if (ChangeRequestTemplateType.length > 0) {
            const template = ChangeRequestTemplateType[0];
            setFormData((prevFormData) => ({
                ...prevFormData,

                CDocumentCode: template.DocumentCode || "",
                CRevisionNumber: template.RevisionNumber,
                CIssueNumber: template.IssueNumber,
                CRevisionDate: template.RevisionDate ? new Date(template.RevisionDate).toLocaleDateString("en-CA") : null,
                CIssueDate: template.IssueDate ? new Date(template.IssueDate).toLocaleDateString("en-CA") : null,

            }));
        }

        //}
        //#endregion


    };

    const onSelectDocCancelledBy = (selectedOptions: any) => {
        setFormData((prev) => ({
            ...prev,
            DocumentCancelledBy: selectedOptions || []
        }));

    };

    const onSelect = async (selectedList: any) => {
        console.log(selectedList, "selectedList");

        if (selectedList) {
            setFormData(prevData => ({
                ...prevData,
                IssueNumber: selectedList.IssueNumber,
                ReferenceNumber: selectedList.ReferenceNumber,
                RevisionNumber: selectedList.RevisionNumber,
                ChangeRequestID: selectedList.ChangeRequestID,
                IssueDate: selectedList.IssueDate,
                LocationId: selectedList.LocationId,
                CustodianId: selectedList.CustodianId,
                SerialNumber: selectedList.SerialNumber,
                RevisionDate: selectedList.RevisionDate,
                AmendmentTypeId: selectedList.AmendmentTypeId,
                ClassificationId: selectedList.ClassificationId,
                ChangeRequestTypeId: selectedList.ChangeRequestTypeId,
                RequestTypeId: RequestTypeId,
                SubmiitedDate: selectedList.SubmiitedDate,
                SubmitStatus: selectedList.SubmitStatus,
                DocumentCode: selectedList.value,
                DocumentTypeId: selectedList.DocumentTypeId,
                Department: selectedList.Department,
                DepartmentName: selectedList.DepartmentName || "",
                AttachmentId: selectedList.AttachmentId || [],
                AttachmentDigitalSignatureId: selectedList.AttachmentDigitalSignatureId || [],
                AttachmentJson: selectedList.AttachmentJson,
                Location: selectedList.Location,
                Custodian: selectedList.Custodian,
                AmendmentType: selectedList.AmendmentType,
                Classification: selectedList.Classification,
                DocumentType: selectedList.DocumentType,
                TemplateTypeId: selectedList.TemplateTypeId,
                TemplateTypeValue: selectedList.TemplateTypeValue || "",

                filename: selectedList?.filename || "",
                PreparedById: selectedList?.PreparedById || [],
                PreparedBy: selectedList?.PreparedBy || [],
            }));
        } else {
            setFormData(prevData => ({
                ...prevData,
                IssueNumber: "",
                ReferenceNumber: "",
                RevisionNumber: "",
                ChangeRequestID: null,
                IssueDate: "",
                LocationId: null,
                CustodianId: null,
                SerialNumber: "",
                RevisionDate: "",
                AmendmentTypeId: null,
                ClassificationId: null,
                ChangeRequestTypeId: [],
                RequestTypeId: 0,
                SubmiitedDate: "",
                SubmitStatus: "",
                DocumentCode: "",
                DocumentTypeId: null,
                Department: null,
                DepartmentName: "",
                AttachmentId: [],
                AttachmentDigitalSignatureId: [],
                AttachmentJson: "",
                Location: "",
                Custodian: "",
                AmendmentType: "",
                Classification: "",
                DocumentType: "",
                TemplateTypeId: null,
                TemplateTypeValue: "",
                filename: "",
                PreparedById: [],
                PreparedBy: [],
            }));
        }
        setSharewithusers(selectedList?.PreparedBy?.map((obj: any) => {
            // const filteredUser = Selectedoptions.find((dept: any) => dept.value === obj.ID);

            return {
                value: obj.ID,
                label: obj.Title,
                UserName: obj.Title,
                UserEmail: obj.Email


            };

        }) || []);
        setSelectedOption(selectedList);
        if (selectedList?.AttachmentId.length) {
            setDocumentLink(await getDocumentLinkByID(sp, selectedList.AttachmentId))
        }
        else if (selectedList?.AttachmentDigitalSignatureId.length) {
            setDocumentLink(await getDocumentLinkByIDSigned(sp, selectedList.AttachmentDigitalSignatureId))
        }
        else {
            setDocumentLink(null);
        }  // Set the selected users
    };

    const onSelectApprovers = (selectedOptions: any, lvl: number) => {
        setForwardToArr((prev) =>
            prev.map((row) =>
                row.level === lvl ? { ...row, approvers: selectedOptions || [] } : row
            )
        );
    };


    // const onSelectRole = (selectedList: any) => {
    //     // console.log(selectedList , "selectedList");
    //     setSelectedRole(selectedList);  // Set the selected users
    // };

    const onSelectRole = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        event.preventDefault();
        const updatedArr = forwardToArr.map(row =>
            row.level === lvl ? { ...row, role: Number(event.target.value) } : row
        );
        setForwardToArr(updatedArr);
        setSelectedRole(forwardToArr.map(r => r.role).filter(role => role))
    };


    // const handleAddRow = () => {
    //     const newRow: ForwardTo = { id: Date.now(), role: 0, level: "", approvers: [] };
    //     setForwardToArr([...forwardToArr, newRow]);
    // };
    const handleAddRow = () => {
        setForwardToArr((prev) => [
            ...prev,
            { id: 0, role: 0, level: prev.length + 1, approvers: [], approvalType: "One", Responsibility: "Signer", IsSignatureRequired: true, IsDisabled: false }
        ]);
    };

    // const handleDeleteRow = (index: number) => {
    //     // const updatedRows = forwardToArr.filter((_, i) => i !== index);
    //     // setForwardToArr(updatedRows);
    //     forwardToArr.splice(index, 1);
    //     setForwardToArr(forwardToArr);
    // };
    // const handleDeleteRow = (index: number) => {
    //     const updatedRows = forwardToArr.filter((_, i) => i !== index);
    //     setForwardToArr([...updatedRows]); // Ensure a new array reference
    // };
    const handleDeleteRow = (index: number) => {
        const updatedRows = forwardToArr
            .filter((_, i) => i !== index) // Remove selected row
            .map((row, newIndex) => ({ ...row, level: newIndex + 1 })); // Reassign levels

        setForwardToArr([...updatedRows]); // Ensure a new array reference
    };



    React.useEffect(() => {

        ApiCallFunc();




        // formData.title = currentUser.Title;

    }, [useHide]);

    const handleCancel = () => {
        // window.location.reload();
        window.history.back();
        // window.location.reload();
        setTimeout(() => {
            location.reload();
        }, 100);

    }

    //#region deleteLocalFile
    const deleteLocalFile = (index: number, filArray: any[]) => {
        // Create a new array without mutating the existing state
        const updatedFiles = [...cancellReason];
        updatedFiles.splice(index, 1);

        // Update the state with the new array
        setcancellReason(updatedFiles);
    };

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
            link.setAttribute("download", (obj.FileLeafRef?.replace(/_\d+(\.\w+)$/, '$1') || obj.FileLeafRef)); // Suggests a filename for download
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        }



    }

    //#endregion


    const addCancelReason = () => {
        setcancellReason([...cancellReason, { id: 0, description: "", reason: "" }]);
    };
    const validateForm = async (fmode: FormSubmissionMode) => {
        const { RequesterName, RequesterDesignation, RequestDate, DocumentCode, IssueNumber, RevisionNumber, ReferenceNumber, DocumentCancelledBy } = formData;
        // const { description } = richTextValues;
        let valid = true;
        let valid1 = true;
        // let validateOverview:boolean = false;
        // let validatetitlelength = false;
        // let validateTitle = false;
        setValidDraft(true);
        setValidSubmit(true);
        setValidCancelReason(true);
        let errormsg = "";

        Array.from(document.getElementsByClassName("border-on-error")).forEach((element: Element) => {
            element.classList.remove("border-on-error");
        });

        if (fmode == FormSubmissionMode.SUBMIT) {
            if (!RequesterName) {
                //Swal.fire('Error', 'Title is required!', 'error');
                valid = false;
            }
            // else if (!RequesterDesignation) {
            //     //Swal.fire('Error', 'Type is required!', 'error');
            //     valid = false;
            // }
            //   else if (!RequestDate) {
            //     //Swal.fire('Error', 'Category is required!', 'error');
            //     valid = false;
            //   }
            if (selectedOption == null || !selectedOption.value) {
                //Swal.fire('Error', 'Entity is required!', 'error');
                document.getElementById("DocumentCode")?.classList.add("border-on-error");
                valid = false;
            }
            if (DocumentCancelledBy.length == 0) {
                // Swal.fire('Error', 'Cancelled By is required!', 'error');
                document.getElementById("DocCancelledBy")?.classList.add("border-on-error");
                valid = false;

            }
            if (cancellReason.length > 0 && cancellReason.every((row: any) => row.description !== null && row.reason !== null && row.description.trim() !== "" && row.reason.trim() !== "") == false) {
                // const isValid = cancellReason.every((row:any) => row.description.trim() !== "" && row.reason.trim() !== "");
                valid1 = false;

                Array.from(document.getElementsByClassName("recommendClsErr")).forEach((element: Element) => {
                    if (element.tagName === "DIV" && (element.textContent?.trim() === "Select" || element.textContent?.trim() === "")) {
                        element.classList.add("border-on-error");
                    }
                    else if ((element.tagName === "INPUT" || element.tagName === "TEXTAREA") && (element as HTMLInputElement).value.trim() === "") {
                        element.classList.add("border-on-error");
                    }


                });
            }
            if (cancellReason.length == 0) {
                // const isValid = rows.every((row:any) => row.description.trim() !== "" && row.reason.trim() !== "");
                valid1 = false;
            }
            else if (IssueNumber === "") {
                //Swal.fire('Error', 'Entity is required!', 'error');
                valid = false;
            }
            else if (RevisionNumber === "") {
                //Swal.fire('Error', 'Entity is required!', 'error');
                valid = false;
            }
            else if (!ReferenceNumber) {
                //Swal.fire('Error', 'Entity is required!', 'error');
                valid = false;
            }


            setValidSubmit(valid);
            // setValidCancelReason(valid1);

        }
        else {
            if (!RequesterName) {
                //Swal.fire('Error', 'Title is required!', 'error');
                valid = false;
            }
            // else if (!RequesterDesignation) {
            //     //Swal.fire('Error', 'Type is required!', 'error');
            //     valid = false;
            // }
            else if (selectedOption == null || !selectedOption.value) {
                //Swal.fire('Error', 'Entity is required!', 'error');
                valid = false;
            }
            // else if (cancellReason.length > 0 && cancellReason.every((row: any) => row.description.trim() !== "" && row.reason.trim() !== "") == false) {
            //     // const isValid = cancellReason.every((row:any) => row.description.trim() !== "" && row.reason.trim() !== "");
            //     valid1 = false;
            // }

            setValidDraft(valid);
            // setValidCancelReason(valid1);

        }

        // console.log("validateTitle", validateTitle,"errormsg", errormsg,"valid,", valid, ImagepostArr.length);
        if (!valid && fmode == FormSubmissionMode.SUBMIT)
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');
        // else if (!valid && fmode == FormSubmissionMode.SUBMIT && rows.length >0){
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');
        // }
        else if (!valid1 && fmode == FormSubmissionMode.SUBMIT)
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields in description section.');
        else if (!valid && fmode == FormSubmissionMode.DRAFT) {
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');
        }
        else if (!valid1 && fmode == FormSubmissionMode.DRAFT) {
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields in description section..');
        }
        if (valid == false || valid1 == false) {
            return false
        }
        else {
            return true
        }
        // return valid;
    };
    //#region  Submit Form
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

                        const DocumentCancelledByIds: any[] = [];
                        formData.DocumentCancelledBy.forEach((user: any) => {
                            if (user?.value) {
                                DocumentCancelledByIds.push(user.value);
                            }
                        });

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {
                            // Title: formData.RequesterName,
                            // RequesterNameId: formData.RequesterNameId,
                            // RequesterDesignation: formData.RequesterDesignation,
                            DepartmentId: formData.Department,
                            RequestDate: formData.RequestDate,
                            IssueDate: formData.IssueDate,
                            LocationId: selectedOption.LocationId,
                            CustodianId: selectedOption.CustodianId,
                            SerialNumber: selectedOption.SerialNumber,
                            IssueNumber: selectedOption.IssueNumber,
                            RevisionNumber: selectedOption.RevisionNumber,
                            RevisionDate: selectedOption.RevisionDate,
                            DocumentCode: selectedOption.value,
                            ReferenceNumber: selectedOption.ReferenceNumber,
                            AmendmentTypeId: selectedOption.AmendmentTypeId,
                            RequestTypeId: RequestTypeId,
                            ClassificationId: selectedOption.ClassificationId,
                            ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                            SubmiitedDate: selectedOption.SubmiitedDate,
                            // SubmitStatus: selectedOption.SubmitStatus,
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            // DocumentName: "",
                            // IsRework: false,
                            // DigitalSignStatus: false,
                            ChangeRequestIDId: formData.ChangeRequestID,
                            DocumentTypeId: selectedOption.DocumentTypeId,
                            OESSubmitStatus: "No",
                            InitiatorSubmitStatus: "Yes",
                            CurrentUserRole: "OES",
                            AttachmentId: selectedOption.AttachmentId,
                            AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                            AttachmentJson: selectedOption.AttachmentJson,
                            TemplateTypeId: selectedOption.TemplateTypeId,

                            FileName: formData.filename,
                            PreparedById: formData.PreparedById,
                            DocumentCancellationById: DocumentCancelledByIds,


                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;


                        for (const row of cancellReason) {

                            const postPayload2 = {
                                ChangeRequestDCIDId: editItemID, // Assuming "Title" column exists
                                ChangeDescription: row.description,
                                ReasonforChange: row.reason,
                            }

                            if (!row.id) {

                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }

                            }
                            else if (row.id > 0) {
                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;
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



                        // /*********** */

                        // Find items that are in cancellReasonEdit but NOT in cancellReason
                        const toDelete = cancellReasonEdit.filter(
                            (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("ChangeRequestReasonDocumentCancellationList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        // ///////************* */

                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////Handleforward to users with document cancelled by changes new


                        const itemsToDelete = await sp.web.lists.getByTitle("AllProcessApprovalLevelList")
                            .items
                            .filter(`MainListID eq '${editItemID}' and ProcessName eq '${CONTENTTYPE_DocumentCancel}'`)
                            .select("Id", "Approvers/Id")
                            .expand("Approvers")
                            .getAll();


                        for (const delItem of itemsToDelete) {
                            await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(delItem.Id).delete();
                        }



                        const filteredCancelledBy = formData.DocumentCancelledBy.filter(
                            (user: any) => user.value !== formData.RequesterNameId
                        );

                        const filteredApprovers = forwardToArrEdit.filter(
                            (user: any) => user.Responsibility !== "Preparer"
                        );
                        let countLevel = 0;
                        for (let i = 0; i < filteredCancelledBy?.length; i++) {
                            let user = filteredCancelledBy[i];
                            if (user && user.value != formData.RequesterNameId) {
                                let level = i + 1;
                                countLevel = level;
                                let RoleID = 0;
                                const identifyexistingitem = forwardToArrEdit.filter(item => {
                                    return (
                                        item.Responsibility === "Preparer" &&
                                        item.approvers.length === 1 &&
                                        //DocumentCancelledByIds.includes(user.value)
                                        item.approvers[0].value === user.value
                                    );
                                });
                                if (identifyexistingitem.length > 0) {
                                    // If an item with Responsibility "Preparer" and one approver exists, skip this iteration
                                    RoleID = identifyexistingitem[0].role;
                                }

                                let approversIds: any[] = [];
                                // item.approvers.forEach((user: any) => {
                                if (user) {
                                    approversIds.push(user.value);
                                }


                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.filename,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: RoleID,
                                    Level: Number(level),
                                    ApproversId: approversIds,
                                    LevelType: "One",
                                    // LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: 1,
                                    // ContentTitle:,
                                    MainListID: String(editItemID),
                                    RequestId: formData.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Document Cancellation",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No",
                                    // RedirectionLink:,
                                    Responsibility: "Preparer",
                                    IsSignatureRequired: "Yes",

                                    CRIssueNumber: formData.IssueNumber,
                                    CRRevisionNumber: formData.RevisionNumber,



                                }

                                if (user && user.value != formData.RequesterNameId) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }


                            }


                        }

                        for (const item of filteredApprovers) {
                            countLevel++;
                            const approversIds: any[] = [];
                            item.approvers.forEach((user: any) => {
                                if (user?.value) {
                                    approversIds.push(user.value);
                                }
                            });

                            let arr2 = {
                                Title: currentUser.Title,
                                // ContentTitle: selectedOption.ReferenceNumber,
                                ContentTitle: formData.filename,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role,
                                Level: Number(countLevel),
                                ApproversId: approversIds,
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "No",
                                Maxlevel: item.approvers?.length,
                                // ContentTitle:,
                                MainListID: String(editItemID),
                                RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Document Cancellation",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                IsApprovalGenerated: "No",
                                // RedirectionLink:,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",

                                CRIssueNumber: formData.IssueNumber,
                                CRRevisionNumber: formData.RevisionNumber,


                            }
                            // if (item.id) {
                            //     const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                            //     const postId2 = postResult2?.data?.ID;

                            // }
                            // else {

                            const postResult2 = await addAllProcessItem(arr2, sp);
                            const postId2 = postResult2?.data?.ID;

                            // }


                        }




                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////

                        // if (boolval == true) {
                        setLoading(false);
                        // Swal.fire('Submitted successfully.', '', 'success');
                        // sessionStorage.removeItem("DocumentCancelId")
                        // setTimeout(() => {

                        //     window.history.back();
                        //     // window.location.reload();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        //     // let url = window.location.href;
                        //     // let baseUrl = url.split("#")[0];
                        // }, 500);
                        // // }
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
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

                        const DocumentCancelledByIds: any[] = [];
                        formData.DocumentCancelledBy.forEach((user: any) => {
                            if (user?.value) {
                                DocumentCancelledByIds.push(user.value);
                            }
                        });

                        const postPayload = {
                            Title: formData.RequesterName,
                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            DepartmentId: formData.Department,
                            RequestDate: formData.RequestDate,
                            IssueDate: formData.IssueDate ? formData.IssueDate : null,
                            LocationId: selectedOption.LocationId,
                            CustodianId: selectedOption.CustodianId,
                            SerialNumber: Number(selectedOption.SerialNumber),
                            IssueNumber: Number(selectedOption.IssueNumber),
                            RevisionNumber: selectedOption.RevisionNumber,
                            RevisionDate: selectedOption.RevisionDate ? selectedOption.RevisionDate : null,
                            DocumentCode: selectedOption.value,
                            ReferenceNumber: selectedOption.ReferenceNumber,
                            AmendmentTypeId: selectedOption.AmendmentTypeId,
                            RequestTypeId: RequestTypeId,
                            ClassificationId: selectedOption.ClassificationId,
                            ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                            SubmiitedDate: selectedOption.SubmiitedDate,
                            // SubmitStatus: selectedOption.SubmitStatus,
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            //   DocumentName: "",
                            //   IsRework: false,
                            //   DigitalSignStatus: false,
                            ChangeRequestIDId: formData.ChangeRequestID,
                            DocumentTypeId: selectedOption.DocumentTypeId,
                            OESSubmitStatus: "No",
                            InitiatorSubmitStatus: "Yes",
                            CurrentUserRole: "OES",
                            TemplateTypeId: selectedOption.TemplateTypeId,
                            AttachmentId: selectedOption.AttachmentId,
                            AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                            AttachmentJson: selectedOption.AttachmentJson ? selectedOption.AttachmentJson : "",





                            CDocumentCode: formData.CDocumentCode,
                            CRevisionNumber: formData.CRevisionNumber,
                            CIssueNumber: formData.CIssueNumber,
                            CRevisionDate: formData.CRevisionDate,
                            CIssueDate: formData.CIssueDate,

                            FileName: formData.filename,
                            PreparedById: formData.PreparedById,
                            DocumentCancellationById: DocumentCancelledByIds,


                        };
                        console.log(postPayload);

                        const postResult = await addItem(postPayload, sp);
                        const postId = postResult?.data?.ID;
                        // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }

                        for (const row of cancellReason) {

                            const postPayload2 = {
                                ChangeRequestDCIDId: postId, // Assuming "Title" column exists
                                ChangeDescription: row.description,
                                ReasonforChange: row.reason,
                            }

                            const postResult2 = await addItem2(postPayload2, sp);
                            const postId2 = postResult2?.data?.ID;
                            // debugger
                            if (!postId2) {
                                console.error("Post creation failed.");
                                return;
                            }
                        }

                        // 888888888888888888888888**********************
                        const filteredCancelledBy = formData.DocumentCancelledBy.filter(
                            (user: any) => user.value !== formData.RequesterNameId
                        );

                        for (let i = 0; i < filteredCancelledBy?.length; i++) {
                            let user = filteredCancelledBy[i];
                            if (user && user.value != formData.RequesterNameId) {
                                let level = i + 1;

                                //  for (const item of forwardToArr) {

                                let approversIds: any[] = [];
                                // item.approvers.forEach((user: any) => {
                                if (user) {
                                    approversIds.push(user.value);
                                }
                                // });

                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.filename,
                                    MainListNameId: ListNameId,
                                    // ApproverRoleId: item.role,
                                    Level: Number(level),
                                    ApproversId: approversIds,
                                    LevelType: "One",
                                    // LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: 1,
                                    // ContentTitle:,
                                    MainListID: String(postId),
                                    RequestId: formData.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Document Cancellation",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No",
                                    // RedirectionLink:,
                                    Responsibility: "Preparer",
                                    IsSignatureRequired: "Yes",

                                    CRIssueNumber: formData.IssueNumber,

                                    CRRevisionNumber: formData.RevisionNumber,

                                }
                                // if (item.id) {
                                //     const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                //     const postId2 = postResult2?.data?.ID;

                                // }
                                // else {
                                if (user && user.value != formData.RequesterNameId) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                                // }

                                // }

                            }


                        }



                        // 888888888888888888888888**********************




                        // let boolval;

                        // if (boolval == true) {
                        setLoading(false);
                        // Swal.fire('Submitted successfully.', '', 'success');
                        // // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 500);
                        // // }
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

                        let arr = {};
                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        const DocumentCancelledByIds: any[] = [];
                        formData.DocumentCancelledBy.forEach((user: any) => {
                            if (user?.value) {
                                DocumentCancelledByIds.push(user.value);
                            }
                        });

                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {
                            arr = {
                                // Title: formData.RequesterName,
                                // RequesterNameId: formData.RequesterNameId,
                                // RequesterDesignation: formData.RequesterDesignation,
                                DepartmentId: formData.Department,
                                RequestDate: formData.RequestDate,
                                IssueDate: formData.IssueDate,
                                LocationId: selectedOption.LocationId,
                                CustodianId: selectedOption.CustodianId,
                                SerialNumber: selectedOption.SerialNumber,
                                IssueNumber: selectedOption.IssueNumber,
                                RevisionNumber: selectedOption.RevisionNumber,
                                RevisionDate: selectedOption.RevisionDate,
                                DocumentCode: selectedOption.value,
                                ReferenceNumber: selectedOption.ReferenceNumber,
                                AmendmentTypeId: selectedOption.AmendmentTypeId,
                                RequestTypeId: RequestTypeId,
                                ClassificationId: selectedOption.ClassificationId,
                                ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                                SubmiitedDate: selectedOption.SubmiitedDate,
                                // SubmitStatus: selectedOption.SubmitStatus,
                                SubmitStatus: "No",
                                Status: "Save as draft",
                                // DocumentName: "",
                                // IsRework: false,
                                // DigitalSignStatus: false,
                                ChangeRequestIDId: formData.ChangeRequestID,
                                DocumentTypeId: selectedOption.DocumentTypeId,
                                OESSubmitStatus: "No",
                                InitiatorSubmitStatus: "No",
                                CurrentUserRole: "OES",
                                AttachmentId: selectedOption.AttachmentId,
                                AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                                AttachmentJson: selectedOption.AttachmentJson,
                                TemplateTypeId: selectedOption.TemplateTypeId,
                                PreparedById: formData.PreparedById,
                                DocumentCancellationById: DocumentCancelledByIds,


                            }
                        }
                        else {
                            arr = {
                                // Title: formData.RequesterName,
                                // RequesterNameId: formData.RequesterNameId,
                                // RequesterDesignation: formData.RequesterDesignation,
                                DepartmentId: formData.Department,
                                RequestDate: formData.RequestDate,
                                IssueDate: formData.IssueDate,
                                LocationId: selectedOption.LocationId,
                                CustodianId: selectedOption.CustodianId,
                                SerialNumber: selectedOption.SerialNumber,
                                IssueNumber: selectedOption.IssueNumber,
                                RevisionNumber: selectedOption.RevisionNumber,
                                RevisionDate: selectedOption.RevisionDate,
                                DocumentCode: selectedOption.value,
                                ReferenceNumber: selectedOption.ReferenceNumber,
                                AmendmentTypeId: selectedOption.AmendmentTypeId,
                                RequestTypeId: RequestTypeId,
                                ClassificationId: selectedOption.ClassificationId,
                                ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                                SubmiitedDate: selectedOption.SubmiitedDate,
                                // SubmitStatus: selectedOption.SubmitStatus,
                                SubmitStatus: "No",
                                Status: "Save as draft",
                                // DocumentName: "",
                                // IsRework: false,
                                // DigitalSignStatus: false,
                                ChangeRequestIDId: formData.ChangeRequestID,
                                DocumentTypeId: selectedOption.DocumentTypeId,
                                OESSubmitStatus: "No",
                                InitiatorSubmitStatus: "No",
                                CurrentUserRole: "OES",
                                AttachmentId: selectedOption.AttachmentId,
                                AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                                AttachmentJson: selectedOption.AttachmentJson,
                                TemplateTypeId: selectedOption.TemplateTypeId,

                                FileName: formData.filename,
                                PreparedById: formData.PreparedById,
                                DocumentCancellationById: DocumentCancelledByIds,


                            }
                        }

                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;


                        for (const row of cancellReason) {

                            const postPayload2 = {
                                ChangeRequestDCIDId: editItemID, // Assuming "Title" column exists
                                ChangeDescription: row.description,
                                ReasonforChange: row.reason,
                            }

                            if (!row.id) {

                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }

                            }
                            else if (row.id > 0) {
                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;
                            }

                        }



                        // if(DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0){

                        //     let arr2 = {
                        //         ActionTakenById: currentUser.Id,
                        //         ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                        //         // ActionTakenRoleId: formData.RequesterDesignation,
                        //         Status: "Save as draft",
                        //         // Remark: remark,

                        //     }
                        //     const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                        //     const postId = postResult?.data?.ID;

                        // }
                        // /*********** */

                        // Find items that are in cancellReasonEdit but NOT in cancellReason
                        const toDelete = cancellReasonEdit.filter(
                            (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("ChangeRequestReasonDocumentCancellationList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }


                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////Handleforward to users with document cancelled by changes new


                        const itemsToDelete = await sp.web.lists.getByTitle("AllProcessApprovalLevelList")
                            .items
                            .filter(`MainListID eq '${editItemID}' and ProcessName eq '${CONTENTTYPE_DocumentCancel}'`)
                            .select("Id", "Approvers/Id")
                            .expand("Approvers")
                            .getAll();


                        for (const delItem of itemsToDelete) {
                            await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(delItem.Id).delete();
                        }



                        const filteredCancelledBy = formData.DocumentCancelledBy.filter(
                            (user: any) => user.value !== formData.RequesterNameId
                        );

                        const filteredApprovers = forwardToArrEdit.filter(
                            (user: any) => user.Responsibility !== "Preparer"
                        );
                        let countLevel = 0;
                        for (let i = 0; i < filteredCancelledBy?.length; i++) {
                            let user = filteredCancelledBy[i];
                            if (user && user.value != formData.RequesterNameId) {
                                let level = i + 1;
                                countLevel = level;
                                let RoleID = 0;
                                const identifyexistingitem = forwardToArrEdit.filter(item => {
                                    return (
                                        item.Responsibility === "Preparer" &&
                                        item.approvers.length === 1 &&
                                        //DocumentCancelledByIds.includes(user.value)
                                        item.approvers[0].value === user.value
                                    );
                                });
                                if (identifyexistingitem.length > 0) {
                                    // If an item with Responsibility "Preparer" and one approver exists, skip this iteration
                                    RoleID = identifyexistingitem[0].role;
                                }

                                let approversIds: any[] = [];
                                // item.approvers.forEach((user: any) => {
                                if (user) {
                                    approversIds.push(user.value);
                                }


                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.filename,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: RoleID,
                                    Level: Number(level),
                                    ApproversId: approversIds,
                                    LevelType: "One",
                                    // LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: 1,
                                    // ContentTitle:,
                                    MainListID: String(editItemID),
                                    RequestId: formData.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Document Cancellation",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No",
                                    // RedirectionLink:,
                                    Responsibility: "Preparer",
                                    IsSignatureRequired: "Yes",

                                    CRIssueNumber: formData.IssueNumber,
                                    CRRevisionNumber: formData.RevisionNumber,


                                }

                                if (user && user.value != formData.RequesterNameId) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }


                            }


                        }

                        for (const item of filteredApprovers) {
                            countLevel++;
                            const approversIds: any[] = [];
                            item.approvers.forEach((user: any) => {
                                if (user?.value) {
                                    approversIds.push(user.value);
                                }
                            });

                            let arr2 = {
                                Title: currentUser.Title,
                                // ContentTitle: selectedOption.ReferenceNumber,
                                ContentTitle: formData.filename,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role,
                                Level: Number(countLevel),
                                ApproversId: approversIds,
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "No",
                                Maxlevel: item.approvers?.length,
                                // ContentTitle:,
                                MainListID: String(editItemID),
                                RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Document Cancellation",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                IsApprovalGenerated: "No",
                                // RedirectionLink:,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",

                                CRIssueNumber: formData.IssueNumber,
                                CRRevisionNumber: formData.RevisionNumber,


                            }
                            // if (item.id) {
                            //     const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                            //     const postId2 = postResult2?.data?.ID;

                            // }
                            // else {

                            const postResult2 = await addAllProcessItem(arr2, sp);
                            const postId2 = postResult2?.data?.ID;

                            // }


                        }




                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////

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
                        // // }
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

                        const DocumentCancelledByIds: any[] = [];
                        formData.DocumentCancelledBy.forEach((user: any) => {
                            if (user?.value) {
                                DocumentCancelledByIds.push(user.value);
                            }
                        });

                        let valid1 = true;
                        const postPayload = {
                            Title: formData.RequesterName,
                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            DepartmentId: formData.Department,
                            RequestDate: formData.RequestDate,
                            IssueDate: formData.IssueDate,
                            LocationId: selectedOption.LocationId,
                            CustodianId: selectedOption.CustodianId,
                            SerialNumber: Number(selectedOption.SerialNumber),
                            IssueNumber: Number(selectedOption.IssueNumber),
                            RevisionNumber: selectedOption.RevisionNumber,
                            RevisionDate: selectedOption.RevisionDate,
                            DocumentCode: selectedOption.value,
                            ReferenceNumber: selectedOption.ReferenceNumber,
                            AmendmentTypeId: selectedOption.AmendmentTypeId,
                            RequestTypeId: RequestTypeId,
                            ClassificationId: selectedOption.ClassificationId,
                            ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                            SubmiitedDate: selectedOption.SubmiitedDate,
                            SubmitStatus: "No",
                            Status: "Save as draft",
                            //   DocumentName: "",
                            //   IsRework: false,
                            //   DigitalSignStatus: false,
                            ChangeRequestIDId: formData.ChangeRequestID,
                            DocumentTypeId: selectedOption.DocumentTypeId,
                            OESSubmitStatus: "No",
                            InitiatorSubmitStatus: "No",
                            CurrentUserRole: "OES",
                            AttachmentId: selectedOption.AttachmentId,
                            AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                            AttachmentJson: selectedOption.AttachmentJson,
                            TemplateTypeId: selectedOption.TemplateTypeId,




                            CDocumentCode: formData.CDocumentCode,
                            CRevisionNumber: formData.CRevisionNumber,
                            CIssueNumber: formData.CIssueNumber,
                            CRevisionDate: formData.CRevisionDate,
                            CIssueDate: formData.CIssueDate,

                            FileName: formData.filename,
                            PreparedById: formData.PreparedById,
                            DocumentCancellationById: DocumentCancelledByIds,


                        };
                        console.log(postPayload);

                        const postResult = await addItem(postPayload, sp);
                        const postId = postResult?.data?.ID;
                        // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }

                        if (cancellReason.length > 0 && cancellReason.every((row: any) => row.description.trim() == "" && row.reason.trim() == "") == false) {
                            // const isValid = cancellReason.every((row:any) => row.description.trim() !== "" && row.reason.trim() !== "");
                            // valid1 = false;
                            for (const row of cancellReason) {

                                const postPayload2 = {
                                    ChangeRequestDCIDId: postId, // Assuming "Title" column exists
                                    ChangeDescription: row.description,
                                    ReasonforChange: row.reason,
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

                        // 888888888888888888888888**********************
                        const filteredCancelledBy = formData.DocumentCancelledBy.filter(
                            (user: any) => user.value !== formData.RequesterNameId
                        );

                        for (let i = 0; i < filteredCancelledBy?.length; i++) {
                            let user = filteredCancelledBy[i];
                            if (user && user.value != formData.RequesterNameId) {
                                let level = i + 1;

                                //  for (const item of forwardToArr) {

                                let approversIds: any[] = [];
                                // item.approvers.forEach((user: any) => {
                                if (user) {
                                    approversIds.push(user.value);
                                }
                                // });

                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.filename,
                                    MainListNameId: ListNameId,
                                    // ApproverRoleId: item.role,
                                    Level: Number(level),
                                    ApproversId: approversIds,
                                    LevelType: "One",
                                    // LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: 1,
                                    // ContentTitle:,
                                    MainListID: String(postId),
                                    RequestId: formData.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Document Cancellation",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No",
                                    // RedirectionLink:,
                                    Responsibility: "Preparer",
                                    IsSignatureRequired: "Yes",

                                    CRIssueNumber: formData.IssueNumber,
                                    CRRevisionNumber: formData.RevisionNumber,


                                }
                                // if (item.id) {
                                //     const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                //     const postId2 = postResult2?.data?.ID;

                                // }
                                // else {
                                if (user && user.value != formData.RequesterNameId) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                                // }

                                // }

                            }


                        }



                        // 888888888888888888888888**********************




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


    // public render(): React.ReactElement<IDocumentCancellationProcessProps> {  
    //     const peoplePickerContext: IPeoplePickerContext = {
    //       absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
    //       msGraphClientFactory: this.props.context.msGraphClientFactory,
    //       spHttpClient: this.props.context.spHttpClient
    //   };

    const ForwardApproval = async (status: string) => {
        let valid = true;
        let actionMessage = "";
        let successMessage = "";
        // setValidSubmit(true);
        setValidForwardTo(true);
        Array.from(document.getElementsByClassName("border-on-error")).forEach((element: Element) => {
            element.classList.remove("border-on-error");
        });

        if ((status === 'Rework' || status === 'Rejected') && formData.Remark === "") {
            //   setValidRemark(false);
            document.getElementById("Remark-textarea")?.classList.add("border-on-error");
            Swal.fire('Please fill the mandatory fields', '', 'warning');
            return;
        }

        switch (status) {
            case "Forward":
                actionMessage = "Do you want to forward this request?";
                successMessage = "Forwarded successfully.";
                break;
            case "Rejected":
                actionMessage = "Do you want to reject this request?";
                successMessage = 'Rejected successfully.'
                break;
            case "Rework":
                actionMessage = "Do you want to rework this request?";
                successMessage = 'Sent for rework.'
                break;
        }

        const IsactionTaken = await CheckIfAlreadyactionTaken(sp, editID.Id);
        if (!IsactionTaken) {
            Swal.fire("Action has already been taken on this record.");
            return;
        }


        if (status == "Forward") {
            if (forwardToArr.length === 0) {
                // alert("At least one row is required.");
                valid = false;
            }

            const isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                row.approvalType.trim() !== "" && row.Responsibility.trim() !== "");

            if (!isValid) {
                // alert("Each row must have a role selected and at least one approver.");
                valid = false;
            }

            if (!valid) {
                Swal.fire('Please fill all the mandatory fields.');
                Array.from(document.getElementsByClassName("ForwardClsErr")).forEach((element: Element) => {
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



                // setValidForwardTo(false);
                return;
            }

            if (valid) {
                Swal.fire({
                    title: actionMessage,
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

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))

                        let arr = {
                            Remark: formData.Remark,
                            ActionTakenById: currentUser.Id,
                            // ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                            ActionTakenOn: new Date().toISOString(),
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            Status: "Approved",
                            // Remark: remark,


                        }
                        const postResult = await updateApprovalItem(arr, sp, editID.Id);
                        const postId = postResult?.data?.ID;

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
                                ContentTitle: formData.filename,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role,
                                Level: Number(item.level),
                                ApproversId: approversIds,
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "Yes",
                                Maxlevel: item.approvers?.length,
                                // ContentTitle:,
                                MainListID: String(editItemID),
                                RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Document Cancellation",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                IsApprovalGenerated: "No",
                                // RedirectionLink:,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",

                                CRIssueNumber: formData.IssueNumber,
                                CRRevisionNumber: formData.RevisionNumber,


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

                        let arr2 = {
                            // ActionTakenById: currentUser.Id,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            // Status: status ==="Rework"?"Rework":"Rejected",
                            // IsRework: status ==="Rework"?"Yes":"No",
                            OESSubmitStatus: "Yes",
                            InitiatorSubmitStatus: "Yes",
                            SubmitStatus: "Yes",


                        }
                        const postResult3 = await updateItem(arr2, sp, editItemID);
                        const postId3 = postResult3?.data?.ID;

                        // //////// if approver row deleted

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




                        setLoading(false);
                        // Swal.fire(successMessage, '', 'success');
                        // sessionStorage.removeItem("DocumentCancelId")
                        // setTimeout(() => {
                        //     // window.location.href = `${siteUrl}/SitePages/MyTasks.aspx`;
                        //     window.history.back();
                        // }, 1000);
                        // // }
                        Swal.fire(successMessage, '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
                    }
                })
            }



        }
        else {
            let isValid = true;

            if (forwardToArr.length) {
                isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                    row.approvalType.trim() !== "" && row.Responsibility.trim() !== "");

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


            if (valid) {
                Swal.fire({
                    title: actionMessage,
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

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {
                            Remark: formData.Remark,
                            ActionTakenById: currentUser.Id,
                            // ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                            ActionTakenOn: new Date().toISOString(),
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            Status: status,
                            // Remark: remark,


                        }
                        const postResult = await updateApprovalItem(arr, sp, editID.Id);
                        const postId = postResult?.data?.ID;

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
                                ContentTitle: formData.filename,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role,
                                Level: Number(item.level),
                                ApproversId: approversIds,
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "Yes",
                                Maxlevel: item.approvers?.length,
                                // ContentTitle:,
                                MainListID: String(editItemID),
                                RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Document Cancellation",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                IsApprovalGenerated: "No",
                                // RedirectionLink:,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",

                                CRIssueNumber: formData.IssueNumber,
                                CRRevisionNumber: formData.RevisionNumber,


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

                        // }



                        let arr2 = {
                            // ActionTakenById: currentUser.Id,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            Status: status === "Rework" ? "Rework" : "Rejected",
                            IsRework: status === "Rework" ? "Yes" : "No",
                            OESSubmitStatus: "No",
                            InitiatorSubmitStatus: "No",
                            SubmitStatus: "No",


                        }
                        const postResult2 = await updateItem(arr2, sp, editItemID);
                        const postId2 = postResult2?.data?.ID;

                        // //////// if approver row deleted

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

                        setLoading(false);
                        // Swal.fire(successMessage, '', 'success');
                        // sessionStorage.removeItem("DocumentCancelId")
                        // setTimeout(() => {
                        //     // window.location.href = `${siteUrl}/SitePages/MyTasks.aspx`;
                        //     window.history.back();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        // }, 1000);
                        // // }
                        Swal.fire(successMessage, '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
                    }

                })
            }
        }


    }


    const ForwardInitiatorApproval = async (status: string) => {
        // let valid = true;
        let actionMessage = "";
        let successMessage = "";
        Array.from(document.getElementsByClassName("border-on-error")).forEach((element: Element) => {
            element.classList.remove("border-on-error");
        });
        switch (status) {
            case "Approved":
                actionMessage = "Do you want to submit this request?";
                successMessage = "Submitted successfully.";
                break;
            case "Save as draft":
                actionMessage = "Do you want to save this request?";
                successMessage = "Saved successfully.";
                break;
            // case "Rework":
            //     actionMessage = "Do you want to send this request for rework?";
            //     successMessage = "Request sent for rework successfully.";
            //     break;
        }


        if (status == "Approved") {
            if (formData.Remark === "") {
                //   setValidRemark(false);
                document.getElementById("Remark-textarea2")?.classList.add("border-on-error");
                Swal.fire('Please fill the mandatory fields', '', 'warning');
                return;
                // valid = false;
            }
            if (await validateForm(FormSubmissionMode.SUBMIT)) {

                //   if (valid) {
                Swal.fire({
                    title: actionMessage,
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

                        const DocumentCancelledByIds: any[] = [];
                        formData.DocumentCancelledBy.forEach((user: any) => {
                            if (user?.value) {
                                DocumentCancelledByIds.push(user.value);
                            }
                        });

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        // //////////////Update Process Approval List when Submitted
                        let arr = {
                            ActionTakenById: currentUser.Id,
                            // ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                            ActionTakenOn: new Date().toISOString(),
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            Status: "Approved",
                            // Remark: remark,
                            Remark: formData.Remark,

                        }
                        const postResult = await updateApprovalItem(arr, sp, editID.Id);
                        const postId = postResult?.data?.ID;

                        // //////////////Update Document cancellation List when Submitted
                        let arr3 = {
                            // Title: formData.RequesterName,
                            // RequesterNameId: formData.RequesterNameId,
                            // RequesterDesignation: formData.RequesterDesignation,
                            // Department: formData.Department,
                            RequestDate: formData.RequestDate,
                            IssueDate: formData.IssueDate,
                            LocationId: selectedOption.LocationId,
                            CustodianId: selectedOption.CustodianId,
                            SerialNumber: selectedOption.SerialNumber,
                            IssueNumber: selectedOption.IssueNumber,
                            RevisionNumber: selectedOption.RevisionNumber,
                            RevisionDate: selectedOption.RevisionDate,
                            DocumentCode: selectedOption.value,
                            ReferenceNumber: selectedOption.ReferenceNumber,
                            AmendmentTypeId: selectedOption.AmendmentTypeId,
                            RequestTypeId: RequestTypeId,
                            ClassificationId: selectedOption.ClassificationId,
                            ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                            SubmiitedDate: selectedOption.SubmiitedDate,
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            // DocumentName: "",
                            // IsRework: false,
                            // DigitalSignStatus: false,
                            ChangeRequestIDId: formData.ChangeRequestID,
                            DocumentTypeId: selectedOption.DocumentTypeId,
                            OESSubmitStatus: "No",
                            InitiatorSubmitStatus: "Yes",
                            CurrentUserRole: "OES",
                            AttachmentId: selectedOption.AttachmentId,
                            AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                            AttachmentJson: selectedOption.AttachmentJson,
                            DocumentCancellationById: DocumentCancelledByIds,


                        }
                        const postResult3 = await updateItem(arr3, sp, editItemID);
                        const postId3 = postResult?.data?.ID;


                        // //////////////Update Document cancellation Reason List when Submitted

                        for (const row of cancellReason) {

                            const postPayload2 = {
                                ChangeRequestDCIDId: editItemID, // Assuming "Title" column exists
                                ChangeDescription: row.description,
                                ReasonforChange: row.reason,
                            }

                            if (!row.id) {

                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }

                            }
                            else if (row.id > 0) {
                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;
                            }

                        }

                        const toDelete = cancellReasonEdit.filter(
                            (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("ChangeRequestReasonDocumentCancellationList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////Handleforward to users with document cancelled by changes new


                        const itemsToDelete = await sp.web.lists.getByTitle("AllProcessApprovalLevelList")
                            .items
                            .filter(`MainListID eq '${editItemID}' and ProcessName eq '${CONTENTTYPE_DocumentCancel}'`)
                            .select("Id", "Approvers/Id")
                            .expand("Approvers")
                            .getAll();


                        for (const delItem of itemsToDelete) {
                            await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(delItem.Id).delete();
                        }



                        const filteredCancelledBy = formData.DocumentCancelledBy.filter(
                            (user: any) => user.value !== formData.RequesterNameId
                        );

                        const filteredApprovers = forwardToArrEdit.filter(
                            (user: any) => user.Responsibility !== "Preparer"
                        );
                        let countLevel = 0;
                        for (let i = 0; i < filteredCancelledBy?.length; i++) {
                            let user = filteredCancelledBy[i];
                            if (user && user.value != formData.RequesterNameId) {
                                let level = i + 1;
                                countLevel = level;
                                let RoleID = 0;
                                const identifyexistingitem = forwardToArrEdit.filter(item => {
                                    return (
                                        item.Responsibility === "Preparer" &&
                                        item.approvers.length === 1 &&
                                        //DocumentCancelledByIds.includes(user.value)
                                        item.approvers[0].value === user.value
                                    );
                                });
                                if (identifyexistingitem.length > 0) {
                                    // If an item with Responsibility "Preparer" and one approver exists, skip this iteration
                                    RoleID = identifyexistingitem[0].role;
                                }

                                let approversIds: any[] = [];
                                // item.approvers.forEach((user: any) => {
                                if (user) {
                                    approversIds.push(user.value);
                                }


                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.filename,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: RoleID,
                                    Level: Number(level),
                                    ApproversId: approversIds,
                                    LevelType: "One",
                                    // LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: 1,
                                    // ContentTitle:,
                                    MainListID: String(editItemID),
                                    RequestId: formData.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Document Cancellation",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No",
                                    // RedirectionLink:,
                                    Responsibility: "Preparer",
                                    IsSignatureRequired: "Yes",


                                    CRIssueNumber: formData.IssueNumber,
                                    CRRevisionNumber: formData.RevisionNumber,


                                }

                                if (user && user.value != formData.RequesterNameId) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }


                            }


                        }

                        for (const item of filteredApprovers) {
                            countLevel++;
                            const approversIds: any[] = [];
                            item.approvers.forEach((user: any) => {
                                if (user?.value) {
                                    approversIds.push(user.value);
                                }
                            });

                            let arr2 = {
                                Title: currentUser.Title,
                                // ContentTitle: selectedOption.ReferenceNumber,
                                ContentTitle: formData.filename,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role,
                                Level: Number(countLevel),
                                ApproversId: approversIds,
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "Yes",
                                Maxlevel: item.approvers?.length,
                                // ContentTitle:,
                                MainListID: String(editItemID),
                                RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Document Cancellation",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                IsApprovalGenerated: "No",
                                // RedirectionLink:,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",

                                CRIssueNumber: formData.IssueNumber,
                                CRRevisionNumber: formData.RevisionNumber,


                            }
                            // if (item.id) {
                            //     const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                            //     const postId2 = postResult2?.data?.ID;

                            // }
                            // else {

                            const postResult2 = await addAllProcessItem(arr2, sp);
                            const postId2 = postResult2?.data?.ID;

                            // }


                        }




                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////

                        setLoading(false);
                        // Swal.fire(successMessage, '', 'success');
                        // sessionStorage.removeItem("DocumentCancelId")
                        // setTimeout(() => {
                        //     // window.location.href = `${siteUrl}/SitePages/MyTasks.aspx`;
                        //     window.history.back();

                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        // }, 1000);
                        Swal.fire(successMessage, '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = modeValue == "approve" ? `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx` : `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });

                    }

                })



            }
        }
        else if (status == "Save as draft") {

            if (await validateForm(FormSubmissionMode.DRAFT)) {
                Swal.fire({
                    title: actionMessage,
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
                        const DocumentCancelledByIds: any[] = [];
                        formData.DocumentCancelledBy.forEach((user: any) => {
                            if (user?.value) {
                                DocumentCancelledByIds.push(user.value);
                            }
                        });

                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {
                            ActionTakenById: currentUser.Id,
                            // ActionTakenOn: new Date().toLocaleDateString("en-CA"),
                            ActionTakenOn: new Date().toISOString(),
                            // ActionTakenRoleId: formData.RequesterDesignation,
                            // Status: status,
                            // Remark: remark,
                            // Status: "Pending",
                            Remark: formData.Remark,

                        }
                        const postResult = await updateApprovalItem(arr, sp, editID.Id);
                        const postId = postResult?.data?.ID;

                        // //////////////Update Document cancellation List when Submitted
                        let arr3 = {
                            // Title: formData.RequesterName,
                            // RequesterNameId: formData.RequesterNameId,
                            // RequesterDesignation: formData.RequesterDesignation,
                            // Department: formData.Department,
                            RequestDate: formData.RequestDate,
                            IssueDate: formData.IssueDate,
                            LocationId: selectedOption.LocationId,
                            CustodianId: selectedOption.CustodianId,
                            SerialNumber: selectedOption.SerialNumber,
                            IssueNumber: selectedOption.IssueNumber,
                            RevisionNumber: selectedOption.RevisionNumber,
                            RevisionDate: selectedOption.RevisionDate,
                            DocumentCode: selectedOption.value,
                            ReferenceNumber: selectedOption.ReferenceNumber,
                            AmendmentTypeId: selectedOption.AmendmentTypeId,
                            RequestTypeId: RequestTypeId,
                            ClassificationId: selectedOption.ClassificationId,
                            ChangeRequestTypeId: selectedOption.ChangeRequestTypeId ? selectedOption.ChangeRequestTypeId : [],
                            SubmiitedDate: selectedOption.SubmiitedDate,
                            SubmitStatus: "No",
                            // Status: "Save as draft",
                            // //////
                            Status: "Rework",
                            IsRework: "Yes",
                            // /////

                            ChangeRequestIDId: formData.ChangeRequestID,
                            DocumentTypeId: selectedOption.DocumentTypeId,
                            OESSubmitStatus: "No",
                            InitiatorSubmitStatus: "No",
                            CurrentUserRole: "OES",
                            AttachmentId: selectedOption.AttachmentId,
                            AttachmentDigitalSignatureId: selectedOption.AttachmentDigitalSignatureId,
                            AttachmentJson: selectedOption.AttachmentJson,

                            DocumentCancellationById: DocumentCancelledByIds,


                        }
                        const postResult3 = await updateItem(arr3, sp, editItemID);
                        const postId3 = postResult?.data?.ID;


                        // //////////////Update Document cancellation Reason List when Submitted
                        if (cancellReason.length > 0 && cancellReason.every((row: any) => row.description.trim() == "" && row.reason.trim() == "") == false) {

                            for (const row of cancellReason) {

                                const postPayload2 = {
                                    ChangeRequestDCIDId: editItemID, // Assuming "Title" column exists
                                    ChangeDescription: row.description,
                                    ReasonforChange: row.reason,
                                }

                                if (!row.id) {

                                    const postResult2 = await addItem2(postPayload2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                    // debugger
                                    if (!postId2) {
                                        console.error("Post creation failed.");
                                        return;
                                    }

                                }
                                else if (row.id > 0) {
                                    const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }
                        }



                        const toDelete = cancellReasonEdit.filter(
                            (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("ChangeRequestReasonDocumentCancellationList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }


                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////Handleforward to users with document cancelled by changes new


                        const itemsToDelete = await sp.web.lists.getByTitle("AllProcessApprovalLevelList")
                            .items
                            .filter(`MainListID eq '${editItemID}' and ProcessName eq '${CONTENTTYPE_DocumentCancel}'`)
                            .select("Id", "Approvers/Id")
                            .expand("Approvers")
                            .getAll();


                        for (const delItem of itemsToDelete) {
                            await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(delItem.Id).delete();
                        }



                        const filteredCancelledBy = formData.DocumentCancelledBy.filter(
                            (user: any) => user.value !== formData.RequesterNameId
                        );

                        const filteredApprovers = forwardToArrEdit.filter(
                            (user: any) => user.Responsibility !== "Preparer"
                        );
                        let countLevel = 0;
                        for (let i = 0; i < filteredCancelledBy?.length; i++) {
                            let user = filteredCancelledBy[i];
                            if (user && user.value != formData.RequesterNameId) {
                                let level = i + 1;
                                countLevel = level;
                                let RoleID = 0;
                                const identifyexistingitem = forwardToArrEdit.filter(item => {
                                    return (
                                        item.Responsibility === "Preparer" &&
                                        item.approvers.length === 1 &&
                                        // item.approvers[0].value === user.value
                                        //DocumentCancelledByIds.includes(user.value)
                                        item.approvers[0].value === user.value
                                    );
                                });
                                if (identifyexistingitem.length > 0) {
                                    // If an item with Responsibility "Preparer" and one approver exists, skip this iteration
                                    RoleID = identifyexistingitem[0].role;
                                }

                                let approversIds: any[] = [];
                                // item.approvers.forEach((user: any) => {
                                if (user) {
                                    approversIds.push(user.value);
                                }


                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: formData.filename,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: RoleID,
                                    Level: Number(level),
                                    ApproversId: approversIds,
                                    LevelType: "One",
                                    // LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: 1,
                                    // ContentTitle:,
                                    MainListID: String(editItemID),
                                    RequestId: formData.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "Document Cancellation",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No",
                                    // RedirectionLink:,
                                    Responsibility: "Preparer",
                                    IsSignatureRequired: "Yes",

                                    CRIssueNumber: formData.IssueNumber,
                                    CRRevisionNumber: formData.RevisionNumber,


                                }

                                if (user && user.value != formData.RequesterNameId) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }


                            }


                        }

                        for (const item of filteredApprovers) {
                            countLevel++;
                            const approversIds: any[] = [];
                            item.approvers.forEach((user: any) => {
                                if (user?.value) {
                                    approversIds.push(user.value);
                                }
                            });

                            let arr2 = {
                                Title: currentUser.Title,
                                // ContentTitle: selectedOption.ReferenceNumber,
                                ContentTitle: formData.filename,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role,
                                Level: Number(countLevel),
                                ApproversId: approversIds,
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "No",
                                Maxlevel: item.approvers?.length,
                                // ContentTitle:,
                                MainListID: String(editItemID),
                                RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "Document Cancellation",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                IsApprovalGenerated: "No",
                                // RedirectionLink:,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",

                                CRIssueNumber: formData.IssueNumber,
                                CRRevisionNumber: formData.RevisionNumber,
                            }
                            // if (item.id) {
                            //     const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                            //     const postId2 = postResult2?.data?.ID;

                            // }
                            // else {

                            const postResult2 = await addAllProcessItem(arr2, sp);
                            const postId2 = postResult2?.data?.ID;

                            // }


                        }




                        // ///%%%%%%%%%%%%%%%%%%%%%%%%%%%%////////////////


                        setLoading(false);

                        Swal.fire(successMessage, '', 'success').then(async (result) => {
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

    const updatedigisignnew = async () => {
        let items = await updateDigitalsign(LIST_TITLE_DocCancel, sp, DigitalsignID);
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
                                                        src={require("../../assets/edc-gif.gif")}
                                                        className="alignrightl"
                                                        alt="Loading..."
                                                    />
                                                </div>
                                                <span>Loading </span>{" "}
                                                <span>
                                                    <img
                                                        src={require("../../assets/edcnew.gif")}
                                                        className="alignrightl"
                                                        alt="Loading..."
                                                    />
                                                </span>
                                            </div>
                                            :


                                            <div style={{ width: '100%' }} className="inbox-rightbar">
                                                <div className="card">
                                                    <div className="card-body">

                                                        {/* <h3 className="text-dark font-16 fw-bold mb-3">Requested By</h3> */}
                                                        <div className="previewIcon">
                                                            <h4 style={{ textAlign: 'left', margin: 'inherit' }} className="text-dark font-16 fw-bold mb-3">Requested By</h4>
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
                                                                                src={require("../../assets/digisign.png")}
                                                                                alt="Signature Icon"
                                                                            />
                                                                        </div>
                                                                    </span>
                                                                )}
                                                                {TemplateDoc && TemplateDoc.length > 0 && <div className='btn btn-primary'
                                                                    // onClick={() => OpenFile(TemplateDoc[0], "Open")}
                                                                    onClick={() => OpenFileTemplate(TemplateDoc[0], "Open")}

                                                                >
                                                                    {(() => {
                                                                        const parts = TemplateDoc[0]?.FileRef?.split('/');
                                                                        const folderName = parts && parts[3] ? parts[3] : null;
                                                                        return folderName === "DocumentCancellationDigitalSignedDocs" ? (
                                                                            // <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../../assets/noun-download-5006210.png")} alt="Download Icon" />
                                                                            // <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../../assets/digisigndownload.png")} alt="Digital Sign Download Icon" />
                                                                            <img style={{ cursor: 'pointer', height: '24px' }} className='mt-0' src={require("../../assets/signicon.png")} alt="Digital Sign Download Icon" />


                                                                        ) : (
                                                                            <img style={{ cursor: 'pointer', height: '24px' }} className='mt-0' src={require("../../assets/noun-download-5006210.png")} alt="Download Icon" />
                                                                            // <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../../assets/digisigndownload.png")} alt="Digital Sign Download Icon" />
                                                                        );
                                                                    })()}


                                                                    {/* <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../../assets/noun-download-5006210.png")} ></img> */}
                                                                    {/* <FontAwesomeIcon icon={faEye} /> */}
                                                                </div>
                                                                }
                                                            </div>
                                                        </div>

                                                        {FormLoading &&

                                                            <div className="loadernewadd mt-10">
                                                                <div>
                                                                    <img
                                                                        src={require("../../assets/edc-gif.gif")}
                                                                        className="alignrightl"
                                                                        alt="Loading..."
                                                                    />
                                                                </div>
                                                                <span>Loading </span>{" "}
                                                                <span>
                                                                    <img
                                                                        src={require("../../assets/edcnew.gif")}
                                                                        className="alignrightl"
                                                                        alt="Loading..."
                                                                    />
                                                                </span>
                                                            </div>
                                                            // :
                                                        }
                                                        <div className="row">
                                                            <div className="col-lg-4">


                                                                <div className="mb-3">
                                                                    <label htmlFor="RequesterName" className="form-label">Name</label>
                                                                    <input style={{ height: '47px' }} type="text" id="Name" name="RequesterName" className="form-control" title={formData.RequesterName} value={formData.RequesterName} disabled={true} />
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">


                                                                <div className="mb-3">
                                                                    <label htmlFor="Department" className="form-label">Department</label>
                                                                    <input style={{ height: '47px' }} type="text" id="Department" name="DepartmentName" className="form-control" title={formData.DepartmentName} value={formData.DepartmentName} disabled={true} />
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">


                                                                <div className="mb-3">
                                                                    <label htmlFor="RequesterDesignation" className="form-label">Designation</label>
                                                                    <input style={{ height: '47px' }} type="text" id="RequesterDesignation" name="RequesterDesignation" className="form-control" title={formData.RequesterDesignation} value={formData.RequesterDesignation} disabled={true} />
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="RequestDate" className="form-label">Request Date</label>
                                                                    <input style={{ height: '47px' }} type="text" id="RequestDate" name="RequestDate" className="form-control" value={formData.RequestDateNew} title={formData.RequestDateNew} onChange={(e) => setFormData({ ...formData, RequestDate: e.target.value })} disabled={true} />

                                                                    {/* <input type="date" id="RequestDate" name="RequestDate" className="form-control" value={formData.RequestDate} onChange={(e) => setFormData({ ...formData, RequestDate: e.target.value })} disabled={InputDisabled} /> */}
                                                                </div>
                                                            </div>


                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="DocumentCode" className="form-label">Document Code <span className="text-danger1"> *</span></label>
                                                                    {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                                                    <div title={formData?.DocumentCode || ""}> <Select
                                                                        // title={selectedOption?.value}
                                                                        title={formData?.DocumentCode || ""}
                                                                        id="DocumentCode"
                                                                        isClearable={true}
                                                                        options={rows}
                                                                        value={selectedOption}
                                                                        name="DocumentCode"
                                                                        className="newse"

                                                                        // className={`newse ${(!ValidDraft && !selectedOption) ? ErrorCls : ""} ${(!ValidSubmit && !selectedOption) ? ErrorCls : ""}`}
                                                                        onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                        placeholder="Search Document Code" isDisabled={InputDisabled || (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0 ? true : false)}
                                                                    /></div>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Issue No</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.IssueNumber} value={formData.IssueNumber} />
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Revision No</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.RevisionNumber} value={formData.RevisionNumber} />
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Reference No</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.ReferenceNumber} value={formData.ReferenceNumber} />
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Document Type</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.DocumentType} value={formData.DocumentType} />
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Location</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.Location} value={formData.Location} />
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Custodian</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.Custodian} value={formData.Custodian} />
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Amendment Type</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.AmendmentType} value={formData.AmendmentType} />
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Classification</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" title={formData.Classification} value={formData.Classification} />
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">

                                                                <div className="mb-3">
                                                                    <label htmlFor="example-email" className="form-label">Template Type</label>
                                                                    <input style={{ height: '47px' }} disabled type="text" id="templateid" name="example-email" className="form-control" placeholder="" title={formData.TemplateTypeValue} value={formData.TemplateTypeValue} />
                                                                </div>
                                                            </div>

                                                            {/* <div className="col-lg-8">

                                                            <div className="mb-3">
                                                                <label htmlFor="example-email" className="form-label">Document Link</label>


                                                                    <div className="text-dark mt-0"> <span onClick={() => OpenFile(DocumentLink)} style={{ color: "blue", cursor: "pointer" }}>{DocumentLink ? `${Tenant_URL}${DocumentLink?.FileRef}` : ""}</span>
                                                                    </div>

                                                                </div>
                                                            </div> */}

                                                            {/* changes  */}

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="RequesterName" className="form-label">Document Name</label>
                                                                    {/* <input type="text" id="Name" name="department" className="form-control" value={formData.Department} disabled={true} />
   */}
                                                                    <div
                                                                        title={formData.filename || ""}
                                                                        style={{ width: "100%" }}
                                                                    >
                                                                        <input style={{ height: '47px' }} type="text"
                                                                            id="example-email"
                                                                            name="example-email"
                                                                            // className={`form-control ${(!ValidDraft && filenameerr) ? "border-on-error" : ""} ${(!ValidSubmit && filenameerr) ? "border-on-error" : ""}`}
                                                                            className={`form-control`}
                                                                            // onChange={(e) => onChangefilename("filename", e.target.value)}
                                                                            // placeholder="Document name"
                                                                            value={formData.filename} disabled={true} />


                                                                    </div>

                                                                </div>
                                                            </div>

                                                            {/* changes */}

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">

                                                                    <div className='d-flex justify-content-between'>
                                                                        <div>
                                                                            <label htmlFor="bannerImage" className="form-label">
                                                                                Attachment
                                                                            </label>
                                                                        </div>
                                                                        <div>
                                                                            <div>
                                                                                {DocumentLink != null && DocumentLink.length != 0 ?
                                                                                    (<a style={{ fontSize: '0.875rem' }} onClick={() => setShowModal(true)}>
                                                                                        <FontAwesomeIcon icon={faPaperclip} /> 1 file Attached
                                                                                    </a>) : ""

                                                                                }
                                                                            </div>
                                                                        </div>
                                                                    </div>
                                                                    <input
                                                                        type="file"

                                                                        id="bannerImage"
                                                                        name="bannerImage"
                                                                        accept=".jpeg,.jpg,.png,.gif"
                                                                        // className={`form-control  ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                        className="form-control inputcss"
                                                                        // onChange={(e) => onFileChange(e, "bannerimg", "Document")}
                                                                        // disabled={ApprovalMode}
                                                                        disabled={true}

                                                                    />
                                                                </div>
                                                            </div>

                                                            {/*  */}

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="Prepared">Prepared By</label>
                                                                    <div className="mt-10"
                                                                    // title={sharewithusers.map(user => user.label).join(', ')}
                                                                    >
                                                                        <Select
                                                                            //onKeyDown={handleKeyDown}
                                                                            isClearable={true}
                                                                            options={rows1}
                                                                            isMulti
                                                                            value={sharewithusers}
                                                                            // value={formData.PreparedById}
                                                                            name="share with"
                                                                            //className={`form-control `}
                                                                            // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                            // onChange={(selectedOptions: any) => onSelectsharewith(selectedOptions)}
                                                                            placeholder=""
                                                                            //styles={getCustomStyles(sharewithusers.length)}
                                                                            isDisabled={true}
                                                                        />
                                                                    </div>
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-4">
                                                                <div className="mb-3">
                                                                    <label htmlFor="Prepared">Document Cancellation Requested By <span className="text-danger1"> *</span></label>
                                                                    <div className="mt-10"
                                                                    // title={sharewithusers.map(user => user.label).join(', ')}
                                                                    >
                                                                        {/* <Select
                                                                            //onKeyDown={handleKeyDown}
                                                                            isClearable={true}
                                                                            options={rows1}
                                                                            isMulti
                                                                            value={sharewithusers}
                                                                            // value={formData.PreparedById}
                                                                            name="share with"
                                                                            //className={`form-control `}
                                                                            // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                            // onChange={(selectedOptions: any) => onSelectsharewith(selectedOptions)}
                                                                            placeholder=""
                                                                            //styles={getCustomStyles(sharewithusers.length)}
                                                                            isDisabled={true}
                                                                        /> */}
                                                                        <Select
                                                                            isClearable={true}
                                                                            options={rows1}
                                                                            id="DocCancelledBy"
                                                                            isMulti
                                                                            value={formData.DocumentCancelledBy}
                                                                            name="DocCancelledBy"
                                                                            className={`newse ${(!ValidSubmit && !formData.DocumentCancelledBy.length) ? ErrorCls : ""}`}
                                                                            // className={`newse ForwardClsErr ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                            // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                            onChange={(selectedOptions: any) => onSelectDocCancelledBy(selectedOptions)}
                                                                            placeholder="Select"
                                                                            // isDisabled={InputDisabled || (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0 ? true : false)}
                                                                            isDisabled={InputDisabled}

                                                                        />
                                                                    </div>
                                                                </div>
                                                            </div>

                                                        </div>
                                                        {/* // } */}
                                                    </div>


                                                </div>






                                                <div className="card mt-3">
                                                    <div className="card-body">
                                                        <div className='row'>
                                                            <div className='col-sm-8'>
                                                                <h3 className="text-dark font-16 fw-bold mb-3">Description</h3>

                                                            </div>

                                                            <div className='col-sm-4'>
                                                                <div style={{ textAlign: "right" }} className="mt-2 float-end text-right">
                                                                    {/* <i style={{ cursor: "pointer" }} onClick={addField}  className="fe-plus-circle  font-20 text-warning"></i> */}
                                                                    {/* <i style={{ cursor: "pointer" }} className="fe-plus-circle  font-20 text-warning"></i> */}
                                                                    {/* {!InputDisabled && <img style={{width:'30px', cursor:'pointer', marginTop:'-7px'}} src={require("../assets/plus.png")} onClick={addCancelReason} className=''></img>} */}
                                                                    {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer', marginTop: '-7px' }} src={require("../../assets/plus.png")} onClick={addCancelReason} className=''></img>}


                                                                </div>
                                                            </div>
                                                        </div>

                                                        {/* <p className="sub-header">
                                                        Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s
                                                    </p> */}

                                                        <div className="row">
                                                            <div style={{ overflow: 'inherit' }} className="table-responsive mt-0 pt-0">
                                                                <table className="mtbalenewscrollnew4 table-centered table-nowrap table-borderless mb-0" id="tbl">
                                                                    <thead>
                                                                        <tr>
                                                                            <th style={{ minWidth: "30px", maxWidth: "30px" }}>S.No</th>
                                                                            <th>Description (Mention the Document Information)<span className="text-danger1"> *</span></th>
                                                                            <th>Reason for Cancellation<span className="text-danger1"> *</span></th>
                                                                            {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <th style={{ minWidth: "40px", maxWidth: "40px" }}>Action</th>}
                                                                        </tr>

                                                                    </thead>

                                                                    <tbody >
                                                                        {cancellReason.map((row, index) => (
                                                                            <tr key={index}>
                                                                                <td style={{ minWidth: "30px", maxWidth: "30px" }}>
                                                                                    <div
                                                                                        style={{ marginLeft: "5px" }}
                                                                                        className="indexdesign"
                                                                                    >
                                                                                        {index + 1}</div></td>
                                                                                <td>
                                                                                    <textarea id="simpleinput" disabled={InputDisabled}
                                                                                        // className="form-control"                                                                      
                                                                                        className={`newse mb-0 recommendClsErr ${(!ValidCancelReason) ? "border-on-error" : ""}`}
                                                                                        title={row.description}
                                                                                        value={row.description}
                                                                                        onChange={(e) => {
                                                                                            const newRowscancellReason = [...cancellReason];
                                                                                            newRowscancellReason[index].description = e.target.value;
                                                                                            setcancellReason(newRowscancellReason);
                                                                                        }}>

                                                                                    </textarea>
                                                                                    {/* <input type="text"
                                                                        /> */}

                                                                                </td>
                                                                                <td>
                                                                                    <textarea id="simpleinput" disabled={InputDisabled}
                                                                                        //  className="form-control"
                                                                                        className={`newse mb-0 recommendClsErr ${(!ValidCancelReason) ? "border-on-error" : ""}`}
                                                                                        title={row.reason}
                                                                                        value={row.reason}
                                                                                        onChange={(e) => {
                                                                                            const newRowscancellReason = [...cancellReason];
                                                                                            newRowscancellReason[index].reason = e.target.value;
                                                                                            setcancellReason(newRowscancellReason);
                                                                                        }}>

                                                                                    </textarea>

                                                                                    {/* <input type="text"
                                                                        /> */}
                                                                                </td>
                                                                                {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td style={{ minWidth: "40px", maxWidth: "40px", textAlign: 'center' }}>
                                                                                    {/* <img src={require("../../../CustomAsset/del.png")} style={{ width: '30px', cursor: 'pointer', marginTop: '-7px' }} onClick={() => deleteLocalFile(index, cancellReason)}></img> */}
                                                                                    <img src={require("../../assets/del.png")} style={{ width: '30px', cursor: 'pointer', marginTop: '-7px' }} onClick={() => deleteLocalFile(index, cancellReason)}></img>

                                                                                    {/* <img src={require("../../../CustomAsset/del.png")} className='sidebariconsmall' style={{ width: '30px', cursor: 'pointer', marginTop: '-7px' }} onClick={() => deleteLocalFile(index, cancellReason)}></img> */}
                                                                                </td>
                                                                                }
                                                                            </tr>
                                                                        ))}
                                                                    </tbody>


                                                                </table>


                                                            </div>
                                                        </div>


                                                    </div>


                                                </div>
                                                {/* //////&&&&& changes*/}
                                                {(editID != null && editID.CurrentUserRole === "Initiator" && editID.IsInitiator == "Yes" && editID.Level == 0) && editID.Status === "Pending" &&
                                                    <div className="card mt-3">
                                                        <div className="col-12 text-center card-body">



                                                            <div className="row">
                                                                <div className="col-lg-12">

                                                                    <div className="mb-0" >

                                                                        <label htmlFor="example-textarea" className="form-label text-dark font-14" style={{ textAlign: 'left' }}>Remarks <span className="text-danger1"> *</span></label>

                                                                        <textarea
                                                                            style={{ height: '80px' }}
                                                                            className={`form-control `}
                                                                            id="Remark-textarea2"
                                                                            rows={5}
                                                                            name="Remark"
                                                                            value={formData.Remark}
                                                                            onChange={(e) => setFormData({ ...formData, Remark: e.target.value })}
                                                                        ></textarea>

                                                                    </div>

                                                                </div>



                                                            </div>




                                                        </div>
                                                    </div>
                                                }

                                                {/* ////////&&&& */}

                                                {/* /////////////////%%%%%%%%%%%%%%%%%%%%%%%% */}

                                                {/* {modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES" && */}
                                                {modeValue === "approve" && editID != null && editID.CurrentUserRole !== "Initiator" && MainEditItem !== null && MainEditItem?.Status !== "Save as draft" &&

                                                    <div className="card mt-3" style={{ marginBottom: '17px' }}>
                                                        <div className="card-body">
                                                            <div className='row'>
                                                                <div className='col-sm-11'>
                                                                    <h4 className="header-title text-dark font-16 fw-bold mb-3 ">Forward Approval To <span >(Preparer (Cancellation Requested By) {'>'} Reviewer {'>'} Endorser {'>'} Signer\Approver)</span></h4>

                                                                </div>
                                                                <div className='col-sm-1'>
                                                                    <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right", paddingRight: "22px" }}>
                                                                        {/* {editID.CurrentUserRole === "OES" && <img style={{ width: '34px' }} src={require("../assets/plus.png")} onClick={handleAddRow} className='' />} */}
                                                                        {(editID?.CurrentUserRole === "OES" && editID?.Status === "Pending") && <img style={{ width: '34px' }} src={require("../../assets/plus.png")} onClick={handleAddRow} className='' />}

                                                                        {/* <i style={{ cursor: "pointer" }} onClick={handleAddRow} className="fe-plus-circle font-20 text-warning"></i> */}
                                                                    </div>
                                                                </div>

                                                            </div>

                                                            <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">
                                                                <table style={{ overflow: 'inherit' }} className="mtbalenew  table-centered table-nowrap table-borderless mb-0 overhi" id="myTabl">
                                                                    <thead >
                                                                        <tr>
                                                                            <th style={{ minWidth: "35px", maxWidth: "35px" }}>S.No</th>
                                                                            <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Role<span className="text-danger1"> *</span></th>
                                                                            <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Responsibility<span className="text-danger1"> *</span></th>
                                                                            <th style={{ minWidth: "40px", maxWidth: "40px" }} title="Use your electronic digital signature to sign this document digitally">E-Sign?</th>
                                                                            <th style={{ minWidth: '40px', maxWidth: '40px' }} >Level</th>
                                                                            {/* <th style={{ minWidth: '120px', maxWidth: '120px' }}>Approver name</th> */}
                                                                            <th>Approver name<span className="text-danger1"> *</span></th>
                                                                            <th style={{ minWidth: '80px', maxWidth: '80px' }} >Approval criteria<span className="text-danger1"> *</span></th>
                                                                            <th style={{ minWidth: '36px', maxWidth: '36px' }}>Action</th>
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
                                                                                <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px', }} className="ng-binding" title={UserRoles.find((role: any) => role.value === row.role)?.label || "Select Role"}>
                                                                                    <select
                                                                                        // className="form-select"

                                                                                        className={`form-select ForwardClsErr newse ${(!ValidForwardTo) ? "border-on-error" : ""} `}
                                                                                        title={UserRoles.find((role: any) => role.value === row.role)?.label || "Select Role"}

                                                                                        onChange={(e) => onSelectRole(e, row.level)} value={row.role} disabled={!(editID.CurrentUserRole == "OES" && editID.Status == "Pending")}>

                                                                                        <option value="" selected>Select Role</option>

                                                                                        {UserRoles.filter((role: any) =>
                                                                                            !forwardToArr.some((r) => r.role === role.value && r.level !== row.level) || role.value === row.role // Allow the current row's role
                                                                                        ).map((role: any, idx: number) => (
                                                                                            <option key={idx} value={role.value} >{role.label}</option>
                                                                                        ))}
                                                                                    </select>

                                                                                </td>
                                                                                <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px' }}>
                                                                                    <div
                                                                                    //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                                                                    >
                                                                                        <select
                                                                                            id="responsibleId"
                                                                                            value={row.Responsibility}
                                                                                            onChange={(e) => handleChangeResp(e, row.level)}
                                                                                            className={`newse ForwardClsErr form-select`}
                                                                                            // disabled={InputDisabled}
                                                                                            disabled={!(editID.CurrentUserRole == "OES" && editID.Status == "Pending") || row.IsDisabled || row.Responsibility === "Preparer"}
                                                                                            title={row.Responsibility ? row.Responsibility : "Select"}
                                                                                        >
                                                                                            <option value="">Select </option>
                                                                                            <option value="Preparer" disabled>Preparer</option>
                                                                                            <option value="Reviewer">Reviewer</option>
                                                                                            <option value="Endorser">Endorser</option>
                                                                                            <option value="Signer">Signer</option>
                                                                                            {/* <option value="Signer">Signer</option>
                                                                                            <option value="Reviewer">Reviewer</option>
                                                                                            <option value="Endorser">Endorser</option>
                                                                                            <option value="Preparer" disabled>Preparer</option> */}
                                                                                        </select>

                                                                                    </div>
                                                                                </td>

                                                                                <td style={{ minWidth: "40px", maxWidth: "40px", overflow: 'inherit' }}>
                                                                                    <div
                                                                                    //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                                                                    >
                                                                                        <input
                                                                                            type="checkbox"
                                                                                            checked={row.IsSignatureRequired}
                                                                                            // disabled={!(editID.CurrentUserRole == "OES" && editID.Status == "Pending") && (row.Responsibility === "Signer" ||row.Responsibility === "Endorser" || row.Responsibility === "") }
                                                                                            // disabled={row.Responsibility !== "Endorser" || row.Responsibility === "" || InputDisabled}
                                                                                            disabled={row.Responsibility !== "Reviewer" || (!(editID.CurrentUserRole == "OES" && editID.Status == "Pending"))}
                                                                                            style={{ marginLeft: '17px', width: "15px" }}
                                                                                            title="Signature Required"
                                                                                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                                                                                                const isChecked = e.target.checked;
                                                                                                const updatedArr = forwardToArr.map(row1 =>
                                                                                                    row1.level === row.level ? { ...row1, IsSignatureRequired: isChecked } : row1
                                                                                                );
                                                                                                setForwardToArr(updatedArr);
                                                                                            }}
                                                                                        />

                                                                                    </div>
                                                                                </td>
                                                                                <td style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>Level {index + 1}</td>
                                                                                <td style={{ overflow: 'inherit' }} title={row.approvers.map((approver: any) => approver.label).join(", ")}>

                                                                                    <Select
                                                                                        options={rows1}

                                                                                        isMulti
                                                                                        value={row.approvers}
                                                                                        name="Approvers"
                                                                                        className={`newse ForwardClsErr ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                        // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                        onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                                                                        placeholder="Enter Approver Name"
                                                                                        isDisabled={!(editID.CurrentUserRole == "OES" && editID.Status == "Pending") || row.IsDisabled || row.Responsibility === "Preparer"}
                                                                                    />



                                                                                </td>



                                                                                <td style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                                                    {/* <label htmlFor="approvalType">Approval Type: </label> */}
                                                                                    <select id="approvalType" value={row.approvalType} onChange={(e) => handleChange(e, row.level)} className={`newse ForwardClsErr form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`} disabled={!(editID.CurrentUserRole == "OES" && editID.Status == "Pending") || row.IsDisabled || row.Responsibility === "Preparer"} title={row.approvalType === "One" ? "Anyone" : row.approvalType === "All" ? "Everyone" : "Select Approval Type"}>
                                                                                        <option value="">Select </option>
                                                                                        <option value="One">Anyone</option>
                                                                                        <option value="All">Everyone</option>
                                                                                    </select>
                                                                                </td>
                                                                                <td style={{ minWidth: '36px', maxWidth: '36px', overflow: 'inherit' }}>
                                                                                    {/* <i className="fe-trash-2 text-danger"></i> */}
                                                                                    {/* {editID.CurrentUserRole === "OES"? <img src={require("../../../CustomAsset/del.png")} onClick={() => handleDeleteRow(index)} />:
                                                                                <img src={require("../assets/recycle-bin.png")}  className='sidebariconsmall' />} */}
                                                                                    {((editID.CurrentUserRole === "OES" && editID.Status === "Pending" && row.Responsibility != "Preparer")) ? <img src={require("../../assets/del.png")} onClick={() => handleDeleteRow(index)} /> :
                                                                                        <img src={require("../../assets/recycle-bin.png")} className='sidebariconsmall' />}

                                                                                </td>
                                                                            </tr>

                                                                        ))}


                                                                    </tbody>
                                                                </table>
                                                            </div>



                                                            {editID.CurrentUserRole === "OES" && editID.Status === "Pending" && <div className="row mt-3">
                                                                <div className="col-12 text-center">



                                                                    <div className="row">
                                                                        <div className="col-lg-12">

                                                                            <div className="mb-0" >

                                                                                <label htmlFor="example-textarea" className="form-label text-dark font-14" style={{ textAlign: 'left' }}>Remarks <span className="text-danger1"> *</span></label>

                                                                                <textarea
                                                                                    style={{ height: '80px' }}
                                                                                    className={`form-control `}
                                                                                    id="Remark-textarea"
                                                                                    rows={5}
                                                                                    name="Remark"
                                                                                    value={formData.Remark}
                                                                                    onChange={(e) => setFormData({ ...formData, Remark: e.target.value })}
                                                                                ></textarea>

                                                                            </div>

                                                                        </div>



                                                                    </div>

                                                                    {/* <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardApproval("Forward")} >
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
                                                                    </button> */}

                                                                    <div className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardApproval("Forward")} >
                                                                        <i className="fe-check-circle me-1"></i> Forward
                                                                    </div>

                                                                    <div className="btn btn-warning waves-effect waves-light m-1" onClick={() => ForwardApproval("Rework")} >
                                                                        <i className="fe-corner-up-left me-1"></i> Rework
                                                                    </div>

                                                                    <div className="btn btn-danger waves-effect waves-light m-1" onClick={() => ForwardApproval("Rejected")} >
                                                                        <i className="fe-x me-1"></i> Reject
                                                                    </div>

                                                                    <div className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}>
                                                                        <i className="fe-x me-1"></i> Cancel
                                                                    </div>

                                                                </div>
                                                            </div>
                                                            }
                                                        </div>
                                                    </div>
                                                }

                                                {(modeValue === "view") && forwardToArr.length > 0 && MainEditItem !== null && MainEditItem?.Status !== "Save as draft" &&

                                                    <div className="card mt-3" style={{ marginBottom: '17px' }}>
                                                        <div className="card-body">
                                                            <div className='row'>
                                                                <div className='col-sm-11'>
                                                                    <h4 className="header-title text-dark font-16 fw-bold mb-3 ">Forward Approval To <span >(Preparer (Cancellation Requested By) {'>'} Reviewer {'>'} Endorser {'>'} Signer\Approver)</span></h4>
                                                                    <label>Define the approval hierarchy to ensure requests are routed to the appropriate approvers.
                                                                    </label>

                                                                </div>
                                                                <div className='col-sm-1'>
                                                                    <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right", paddingRight: "22px" }}>
                                                                        {/* {editID.CurrentUserRole === "OES" && <img style={{ width: '34px' }} src={require("../assets/plus.png")} onClick={handleAddRow} className='' />} */}
                                                                        {/* {(editID?.CurrentUserRole === "OES" ||MainEditItem?.Status == "Rework") && <img style={{ width: '34px' }} src={require("../../assets/plus.png")} onClick={handleAddRow} className='' />} */}

                                                                        {/* <i style={{ cursor: "pointer" }} onClick={handleAddRow} className="fe-plus-circle font-20 text-warning"></i> */}
                                                                    </div>
                                                                </div>

                                                            </div>

                                                            <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">
                                                                <table style={{ overflow: 'inherit' }} className="mtbalenew  table-centered table-nowrap table-borderless mb-0 overhi" id="myTabl">
                                                                    <thead >
                                                                        <tr>
                                                                            <th style={{ minWidth: "35px", maxWidth: "35px" }}>S.No</th>
                                                                            <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Role<span className="text-danger1"> *</span></th>
                                                                            <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Responsibility<span className="text-danger1"> *</span></th>
                                                                            <th style={{ minWidth: "40px", maxWidth: "40px" }} title="Use your electronic digital signature to sign this document digitally">E-Sign?</th>
                                                                            <th style={{ minWidth: '40px', maxWidth: '40px' }} >Level</th>
                                                                            {/* <th style={{ minWidth: '120px', maxWidth: '120px' }}>Approver name</th> */}
                                                                            <th>Approver name<span className="text-danger1"> *</span></th>
                                                                            <th style={{ minWidth: '80px', maxWidth: '80px' }} >Approval criteria<span className="text-danger1"> *</span></th>
                                                                            <th style={{ minWidth: '36px', maxWidth: '36px' }}>Action</th>
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
                                                                                <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px', }} className="ng-binding" >
                                                                                    <select
                                                                                        // className="form-select"

                                                                                        className={`form-select newse ${(!ValidForwardTo) ? "border-on-error" : ""} `}
                                                                                        title={UserRoles.find((role: any) => role.value === row.role)?.label || "Select Role"}

                                                                                        onChange={(e) => onSelectRole(e, row.level)} value={row.role} disabled={true}>

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
                                                                                <td style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px' }}>
                                                                                    <div
                                                                                    //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                                                                    >
                                                                                        <select
                                                                                            id="responsibleId"
                                                                                            value={row.Responsibility}
                                                                                            onChange={(e) => handleChangeResp(e, row.level)}
                                                                                            className={`newse HierarchyClsErr form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                            // disabled={InputDisabled}
                                                                                            disabled={true}
                                                                                            title={row.Responsibility ? row.Responsibility : "Select"}
                                                                                        >
                                                                                            <option value="">Select </option>
                                                                                            <option value="Signer">Signer</option>
                                                                                            <option value="Reviewer">Reviewer</option>
                                                                                            <option value="Endorser">Endorser</option>
                                                                                            <option value="Preparer">Preparer</option>
                                                                                        </select>

                                                                                    </div>
                                                                                </td>
                                                                                <td style={{ minWidth: "40px", maxWidth: "40px", overflow: 'inherit' }}>
                                                                                    <div
                                                                                    //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                                                                    >
                                                                                        <input
                                                                                            type="checkbox"
                                                                                            checked={row.IsSignatureRequired}
                                                                                            disabled={true}
                                                                                            // disabled={row.Responsibility === "Signer" || row.Responsibility === "" || InputDisabled}
                                                                                            style={{ marginLeft: '17px', width: "15px" }}
                                                                                            title="Signature Required"
                                                                                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                                                                                                const isChecked = e.target.checked;
                                                                                                const updatedArr = forwardToArr.map(row1 =>
                                                                                                    row1.level === row.level ? { ...row1, IsSignatureRequired: isChecked } : row1
                                                                                                );
                                                                                                setForwardToArr(updatedArr);
                                                                                            }}
                                                                                        />

                                                                                    </div>
                                                                                </td>
                                                                                <td style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>Level {index + 1}</td>
                                                                                <td style={{ overflow: 'inherit', minWidth: '120px', maxWidth: '120px', }} title={row.approvers.map((approver: any) => approver.label).join(", ")}>

                                                                                    <Select
                                                                                        options={rows1}

                                                                                        isMulti
                                                                                        value={row.approvers}
                                                                                        name="Approvers"
                                                                                        className={`newse ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                        // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                        onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                                                                        placeholder="Enter Approver Name"
                                                                                        isDisabled={true}
                                                                                    />



                                                                                </td>
                                                                                <td style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                                                    {/* <label htmlFor="approvalType">Approval Type: </label> */}
                                                                                    <select id="approvalType" value={row.approvalType} onChange={(e) => handleChange(e, row.level)} className={`newse form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`} disabled={true} title={row.approvalType === "One" ? "Anyone" : row.approvalType === "All" ? "Everyone" : "Select Approval Type"} >
                                                                                        <option value="">Select</option>
                                                                                        <option value="One">Anyone</option>
                                                                                        <option value="All">Everyone</option>
                                                                                    </select>
                                                                                </td>
                                                                                <td style={{ minWidth: '36px', maxWidth: '36px', overflow: 'inherit' }}>
                                                                                    {/* <i className="fe-trash-2 text-danger"></i> */}
                                                                                    {/* {editID.CurrentUserRole === "OES"? <img src={require("../../../CustomAsset/del.png")} onClick={() => handleDeleteRow(index)} />:
                            <img src={require("../assets/recycle-bin.png")}  className='sidebariconsmall' />} */}
                                                                                    {/* {(editID.CurrentUserRole === "OES"||MainEditItem?.Status !== "Rework") ? <img src={require("../../assets/del.png")} onClick={() => handleDeleteRow(index)} /> : */}
                                                                                    <img src={require("../../assets/recycle-bin.png")} className='sidebariconsmall' />
                                                                                    {/* } */}

                                                                                </td>
                                                                            </tr>

                                                                        ))}


                                                                    </tbody>
                                                                </table>
                                                            </div>




                                                        </div>
                                                    </div>
                                                }

                                                {/* {modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES" && (
                                                <div className="card">
                                                    <div className="card-body">
                                                        <h4 className="header-title mb-0">Remarks</h4>
                                                        <textarea
                                                            className="form-control"
                                                            value={remark}
                                                            onChange={(e) => setRemark(e.target.value)}
                                                            placeholder="Enter your remarks here..."
                                                        ></textarea>
                                                    </div>
                                                </div>
                                            )} */}

                                                {/* ////////////Approval card */}

                                                {
                                                    //let forrework=ApprovalRequestItem && ApprovalRequestItem.IsRework=='Yes'&& ApprovalRequestItem.LevelSequence!=0;
                                                    // (InputDisabled && ApprovalRequestItem) || (ApprovalRequestItem && ApprovalRequestItem.IsRework == 'Yes' && ApprovalRequestItem.LevelSequence != 0) ? (
                                                    (InputDisabled && editID != null && modeValue === "approve" && editID.ApprovalType === "Approval" && editID.Status === "Pending") ? (
                                                        <WorkflowAction currentItem={editID} ctx={props.context} ContentType={CONTENTTYPE_DocumentCancel}
                                                            DisableApproval={false} DisableCancel={false}
                                                        // DisableApproval={ApprovalRequestItem && ApprovalRequestItem.IsRework == 'Yes' && ApprovalRequestItem.LevelSequence != 0}
                                                        // DisableCancel={ApprovalRequestItem && ApprovalRequestItem.IsRework == 'Yes' && ApprovalRequestItem.LevelSequence != 0}
                                                        //DisableReject={ApprovalRequestItem && ApprovalRequestItem.IsRework=='Yes'&& ApprovalRequestItem.LevelSequence!=0}
                                                        />
                                                    ) : (<div></div>)
                                                }

                                                {/* ////////////Audit History card */}
                                                {/* {editID !== null && editID.length != 0 && modeValue === "approve" && */}
                                                {MainEditItem !== null && MainEditItem.length != 0 && MainEditItem?.Status !== "Save as draft" &&
                                                    <WorkflowAuditHistory ContentItemId={MainEditItem} ContentType={CONTENTTYPE_DocumentCancel} ctx={props.context} />
                                                }
                                                {/* ////////////Audit History card */}

                                                {/* ////////////Approval card */}


                                                {/* </div> */}
                                                {/* /////////////////%%%%%%%%%%%%%%%%%%%%%%%% */}

                                                <div className="row mt-3">
                                                    {/* <div className="col-12 text-center">

                                                        {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                            <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                            Save As Draft</button>}

                                                        {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                            <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                            Submit</button>}

                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</button>}


                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</button>}


                                                       
                                                        {((modeValue === "" || modeValue === "edit" || modeValue === "view") || (editID !== null && editID.IsInitiator == "Yes")) &&
                                                            <button style={{ width: '145px' }} type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</button>
                                                        }
                                                        {((modeValue === "approve") && (editID !== null && editID.Status == "Approved")) &&
                                                            <button style={{ width: '145px' }} type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</button>
                                                        }
                                                       
                                                    </div> */}
                                                    <div className="col-12 text-center">

                                                        {/* {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                        <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                        Save As Draft</div>}

                                                    {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                        <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                        Submit</div>} */}

                                                        {(!editID && (modeValue === "" || modeValue === "edit") && (MainEditItem == null || MainEditItem?.Status === "Save as draft")) && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                            <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                            Save As Draft</div>}

                                                        {(!editID && (modeValue === "" || modeValue === "edit") && (MainEditItem == null || MainEditItem?.Status === "Save as draft")) && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                            <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                            Submit</div>}

                                                        {/* {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</div>}


                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</div>} */}

                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</div>}


                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <div style={{ width: '145px' }} className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</div>}


                                                        {/* {((modeValue === "" || modeValue === "edit" || modeValue === "view") || (editID !== null && editID.IsInitiator == "Yes" && (editID.Status == "Pending"))) &&
                                                            <div style={{ width: '145px' }} className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</div>
                                                        }
                                                        {((modeValue === "approve") && ((editID !== null && (editID.Status == "Approved" || editID.Status == "Rework")))) &&
                                                            <div style={{ width: '145px' }} className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</div>
                                                        }  */}

                                                        {((modeValue == "" || modeValue == "edit" || modeValue === "view") || (modeValue === "approve" && (!editID || (editID !== null && (editID.Status == "Approved" || editID.Status == "Rework" || (editID.Status == "Pending" && editID.CurrentUserRole == "Initiator")))))) &&
                                                            <div style={{ width: '145px' }} className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</div>
                                                        }

                                                    </div>
                                                </div>

                                                {/* /////////// */}

                                                <Modal show={showModal} onHide={() => setShowModal(false)} size={Showfile ? "xl" : "lg"} className='filemodal'>
                                                    <Modal.Header closeButton>
                                                        <Modal.Title> <h4 className='font-16 fw-bold text-dark'></h4> Attachment Details <br />
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for Document Cancellation
                                                            </p>

                                                        </Modal.Title>


                                                    </Modal.Header>
                                                    <Modal.Body className="" id="style-5">

                                                        {DocumentLink &&
                                                            (
                                                                <>

                                                                    {Showfile ?

                                                                        <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                                                        :
                                                                        <table className="mtbalenew">
                                                                            <thead style={{ background: '#eef6f7' }}>
                                                                                <tr>
                                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                                                                    <th>File Name</th>
                                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>Upload date</th>
                                                                                    <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Action</th>
                                                                                </tr>
                                                                            </thead>
                                                                            <tbody>
                                                                                {DocumentLink != null && DocumentLink.length != 0 && (
                                                                                    <tr>
                                                                                        <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>1</td>
                                                                                        {/* <td title={DocumentLink?.FileLeafRef?.includes('_') && DocumentLink?.FileLeafRef?.includes('.') ? DocumentLink.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1') : DocumentLink?.FileLeafRef}>
                                                                                        {DocumentLink?.FileLeafRef?.includes('_') && DocumentLink?.FileLeafRef?.includes('.') ? DocumentLink.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1') : DocumentLink?.FileLeafRef}
                                                                                    </td> */}
                                                                                        <td title={DocumentLink?.FileLeafRef?.replace(/_\d+(\.\w+)$/, '$1') || DocumentLink?.FileLeafRef}>{DocumentLink?.FileLeafRef?.replace(/_\d+(\.\w+)$/, '$1') || DocumentLink?.FileLeafRef}</td>

                                                                                        {/* <td title={DocumentLink?.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1')}>{DocumentLink?.FileLeafRef.replace(/_\d+(\.\w+)$/, '$1')}</td> */}

                                                                                        <td style={{ minWidth: '70px', maxWidth: '70px' }}>{DocumentLink.Created
                                                                                            ? new Intl.DateTimeFormat('en-GB', {
                                                                                                day: '2-digit',
                                                                                                month: 'short',
                                                                                                year: 'numeric'
                                                                                            }).format(new Date(DocumentLink.Created)).replace(/ /g, "/")
                                                                                            : ""}</td>
                                                                                        <td style={{ textAlign: 'center', minWidth: '50px', maxWidth: '50px' }} >
                                                                                            <span onClick={() => OpenFile(DocumentLink, "Open")} title='preview file' style={{ color: "blue", cursor: "pointer", margin: "10px" }}><FontAwesomeIcon icon={faEye} /></span>

                                                                                            <span onClick={() => OpenFile(DocumentLink, "Download")} title='download file' style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                                <FontAwesomeIcon icon={faDownload} /></span>
                                                                                        </td>


                                                                                    </tr>
                                                                                )}
                                                                            </tbody>
                                                                        </table>
                                                                    }</>
                                                            )
                                                        }

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

const DocumentCancellationProcess: React.FC<IDocumentCancellationProcessProps> = (props) => (
    <Provider>
        <DocumentCancellationProcessContext props={props} />
    </Provider>
);


export default DocumentCancellationProcess
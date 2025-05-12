// this is our main file for CR
import * as React from 'react';

// import type { IChangeDocumentRequestProps } from './IChangeDocumentRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import Provider from '../../../GlobalContext/provider';
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import HorizontalNavbar from '../../horizontalNavBar/components/HorizontalNavBar';
import CustomBreadcrumb from './CustomBreadcrumb/CustomBreadcrumb';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import UserContext from "../../../GlobalContext/context";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import "./changeDocumentRequest.scss";
import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import {
  addAllProcessItem, addApprovalItem, addItem, addItemChangeRequestReasonlist,
  addItemChangeRequestList, getAllAmendmentType, getAllClassificationMaster, getAllDocumentCode,
  getAllProcessData, getAllRequestType, getApprovalByID, getApprovalByID2, getDataRoles,
  getDocumentLinkByID, getFormNameID, getItemByID, getItemByIDChangeRequest, getItemByIDCR,
  getListNameID, GetQueryString, getRequesterID, UpdateAllProcessItem, updateApprovalItem,
  updateItem, updateItemChangeRequestReasonList, updateItemChangeRequestList,
  getDocumentCodeselected, getDocumentLinkByIDarr,
  getAllDepartment,
  getAllTemplateType,
  getGeneratedTemplateDocCR
} from './DocumentCancellation';
import Select from "react-select";
import Swal from 'sweetalert2';
// import { FormSubmissionMode } from '../../../Shared/Interfaces';
import { decryptId } from '../../../APISearvice/CryptoService';
// import { getUrlParameterValue } from '../../../Shared/Helper';
import { WorkflowAction } from './WorkflowAction';
// import { WorkflowAuditHistory } from './WorkflowAuditHistory/WorkflowAuditHistory';
import { WorkflowAuditHistory } from '../../../CustomJSComponents/WorkflowAuditHistory/WorkflowAuditHistory';
import { CONTENTTYPE_ChangeDocument, CONTENTTYPE_DocumentCancel, LIST_TITLE_ChangeRequest, Tenant_URL } from './Constants';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faDownload, faEye, faPaperclip } from '@fortawesome/free-solid-svg-icons';
// import { uploadFile } from '../../../APISearvice/MediaService';
import { Modal } from 'react-bootstrap';
import { Link } from '@fluentui/react';
import moment from 'moment';
import { DatePicker } from 'office-ui-fabric-react';
import * as XLSX from 'xlsx';
import { SITE_URL } from '../../../Shared/Constants';
let newfileupload: any
let newfilepreview: any;
let filechanged: boolean = false;
let locationPath: any;
export enum FormSubmissionMode {
  DRAFT, SUBMIT
}

export interface IChangeDocumentRequestProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: any;
  siteUrl: string;
}
interface ForwardTo {
  id: number;
  role: number;
  level: number;
  approvers: any[]; // Or a more specific type like `string[]` or `SPUser[]`
  leveltype: string
}
interface ChangeRequestCheckbox {
  id: number;
  name: string;
  isActive: boolean;
}
interface Location {
  locationId: number; // ID for Location lookup
  locationName: string;
  locationCode: string;
}
interface Custodian {
  custodianId: number; // ID for Custodian lookup
  custodianName: string;
  custodianCode: string;
}

interface DocumentType {
  documentTypeId: number; // ID for Document Type lookup
  documentTypeName: string;
  documentTypeCode: string;
}
interface IEmployeeDetails {
  userId: number;
  preffedName?: string;
  deptName?: string;
  office?: string;
  Title?: string;
}
const ChangeDocumentRequestContext = ({ props }: any) => {
  const sp: SPFI = getSP();
  const elementRef = React.useRef<HTMLDivElement>(null);
  const elementRef1 = React.useRef<HTMLDivElement>(null);
  const siteUrl = props.siteUrl;
  const tenantUrl = props.siteUrl?.split("/sites/")[0];
  const { useHide }: any = React.useContext(UserContext);
  const [InputDisabled, setInputDisabled] = React.useState(false);
  const selectedTextDiv = document.getElementById('selectedText');

  selectedTextDiv.style.display = 'none';
  const Breadcrumb = [
    {
      MainComponent: "My Request",
      MainComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx`,
    },
    {
      ChildComponent: "Change Request",
      ChildComponentURl: `${siteUrl}/SitePages/EDCMAIN.aspx#/ChangeDocumentRequest`,
    },
  ];
  const [requesttypeerr, setrequesttypeerr] = React.useState(false);
  const [documentcodeerr, setdocumentcodeerr] = React.useState(false);
  const [amendmenterr, setamendmenterr] = React.useState(false);
  const [departmenterr, setdepartmenterr] = React.useState(false);
  const [templatetypeerr, settemplatetypeerr] = React.useState(false);
  const [classificationerr, setclassificationerr] = React.useState(false);
  const [custodianerr, setcustodianerr] = React.useState(false);
  const [locationerr, setlocationerr] = React.useState(false);
  const [documenttypeerr, setdocumenttypeerr] = React.useState(false);
  const [attachmenterr, setattachmenterr] = React.useState(false);
  const [changedescriptionerr, setchangedescriptionerr] = React.useState(false);
  const [changereasonerr, setchangereasonerr] = React.useState(false);
  const [changerequesttypeerr, setchangerequesttypeerr] = React.useState(false);
  const [Loading, setLoading] = React.useState(false);
  const [FormItemId, setFormItemId] = React.useState(null);
  const [editID, setEditID] = React.useState(null);
  const [editItemID, setEditItemID] = React.useState(null);
  const [MainEditItem, setMainEditItem] = React.useState(null);
  const [rows, setRows] = React.useState<any>([]);
  const [doccoderows, setdoccoderows] = React.useState<any>([]);
  const [ReqType, setReqType] = React.useState<any>([]);
  const [Departopt, setDepartment] = React.useState<any>([]);
  const [TemplateTypeopt, setTemplateType] = React.useState<any>([]);
  const [Amendtype, setAmendtype] = React.useState<any>([]);
  const [Classificationopt, setClassificationopt] = React.useState<any>([]);
  const [LocationOpt, setLocationOpt] = React.useState<any>([]);
  const [Custodianopt, setCustodianopt] = React.useState<any>([]);
  const [DocumentTypeOpt, setDocumentTypeOpt] = React.useState<any>([]);
  const [UserRoles, setUserRoles] = React.useState<any>([]);
  const [rows1, setRows1] = React.useState<any>([]);
  const [DocumentLink, setDocumentLink] = React.useState(null);
  const [currentUser, setCurrentUser] = React.useState(null);
  const [selectedOption, setSelectedOption] = React.useState(null);
  const [selectedOptionReq, setSelectedOptionReq] = React.useState(null);
  const [selectedOptionAmend, setSelectedOptionAmend] = React.useState(null);
  const [SelectedOptionDepart, setSelectedOptionDepart] = React.useState(null);
  const [SelectedOptionTemplate, setSelectedOptionTemplate] = React.useState(null);

  const [selectedOptionClass, setSelectedOptionClassification] = React.useState(null);
  const [selectedOptionLoc, setselectedOptionLoc] = React.useState(null);
  const [selectedOptionCusto, setselectedOptionCusto] = React.useState(null);
  const [selectedOptionDoctype, setselectedOptionDoctype] = React.useState(null);
  const [selectedPeople, setSelectedPeople] = React.useState(null);
  const [selectedRole, setSelectedRole] = React.useState(null);
  const [ValidDraft, setValidDraft] = React.useState(true);
  const [ValidSubmit, setValidSubmit] = React.useState(true);
  const [Validforward, setValidforward] = React.useState(true);
  const [showviewdownload, setshowviewdownload] = React.useState(true);

  const [RequesterRoleId, setRequesterRoleId] = React.useState(null);
  const [FormNameId, setFormNameId] = React.useState(null);
  const [ListNameId, setListNameId] = React.useState(null);
  const [ValidCancelReason, setValidCancelReason] = React.useState(true);
  const [showdate, setshowdate] = React.useState(false);
  const [editForm, setEditForm] = React.useState(false);
  const [showpreviousattachment, setshowpreviousattachment] = React.useState(false);
  const [disabledforwardarr, setdisabledforwardarr] = React.useState(false);
  const [modeValue, setmode] = React.useState("");
  const [referencedocCode, setreferencedocCode] = React.useState("");
  const [docCode, setdocCode] = React.useState("");
  const [serialNo, setserialNo] = React.useState("");
  const [issueNo, setissueNo] = React.useState("");
  const [revisionNo, setrevisionNo] = React.useState("");
  const [pageValue, setpage] = React.useState("");
  const [TemplateDoc, setTemplateDoc] = React.useState<any>([]);
  const [cancellReason, setcancellReason] = React.useState([{ id: 0, description: "", reason: "" }]);
  const [cancellReasonEdit, setcancellReasonEdit] = React.useState([]);
  const [formData, setFormData] = React.useState({
    RequesterNameId: 0,
    RequesterName: "",
    RequesterDesignation: "",
    DepartmentId: 0,
    TemplateTypeId: 0,
    RequestDate: "",
    IssueDate: "",
    LocationId: 0,
    CustodianId: 0,
    SerialNumber: "",
    IssueNumber: "",
    RevisionNumber: "",
    RevisionDate: "",
    DocumentCode: "",
    ReferenceNumber: "",
    AmendmentTypeId: 0,
    RequestTypeId: 0,
    ClassificationId: 0,
    ChangeRequestTypeId: [],
    SubmiitedDate: "",
    SubmitStatus: "",
    Status: "",
    DocumentName: "",
    IsRework: false,
    DigitalSignStatus: false,
    ChangeRequestID: 0,
    AttachmentId: [],
    AttachmentJson: "",
    DocumentTypeId: 0
  });
  const [Attachmentarr, setAttachmentarr] = React.useState([]);
  const [previewUrl, setPreviewUrl] = React.useState<string | null>(null); // To store the file preview URL
  const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
  const [showModal, setShowModal] = React.useState(false);
  const [ShowModalAtt, setShowModalAtt] = React.useState(false);

  const [forwardToArr, setForwardToArr] = React.useState<ForwardTo[]>([
    { id: 0, role: 0, level: 1, approvers: [], leveltype: "One" } // Default row
  ]);
  const [currentUserDept, setcurrentUserDept] = React.useState("");
  const [forwardToArrEdit, setForwardToArrEdit] = React.useState<ForwardTo[]>([]);
  const [changeRequestCheckboxes, setChangeRequestCheckboxes] = React.useState<ChangeRequestCheckbox[]>([]);
  const [employeeDetails, setemployeeDetails] = React.useState<IEmployeeDetails[] | null>([]);
  const [selectedCheckboxIds, setselectedCheckboxIds] = React.useState<number[]>([]);
  const [isCheckboxSectionHighlighted, setisCheckboxSectionHighlighted] = React.useState<boolean>(false);
  const [fileType, setFileType] = React.useState("");
  const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);
  const [remark, setRemark] = React.useState("");
  const [showview, setshowview] = React.useState(false);
  const [showButton, setShowButton] = React.useState(false);
  // Function to handle People Picker selection
  const onPeoplePickerChange = (items: any[]) => {
    setSelectedUsers(items);
  };

  const ApiCallFunc = async () => {
    const path1 = window.location.href;
    debugger
    locationPath = window.location.href.match(/\/sites\/[^\/]+/)[0];
    if (path1.includes("/view/") || path1.includes("/approve/")) {
      setLoading(true);
      setInputDisabled(true);
      setshowview(true);
    }
    else {
      setInputDisabled(false);
    }
    if (path1.includes("/edit/")) {
      setLoading(true); ////
      setshowview(true);
    }

    console.log("inpt diasba", InputDisabled, path1, path1.includes("/view/"))
    //setLoading(true);
    var ReqTypeArr = await getAllRequestType(sp);
    ReqTypeArr.sort((a, b) => a.RequestType.localeCompare(b.RequestType));
    const optionsreq = ReqTypeArr.map((item: any) => ({
      value: item.ID,
      label: item.RequestType,
      itemId: item.ID
    }));
    setReqType(optionsreq);
    var DepartmentArr = await getAllDepartment(sp);
    DepartmentArr.sort((a, b) => a.Department.localeCompare(b.Department));
    const optionsDepartment = DepartmentArr.map((item: any) => ({
      value: item.ID,
      label: item.Department,
      itemId: item.ID,
      department: item.Department,
      adDepartmentName: item.ADDepartmentName,
      departmentcode: item.DepartmentCode
    }));
    setDepartment(optionsDepartment);
    var TemplateTypeArr = await getAllTemplateType(sp);
    TemplateTypeArr.sort((a, b) => a.TemplateTypeName.localeCompare(b.TemplateTypeName));
    const optionsTemplateType = TemplateTypeArr.map((item: any) => ({
      value: item.ID,
      label: item.TemplateTypeName,
      itemId: item.ID,
      TemplateTypeName: item.TemplateTypeName,
      TemplateTypeValue: item.TemplateTypeValue
    }));
    setTemplateType(optionsTemplateType);
    var AmendmentTypeArr = await getAllAmendmentType(sp);
    AmendmentTypeArr.sort((a, b) => a.AmendmentType.localeCompare(b.AmendmentType));
    const optionsamendment = AmendmentTypeArr.map((item: any) => ({
      value: item.ID,
      label: item.AmendmentType,
      itemId: item.ID
    }));
    setAmendtype(optionsamendment);
    var ClassificationArr = await getAllClassificationMaster(sp);
    ClassificationArr.sort((a, b) => a.Classification.localeCompare(b.Classification));
    const optionsclassification = ClassificationArr.map((item: any) => ({
      value: item.ID,
      label: item.Classification,
      itemId: item.ID
    }));
    setClassificationopt(optionsclassification);
    const Currusers: any = await getCurrentUser(sp, siteUrl);
    setCurrentUser(await getCurrentUser(sp, siteUrl));
    const userProfile = await sp.profiles.myProperties();
    let requesttypesoptions = await fetchChangeRequestTypes();
    let locationoptions = await fetchLocations();
    let custodianoptions = await fetchCustodian();
    let documenttypeoptions = await fetchDocumentTypes();
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
    //const userProfile = await sp.profiles.myProperties();
    setcurrentUserDept(userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "")
    const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
    let currentuserdepartment = UserDept == "IT" ? "Information Technology" : UserDept;
    let optionsfilterdepart = optionsDepartment.filter((user) => user.label === currentuserdepartment);
    setSelectedOptionDepart(optionsDepartment.filter((user) => user.label === currentuserdepartment));
    setRows1(Selectedoptions);
    //Swal.fire('Error', 'Entity is required!', 'error');
    console.log("rerere", optionsDepartment.filter((user) => user.label === currentuserdepartment));
    setFormData(prevData => ({
      ...prevData,
      RequesterNameId: Currusers?.Id || "",
      RequesterDesignation: userProfile?.Title || "",
      RequesterName: userProfile?.DisplayName || "",
      RequestDate: new Date().toLocaleDateString("en-CA"),
      // Department: UserDept
      //RequestedDate: new Date().toISOString().split("T")[0] // Format as YYYY-MM-DD

    }));

    const selectedTemplatefirst = optionsTemplateType.filter((cust: { label: any; }) => cust.label === "Others")[0] || null;
    setSelectedOptionTemplate(selectedTemplatefirst);
    setFormData(prevData => ({
      ...prevData,
      TemplateTypeId: selectedTemplatefirst.value
      // Format as YYYY-MM-DD
    }));
    var DocCodeArr = await getAllDocumentCode(sp);
    let doccodearrew:any;
    const options = DocCodeArr.map((item: any) => ({
      value: item.DocumentCode,
      label: item.DocumentCode,
      IssueNumber: item.IssueNumber,
      ReferenceNumber: item.ReferenceNumber,
      RevisionNumber: item.RevisionNumber,
      ChangeRequestID: item.ID,
      IssueDate: item.IssueDate,
      LocationId: item.LocationId,
      TemplateTypeId: item.TemplateTypeId,
      CustodianId: item.CustodianId,
      SerialNumber: item.SerialNumber,
      RevisionDate: item.RevisionDate,
      AmendmentTypeId: item.AmendmentTypeId,
      DocumentCode: item.DocumentCode,
      ClassificationId: item.ClassificationId,
      ChangeRequestTypeId: item.ChangeRequestTypeId,
      SubmiitedDate: item.SubmiitedDate,
      SubmitStatus: item.SubmitStatus,
      DocumentTypeId: item.DocumentTypeId,
      DepartmentId: item.DepartmentId,
      AttachmentId: item.AttachmentId,
      AttachmentJson: item.AttachmentJson,
      ID: item.ID,
      Status: item.Status
      // DocumentName: "",
      // IsRework: false,
      // DigitalSignStatus: false,


    }));
    setdoccoderows(options);
    console.log("DocCodeArr", DocCodeArr);
    if (optionsfilterdepart.length > 0){
       doccodearrew = options.filter((x: any) => x.DepartmentId == optionsfilterdepart[0].value)
    }
   
    setRows(doccodearrew);



    let formitemid;
    let ProcessItemId: any
    //#region getdataByID
    if (sessionStorage.getItem("ChangeRequestId") != undefined) {
      const iD = sessionStorage.getItem("ChangeRequestId")
      let iDs = decryptId(iD)
      formitemid = Number(iDs);
      setFormItemId(Number(iDs));
      setEditID(await getApprovalByID(sp, Number(iDs), CONTENTTYPE_ChangeDocument));
      ProcessItemId = await getApprovalByID(sp, Number(iDs), CONTENTTYPE_ChangeDocument);
      setInputDisabled(await getApprovalByID2(sp, Number(iDs), CONTENTTYPE_ChangeDocument));
    }
    else {
      const path = window.location.hash;
      const segments = path.split('/').filter(Boolean); // Remove empty elements
      let pathnew = window.location.href;
      debugger

      if (pathnew.includes("/view/") || pathnew.includes("/approve/")) {
        setInputDisabled(true);
      }
      else {
        setInputDisabled(false);
      }
      // Check if "edit" or "view" exists in the URL
      const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");
      console.log("segmentssegments", segments, paramIndex)
      //let segmentsnew = GetQueryString("mode");
      if (paramIndex !== -1 && segments[paramIndex + 1]) {
        setmode(segments[paramIndex])
        // mode = segments[paramIndex]; // Will be "edit" or "view"
        formitemid = segments[paramIndex + 1]; // Get the ID
        setFormItemId(segments[paramIndex + 1]);
        if (segments[paramIndex + 2] !== undefined) {

          setEditID(await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_ChangeDocument));
          ProcessItemId = await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_ChangeDocument);
          setInputDisabled(await getApprovalByID2(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_ChangeDocument));
        }
      }


    }
    // formitemid =20;
    if (formitemid) {
      setEditItemID(Number(formitemid));
      let pathnew1 = window.location.href;
      debugger
      if (pathnew1.includes("/view/") || pathnew1.includes("/approve/")) {
        setInputDisabled(true);
      }
      else {
        setInputDisabled(false);
      }
      const setBannerById = await getItemByIDCR(sp, Number(formitemid))
      // setEditID(Number(setBannerById[0].ID))
      if (setBannerById.length > 0) {
        debugger
        setEditForm(true);
        setMainEditItem(setBannerById[0]);
        debugger
        // setCategoryData(await getCategory(sp, Number(setBannerById[0]?.TypeMaster))) // Category
        if (setBannerById[0].AttachmentId.length > 0) {
          let arrn = await getDocumentLinkByIDarr(sp, setBannerById[0].AttachmentId[0]);
          //let arraynew: any[];
          //arraynew.push(arrn)
          console.log("arrrrrrnh", arrn);
          if (setBannerById[0].Status != "Rework") {
            setAttachmentarr(arrn);
          }

          setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].AttachmentId[0]))
        }

        if (ProcessItemId && ProcessItemId.Level === 0 && ProcessItemId.CurrentUserRole === "OES" && ProcessItemId.IsInitiator == "No") {
          const ApprowData: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_ChangeDocument, setBannerById[0].DocumentCode)

          if (ApprowData.length > 0) {

            const EditApprowData = ApprowData.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id || 0, // Assuming role comes from ApproverRole
              level: item.Level || 1, // Default to 1 if missing
              approvers: item.Approvers?.map((approver: any) => ({
                value: approver.Id,
                label: approver.Title,

              })) || []
            }));

            setForwardToArr(EditApprowData);
            setForwardToArrEdit(EditApprowData);

          }

          // MainListID
        }
        let arr = {

          RequesterName: setBannerById[0].Title,
          //RequesterNameId: setBannerById[0].RequesterNameId,
          RequesterDesignation: setBannerById[0].RequesterDesignation,
          DepartmentId: setBannerById[0].DepartmentId,
          TemplateTypeId: setBannerById[0].TemplateTypeId,
          RequestDate: setBannerById[0].RequestDate,
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
          SubmiitedDate: setBannerById[0].SubmiitedDate,
          SubmitStatus: setBannerById[0].SubmitStatus,
          value: setBannerById[0].DocumentCode,
          label: setBannerById[0].DocumentCode,
          RequestTypeId: setBannerById[0].RequestTypeId,
          Status: setBannerById[0].Status,
          // Status: "Pending",
          // DocumentName: "",
          // IsRework: false,
          // DigitalSignStatus: false,
          //ChangeRequestIDId: setBannerById[0].ID,
          DocumentTypeId: setBannerById[0].DocumentTypeId,
          AttachmentId: setBannerById[0].AttachmentId,
          AttachmentJson: setBannerById[0].AttachmentJson


        }
        if (ProcessItemId && ProcessItemId.CurrentUserRole !== "OES" && ProcessItemId.IsInitiator == "No") {
          const ApprowData1: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_ChangeDocument, setBannerById[0].DocumentCode)

          if (ApprowData1.length > 0) {

            const EditApprowData1 = ApprowData1.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id, // Assuming role comes from ApproverRole
              level: item.Level, // Default to 1 if missing
              approvers: item.Approvers?.map((approver: any) => ({
                value: approver.Id,
                label: approver.Title,

              }))
            }));
            setdisabledforwardarr(true);
            setForwardToArr(EditApprowData1);
            setForwardToArrEdit(EditApprowData1);

          }

          // MainListID
        }
        setissueNo(setBannerById[0].IssueNumber);
        setserialNo(setBannerById[0].SerialNumber);
        setFormData(prevData => ({
          ...prevData,
          IssueNumber: setBannerById[0].IssueNumber,
          ReferenceNumber: setBannerById[0].ReferenceNumber,
          RevisionNumber: setBannerById[0].RevisionNumber,
          ChangeRequestID: setBannerById[0].ID,
          RequestDate: setBannerById[0].RequestDate,
          IssueDate: setBannerById[0].IssueDate,
          LocationId: setBannerById[0].LocationId,
          CustodianId: setBannerById[0].CustodianId,
          SerialNumber: setBannerById[0].SerialNumber,
          RevisionDate: setBannerById[0].RevisionDate,
          AmendmentTypeId: setBannerById[0].AmendmentTypeId,
          ClassificationId: setBannerById[0].ClassificationId,
          ChangeRequestTypeId: setBannerById[0].ChangeRequestTypeId,
          SubmiitedDate: setBannerById[0].SubmiitedDate,
          SubmitStatus: setBannerById[0].SubmitStatus,
          DocumentCode: setBannerById[0].DocumentCode,
          RequestTypeId: setBannerById[0].RequestTypeId,
          DocumentTypeId: setBannerById[0].DocumentTypeId,
          DepartmentId: setBannerById[0].DepartmentId,
          TemplateTypeId: setBannerById[0].TemplateTypeId,
          AttachmentId: setBannerById[0].AttachmentId,
          AttachmentJson: setBannerById[0].AttachmentJson,
          Status: setBannerById[0].Status,
          RequesterName: setBannerById[0].Title,
          //RequesterNameId: setBannerById[0].RequesterNameId,
          RequesterDesignation: setBannerById[0].RequesterDesignation,
          // Format as YYYY-MM-DD
        }));

        // setFormData(arr2);

        debugger
        const selectedDocCode = options.filter((code: { value: any; }) => code.value === setBannerById[0].DocumentCode) || null;
        const selectedLocation = locationoptions.filter((loc: { locationId: any; }) => loc.locationId === setBannerById[0].LocationId)[0] || null;
        const selectedCustodian = custodianoptions.filter((cust: { custodianId: any; }) => cust.custodianId === setBannerById[0].CustodianId)[0] || null;
        const selectedDocumentType = documenttypeoptions.filter((docType: { documentTypeId: any; }) => docType.documentTypeId === setBannerById[0].DocumentTypeId)[0] || null;
        const selectedAmendment = optionsamendment.filter((amend: { value: any; }) => amend.value === setBannerById[0].AmendmentTypeId)[0] || null;
        const selectedClassifiction = optionsclassification.filter((classi: { value: any; }) => classi.value === setBannerById[0].ClassificationId)[0] || null;
        const selectedRequesttype = optionsreq.filter((cust: { value: any; }) => cust.value === setBannerById[0].RequestTypeId)[0] || null;
        const selecteddepart = optionsDepartment.filter((cust: { value: any; }) => cust.value === setBannerById[0].DepartmentId)[0] || null;
        const selectedTemplate = optionsTemplateType.filter((cust: { value: any; }) => cust.value === setBannerById[0].TemplateTypeId)[0] || null;
        setSelectedOption(selectedDocCode.length > 0 && selectedDocCode[0]);
        setselectedOptionDoctype(selectedDocumentType);
        setSelectedOptionClassification(selectedClassifiction);
        setSelectedOptionAmend(selectedAmendment);
        setSelectedOptionDepart(selecteddepart);
        setSelectedOptionTemplate(selectedTemplate);
        setselectedOptionCusto(selectedCustodian);
        setselectedOptionLoc(selectedLocation);
        setselectedCheckboxIds(setBannerById[0].ChangeRequestTypeId);
        setSelectedOptionReq(selectedRequesttype);

        if (setBannerById[0].AttachmentId.length > 0) {
          setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].AttachmentId[0]));
          let arrn = await getDocumentLinkByIDarr(sp, setBannerById[0].AttachmentId[0]);
          //let arraynew: any[];
          //arraynew.push(arrn)
          console.log("arrrrrrnty", arrn);

          if (setBannerById[0].Status != "Rework") {
            setAttachmentarr(arrn);
          }
        }
        else {
          setDocumentLink(null);
        }  // Set the selected users
        setTemplateDoc(await getGeneratedTemplateDocCR(sp, Number(formitemid)));
        const rowData: any[] = await getItemByIDChangeRequest(sp, Number(setBannerById[0].ID)) //baseUrl
        const initialRows = rowData.map((item: any) => ({
          id: item.Id,
          description: item.ChangeDescription,
          reason: item.ReasonforChange,
        }));
        setcancellReason(initialRows.length > 0 ? initialRows : [{ id: 0, description: "", reason: "" }]);
        setcancellReasonEdit(initialRows.length > 0 ? initialRows : [{ id: 0, description: "", reason: "" }]);


      }
      setRequesterRoleId(await getRequesterID(sp))
      setFormNameId(await getFormNameID(sp, CONTENTTYPE_ChangeDocument))
      setListNameId(await getListNameID(sp, LIST_TITLE_ChangeRequest))

    }
    //}
    setLoading(false);
    //#endregion


  };
  const onSelectDocCode = async (selectedList: any) => {
    debugger
    setshowpreviousattachment(true);
    console.log(selectedList, "selectedList");
    if (selectedList != null) {
      setLoading(true);
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
        SubmiitedDate: selectedList.SubmiitedDate,
        SubmitStatus: selectedList.SubmitStatus,
        DocumentCode: selectedList.value,
        DocumentTypeId: selectedList.DocumentTypeId,
        DepartmentId: selectedList.DepartmentId,
        TemplateTypeId: selectedList.TemplateTypeId,
        AttachmentId: selectedList.AttachmentId,
        AttachmentJson: selectedList.AttachmentJson
        // Format as YYYY-MM-DD
      }));
      setSelectedOption(selectedList);
      const rowData: any[] = await getItemByIDChangeRequest(sp, Number(selectedList.ID)) //baseUrl
      const initialRows = rowData.map((item: any) => ({
        id: item.Id,
        description: item.ChangeDescription,
        reason: item.ReasonforChange,
      }));
      // setcancellReason(initialRows);
      // setcancellReasonEdit(initialRows);
      //setcancellReason([{ id: 0, description: "", reason: "" }]);
      //setcancellReasonEdit([{ id: 0, description: "", reason: "" }]);
      const selectedLocation = LocationOpt.filter((loc: { locationId: any; }) => loc.locationId === selectedList.LocationId)[0] || null;
      const selectedCustodian = Custodianopt.filter((cust: { custodianId: any; }) => cust.custodianId === selectedList.CustodianId)[0] || null;
      const selectedDocumentType = DocumentTypeOpt.filter((docType: { documentTypeId: any; }) => docType.documentTypeId === selectedList.DocumentTypeId)[0] || null;
      const selectedAmendment = Amendtype.filter((cust: { value: any; }) => cust.value === selectedList.AmendmentTypeId)[0] || null;
      const selectedClassifiction = Classificationopt.filter((docType: { value: any; }) => docType.value === selectedList.ClassificationId)[0] || null;
      //const selectedRequesttype = optionsreq.filter((cust: { value: any; }) => cust.value === selectedList.RequestTypeId)[0] || null;
      const selecteddepart = Departopt.filter((cust: { value: any; }) => cust.value === selectedList.DepartmentId)[0] || null;
      const selectedTemplate = TemplateTypeopt.filter((cust: { value: any; }) => cust.value === selectedList.TemplateTypeId)[0] || null;
      setselectedOptionDoctype(selectedDocumentType);
      setSelectedOptionClassification(selectedClassifiction);
      setSelectedOptionAmend(selectedAmendment);
      setSelectedOptionDepart(selecteddepart);
      setSelectedOptionTemplate(selectedTemplate);
      setselectedOptionCusto(selectedCustodian);
      setselectedOptionLoc(selectedLocation);
      //setselectedCheckboxIds(selectedList.ChangeRequestTypeId);
      //setselectedCheckboxIds([])
      if (selectedList.AttachmentId.length > 0) {
        let arrn = await getDocumentLinkByIDarr(sp, selectedList.AttachmentId[0]);
        //let arraynew: any[];
        //arraynew.push(arrn)
        console.log("arrrrrrn56", arrn);
        //setAttachmentarr(arrn);
        setDocumentLink(await getDocumentLinkByID(sp, selectedList.AttachmentId[0]))
      }
      else {
        setDocumentLink(null);
      }  // Set the selected users
      setLoading(false);
    };
  }

  const scrollToTop = () => {
    if (elementRef1.current) {
      elementRef1.current.scrollTo({
        top: 0,
        behavior: 'smooth', // Smooth scroll to top
      });
    }
  };
  const onSelectReq = (selectedList: any) => {
    setFormData(prevData => ({
      ...prevData,
      SerialNumber: "",
      IssueNumber: "",
      RevisionNumber: "",
      RevisionDate: "",
      DocumentCode: "",
      ReferenceNumber: "",
    }));
    //setselectedCheckboxIds([]);
    if (selectedList?.label != "Change Request for New Addition") {
      setselectedOptionDoctype(null);
      setselectedOptionCusto(null);
      setselectedOptionLoc(null);
    }

    //setSelectedOptionClassification(null);
    //setSelectedOptionAmend(null);
    //setSelectedOptionDepart(null);

    //setSelectedOptionReq(null);
    setSelectedOption(null);
    //setAttachmentarr([]);
    //setcancellReason([{ id: 0, description: "", reason: "" }]);
    console.log(selectedList, "selectedListreq");
    setFormData(prevData => ({
      ...prevData,
      RequestTypeId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionReq(selectedList);  // Set the selected users
  };
  const onSelectAmend = (selectedList: any) => {
    console.log(selectedList, "selectedListamenddd");
    setFormData(prevData => ({
      ...prevData,
      AmendmentTypeId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionAmend(selectedList);  // Set the selected users
  };
  const onSelectDepart = (selectedList: any) => {
    console.log(selectedList, "selectedListadepartttt");
    setFormData(prevData => ({
      ...prevData,
      DepartmentId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    let doccodearrew:any;
    if (selectedList) {
      doccodearrew = doccoderows.filter((x: any) => x.DepartmentId == selectedList.value)
    }

    setRows(doccodearrew);

    setSelectedOptionDepart(selectedList);  // Set the selected users
  };
  const onSelectTemplatetype = (selectedList: any) => {
    console.log(selectedList, "selectedListadtemplatetype");
    setFormData(prevData => ({
      ...prevData,
      TemplateTypeId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionTemplate(selectedList);  // Set the selected users
  };
  const onSelectClassification = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      ClassificationId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionClassification(selectedList);  // Set the selected users
  };

  const onSelectLocation = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      LocationId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setselectedOptionLoc(selectedList);  // Set the selected users
  };

  const onSelectCustodian = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      CustodianId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setselectedOptionCusto(selectedList);  // Set the selected users
  };

  const onSelectDocumentType = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      DocumentTypeId: selectedList.value
      // Format as YYYY-MM-DD
    }));
    setselectedOptionDoctype(selectedList);  // Set the selected users
  };
  const onSelectApprovers = (selectedOptions: any, id: number) => {
    setForwardToArr((prev) =>
      prev.map((row) =>
        row.level === id ? { ...row, approvers: selectedOptions || [] } : row
      )
    );
  };

  const fetchLocations = async () => {
    try {
      // Fetch the items
      const items = await sp.web.lists
        .getByTitle("LocationMaster") // Your list name
        .items.filter("IsActive eq 'Yes'") // Filter active items
        .select("ID", "Location", "LocationCode") // Select required fields
        .top(5000) // Limit number of records
        (); // Call get() to fetch data
      items.sort((a, b) => a.Location.localeCompare(b.Location));
      // Use map on the result to create the desired array structure
      const locations: Location[] = items.map((item: any) => ({
        locationId: item.ID, // Store the ID for lookup
        locationName: item.Location, // Name of the location
        locationCode: item.LocationCode, // Code of the location
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

  const fetchCustodian = async () => {
    // debugger
    try {
      const items = await sp.web.lists
        .getByTitle("CustodianMaster")
        .items.filter("IsActive eq 'Yes'")
        .select("ID", "Custodian", "CustodianCode")
        .top(5000) // Limit number of records
        ();
      items.sort((a, b) => a.Custodian.localeCompare(b.Custodian));
      const custodians: Custodian[] = items.map((item: any) => ({
        custodianId: item.ID, // Store the Custodian ID for lookup
        custodianName: item.Custodian,
        custodianCode: item.CustodianCode,
        label: item.Custodian,
        value: item.ID
      }));

      setCustodianopt(custodians);
      return custodians;
    } catch (error) {
      console.error("Error fetching Custodian: ", error);
    }
  };

  const fetchDocumentTypes = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle("DocumentTypeMaster")
        .items.select("ID", "DocumentType", "DocumentTypeCode")
        .top(5000)();
      items.sort((a, b) => a.DocumentType.localeCompare(b.DocumentType));
      const documentTypes: DocumentType[] = items.map((item: any) => ({
        documentTypeId: item.ID, // Store the Document Type ID for lookup
        documentTypeName: item.DocumentType,
        documentTypeCode: item.DocumentTypeCode,
        label: item.DocumentType,
        value: item.ID
      }));

      setDocumentTypeOpt(documentTypes);
      return documentTypes;
    } catch (error) {
      console.error("Error fetching Document Types: ", error);
    }
  };


  // const onSelectRole = (selectedList: any) => {
  //     // console.log(selectedList , "selectedList");
  //     setSelectedRole(selectedList);  // Set the selected users
  // };

  const onSelectRole = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
    debugger
    const updatedArr = forwardToArr.map(row =>
      row.level === lvl ? { ...row, role: Number(event.target.value) } : row
    );
    setForwardToArr(updatedArr);
    setSelectedRole(forwardToArr.map(r => r.role).filter(role => role))
  };

  const onSelectApprovalType = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
    debugger
    const updatedArr1 = forwardToArr.map(row =>
      row.level === lvl ? { ...row, leveltype: event.target.value } : row
    );
    setForwardToArr(updatedArr1);
    //setUserRoles(UserRoles.filter((x: any) => x.label !== event.target.value))
  };

  // const handleAddRow = () => {
  //     const newRow: ForwardTo = { id: Date.now(), role: 0, level: "", approvers: [] };
  //     setForwardToArr([...forwardToArr, newRow]);
  // };
  const handleAddRow = () => {
    setForwardToArr((prev) => [
      ...prev,
      { id: 0, role: 0, level: prev.length + 1, approvers: [], leveltype: "One" }
    ]);
  };

  const handleDeleteRow = (index: number) => {
    // const updatedRows = forwardToArr.filter((_, i) => i !== index);
    // setForwardToArr(updatedRows);
    const updatedRows = forwardToArr
      .filter((_, i) => i !== index) // Remove selected row
      .map((row, newIndex) => ({ ...row, level: newIndex + 1 })); // Reassign levels

    setForwardToArr([...updatedRows]); // Ensure a new array reference
    //forwardToArr.splice(index, 1);
    //setForwardToArr(forwardToArr);
  };

  React.useEffect(() => {
    const handleScroll = () => {
      console.log("elementRef.current", elementRef.current)
      // If the user has scrolled down 20px or more, show the button
      if (elementRef1.current) {
        // Check if scrolled down more than 20px
        if (elementRef1.current.scrollTop > 20) {
          setShowButton(true);
        } else {
          setShowButton(false);
        }
      }
    };

    // Add scroll event listener
    window.addEventListener('scroll', handleScroll);

    // Clean up the event listener on component unmount
    return () => {
      window.removeEventListener('scroll', handleScroll);
    };
  }, []);

  React.useEffect(() => {
    ApiCallFunc();
    const path = window.location.href;
    if (path.includes("/view/") || path.includes("/approve/")) {
      setInputDisabled(true);
    }
    else {
      setInputDisabled(false);
    }
    // formData.title = currentUser.Title;
  }, [useHide]);

  const handleCancel = () => {
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

  //#endregion
  const OpenFile = (obj: any, sts: string) => {
    debugger
    console.log("ttrtrtrtt", obj)
    const fileUrl = `${Tenant_URL}${obj?.FileRef != "" ? obj.FileRef : obj.fileUrl}`;
    if (sts == "Open") {
      if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {

        window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=default`, "_blank");
      } else {
        window.open(fileUrl, "_blank"); // Open PDF and other files normally
      }

    } else if (sts == "Download") {
      const link = document.createElement("a");
      link.href = fileUrl;
      link.setAttribute("download", obj?.FileLeafRef != "" ? obj.FileLeafRef : obj.name); // Suggests a filename for download
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

    }
    // if (obj.FileRef.endsWith(".docx")) {
    //     window.open(`https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(fileUrl)}`, "_blank");
    //   } else if (obj.FileRef.endsWith(".xlsx")) {
    //     window.open(`https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(fileUrl)}`, "_blank");
    //   } else {
    //     window.open(fileUrl, "_blank"); // Open PDF and other files normally
    //   }

    // if (fileUrl.endsWith(".docx") || fileUrl.endsWith(".xlsx") || fileUrl.endsWith(".pptx")) {
    //   //window.open(`${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=default`, "_blank");
    //   window.open(obj, "_blank");
    // } else {
    //   window.open(fileUrl, "_blank"); // Open PDF and other files normally
    // }
  }
  const ApprovalTypeOptions = [
    { value: 'One', label: 'Anyone' },
    { value: 'All', label: 'Everyone' }

  ];
  const addCancelReason = () => {
    setcancellReason([...cancellReason, { id: 0, description: "", reason: "" }]);
  };
  const validateForm = async (fmode: FormSubmissionMode) => {
    debugger
    const { RequesterName, RequesterDesignation, RequestDate, DocumentCode, IssueNumber, RevisionNumber, ReferenceNumber } = formData;
    // const { description } = richTextValues;
    let valid = true;
    let valid1 = true;
    // let validateOverview:boolean = false;
    // let validatetitlelength = false;
    // let validateTitle = false;
    setlocationerr(false);
    setcustodianerr(false);
    setdocumenttypeerr(false);
    setclassificationerr(false);
    setrequesttypeerr(false);
    setdocumentcodeerr(false);
    setamendmenterr(false);
    setdepartmenterr(false);
    settemplatetypeerr(false);
    setattachmenterr(false);
    setchangedescriptionerr(false);
    setchangereasonerr(false);
    setchangerequesttypeerr(false);
    setValidDraft(true);
    setValidSubmit(true);
    setValidCancelReason(true);
    let errormsg = "";

    if (fmode == FormSubmissionMode.SUBMIT) {
      if (!RequesterName) {
        valid = false;
      }
      if (!selectedOptionLoc) {
        setlocationerr(true);
        valid = false;
      }
      if (!selectedOptionCusto) {
        setcustodianerr(true);
        valid = false;
      }
      if (!selectedOptionDoctype) {
        setdocumenttypeerr(true);
        valid = false;
      }
      if (!selectedOptionClass) {
        setclassificationerr(true);
        valid = false;
      }
      if (!selectedOptionReq) {
        setrequesttypeerr(true);
        valid = false;
      }
      if (selectedOptionReq && selectedOptionReq.label != "Change Request for New Addition" && !selectedOption) {
        setdocumentcodeerr(true);
        valid = false;
      }
      if (!selectedOptionAmend) {
        setamendmenterr(true);
        valid = false;
      }
      if (!SelectedOptionDepart) {
        setdepartmenterr(true);
        valid = false;
      }
      if (!SelectedOptionTemplate) {
        settemplatetypeerr(true);
        valid = false;
      }
      if (cancellReason.length > 0) {
        let descriptionError = false;
        let reasonError = false;

        cancellReason.forEach((row: any) => {
          if (row.description === null || row.reason === null) {
            if (row.description === null) {
              descriptionError = true;
            }
            if (row.reason === null) {
              reasonError = true;
            }
          } else {
            if (row.description != null && row.description.trim() === "") {
              descriptionError = true;
            }
            if (row.reason != null && row.reason.trim() === "") {
              reasonError = true;
            }
          }

        });

        // If any description or reason is blank, set the respective error flags to true
        if (descriptionError) {
          setchangedescriptionerr(true);
        }
        if (reasonError) {
          setchangereasonerr(true);
        }
        if (descriptionError || reasonError) {
          valid1 = false;
        }
      }
      // if (cancellReason.length > 0 && cancellReason.every((row: any) => row.description.trim() !== "" || row.reason.trim() !== "") == false) {
      //   setchangedescriptionerr(true);
      //   setchangereasonerr(true);
      //   valid1 = false;
      // }
      if (cancellReason.length == 0) {
        setchangedescriptionerr(true);
        setchangereasonerr(true);
        valid1 = false;
      }
      if (selectedCheckboxIds.length == 0) {
        setchangerequesttypeerr(true);
        valid = false;
      }
      if (Attachmentarr.length == 0 && DocumentLink == null) {
        setattachmenterr(true);
        valid = false;
      }

      // return true;

      setValidSubmit(valid);
      setValidCancelReason(valid1);
    }
    else {
      if (!RequesterName) {
        valid = false;
      }
      if (!selectedOptionReq) {
        setrequesttypeerr(true);
        valid = false;
      }
      if (selectedOptionReq && selectedOptionReq.label != "Change Request for New Addition" && !selectedOption) {
        setdocumentcodeerr(true);
        valid = false;
      }
      if (!SelectedOptionDepart) {
        setdepartmenterr(true);
        valid = false;
      }
      // if (cancellReason.length > 0) {
      //   let descriptionError = false;
      //   let reasonError = false;

      //   cancellReason.forEach((row: any) => {
      //     if (row.description.trim() === "") {
      //       descriptionError = true;
      //     }
      //     if (row.reason.trim() === "") {
      //       reasonError = true;
      //     }
      //   });

      //   // If any description or reason is blank, set the respective error flags to true
      //   if (descriptionError) {
      //     setchangedescriptionerr(true);
      //   }
      //   if (reasonError) {
      //     setchangereasonerr(true);
      //   }
      //   if (descriptionError || reasonError) {
      //     valid1 = false;
      //   }
      // }
      // if (cancellReason.length > 0 && cancellReason.every((row: any) => row.description.trim() !== "" || row.reason.trim() !== "") == false) {
      //   setchangedescriptionerr(true);
      //   setchangereasonerr(true);
      //   valid1 = false;
      // }
      setValidDraft(valid);
      setValidCancelReason(valid1);
    }

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
  };
  //#region  Submit Form
  const handleFormSubmit = async () => {
    debugger
    scrollToTop();
    let url = window.location.href.split('/sites/')[0];
    console.log("topp submit", editItemID, cancellReason);
    if (await validateForm(FormSubmissionMode.SUBMIT)) {
      debugger
      let serialnumber = await getDocumentCodeselected(sp, selectedOptionLoc.locationId, selectedOptionCusto.custodianId, selectedOptionDoctype.documentTypeId)
      let issueno = "";
      let serialno = "";
      let revisionno = "";
      if (selectedOptionReq.label == "Change Request for New Addition") {
        if (serialnumber.length > 0) {
          issueno = serialnumber[0].IssueNo;
          serialno = (Number(serialnumber[0].SerialNo) + 1).toString();
          revisionno = serialnumber[0].RevisionNo;
          setissueNo(serialnumber[0].IssueNo);
          setserialNo(serialno);
          setrevisionNo(serialnumber[0].RevisionNo);
        } else {
          issueno = issueNo == "" || issueNo == null ? "01" : issueNo;
          serialno = serialNo == "" || issueNo == null ? "01" : serialNo;
          revisionno = revisionNo == "" || issueNo == null ? "00" : revisionNo;
          setissueNo(issueNo == "" || issueNo == null ? "01" : issueNo);
          setserialNo(serialNo == "" || issueNo == null ? "01" : serialNo);
        }
      } else {
        issueno = (Number(selectedOption?.IssueNumber) + 1).toString();
        serialno = selectedOption?.SerialNumber;
        revisionno = (Number(selectedOption?.RevisionNumber) + 1).toString();
        setissueNo(issueno);
        setserialNo(serialno);
        setrevisionNo(revisionno);
      }
      let doccode = selectedOptionReq.label == "Change Request for New Addition" ? await generateDocCode(serialno) : selectedOption?.DocumentCode;
      let referencecode = await generateReferenceCode(serialno, issueno);
      console.log("doccode doccode", doccode, referencecode);
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
            let bannerImageArray: any = {};
            let DocumentName: string = "";
            let attachmentIds = [];
            debugger
            const folder = sp.web.getFolderByServerRelativePath('/sites/EDeDMS/ChangeRequestDocs');
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                DocumentName = file.name;
                const fileAddResult = await folder.files.addChunked(file.name, file);
                const fileNew = fileAddResult.file;
                const documentName = fileAddResult.data.Name;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: documentName, // Assuming FileName is the internal name of the column
                  DocumentCode: selectedOptionReq.label == "Change Request for New Addition" ? doccode : selectedOption?.DocumentCode,
                });

                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            }
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            let arr = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              //IssueDate: new Date(formData.IssueDate).toISOString(),
              LocationId: formData.LocationId,
              CustodianId: formData.CustodianId,
              SerialNumber: Number(serialno),
              IssueNumber: Number(issueno),
              RevisionNumber: Number(revisionno),
              //RevisionDate: formData.RevisionDate,
              DocumentCode: selectedOptionReq.label == "Change Request for New Addition" ? doccode : selectedOption?.DocumentCode,
              ReferenceNumber: referencecode,
              RequestTypeId: formData.RequestTypeId,
              AmendmentTypeId: formData.AmendmentTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              ClassificationId: formData.ClassificationId,
              SubmiitedDate: new Date(formData.RequestDate).toISOString(),
              SubmitStatus: "Yes",
              Status: "Pending",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              DocumentName: attachmentIds.length != 0 ? DocumentName : formData.DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              //AttachmentId: selectedOptionReq.label == "Change Request for New Addition" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "Change Request for New Addition" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso
            }
            const postResult = await updateItemChangeRequestList(arr, sp, editItemID);
            const postId = postResult?.data?.ID;
            console.log("postPayload edit arr", arr, postResult);
            debugger
            for (const row of cancellReason) {

              const postPayload2 = {
                ChangeRequestIDId: editItemID, // Assuming "Title" column exists
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              }

              if (!row.id || row.id == 0) {

                const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                const postId2 = postResult2?.data?.ID;
                // debugger
                if (!postId2) {
                  console.error("Post creation failed.");
                  return;
                }

              }
              else if (row.id > 0) {
                const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                const postId2 = postResult2?.data?.ID;
              }

            }

            const toDelete = cancellReasonEdit.filter(
              (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
            );

            // Delete each item from SharePoint
            for (const item of toDelete) {
              try {
                await sp.web.lists.getByTitle("ChangeRequestReasonList").items.getById(item.id).delete();
                // console.log(`Deleted item with ID: ${item.ID}`);
              } catch (error) {
                console.error(`Error deleting item with ID: ${item.id}`, error);
              }
            }

            // ///////************* */
            let boolval = false;

            // if (boolval == true) {
            setLoading(false);
            sessionStorage.removeItem("ChangeRequestId")
            Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
              }
            });
            //sessionStorage.removeItem("ChangeRequestId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
            // }, 1000);
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
            let bannerImageArray: any = {};
            let DocumentName: string = "";
            let attachmentIds = [];
            debugger
            const folder = sp.web.getFolderByServerRelativePath('/sites/EDeDMS/ChangeRequestDocs');
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                DocumentName = file.name;
                const fileAddResult = await folder.files.addChunked(file.name, file);
                const fileNew = fileAddResult.file;
                const documentName = fileAddResult.data.Name;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: documentName, // Assuming FileName is the internal name of the column
                  DocumentCode: selectedOptionReq.label == "Change Request for New Addition" ? doccode : selectedOption?.DocumentCode,
                });

                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            }
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            const postPayload = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: formData.RequestDate != "" ? new Date(formData.RequestDate).toISOString() : new Date().toISOString(),
              //IssueDate: formData.IssueDate,
              LocationId: formData.LocationId,
              CustodianId: formData.CustodianId,
              SerialNumber: Number(serialno),
              IssueNumber: Number(issueno),
              RevisionNumber: Number(revisionno),
              //RevisionNumber: selectedOption?.RevisionNumber,
              //RevisionDate: new Date().toISOString(),
              DocumentCode: selectedOptionReq.label == "Change Request for New Addition" ? doccode : selectedOption?.DocumentCode,
              ReferenceNumber: referencecode,
              AmendmentTypeId: formData.AmendmentTypeId,
              RequestTypeId: formData.RequestTypeId,
              ClassificationId: formData.ClassificationId,
              ChangeRequestTypeId: selectedCheckboxIds,
              //SubmiitedDate: selectedOption?.SubmiitedDate,
              SubmiitedDate: new Date().toISOString(),
              SubmitStatus: "Yes",
              Status: "Pending",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              DocumentName: DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              //AttachmentId: selectedOptionReq.label == "Change Request for New Addition" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "Change Request for New Addition" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso
            };
            console.log("postPayload new", postPayload);

            const postResult = await addItemChangeRequestList(postPayload, sp);
            const postId = postResult?.data?.ID;
            // debugger
            if (!postId) {
              console.error("Post creation failed.");
              return;
            }

            for (const row of cancellReason) {

              const postPayload2 = {
                ChangeRequestIDId: postId, // Assuming "Title" column exists
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              }

              const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
              const postId2 = postResult2?.data?.ID;
              // debugger
              if (!postId2) {
                console.error("Post creation failed.");
                return;
              }
            }
            let boolval;

            // if (boolval == true) {
            setLoading(false);
            sessionStorage.removeItem("ChangeRequestId")
            Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
              }
            });
            // Swal.fire('Submitted successfully.', '', 'success');
            // // sessionStorage.removeItem("bannerId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
            // }, 1000);
            // }

          }
        })

      }
    }

  }
  const generateDocCode = async (serialno: any) => {
    debugger
    //const { seconsole.log("iiii",issueno,issueNo,serialno,serialNo)lectedLocationId, selectedCustodianId, selectedDocumentTypeId, locations, custodians, documentTypes } = this.state;
    console.log("doccode", serialno, serialNo)
    // Use filter to find the selected items
    const selectedLocation = LocationOpt.filter((loc: { locationId: any; }) => loc.locationId === selectedOptionLoc.locationId)[0] || null;
    const selectedCustodian = Custodianopt.filter((cust: { custodianId: any; }) => cust.custodianId === selectedOptionCusto.custodianId)[0] || null;
    const selectedDocumentType = DocumentTypeOpt.filter((docType: { documentTypeId: any; }) => docType.documentTypeId === selectedOptionDoctype.documentTypeId)[0] || null;

    // If all selections are available, generate the document code
    if (selectedLocation && selectedCustodian && selectedDocumentType) {
      const docCode = `${selectedDocumentType.documentTypeCode}.${selectedLocation.locationCode}.${selectedCustodian.custodianCode}.${serialno.toString().padStart(2, '0')}`;
      setdocCode(docCode)
      return docCode;
      // Store docCode in state
    } else {
      // Reset docCode if any selection is missing
      setdocCode("")
      return;
    }
  };
  const generateReferenceCode = async (serialno: any, issueno: any) => {
    debugger
    console.log("iiii", issueno, issueNo, serialno, serialNo)
    //const { selectedLocationId, selectedCustodianId, selectedDocumentTypeId, locations, custodians, documentTypes } = this.state;
    const selectedLocation = LocationOpt.filter((loc: { locationId: any; }) => loc.locationId === selectedOptionLoc.locationId)[0] || null;
    const selectedCustodian = Custodianopt.filter((cust: { custodianId: any; }) => cust.custodianId === selectedOptionCusto.custodianId)[0] || null;
    const selectedDocumentType = DocumentTypeOpt.filter((docType: { documentTypeId: any; }) => docType.documentTypeId === selectedOptionDoctype.documentTypeId)[0] || null;
    if (selectedLocation && selectedCustodian && selectedDocumentType) {
      const referencedocCode = `${selectedDocumentType.documentTypeCode}.${selectedLocation.locationCode}.${selectedCustodian.custodianCode}.TMP-${serialno.toString().padStart(2, '0')}.${issueno.toString().padStart(2, '0')}`;
      setreferencedocCode(referencedocCode);
      return referencedocCode;
      // Store docCode in state
    } else {
      // Reset docCode if any selection is missing.TMP-01.01

      setreferencedocCode("")
      return;
    }

  }
  const handleSaveAsDraft = async () => {
    debugger
    scrollToTop();
    let url = window.location.href.split('/sites/')[0];
    console.log("topp draft", editItemID, cancellReason, selectedOption);
    let currentReferenceNo = selectedOption?.ReferenceNumber;
    let arrrr = currentReferenceNo && currentReferenceNo.split('.')
    if (selectedOption != null) {
      for (let i = 0; i < arrrr && arrrr.length; i++) {
        if (arrrr[i].includes("RRF")) {
          arrrr[i] = arrrr[i].replace("RRF", "TMP");
        }
      }
    }

    console.log("props.currentItem", props.currentItem);
    let test = arrrr && arrrr.join('.');
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
            let bannerImageArray: any = {};
            let DocumentName: string = "";
            let attachmentIds = [];
            debugger
            const folder = sp.web.getFolderByServerRelativePath('/sites/EDeDMS/ChangeRequestDocs');
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                DocumentName = file.name;
                const fileAddResult = await folder.files.addChunked(file.name, file);
                const fileNew = fileAddResult.file;
                const documentName = fileAddResult.data.Name;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: documentName, // Assuming FileName is the internal name of the column
                });

                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            }
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            let arr = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              LocationId: formData.LocationId,
              CustodianId: formData.CustodianId,
              SerialNumber: selectedOptionReq.label == "Change in Existing Content" && selectedOption ? Number(selectedOption?.SerialNumber) : null,
              IssueNumber: selectedOptionReq.label == "Change in Existing Content" && selectedOption ? Number(selectedOption?.IssueNumber) : null,
              RevisionNumber: selectedOptionReq.label == "Change in Existing Content" && selectedOption ? Number(selectedOption?.RevisionNumber) : null,
              DocumentCode: selectedOptionReq.label == "Change Request for New Addition" ? "" : selectedOption && selectedOption?.DocumentCode,
              ReferenceNumber: selectedOptionReq.label == "Change in Existing Content" ? test : "",
              AmendmentTypeId: formData.AmendmentTypeId,
              RequestTypeId: formData.RequestTypeId,
              ClassificationId: formData.ClassificationId,
              //ChangeRequestTypeId: formData.ChangeRequestTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              //SubmiitedDate: selectedOption?.SubmiitedDate,
              SubmiitedDate: new Date(formData.SubmiitedDate).toISOString(),
              SubmitStatus: "No",
              Status: "Save as draft",
              DocumentName: DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "No",
              CurrentUserRole: "OES",
              //AttachmentId: selectedOptionReq.label == "Change Request for New Addition" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "Change Request for New Addition" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso

            }
            let descriptionError = false;
            let reasonError = false;
            console.log("postPayloaddrafttedit", arr, editItemID, cancellReason);
            const postResult = await updateItemChangeRequestList(arr, sp, editItemID);
            const postId = postResult?.data?.ID;
            for (const row of cancellReason) {
              if (row.description.trim() === "") {
                descriptionError = true;
              }
              if (row.reason.trim() === "") {
                reasonError = true;
              }
              const postPayload2 = {
                ChangeRequestIDId: editItemID, // Assuming "Title" column exists
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              }
              if (!descriptionError || !reasonError) {
                if (!row.id) {
                  const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                  const postId2 = postResult2?.data?.ID;
                  // debugger
                  if (!postId2) {
                    console.error("Post creation failed.");
                    return;
                  }

                }
                else if (row.id > 0) {
                  const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                  const postId2 = postResult2?.data?.ID;
                }
              }


            }
            // await AddContentMaster(sp, arr)

            // const boolval = await handleClick(editID, TypeMasterData?.TypeMaster, Number(formData.entity))
            // /*********** */

            // Find items that are in cancellReasonEdit but NOT in cancellReason
            const toDelete = cancellReasonEdit.filter(
              (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
            );

            // Delete each item from SharePoint
            for (const item of toDelete) {
              try {
                await sp.web.lists.getByTitle("ChangeRequestReasonList").items.getById(item.id).delete();
                // console.log(`Deleted item with ID: ${item.ID}`);
              } catch (error) {
                console.error(`Error deleting item with ID: ${item.id}`, error);
              }
            }

            // ///////************* */
            let boolval = false;

            // if (boolval == true) {
            setLoading(false);
            sessionStorage.removeItem("ChangeRequestId")
            Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
              }
            });
            // Swal.fire('Saved successfully.', '', 'success');
            // sessionStorage.removeItem("ChangeRequestId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
            // }, 2000);
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
            let bannerImageArray: any = {};
            let DocumentName: string = "";
            let attachmentIds = [];
            const folder = sp.web.getFolderByServerRelativePath('/sites/EDeDMS/ChangeRequestDocs');
            debugger
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                DocumentName = file.name;
                const fileAddResult = await folder.files.addChunked(file.name, file);
                const fileNew = fileAddResult.file;
                const documentName = fileAddResult.data.Name;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: documentName, // Assuming FileName is the internal name of the column
                });
                console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            }

            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            const postPayload = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              IssueDate: new Date().toISOString(),
              LocationId: formData.LocationId,
              CustodianId: formData.CustodianId,
              SerialNumber: selectedOptionReq.label == "Change in Existing Content" && selectedOption ? Number(selectedOption?.SerialNumber) : null,
              IssueNumber: selectedOptionReq.label == "Change in Existing Content" && selectedOption ? Number(selectedOption?.IssueNumber) : null,
              RevisionNumber: selectedOptionReq.label == "Change in Existing Content" && selectedOption ? Number(selectedOption?.RevisionNumber) : null,
              DocumentCode: selectedOptionReq.label == "Change Request for New Addition" ? "" : selectedOption && selectedOption?.DocumentCode,
              ReferenceNumber: selectedOptionReq.label == "Change in Existing Content" ? test : "",
              RequestTypeId: formData.RequestTypeId,
              AmendmentTypeId: formData.AmendmentTypeId,
              ClassificationId: formData.ClassificationId,
              ChangeRequestTypeId: selectedCheckboxIds,
              SubmiitedDate: new Date().toISOString(),
              SubmitStatus: "No",
              Status: "Save as draft",
              DocumentName: DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              //AttachmentId: selectedOptionReq.label == "Change Request for New Addition" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "Change Request for New Addition" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso

            };
            console.log("postPayloaddraftttt", postPayload);

            const postResult = await addItemChangeRequestList(postPayload, sp);
            const postId = postResult?.data?.ID;
            // debugger
            if (!postId) {
              console.error("Post creation failed.");
              return;
            }
            let descriptionError = false;
            let reasonError = false;

            for (const row of cancellReason) {
              if (row.description.trim() === "") {
                descriptionError = true;
              }
              if (row.reason.trim() === "") {
                reasonError = true;
              }
              const postPayload2 = {
                ChangeRequestIDId: postId, // Assuming "Title" column exists
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              }
              if (!descriptionError || !reasonError) {
                const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                const postId2 = postResult2?.data?.ID;
                // debugger
                if (!postId2) {
                  console.error("Post creation failed.");
                  return;
                }
              }
            }


            setLoading(false);
            sessionStorage.removeItem("ChangeRequestId")
            Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
              }
            });
            // Swal.fire('Saved successfully.', '', 'success');
            // // sessionStorage.removeItem("bannerId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
            // }, 1000);
          }
        })

      }
    }

  }
  const GetUserProperties = async () => {
    // let employeeDetails: IEmployeeDetails = {
    //   userId: 0, // Initialize with default value
    // };
    let employeeDetails: any = {
      userId: 0,
      preffedName: '',
      deptName: '',
      office: '',
      Title: ''
    };
    debugger;
    try {
      // Fetch the current user's profile details
      const userProfile = await sp.profiles.myProperties();

      // Fetch the current user's ID
      const currentUser = await sp.web.currentUser();

      // Add the user ID to the employeeDetails object
      employeeDetails.userId = currentUser.Id;

      // Assuming the userProfile has the UserProfileProperties you need
      const userProperties = userProfile.UserProfileProperties;

      // Loop through properties and update the employeeDetails object
      for (let property of userProperties) {
        if (property.Key === "PreferredName") {
          employeeDetails.preffedName = property.Value;
        }
        if (property.Key === "Department") {
          employeeDetails.deptName = property.Value;
        }
        if (property.Key === "Office") {
          employeeDetails.office = property.Value;
        }
        if (property.Key === "Title") {
          employeeDetails.Title = property.Value;
        }
      }
      //   this.setState({
      //     name: employeeDetails.preffedName,
      //     designation: employeeDetails.Title,
      // });

      setemployeeDetails(employeeDetails)

    } catch (error) {
      console.error("Error fetching user properties:", error);

    }
  }
  const ForwardApproval = (status: string) => {
    scrollToTop();
    let url = window.location.href.split('/sites/')[0];
    console.log("topp draf fprt", editItemID, cancellReason, url);
    let valid = true;
    let actionMessage = "";
    let successMessage = "";
    switch (status) {
      case "Forward":
        actionMessage = "Do you want to forward this request?";
        successMessage = "Request forwarded successfully.";
        break;
      case "Rejected":
        actionMessage = "Do you want to reject this request?";
        successMessage = "Request rejected successfully.";
        break;
      case "Rework":
        actionMessage = "Do you want to send this request for rework?";
        successMessage = "Request sent for rework successfully.";
        break;
    }
    if (status == "Forward") {
      if (forwardToArr.length === 0) {
        // alert("At least Anyone row is required.");
        valid = false;
      }
      debugger
      const isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 && (row.leveltype == "One" || row.leveltype == "All"));

      if (!isValid) {
        // alert("Each row must have a role selected and at least Anyone approver.");
        valid = false;
      }
      setValidforward(valid)
      if (!valid) {
        Swal.fire('Please fill all the mandatory fields.');
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
            debugger
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))

            let arr = {
              // ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toLocaleDateString("en-CA"),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: "Approved",
              Remark: remark,


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
                ContentTitle: formData.ReferenceNumber,
                MainListNameId: ListNameId,
                ApproverRoleId: item.role,
                Level: Number(item.level),
                ApproversId: approversIds,
                LevelType: item.leveltype,
                SubmitStatus: "Yes",
                Maxlevel: item.approvers?.length,
                // ContentTitle:,
                MainListID: String(editItemID),
                RequestId: formData.DocumentCode,
                // RequestId:String(editID.Id),
                RequesterNameId: currentUser.Id,
                RequestedDate: new Date().toLocaleDateString("en-CA"),
                RequesterRoleId: RequesterRoleId,
                ProcessName: "Change Request",
                FormNameId: FormNameId,
                ApprovalType: "Approval",
                IsApprovalGenerated: "No"
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
            const postResult3 = await updateItemChangeRequestList(arr2, sp, editItemID);
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
            Swal.fire(successMessage, '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                sessionStorage.removeItem("ChangeRequestId")
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/MyApprovals.aspx`;
              }
            });

            // }
          }

        })
      }



    }
    else {

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
              ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toLocaleDateString("en-CA"),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: status,
              Remark: remark,


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
                ContentTitle: selectedOption?.ReferenceNumber,
                MainListNameId: ListNameId,
                ApproverRoleId: item.role,
                Level: Number(item.level),
                ApproversId: approversIds,
                LevelType: item.leveltype,
                SubmitStatus: "Yes",
                Maxlevel: item.approvers?.length,
                // ContentTitle:,
                MainListID: String(editItemID),
                RequestId: selectedOption?.DocumentCode,
                // RequestId:String(editID.Id),
                RequesterNameId: currentUser.Id,
                RequestedDate: new Date().toLocaleDateString("en-CA"),
                RequesterRoleId: RequesterRoleId,
                ProcessName: "Change Request",
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
            let arr2 = {
              SubmiitedDate: new Date().toLocaleDateString("en-CA"),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: status === "Rework" ? "Rework" : "Rejected",
              IsRework: status === "Rework" ? "Yes" : "No",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "No",
              SubmitStatus: "No",
            }
            const postResult2 = await updateItemChangeRequestList(arr2, sp, editItemID);
            const postId2 = postResult2?.data?.ID;
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
            Swal.fire(successMessage, '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                sessionStorage.removeItem("ChangeRequestId")
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/MyApprovals.aspx`;
              }
            });
            // }
          }
        })
      }
    }
  }

  const ForwardInitiatorApproval = async (status: string) => {
    let url = window.location.href.split('/sites/')[0];
    // let valid = true;
    let actionMessage = "";
    let successMessage = "";
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

            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            // //////////////Update Process Approval List when Submitted
            let arr = {
              ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toLocaleDateString("en-CA"),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: "Approved",
              // Remark: remark,

            }
            const postResult = await updateApprovalItem(arr, sp, editID.Id);
            const postId = postResult?.data?.ID;

            // //////////////Update Document cancellation List when Submitted
            let arr3 = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: formData.RequestDate,
              //IssueDate: formData.IssueDate,
              LocationId: selectedOption?.LocationId,
              CustodianId: selectedOption?.CustodianId,
              SerialNumber: selectedOption?.SerialNumber,
              IssueNumber: selectedOption?.IssueNumber,
              RevisionNumber: selectedOption?.RevisionNumber,
              //RevisionDate: selectedOption?.RevisionDate,
              DocumentCode: selectedOption?.value,
              ReferenceNumber: selectedOption?.ReferenceNumber,
              AmendmentTypeId: selectedOption?.AmendmentTypeId,
              RequestTypeId: selectedOption?.RequestTypeId,
              ClassificationId: selectedOption?.ClassificationId,
              ChangeRequestTypeId: selectedOption?.ChangeRequestTypeId,
              SubmiitedDate: selectedOption?.SubmiitedDate,
              SubmitStatus: "Yes",
              Status: "Pending",
              // DocumentName: "",
              // IsRework: false,
              // DigitalSignStatus: false,
              //ChangeRequestIDId: formData.ChangeRequestID,
              DocumentTypeId: selectedOption?.DocumentTypeId,
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              AttachmentId: selectedOption?.AttachmentId,
              AttachmentJson: selectedOption?.AttachmentJson


            }
            const postResult3 = await updateItemChangeRequestList(arr3, sp, editItemID);
            const postId3 = postResult?.data?.ID;


            // //////////////Update Document cancellation Reason List when Submitted

            for (const row of cancellReason) {

              const postPayload2 = {
                ChangeRequestIDId: formData.ChangeRequestID, // Assuming "Title" column exists
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              }

              if (!row.id) {

                const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                const postId2 = postResult2?.data?.ID;
                // debugger
                if (!postId2) {
                  console.error("Post creation failed.");
                  return;
                }

              }
              else if (row.id > 0) {
                const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                const postId2 = postResult2?.data?.ID;
              }

            }

            const toDelete = cancellReasonEdit.filter(
              (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
            );

            // Delete each item from SharePoint
            for (const item of toDelete) {
              try {
                await sp.web.lists.getByTitle("ChangeRequestReasonList").items.getById(item.id).delete();
                // console.log(`Deleted item with ID: ${item.ID}`);
              } catch (error) {
                console.error(`Error deleting item with ID: ${item.id}`, error);
              }
            }

            setLoading(false);
            Swal.fire(successMessage, '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                sessionStorage.removeItem("ChangeRequestId")
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
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

            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            let arr = {
              ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toLocaleDateString("en-CA"),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: status,
              // Remark: remark,

            }
            const postResult = await updateApprovalItem(arr, sp, editID.Id);
            const postId = postResult?.data?.ID;

            // //////////////Update Document cancellation List when Submitted
            let arr3 = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: formData.RequestDate,
              IssueDate: formData.IssueDate,
              LocationId: selectedOption?.LocationId,
              CustodianId: selectedOption?.CustodianId,
              SerialNumber: selectedOption?.SerialNumber,
              IssueNumber: selectedOption?.IssueNumber,
              RevisionNumber: selectedOption?.RevisionNumber,
              //RevisionDate: selectedOption?.RevisionDate,
              DocumentCode: selectedOption?.value,
              ReferenceNumber: selectedOption?.ReferenceNumber,
              AmendmentTypeId: selectedOption?.AmendmentTypeId,
              RequestTypeId: selectedOption?.RequestTypeId,
              ClassificationId: selectedOption?.ClassificationId,
              ChangeRequestTypeId: selectedOption?.ChangeRequestTypeId,
              SubmiitedDate: selectedOption?.SubmiitedDate,
              SubmitStatus: "No",
              Status: "Save as draft",
              // DocumentName: "",
              // IsRework: false,
              // DigitalSignStatus: false,
              //ChangeRequestIDId: formData.ChangeRequestID,
              DocumentTypeId: selectedOption?.DocumentTypeId,
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "No",
              CurrentUserRole: "OES",
              AttachmentId: selectedOption?.AttachmentId,
              AttachmentJson: selectedOption?.AttachmentJson


            }
            const postResult3 = await updateItemChangeRequestList(arr3, sp, editItemID);
            const postId3 = postResult?.data?.ID;


            // //////////////Update Document cancellation Reason List when Submitted

            for (const row of cancellReason) {

              const postPayload2 = {
                ChangeRequestIDId: formData.ChangeRequestID, // Assuming "Title" column exists
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              }

              if (!row.id) {

                const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                const postId2 = postResult2?.data?.ID;
                // debugger
                if (!postId2) {
                  console.error("Post creation failed.");
                  return;
                }

              }
              else if (row.id > 0) {
                const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                const postId2 = postResult2?.data?.ID;
              }

            }

            const toDelete = cancellReasonEdit.filter(
              (itemEdit) => !cancellReason.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
            );

            // Delete each item from SharePoint
            for (const item of toDelete) {
              try {
                await sp.web.lists.getByTitle("ChangeRequestReasonList").items.getById(item.id).delete();
                // console.log(`Deleted item with ID: ${item.ID}`);
              } catch (error) {
                console.error(`Error deleting item with ID: ${item.id}`, error);
              }
            }


            setLoading(false);
            Swal.fire(successMessage, '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                sessionStorage.removeItem("ChangeRequestId")
                window.location.href = `https://edcadae.sharepoint.com/sites/EDeDMS/SitePages/EDCMAIN.aspx`;
              }
            });

            // }
          }

        })
      }

    }


  }

  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>, libraryName: string, docLib: string) => {
    event.preventDefault();
    setAttachmentarr([]);
    setshowviewdownload(false);
    debugger
    //setDocumentLink(null);
    filechanged = true;
    newfileupload = true;
    let uloadBannerImageFiles: any[] = [];
    let uloadImageFiles: any[] = [];
    let uloadImageFiles1: any[] = [];
    if (event.target.files && event.target.files.length > 0) {
      const files = Array.from(event.target.files);
      (event.target as HTMLInputElement).value = '';
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

      if (files.length > 0) {
        const file = files[0];
        if (!allowedTypes.includes(file.type)) {
          Swal.fire({
            icon: "error",
            title: "Invalid File Type",
            text: "Only images and document files are allowed.",
          });
          return;
        }
        var arr = {};
        arr = {
          files: files,
          libraryName: libraryName,
          docLib: docLib,
          name: files[0].name,
          fileName: files[0].name,
          FileName: files[0].name,
          fileSize: files[0].size,
          date: new Date().toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric"
          }).replace(/ /g, "/"),
          //FileRef: previewUrl,
          FileLeafRef: files[0].name,
          //fileUrl: previewUrl,
          //fileType: fileType,
          //previewUrl: previewUrl
        };
        uloadBannerImageFiles.push(arr);
        setAttachmentarr(uloadBannerImageFiles);
        // const fileType = file.type.split("/")[0]; // Extract file type (image, pdf, etc.)
        // const folder = sp.web.getFolderByServerRelativePath('Socialfeedimages');
        // const uploadResult = await folder.files.addChunked(file.name, file);
        // console.log("File uploaded successfully", uploadResult);
        // let previewUrl: any;
        // // Generate the preview URL dynamically
        // if (uploadResult) {
        //   previewUrl = uploadResult.data.ServerRelativeUrl;
        // }

        // //await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);

        // //previewFile(previewUrl);
        // const preview = URL.createObjectURL(file);

        // newfilepreview = preview
        // setPreviewUrl(preview);
        // setFileType(fileType);
        // var arr = {};
        // arr = {
        //   files: files,
        //   libraryName: libraryName,
        //   docLib: docLib,
        //   name: files[0].name,
        //   fileName: files[0].name,
        //   FileName: files[0].name,
        //   fileSize: files[0].size,
        //   date: new Date().toLocaleDateString("en-GB", {
        //     day: "2-digit",
        //     month: "short",
        //     year: "numeric"
        //   }).replace(/ /g, "/"),
        //   FileRef: previewUrl,
        //   FileLeafRef: files[0].name,
        //   fileUrl: previewUrl,
        //   fileType: fileType,
        //   previewUrl: previewUrl
        // };
        // uloadBannerImageFiles.push(arr);
        // setAttachmentarr(uloadBannerImageFiles);
      } else {
        Swal.fire("upload a document")
      }
    }
  };
  const generatePreviewUrl = async (serverRelativeUrl: string) => {
    // Encode the file name and construct the preview URL
    const encodedFilePath = encodeURIComponent(serverRelativeUrl);

    // Example: 
    // serverRelativeUrl = "/sites/EDeDMS/test/DocumentLibraryInsideTest/Book.xlsx"
    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;

    // const previewUrl = `${siteUrl}/sites/EDeDMS/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    const previewUrl = `${siteUrl}${locationPath}/ChangeRequestDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    // const previewUrl = `${siteUrl}/sites/SPFXDemo/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    console.log("Generated Preview URL:", previewUrl);
    if (previewUrl) {
      console.log("enter herr")
      const deletebut = document.getElementById('closeCommand') as HTMLElement
      if (deletebut) {
        console.log(" here ", deletebut)
      }
    }
    return previewUrl;
  };
  const handlePreviewClick = (fileObj: any) => {
    debugger
    if (newfileupload === true) {
      console.log(newfilepreview, "here is newfilepreview")
      //alert(`new file ${newfilepreview}`)
      setPreviewUrl(newfilepreview); // Set the preview URL
      setIsModalOpen(true);   // Open the modal
    } else {
      if (fileObj && fileObj.serverUrl && fileObj.serverRelativeUrl) {
        const fileUrl = `${fileObj.serverUrl.trim()}${fileObj.serverRelativeUrl.trim()}`;
        setPreviewUrl(fileUrl); // Set the preview URL
        setIsModalOpen(true);   // Open the modal
      } else {
        //alert("Invalid file object. Cannot generate preview URL.");
      }
    }
  };
  const fetchChangeRequestTypes = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle('ChangeRequestTypeMaster')
        .items.select('ID', 'ChangeRequestType', 'IsActive')
        // .filter('IsActive eq Yes') // Fetch only active items
        ();

      const changeRequestCheckboxes = items.map((item: any) => ({
        id: item.ID,
        name: item.ChangeRequestType,
        isActive: item.IsActive,
      }));
      // alert(changeRequestCheckboxes)
      setChangeRequestCheckboxes(changeRequestCheckboxes);
      return changeRequestCheckboxes;
    } catch (error) {
      console.error('Error fetching ChangeRequestTypeMaster:', error);
    }
  };
  const validateCheckboxSelection = () => {
    if (selectedCheckboxIds.length === 0) {
      setisCheckboxSectionHighlighted(true);
      return false;
    }
    return true;
  };
  // Handle checkbox selection
  const handleCheckboxChange = (id: number) => {
    const updatedSelectedIds =
      selectedCheckboxIds.indexOf(id) !== -1
        ? selectedCheckboxIds.filter((selectedId) => selectedId !== id)
        : [...selectedCheckboxIds, id];
    setselectedCheckboxIds(updatedSelectedIds);
    setisCheckboxSectionHighlighted(false);
  };
  const renderCheckboxes = () => {
    //const { changeRequestCheckboxes, selectedCheckboxIds } = this.state;

    return changeRequestCheckboxes.map((checkbox) => (

      <div className="col-lg-3">
        <div key={checkbox.id} className="form-check mb-3">
          <input
            type="checkbox"
            // className="form-check-input"
            className={`form-check-input ${(!ValidSubmit && changerequesttypeerr) ? "border-on-error" : ""}`}
            id={`checkbox-${checkbox.id}`}
            // disabled={this.state.isReadonly} // Make the checkbox readonly if the condition is true
            // Use indexOf instead of includes
            disabled={InputDisabled && formData?.Status != "Rework"}
            checked={selectedCheckboxIds.indexOf(checkbox.id) !== -1}
            onChange={() => handleCheckboxChange(checkbox.id)}
          />
          <label className="form-check-label" htmlFor={`checkbox-${checkbox.id}`}>
            {checkbox.name}
          </label>
        </div>
      </div>
    ));
  };
  // const handleDelete = (index: number) => {
  //   setFilesArr((prevFiles: any[]) => prevFiles.filter((_file: any, i: number) => i !== index));
  // };
  const deleteLocalFileAttachment = async (index: number, ImagepostArr: any[]) => {
    try {
      // Extract the file information from the array
      // const fileToDelete = ImagepostArr[index];

      // if (!fileToDelete || (!fileToDelete.fileUrl && !fileToDelete.FileRef)) {
      //   throw new Error("File URL not found");
      // }

      // // Delete the file from SharePoint document library
      // const fileUrl = !fileToDelete.fileUrl ? fileToDelete.FileRef : fileToDelete.fileUrl;
      // console.log(fileUrl, "fileUrl", editID)
      // debugger
      // const removeimage = await sp.web.getFileByServerRelativePath(fileUrl).recycle(); // Sends the file to the recycle bin
      // console.log(removeimage, "removeimage")
      debugger
      const updatedArray = [...ImagepostArr];
      updatedArray.splice(index, 1);
      //setAttachmentarr(updatedArray);
      setAttachmentarr([]);
      Swal.fire("Deleted successfully", "", "success");
    } catch (error) {
      console.error("Error deleting file:", error);
      Swal.fire("Error", error.message, "error");
    }
  };
  const renderFilePreview = () => {
    if (fileType === "image") {
      return <img src={previewUrl} alt="File Preview" style={{ width: "100%", maxHeight: "400px" }} />;
    }

    if (fileType === "application" && previewUrl.endsWith(".pdf")) {
      return <embed src={previewUrl} type="application/pdf" width="100%" height="500px" />;
    }

    if (fileType === "application" && previewUrl.endsWith(".vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      // Excel File (XLSX) preview
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const sheetNames = workbook.SheetNames;
        const sheet = workbook.Sheets[sheetNames[0]];
        const html = XLSX.utils.sheet_to_html(sheet);
        document.getElementById("excel-preview").innerHTML = html;
      };
      reader.readAsBinaryString(Attachmentarr[0].files[0]);

      return <div id="excel-preview"></div>;
    }

    return (
      <div>
        <h4>Preview not available for this file type</h4>
        <p>{Attachmentarr[0].fileName}</p>
      </div>
    );
  };
  const onDateChange = (date: Date) => {
    // Format the selected date to DD-MMM-YYYY format
    const formattedDate = moment(date).format('DD/MMM/YYYY');
    setFormData({ ...formData, RequestDate: formattedDate });
  };
  return (
    <div id="wrapper" ref={elementRef}>
      {/* <div
        className="app-menu"
        id="myHeader">
        {/* <VerticalSideBar _context={sp} /> 
      </div> */}
      <div className="content-page">
        {/* <HorizontalNavbar _context={sp} siteUrl={siteUrl} /> */}
        {/* <div className="content" style={{ marginLeft: `${!useHide ? '240px' : '80px'}`, marginTop: '2.3rem' }}> */}
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

                    {Loading ?
                      // <div className="loadercss" role="status">Loading...
                      //   <img src={require('../../../Assets/ExtraImage/loader.gif')} style={{ height: '80px', width: '70px' }} alt="Check" />
                      // </div>
                      <div style={{ position: 'fixed', zIndex: '9', left: '0%', top: '0%' }} className="loadernewadd mt-10">
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

                      <div style={{ width: '100%' }} className="inbox-rightbar" ref={elementRef1}>

                        <div className="card">
                          <div className="card-body">

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
                            {/* <p className="sub-header">
                                                        Lorem ipsum dolor sit amet, consectetur adipisicing elit. Aspernatur, itaque.
                                                    </p> */}

                            <div className="row">
                              <div className="col-lg-4">


                                <div className="mb-3">
                                  <label htmlFor="RequesterName" className="form-label">Name:</label>
                                  <input type="text" id="Name" name="RequesterName" className="form-control" value={formData.RequesterName} disabled={true} />
                                </div>
                              </div>
                              <div className="col-lg-4">


                                <div className="mb-3">
                                  <label htmlFor="RequesterName" className="form-label">Department:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="Name" name="department" className="form-control" value={formData.Department} disabled={true} />
                                   */}
                                  <div
                                    title={SelectedOptionDepart?.label || "Select a department"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={Departopt}
                                      value={SelectedOptionDepart}
                                      name="Department"
                                      className={`${(!ValidDraft && departmenterr) ? "border-on-error" : ""} ${(!ValidSubmit && departmenterr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectDepart(selectedOption)}
                                      placeholder="Search Department"
                                      isDisabled={InputDisabled || formData?.Status == "Rework"}
                                    />
                                  </div>


                                </div>
                              </div>
                              <div className="col-lg-4">


                                <div className="mb-3">
                                  <label htmlFor="RequesterDesignation" className="form-label">Designation:</label>
                                  <input type="text" id="RequesterDesignation" name="RequesterDesignation" className="form-control" value={formData.RequesterDesignation} disabled={true} />
                                </div>
                              </div>

                              <div className="col-lg-4">
                                {console.log("formDataformData", formData, formData?.RequestDate)}
                                <div className="mb-3">
                                  <label htmlFor="RequestDate" className="form-label">Request Date:</label>

                                  <div
                                    title={
                                      formData?.RequestDate
                                        ? moment(formData?.RequestDate).format('DD/MMM/YYYY')
                                        : "Select a request date"
                                    }
                                  >
                                    <DatePicker
                                      value={
                                        formData?.RequestDate
                                          ? new Date(moment(formData?.RequestDate).format('YYYY-MM-DD'))
                                          : null
                                      }
                                      onSelectDate={onDateChange}
                                      maxDate={new Date()}
                                      minDate={new Date()}
                                      disabled={InputDisabled && formData?.Status !== "Rework"}
                                      formatDate={(date) => moment(date).format('DD/MMM/YYYY')}
                                    />
                                  </div>
                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Request Type:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                  <div
                                    title={selectedOptionReq?.label || "Select a request type"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={ReqType}
                                      value={selectedOptionReq}
                                      name="Request Type"
                                      className={`${(!ValidDraft && requesttypeerr) ? "border-on-error" : ""} ${(!ValidSubmit && requesttypeerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectReq(selectedOption)}
                                      placeholder="Search Request Type"
                                      isDisabled={InputDisabled || formData?.Status == "Rework"}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">
                                {console.log("selectedOptionReqselectedOptionReq", selectedOptionReq)}
                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Document Code:
                                    {selectedOptionReq?.label != "Change Request for New Addition" && <span className="text-danger1">*</span>}

                                  </label>
                                  {editItemID > 0 ?
                                    <input type="text" id="RequesterDesignation" name="RequesterDesignation" className="form-control" value={formData.DocumentCode} disabled={true} />
                                    :
                                    <div
                                      title={selectedOption?.label || "Select a document code"}
                                      style={{ width: "100%" }}
                                    >
                                      <Select
                                        options={rows}
                                        value={selectedOption}
                                        name="DocumentCode"
                                        isClearable={true}
                                        //isOptionDisabled={() => selectedOptionReq.label == "Change Request for New Addition"}
                                        isSearchable={true}
                                        className={`${(selectedOptionReq?.label != "Change Request for New Addition" && !ValidDraft && documentcodeerr) ? "border-on-error" : ""} ${(selectedOptionReq?.label != "Change Request for New Addition" && !ValidSubmit && documentcodeerr) ? "border-on-error" : ""}`}
                                        //className={`${(selectedOptionReq?.label != "Change Request for New Addition" && !ValidSubmit && documentcodeerr) ? "border-on-error" : ""}`}
                                        onChange={(selectedOption: any) => onSelectDocCode(selectedOption)}
                                        placeholder={selectedOptionReq == null || (selectedOptionReq != null && selectedOptionReq?.label == "Change Request for New Addition")
                                          || InputDisabled ? "" : "Search Document Code"}
                                        isDisabled={selectedOptionReq == null || (selectedOptionReq != null && selectedOptionReq?.label == "Change Request for New Addition")
                                          || InputDisabled
                                        }
                                      />
                                    </div>
                                  }
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}

                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="example-email" className="form-label">Issue No:</label>
                                  <input disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" value={formData.IssueNumber} />
                                </div>
                              </div>

                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="example-email" className="form-label">Revision No:</label>
                                  <input disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" value={formData.RevisionNumber} />
                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="example-email" className="form-label">Reference No:</label>
                                  <input disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" value={formData.ReferenceNumber} />
                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Document Type:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                  <div
                                    title={selectedOptionDoctype?.label || "Select a document type"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={DocumentTypeOpt}
                                      value={selectedOptionDoctype}
                                      name="Document Type"
                                      className={` ${(!ValidSubmit && documenttypeerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectDocumentType(selectedOption)}
                                      placeholder="Search Document Type"
                                      isDisabled={InputDisabled || (selectedOptionReq != null && selectedOptionReq?.label != "Change Request for New Addition") || formData?.Status == "Rework"}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Location:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                  <div
                                    title={selectedOptionLoc?.label || "Select a Location"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={LocationOpt}
                                      value={selectedOptionLoc}
                                      name="Location"
                                      className={` ${(!ValidSubmit && locationerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectLocation(selectedOption)}
                                      placeholder="Search Location"
                                      isDisabled={InputDisabled || (selectedOptionReq != null && selectedOptionReq?.label != "Change Request for New Addition") || formData?.Status == "Rework"}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Custodian:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                  <div
                                    title={selectedOptionCusto?.label || "Select a custodian"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={Custodianopt}
                                      value={selectedOptionCusto}
                                      name="Custodian"
                                      className={` ${(!ValidSubmit && custodianerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectCustodian(selectedOption)}
                                      placeholder="Search Custodian"
                                      isDisabled={InputDisabled || (selectedOptionReq != null && selectedOptionReq?.label != "Change Request for New Addition") || formData?.Status == "Rework"}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Amendment Type:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                  <div
                                    title={selectedOptionAmend?.label || "Select a Amendment type"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={Amendtype}
                                      value={selectedOptionAmend}
                                      name="Amendment Type"
                                      className={`${(!ValidSubmit && amendmenterr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectAmend(selectedOption)}
                                      placeholder="Search" isDisabled={InputDisabled && formData?.Status != "Rework"}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Classification:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}
                                  <div
                                    title={selectedOptionClass?.label || "Select a classification"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={Classificationopt}
                                      value={selectedOptionClass}
                                      name="Classification"
                                      className={`${(!ValidSubmit && classificationerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectClassification(selectedOption)}
                                      placeholder="Search Classification" isDisabled={InputDisabled && formData?.Status != "Rework"}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">


                                <div className="mb-3">
                                  <label htmlFor="RequesterName" className="form-label">Template Type:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="Name" name="department" className="form-control" value={formData.Department} disabled={true} />
   */}
                                  <div
                                    title={SelectedOptionTemplate?.label || "Select a template type"}
                                    style={{ width: "100%" }}
                                  >
                                    <Select
                                      options={TemplateTypeopt}
                                      value={SelectedOptionTemplate}
                                      name="Template Type"
                                      className={`${(!ValidSubmit && templatetypeerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectTemplatetype(selectedOption)}
                                      placeholder="Search Template type"
                                      isDisabled={InputDisabled}
                                    />
                                  </div>

                                </div>
                              </div>
                              {console.log("FormItemIdFormItemIdFormItemId", FormItemId, modeValue, selectedOption, Attachmentarr, DocumentLink)}
                              {/* //modeValue != "view" || modeValue == "edit" || modeValue != "approve"  && */}
                              {(FormItemId == null || (FormItemId != null && modeValue == "edit") || (modeValue == "view" || modeValue == "approve")
                                || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                <div className="col-lg-4">

                                  <div className="mb-3">
                                    <div className='d-flex justify-content-between'>
                                      <div>
                                        <label htmlFor="bannerImage" className="form-label">
                                          Attachment<span className="text-danger1">*</span>
                                        </label>
                                      </div>
                                      <div>
                                        <div>
                                          {Attachmentarr[0] != false && Attachmentarr.length > 0 &&
                                            Attachmentarr != undefined ? Attachmentarr.length == 1 &&
                                          (<a style={{ fontSize: '0.875rem' }} onClick={() => setShowModalAtt(true)}>
                                            <FontAwesomeIcon icon={faPaperclip} />1 file Attached
                                          </a>) : ""

                                          }
                                        </div>
                                      </div>
                                    </div>
                                    {/* <label htmlFor="DocumentCode" className="form-label">Attachment:<span className="text-danger1">*</span></label>
                                    {Attachmentarr[0] != false && Attachmentarr.length > 0 &&
                                      Attachmentarr != undefined ? Attachmentarr.length == 1 &&
                                    (<a style={{ fontSize: '0.875rem' }}
                                      //onClick={() => handlePreviewClick(Attachmentarr[0])}
                                      onClick={() => setShowModal(true)}>
                                      <FontAwesomeIcon icon={faPaperclip} />1 file Attached
                                    </a>) : ""
                                    } */}

                                    <input
                                      type="file"
                                      id="attachment"
                                      name="attachment"
                                      disabled={InputDisabled && formData?.Status != "Rework"}
                                      //disabled={handleSectionState('requestedBySection')}
                                      accept=".jpg,.jpeg,.png,.gif,.bmp,.svg,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx"
                                      className={`form-control ${(!ValidSubmit && attachmenterr) ? "border-on-error" : ""}`}
                                      onChange={(e) => onFileChange(e, "bannerimg", "Document")}
                                    />

                                  </div>
                                </div>
                              }
                              {console.log("ghghghghghghgh", showpreviousattachment, "jjjj", (showpreviousattachment && (((modeValue == "view" || modeValue == "approve" ||
                                (selectedOptionReq?.label != "Change Request for New Addition" && selectedOption)) ||
                                (modeValue == "edit" && formData?.Status == "Save as draft")) && DocumentLink && Attachmentarr.length == 0)))}
                              {(showpreviousattachment && selectedOptionReq?.label != "Change Request for New Addition" || (((modeValue == "view" || modeValue == "approve" ||
                                (selectedOptionReq?.label != "Change Request for New Addition" && selectedOption)) ||
                                (modeValue == "edit" && formData?.Status == "Save as draft")) && DocumentLink && Attachmentarr.length == 0)) &&

                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label htmlFor="DocumentCode" className="form-label">Previous Document:</label>
                                    {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}

                                    <div className="text-dark mt-0"> <span >
                                      <a onClick={() => setShowModal(true)} ><FontAwesomeIcon icon={faPaperclip} />{DocumentLink && "1 file Attached"}</a>

                                    </span>
                                    </div>

                                  </div>
                                </div>
                              }
                            </div>

                          </div>
                        </div>


                        <div className="card mt-2">
                          <div className="card-body">
                            <div className='row'>
                              <div className='col-sm-12'>
                                <h3 className="text-dark font-16 fw-bold mb-3">Change Request Type<span className="text-danger1">*</span></h3>
                                {/* <label className="form-label text-muted font-16">Change Request Type</label> */}
                                <div className="row"> {renderCheckboxes()}</div>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="card mt-2">
                          <div className="card-body">
                            <div className='row'>
                              <div className='col-sm-8'>
                                <h3 className="text-dark font-16 fw-bold mb-3">Description</h3>

                              </div>
                              {console.log("formData?.Status ", InputDisabled, formData)}
                              <div className='col-sm-4'>
                                <div style={{ textAlign: "right" }} className="mt-2 float-end text-right">
                                  {/* <i style={{ cursor: "pointer" }} onClick={addField}  className="fe-plus-circle  font-20 text-warning"></i> */}
                                  {/* <i style={{ cursor: "pointer" }} className="fe-plus-circle  font-20 text-warning"></i> */}
                                  {(modeValue === "" || modeValue === "edit" || InputDisabled != true || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                    // {(InputDisabled != true && (formData?.Status == "Rework" || formData?.Status == "" || formData?.Status == "Save as draft")) && 
                                    <img style={{ width: '30px', cursor: 'pointer', marginTop: '-7px' }}
                                      src={require("../assets/plus.png")} onClick={addCancelReason} className=''></img>}


                                </div>
                              </div>
                            </div>

                            {/* <p className="sub-header">
                                                        Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s
                                                    </p> */}

                            <div className="row">
                              <table className="mtbalenew table-centered table-nowrap table-borderless mb-0" id="tbl">
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: "40px", maxWidth: "40px" }}>S.No</th>
                                    <th>Change Description<span className="text-danger1">*</span></th>
                                    <th>Reason for Change<span className="text-danger1">*</span></th>
                                    {(modeValue === "" || modeValue === "edit" || InputDisabled != true || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                      <th style={{ minWidth: "30px", maxWidth: "30px" }}>Action</th>
                                    }
                                  </tr>

                                </thead>
                                <tbody >
                                  {console.log("cancellReasonn", cancellReason)}
                                  {cancellReason.map((row, index) => (
                                    <tr key={index}> <td style={{ minWidth: "30px", maxWidth: "30px" }}>
                                      <div
                                        style={{ marginLeft: "0px" }}
                                        className="indexdesign"
                                      >
                                        {index + 1}</div></td>
                                      <td title={row.description}>
                                        {/* <input type="text" id="simpleinput" disabled={InputDisabled && formData?.Status != "Rework"}
                                        value={row.description}
                                        className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                        onChange={(e) => {
                                          const newRowscancellReason = [...cancellReason];
                                          newRowscancellReason[index].description = e.target.value;
                                          setcancellReason(newRowscancellReason);
                                        }}
                                      /> */}
                                        <textarea
                                          id="simpleinput"
                                          disabled={InputDisabled && formData?.Status !== "Rework"}
                                          value={row.description}
                                          className={`form-control mb-0 ${(!ValidCancelReason && changedescriptionerr) ? "border-on-error" : ""}`}
                                          onChange={(e) => {
                                            const newRowscancellReason = [...cancellReason];
                                            newRowscancellReason[index].description = e.target.value;
                                            setcancellReason(newRowscancellReason);
                                          }}
                                        />

                                      </td>
                                      <td title={row.reason}>
                                        {/* <input type="text" id="simpleinput" disabled={InputDisabled && formData?.Status != "Rework"}
                                        className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                        value={row.reason}
                                        onChange={(e) => {
                                          const newRowscancellReason = [...cancellReason];
                                          newRowscancellReason[index].reason = e.target.value;
                                          setcancellReason(newRowfscancellReason);
                                        }}
                                      /> */}
                                        <textarea
                                          id="simpleinput"
                                          disabled={InputDisabled && formData?.Status !== "Rework"}
                                          className={`form-control mb-0 ${(!ValidCancelReason && changereasonerr) ? "border-on-error" : ""}`}
                                          value={row.reason}
                                          onChange={(e) => {
                                            const newRowscancellReason = [...cancellReason];
                                            newRowscancellReason[index].reason = e.target.value;
                                            setcancellReason(newRowscancellReason);
                                          }}
                                        />
                                      </td>
                                      {(modeValue === "" || modeValue === "edit" || InputDisabled != true || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                        <td style={{ minWidth: "30px", maxWidth: "30px", textAlign: 'center' }}>
                                          <img src={require("../assets/del.png")} className='' onClick={() => deleteLocalFile(index, cancellReason)}></img>
                                        </td>
                                      }
                                    </tr>
                                  ))}
                                </tbody>
                              </table>


                            </div>


                          </div>


                        </div>

                        {console.log("editiiiiifhifassignmentt", editID, modeValue, InputDisabled, ApprovalTypeOptions, MainEditItem,
                          (modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES"),
                          (MainEditItem !== null && MainEditItem.length != 0 && MainEditItem.Status != "Save as draft" && MainEditItem.Status != "Rework" && modeValue !== "view"))}
                        {/* {((modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES") ||
                          (MainEditItem !== null && MainEditItem.length != 0 && MainEditItem.Status != "Save as draft" && MainEditItem.Status != "Rework" && modeValue !== "view"))
                          && */}
                        {modeValue === "approve" && editID != null && editID.Status === "Pending" && editID.CurrentUserRole !== "Initiator" &&
                          <div className="card mt-2" style={{ marginBottom: '17px' }}>
                            <div className="card-body">
                              <div className='row'>
                                <div className='col-sm-8'>
                                  <h3 className="header-title text-dark font-16 fw-bold mb-3 ">Forward Approval To</h3>
                                  <label>Define the approval hierarchy to ensure requests are routed to the appropriate approvers.
                                  </label>
                                </div>

                                <div className='col-sm-4'>
                                  {modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES" &&

                                    <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right", paddingRight: "22px" }}>
                                      <img style={{ width: '30px', cursor: 'pointer' }} src={require("../assets/plus.png")}
                                        onClick={handleAddRow} className='' />

                                      {/*  {/* {modeValue === "approve" && editID != null && editID.CurrentUserRole !== "Initiator" && MainEditItem !== null && MainEditItem?.Status !== "Save as draft" && 
                                      <i style={{ cursor: "pointer" }} onClick={handleAddRow} className="fe-plus-circle font-20 text-warning"></i> */}
                                    </div>
                                  }
                                </div>

                              </div>

                              <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">
                                <table style={{ overflow: 'inherit' }} className="mtbalenew tpnew table-centered table-nowrap table-borderless mb-0 newtabledc" id="myTabl">
                                  <thead >
                                    <tr>
                                      <th style={{ minWidth: "40px", maxWidth: "40px" }}>S.No</th>
                                      <th style={{ borderBottomLeftRadius: "0px" }}>Role</th>
                                      <th style={{ minWidth: '70px', maxWidth: '70px' }} >Level</th>
                                      <th >Approver Name</th>
                                      <th >Approval Criteria</th>
                                      <th style={{ minWidth: '70px', maxWidth: '70px' }}>Action</th>
                                    </tr>
                                  </thead>
                                  <tbody style={{ maxHeight: "8007px", overflow: 'inherit' }}>
                                    {console.log("forwardToArrforwardToArrforwardToArr", forwardToArr, UserRoles, ApprovalTypeOptions)}
                                    {forwardToArr.map((row, index) => (

                                      <tr key={index}> <td style={{ minWidth: "30px", maxWidth: "30px" }}>
                                        <div
                                          style={{ marginLeft: "0px", overflow: 'inherit' }}
                                          className="indexdesign"
                                        >
                                          {index + 1}</div></td>
                                        <td style={{ overflow: 'inherit' }} className="ng-binding">
                                          <select onChange={(e) => onSelectRole(e, row.level)} value={row.role}
                                            disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                            className={`form-select ${(!Validforward) ? "border-on-error" : ""}`}>
                                            <option value="" selected>Select Role</option>
                                            {UserRoles.filter((role: any) =>
                                              !forwardToArr.some((r) => r.role === role.value && r.level !== row.level) || role.value === row.role // Allow the current row's role
                                            ).map((role: any, index: number) => (
                                              <option key={index} value={role.value}

                                              //disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                              >
                                                {role.label}</option>
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
                                            className={` ${(!Validforward) ? "border-on-error" : ""}`}
                                            //className={`form-control ${(!ValidDraft) ? "border-on-error" : ""} ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                            // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                            onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                            placeholder="Enter Approver Name"
                                            isDisabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                          />



                                        </td>
                                        <td style={{ overflow: 'inherit' }} className="ng-binding">
                                          <select className={`form-select ${(!Validforward) ? "border-on-error" : ""}`}
                                            onChange={(e) => onSelectApprovalType(e, row.level)}
                                            value={row.leveltype}
                                            disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                          >
                                            <option value="" selected>Select Type</option>
                                            {ApprovalTypeOptions.map((x: any, index: number) => (
                                              <option key={index} value={x.value}

                                              >{x.label}</option>
                                            ))}
                                          </select>

                                        </td>
                                        <td style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>
                                          {/* <i className="fe-trash-2 text-danger"></i> */}
                                          {editID.CurrentUserRole === "OES" ? <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} /> :
                                            <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' />}
                                          {/* <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} className='sidebariconsmall' /> */}

                                        </td>
                                      </tr>

                                    ))}

                                  </tbody>
                                </table>
                              </div>
                              {modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES" &&
                                <div className="row mt-3">
                                  <div className="col-12 text-center">
                                    <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardApproval("Forward")} >
                                      <i className="fe-check-circle me-1"></i> Forward
                                    </button>                                  {/* <a href="#"> */}
                                    <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={() => ForwardApproval("Rework")} >
                                      <i className="fe-corner-up-left me-1"></i> Rework
                                    </button>
                                    <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={() => ForwardApproval("Rejected")} >
                                      <i className="fe-x me-1"></i> Reject
                                    </button>
                                    <button type="button" className="btn btn-light waves-effect waves-light m-1" onClick={handleCancel}>
                                      <i className="fe-x me-1"></i> Cancel
                                    </button>
                                    {/* </a> */}
                                  </div>
                                </div>
                              }
                            </div>
                          </div>
                        }
                        {console.log("editiiiiifhif", editID, modeValue, InputDisabled)}
                        {
                          (InputDisabled && editID != null && modeValue === "approve" && editID.ApprovalType === "Approval" && editID.Status === "Pending") ? (
                            <WorkflowAction currentItem={editID} ctx={props.context} ContentType={CONTENTTYPE_ChangeDocument}
                              DisableApproval={false} DisableCancel={false}
                            />
                          ) : (<div></div>)
                        }
                        {console.log("MainEditItemMainEditItem", MainEditItem)}
                        {/* ////////////Audit History card */}
                        {MainEditItem !== null && MainEditItem.length != 0 && MainEditItem.Status != "Save as draft" &&
                          <WorkflowAuditHistory ContentItemId={MainEditItem} ContentType={CONTENTTYPE_ChangeDocument} ctx={props.context} />
                        }

                        {console.log("ediiiiitiitiititID", editID, InputDisabled, editItemID, MainEditItem, modeValue)}

                        {/* <div className="row mt-3">
                    <div className="col-12 text-center mt-2">
                      {/* <a href="my-approval.html">   */}
                        {/* {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}><i className="fe-check-circle me-1"></i> Save As Draft</button>}

                      {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}><i className="fe-check-circle me-1"></i> Submit</button>}
                      {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}><i className="fe-check-circle me-1"></i> Save As Draft</button>}

                      {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><i className="fe-check-circle me-1"></i> Submit</button>}

                      {/* </a> */}
                        {/* <a href="../sites/EDeDMS/SitePages/EDCMAIN.aspx">       */}
                        {/* <button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}><i className="fe-x me-1"></i> Cancel</button>
                      {/* </a> */}
                        {/* </div>
                  </div>  */}
                        <div className="row mt-3">
                          <div className="col-12 text-center">
                            {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                              <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                              Save As Draft</button>}

                            {(((InputDisabled != true && editItemID == null && MainEditItem == null) || (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                              <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                              Submit</button>}
                            {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</button>}
                            {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</button>}
                            {((modeValue === "" || modeValue === "edit" || modeValue === "view") ||
                              (InputDisabled && editID != null && modeValue === "approve" && editID.Status === "Approved") ||
                              (editID !== null && editID.IsInitiator == "Yes")) &&
                              <button style={{ width: '145px' }} type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                className='me-1' alt="x" /> Cancel</button>
                            }
                            {/* </a> */}
                          </div>
                        </div>
                        {console.log("Attachmentarrnew", Attachmentarr, isModalOpen)}
                        {Attachmentarr.length > 0 && isModalOpen &&
                          <Modal show={isModalOpen} onHide={() => setIsModalOpen(false)} size="lg" className='newmobmodal'>
                            <Modal.Header closeButton>
                              {previewUrl && (
                                <Modal.Title>Preview of {Attachmentarr.length > 0 && Attachmentarr[0].name}</Modal.Title>
                              )}
                            </Modal.Header>
                            <Modal.Body>
                              {previewUrl && (
                                renderFilePreview()
                              )}
                            </Modal.Body>

                          </Modal>
                        }

                        <Modal show={showModal} onHide={() => setShowModal(false)} size="lg" className='newmobmodal'>
                          <Modal.Header closeButton>
                            <Modal.Title> Attachment Details <br></br>
                              {/* <p className='text-muted font-14 fw-400'>Below are the attachment details for Change Request
                              </p> */}

                            </Modal.Title>
                            {/* {ImagepostArr1.length > 0 && showBannerModal && <Modal.Title>Media Images</Modal.Title>} */}
                          </Modal.Header>
                          <Modal.Body className="" id="style-5">
                            <>
                              <table className="mtbalenew" >
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                    <th>File Name</th>
                                    <th > File Link </th>
                                    <th className='text-center'>Upload date</th>
                                    {/* <th > Action </th> */}
                                  </tr>
                                </thead>
                                <tbody>
                                  {console.log("Attachmentarrnmnm doc link", DocumentLink, DocumentLink != null)}
                                  <tr >
                                    <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>1</td>
                                    <td title={DocumentLink != null && `${DocumentLink?.FileLeafRef}`}>{DocumentLink != null && `${DocumentLink?.FileLeafRef}`}</td>
                                    <td style={{ textAlign: 'center' }}>
                                      {/* <span onClick={() => OpenFile(DocumentLink != null && DocumentLink, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                        <FontAwesomeIcon icon={faEye} /></span> */}
                                      <span onClick={() => OpenFile(DocumentLink != null && DocumentLink, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                        <FontAwesomeIcon icon={faEye} /></span>
                                      <span onClick={() => OpenFile(DocumentLink != null && DocumentLink, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                        <FontAwesomeIcon icon={faDownload} /></span>
                                    </td>
                                    <td title={DocumentLink && moment(DocumentLink?.Created).format("DD/MMM/YYYY")}>{DocumentLink && moment(DocumentLink?.Created).format("DD/MMM/YYYY")}</td>
                                    {/* <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
                                      <img src={require("../assets/del.png")} className='' onClick={() => deleteLocalFileAttachment(0, Attachmentarr)}></img>
                                    </td> */}
                                  </tr>
                                </tbody>

                              </table>
                            </>
                          </Modal.Body>
                        </Modal>
                        <Modal show={ShowModalAtt} onHide={() => setShowModalAtt(false)} size="lg" className='newmobmodal'>
                          <Modal.Header closeButton>
                            <Modal.Title> Attachment Details <br></br>
                              {/* <p className='text-muted font-14 fw-400'>Below are the attachment details for Change Request
                              </p> */}

                            </Modal.Title>
                            {/* {ImagepostArr1.length > 0 && showBannerModal && <Modal.Title>Media Images</Modal.Title>} */}
                          </Modal.Header>
                          <Modal.Body className="" id="style-5">
                            <>
                              <table className="mtbalenew" >
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                    <th>File Name</th>
                                    {/* {showButton && */}
                                    {((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                      || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                      <th > File Link </th>
                                    }
                                    <th className='text-center'>Upload date</th>
                                    {(modeValue == "edit" || modeValue == null || modeValue == ""
                                      || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                      <th style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}> Action </th>
                                    }
                                  </tr>
                                </thead>
                                <tbody>
                                  {console.log("Attachmentarrnmnm attach only", Attachmentarr)}
                                  {Attachmentarr.length > 0 &&
                                    <tr >
                                      <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>1</td>
                                      <td title={Attachmentarr && Attachmentarr[0]?.FileName}>{Attachmentarr && Attachmentarr[0]?.FileName}</td>
                                      {/* {showButton && */}
                                      {((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                        || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                        <td style={{ textAlign: 'center' }}>
                                          {/* <span onClick={() => OpenFile(Attachmentarr && Attachmentarr[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                            <FontAwesomeIcon icon={faEye} /></span> */}
                                          <span onClick={() => OpenFile(Attachmentarr && Attachmentarr[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                            <FontAwesomeIcon icon={faEye} /></span>
                                          <span onClick={() => OpenFile(Attachmentarr && Attachmentarr[0], "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                            <FontAwesomeIcon icon={faDownload} /></span>
                                        </td>
                                      }
                                      <td title={Attachmentarr && moment(Attachmentarr[0]?.Created).format("DD/MMM/YYYY")}>{Attachmentarr && moment(Attachmentarr[0]?.Created).format("DD/MMM/YYYY")}</td>
                                      {(modeValue == "edit" || modeValue == null || modeValue == ""
                                        || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                        <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
                                          <img src={require("../assets/del.png")} className='' onClick={() => deleteLocalFileAttachment(0, Attachmentarr)}></img>
                                        </td>
                                      }
                                    </tr>
                                  }
                                </tbody>
                              </table>
                            </>
                          </Modal.Body>
                        </Modal>
                      </div>
                    }
                  </div>

                </div>
              </div>


            </div>
          </div>
        </div >
      </div >
    </div>

  );
}

const ChangeDocumentRequest: React.FC<IChangeDocumentRequestProps> = (props) => (
  <Provider>
    <ChangeDocumentRequestContext props={props} />
  </Provider>
);


export default ChangeDocumentRequest

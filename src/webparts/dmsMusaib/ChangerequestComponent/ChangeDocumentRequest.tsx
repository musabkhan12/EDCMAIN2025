// this is our main file for CR
import * as React from 'react';

// import type { IChangeDocumentRequestProps } from './IChangeDocumentRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import "@pnp/sp/profiles";
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
import "@pnp/sp/files";
import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import {
  addAllProcessItem, addApprovalItem, addItem, addItemChangeRequestReasonlist,
  addItemChangeRequestList, getAllAmendmentType, getAllClassificationMaster, getAllDocumentCode,
  getAllProcessData, getAllRequestType, getApprovalByID, getApprovalByID2, getDataRoles,
  getDocumentLinkByID, getFormNameID, getItemByID, getItemByIDChangeRequest, getItemByIDCR,
  getListNameID, GetQueryString, getRequesterID, UpdateAllProcessItem, updateApprovalItem, getdigitalsignaturerequestbyID,
  updateItem, updateItemChangeRequestReasonList, updateItemChangeRequestList,
  getDocumentCodeselected, getDocumentLinkByIDarr,
  getAllDepartment,
  getAllTemplateType,
  getGeneratedTemplateDocCR,
  getchangerequesttemp,
  getDocumentCodeselectedApproved,
  getTemplatelink,
  updateDigitalsign,
  getdigitalsignaturerequestbyIDYes,
  CheckIfAlreadyactionTaken,
  getallProcessApprovalitemsLevel,
   getDelegateduser,
  getchangerequestnotes,
  CheckifDocumentisApproved,
  getChangeRequestTypeMaster
  
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
import { faDownload, faEye, faPaperclip, faStickyNote } from '@fortawesome/free-solid-svg-icons';
// import { uploadFile } from '../../../APISearvice/MediaService';
import { Modal } from 'react-bootstrap';
import { Link } from '@fluentui/react';
import moment from 'moment';
import { DatePicker } from 'office-ui-fabric-react';
import * as XLSX from 'xlsx';
import { SITE_URL } from '../../../Shared/Constants';
import { set } from 'date-fns';
import FileViewer from './fileviewer';
let newfileupload: any
let newfilepreview: any;
let filechanged: boolean = false;
let locationPath: any;
let enableTemplatetype: boolean = false;
let Showfile: boolean = false;
let IsRecorddisabled: boolean = false;
let isCurrentuserDelegated: boolean = false;
let IsDepartmentEditable: boolean = false;
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
  leveltype: string;
  roleError?: boolean;
  approverError?: boolean;
  responsibilityerror?: boolean;
  typeError?: boolean;
  rowError?: boolean;
  Responsibility?: string;
  IsSignatureRequired?: boolean;
  IsPreparedBy?: boolean;
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
  const iframeRef = React.useRef<HTMLIFrameElement>(null);
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
  const [filenameerr, setfilenameerr] = React.useState(false);
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
  const [hidedigisign, sethidedigisign] = React.useState(false);
  const [showdigisign, setshowdigisign] = React.useState(false);
  const [Loading, setLoading] = React.useState(false);
  const [FormItemId, setFormItemId] = React.useState(null);
  const [editID, setEditID] = React.useState(null);
  const [editItemID, setEditItemID] = React.useState(null);
  const [MainEditItem, setMainEditItem] = React.useState(null);
  const [rows, setRows] = React.useState<any>([]);
  const [doccoderows, setdoccoderows] = React.useState<any>([]);
  const [ReqType, setReqType] = React.useState<any>([]);
  const [changereqNotes, setchangereqNotes] = React.useState<any>([]);
  const [Departopt, setDepartment] = React.useState<any>([]);
  const [TemplateTypeopt, setTemplateType] = React.useState<any>([]);
  const [Amendtype, setAmendtype] = React.useState<any>([]);
  const [Classificationopt, setClassificationopt] = React.useState<any>([]);
  const [LocationOpt, setLocationOpt] = React.useState<any>([]);
  const [rowErrors, setRowErrors] = React.useState([]); // Will store an array of { descriptionError, reasonError }

  const [Custodianopt, setCustodianopt] = React.useState<any>([]);
  const [DocumentTypeOpt, setDocumentTypeOpt] = React.useState<any>([]);
  const [UserRoles, setUserRoles] = React.useState<any>([]);
  const [rows1, setRows1] = React.useState<any>([]);
  const [DocumentLink, setDocumentLink] = React.useState(null);
  const [Templatelink, setTemplatelink] = React.useState(null);
  const [DigitalsignID, setDigitalsignID] = React.useState(null);
  const [redirecturl, setredirecturl] = React.useState(null);
  const [sharewitherr, setsharewitherr] = React.useState(false);
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
  const [maxlevelAllprocess, setmaxlevelAllprocess] = React.useState("");

  const [ValidRemark, setValidRemark] = React.useState(true);
  const [MandatRemark, setMandatRemark] = React.useState(false);

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
    Remark: "",
    RequesterNameId: 0,
    RequesterName: "",
    RequesterDesignation: "",
    DepartmentId: 0,
    filename: "",
    TemplateTypeId: 0,
    RequestDate: "",
    IssueDate: "",
    LocationId: 0,
    CustodianId: 0,
    PreparedById: [],
    SerialNumber: "",
    IssueNumber: "",
    RevisionNumber: "",
    RevisionDate: "",
    DocumentCode: "",
    ReferenceNumber: "",
    FileName: "",
    AmendmentTypeId: 0,
    RequestTypeId: 0,
    ClassificationId: 0,
    ChangeRequestTypeId: [],
    SubmiitedDate: "",
    SubmitStatus: "",
    Status: "",
    DocumentName: "",
    IsRework: false,
    PreviousAttachmentID: "",
    DigitalSignStatus: false,
    ChangeRequestID: 0,
    AttachmentId: [],
    AttachmentJson: "",
    DocumentTypeId: 0
  });
  const [Attachmentarr, setAttachmentarr] = React.useState([]);
  const [previewUrl, setPreviewUrl] = React.useState<string | null>(null); // To store the file preview URL
  const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
  const [showeditview, setshoweditview] = React.useState<boolean>(false);
  const [showModal, setShowModal] = React.useState(false);
  const [showModaltemp, setShowModaltemp] = React.useState(false);
  const [ShowModalAtt, setShowModalAtt] = React.useState(false);
  const [ShowModalNotes, setShowModalNotes] = React.useState(false);

  const [ShowModalTemplateDoc, setShowModalTemplateDoc] = React.useState(false);

  const [forwardToArr, setForwardToArr] = React.useState<ForwardTo[]>([
    {
      id: 0, role: 0, level: 1, approvers: [], leveltype: "One", roleError: false, Responsibility: "Signer", IsSignatureRequired: true,
      approverError: false,
      responsibilityerror: false,
      typeError: false,
      rowError: false, IsPreparedBy: false
    } // Default row
  ]);
  const [currentUserDept, setcurrentUserDept] = React.useState("");
  const [forwardToArrEdit, setForwardToArrEdit] = React.useState<ForwardTo[]>([]);
  const [changeRequestCheckboxes, setChangeRequestCheckboxes] = React.useState<ChangeRequestCheckbox[]>([]);
  const [enabledCheckboxIds, setEnabledCheckboxIds] = React.useState<number[]>([]);
  const [employeeDetails, setemployeeDetails] = React.useState<IEmployeeDetails[] | null>([]);
  const [selectedCheckboxIds, setselectedCheckboxIds] = React.useState<number[]>([]);
  const [isCheckboxSectionHighlighted, setisCheckboxSectionHighlighted] = React.useState<boolean>(false);
  const [fileType, setFileType] = React.useState("");
  const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);
  const [remark, setRemark] = React.useState("");
  const [showview, setshowview] = React.useState(false);
  const [showButton, setShowButton] = React.useState(false);
  const [sharewithusers, setSharewithusers] = React.useState([]);
  // Function to handle People Picker selection
  const onPeoplePickerChange = (items: any[]) => {
    setSelectedUsers(items);
  };

  const ApiCallFunc = async () => {
    const path1 = window.location.href;
    debugger
    const currentUser = await sp.web.currentUser();
    const userGroups = await sp.web.siteUsers.getById(currentUser.Id).groups();
    const isMemberOfStrategyandSustainableGrowth = userGroups.some(group => group.Title === `Strategy and Sustainable Growth`);
    const isMemberOfSuperAdmin = userGroups.some(group => group.Title === `DMSSuper_Admin`);
    if (isMemberOfStrategyandSustainableGrowth || isMemberOfSuperAdmin) {
      enableTemplatetype = true;
    }
    const IsDepartmentPermission = userGroups.some(group => group.Title === `DepartmentFieldPermission`);
    if (IsDepartmentPermission) {
      IsDepartmentEditable = true;
    }
    locationPath = window.location.href.match(/\/sites\/[^\/]+/)[0];
    if (path1.includes("/view/") || path1.includes("/approve/")) {
      //setLoading(true);
      setInputDisabled(true);
      setshowview(true);
    }
    else {
      setInputDisabled(false);
    }
    if (path1.includes("/edit/")) {
      //setLoading(true); ////
      setshowview(true);
    }

    console.log("inpt diasba", InputDisabled, path1, path1.includes("/view/"))
    //setLoading(true);
    var changerequestnotes = await getchangerequestnotes(sp);
    const changerequestNotesCH = changerequestnotes.map((item: any) => ({
      value: item.ID,
      label: item.RequestType,
      itemId: item.ID,
      Notes: item.Notes
    }));
    setchangereqNotes(changerequestNotesCH);
    console.log("changereqNote", InputDisabled, path1, path1.includes("/view/"), changerequestNotesCH)
    var ReqTypeArr = await getAllRequestType(sp);
    ReqTypeArr.sort((a, b) => a.RequestType.localeCompare(b.RequestType));
    const optionsreq = ReqTypeArr.map((item: any) => ({
      value: item.ID,
      label: item.RequestType,
      itemId: item.ID,
      requestcode: item.RequestCode
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
    // let currentuserdepartment = UserDept == "IT" ? "Information Technology" : UserDept;
    let currentuserdepartment = UserDept;
    let optionsfilterdepart = currentuserdepartment != "" && optionsDepartment.filter((user) => user.adDepartmentName === currentuserdepartment);
    if (currentuserdepartment != "") {
      setSelectedOptionDepart(optionsDepartment.filter((user) => user.adDepartmentName === currentuserdepartment));
    }
    console.log("Current user department", currentuserdepartment, optionsDepartment.filter((user) => user.adDepartmentName === currentuserdepartment))
    // let optionsfilterdepart = optionsDepartment.filter((user) => user.label === currentuserdepartment);
    // setSelectedOptionDepart(optionsDepartment.filter((user) => user.label === currentuserdepartment));
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
      DepartmentId: optionsfilterdepart && optionsfilterdepart[0]?.value
      //RequestedDate: new Date().toISOString().split("T")[0] // Format as YYYY-MM-DD

    }));

    const selectedTemplatefirst = optionsTemplateType.filter((cust: { label: any; }) => cust.label === "Others")[0] || null;
    setSelectedOptionTemplate(selectedTemplatefirst);
    setFormData(prevData => ({
      ...prevData,
      TemplateTypeId: selectedTemplatefirst?.value
      // Format as YYYY-MM-DD
    }));

    let currentuserinpreparedby: any = [];

    currentuserinpreparedby = [{
      value: currentUser.Id,
      label: currentUser.Title,
      UserName: currentUser.Title
    }];

    setSharewithusers([]);
    var DocCodeArr = await getAllDocumentCode(sp);
    let doccodearrew: any;
    const options = DocCodeArr.map((item: any) => ({
      value: item.DocumentCode,
      label: item.DocumentCode,
      IssueNumber: item.IssueNumber,
      FileName: item.FileName,
      ReferenceNumber: item.ReferenceNumber,
      RevisionNumber: item.RevisionNumber,
      ChangeRequestID: item.ID,
      IssueDate: item.IssueDate,
      LocationId: item.LocationId,
      filename: item.FileName,
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
      Status: item.Status,
      PreviousAttachmentID: item.AttacmentId ? item.AttacmentId[0] : "", // Assuming AttachmentId is an array and we want the first element
      // DocumentName: "",
      // IsRework: false,
      // DigitalSignStatus: false,


    }));
    setdoccoderows(options);
    console.log("DocCodeArr", DocCodeArr);
    if (optionsfilterdepart.length > 0) {
      doccodearrew = options.filter((x: any) => x.DepartmentId == optionsfilterdepart[0]?.value)
    }

    setRows(doccodearrew);
    let alltemplates = await getTemplatelink(sp);
    if (alltemplates.length > 0) {
      setTemplatelink(alltemplates[0]);
    }



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
      let disablerecord = ProcessItemId && ProcessItemId?.Id > 0 && await CheckIfAlreadyactionTaken(sp, ProcessItemId?.Id, CONTENTTYPE_ChangeDocument);
      IsRecorddisabled = !disablerecord;
      isCurrentuserDelegated = ProcessItemId && ProcessItemId?.AssignedToId > 0 && await getDelegateduser(sp, ProcessItemId?.AssignedToId)
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
          let disablerecord = ProcessItemId && ProcessItemId?.Id > 0 && await CheckIfAlreadyactionTaken(sp, ProcessItemId?.Id, CONTENTTYPE_ChangeDocument);
          IsRecorddisabled = !disablerecord;
          isCurrentuserDelegated = ProcessItemId && ProcessItemId?.AssignedToId > 0 && await getDelegateduser(sp, ProcessItemId?.AssignedToId)
          setInputDisabled(await getApprovalByID2(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_ChangeDocument));
        }
      }


    }
    setLoading(false);
    // formitemid =20;
    // setLoading(true);
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
      const setBannerById = await getItemByIDCR(sp, Number(formitemid));
      let maxlevelAllprocess = await getallProcessApprovalitemsLevel(sp, Number(formitemid));
      setmaxlevelAllprocess(maxlevelAllprocess);
      const newItem1 = await getdigitalsignaturerequestbyID("ChangeRequestList", sp, Number(formitemid));
      const isRecordExist = await getdigitalsignaturerequestbyIDYes("ChangeRequestList", sp, Number(formitemid));
      console.log("newItem1newItem1", newItem1);

      if (newItem1.length > 0) {
        setDigitalsignID(newItem1[0].ID)
      }

      if (isRecordExist == "Yes" || isRecordExist == "NoRecord") {
        setshowdigisign(false);
      } else {
        setshowdigisign(true);
      }


      // setEditID(Number(setBannerById[0].ID))
      if (setBannerById.length > 0) {
        debugger
        setEditForm(true);
        setMainEditItem(setBannerById[0]);
        debugger
        // setCategoryData(await getCategory(sp, Number(setBannerById[0]?.TypeMaster))) // Category
        if (((setBannerById[0].PreviousAttachmentID != null || setBannerById[0].PreviousAttachmentID != undefined
          || setBannerById[0].PreviousAttachmentID != "") && setBannerById[0].PreviousAttachmentID > 0) || setBannerById[0].AttachmentId.length > 0) {
          if (setBannerById[0].AttachmentId.length > 0) {
            let arrn = await getDocumentLinkByIDarr(sp, setBannerById[0].AttachmentId[0], setBannerById[0].ID);
            //let arraynew: any[];
            //arraynew.push(arrn)
            console.log("arrrrrrnh", arrn);
            // if (setBannerById[0].Status != "Rework") {
            setAttachmentarr(arrn);
            // }
          }
          if ((setBannerById[0].PreviousAttachmentID != null || setBannerById[0].PreviousAttachmentID != undefined
            || setBannerById[0].PreviousAttachmentID != "") && setBannerById[0].PreviousAttachmentID > 0) {
            setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].PreviousAttachmentID, setBannerById[0].ID))
          }

        }
        let sharewithuser: any[] = [];

        if (setBannerById[0].PreparedBy && setBannerById[0].PreparedBy?.length > 0) {
          sharewithuser = setBannerById[0].PreparedBy?.map((approver: any) => ({
            value: approver.ID,
            label: approver.Title,
            UserName: approver.Title,
            UserEmail: approver.EMail
          }));
        }
        //  else {
        //   sharewithuser = [{
        //     value: currentUser.Id,
        //     label: currentUser.Title,
        //     UserName: currentUser.Title
        //   }];
        // }

        setSharewithusers(sharewithuser);

        if (ProcessItemId && ProcessItemId.Level === 0 && ProcessItemId.CurrentUserRole === "OES" && ProcessItemId.IsInitiator == "No") {
          const ApprowData: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_ChangeDocument, setBannerById[0].DocumentCode)

          if (ApprowData.length > 0 && ApprowData.length == sharewithuser.length) {

            const EditApprowData = ApprowData.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id || 0, // Assuming role comes from ApproverRole
              level: item.Level || 1, // Default to 1 if missing
              Responsibility: item.Responsibility || "",
              IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
              IsPreparedBy: item.Responsibility == "Preparer" ? true : false,
              approvers: item.Approvers?.map((approver: any) => ({
                value: approver.Id,
                label: approver.Title,

              })) || []
            }));

            setForwardToArr(EditApprowData);
            setForwardToArrEdit(EditApprowData);

          }
          if (ApprowData.length > 0 && ApprowData.length < sharewithuser.length) {
            // Flatten all approver IDs from ApprowData
            // const existingApproverIds = ApprowData
            //   .flatMap(item => item.approvers.map((approver: any) => approver.value));

            const approverIdArrays = ApprowData.map(item =>
              (item.ApproversId[0])
            );

            // Step 2: Flatten the array of ID arrays manually using reduce
            const existingApproverIds = approverIdArrays.reduce((acc, val) => acc.concat(val), []);

            // Step 3: Filter sharewithuser to get only users NOT in existingApproverIds
            const missingUsers = (sharewithuser || []).filter(user =>
              !existingApproverIds.includes(user.value)
            );
            const EditApprowData = ApprowData.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id || 0, // Assuming role comes from ApproverRole
              level: item.Level || 1, // Default to 1 if missing
              Responsibility: item.Responsibility || "",
              IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
              IsPreparedBy: item.Responsibility == "Preparer" ? true : false,
              approvers: item.Approvers?.map((approver: any) => ({
                value: approver.Id,
                label: approver.Title,

              })) || []
            }));

            // Create new rows for missing users
            // const missingApproverRows: ForwardTo[] = missingUsers.map((user: any, index: number) => ({
            //   id: 0,
            //   leveltype: "One",
            //   role: 0,
            //   level: ApprowData.length + index + 1, // Continue levels from existing data
            //   Responsibility: "Preparer",
            //   IsSignatureRequired: true,
            //   IsPreparedBy: true,
            //   approvers: [
            //     {
            //       value: user.value,
            //       label: user.label
            //     }
            //   ],
            //   roleError: false,
            //   approverError: false,
            //   responsibilityerror: false,
            //   typeError: false,
            //   rowError: false
            // }));
            const buildMissingApproverRows = async (): Promise<ForwardTo[]> => {
              const missingApproverRows: ForwardTo[] = await Promise.all(
                missingUsers.map(async (user: any, index: number) => {
                  // Fallback to user.value if login name format not available
                  const userLogin = user.UserEmail || user.UserEmail || `i:0#.f|membership|${user.UserEmail}`;

                  let designation = "";

                  try {
                    const profile = await sp.profiles.getPropertiesFor(userLogin);
                    designation = profile.UserProfileProperties.find((p: any) => p.Key === "Title")?.Value || "";
                  } catch (err) {
                    console.warn(`Could not retrieve profile for ${userLogin}`, err);
                  }
                  let role = setRolesValue.filter((role: any) => role.label == designation)[0]?.value
                  return {
                    id: 0,
                    leveltype: "One",
                    role: role || 0, // Use job title as role if found
                    level: ApprowData.length + index + 1,
                    Responsibility: "Preparer",
                    IsSignatureRequired: true,
                    IsPreparedBy: true,
                    approvers: [
                      {
                        value: user.value,
                        label: user.label
                      }
                    ],
                    roleError: false,
                    approverError: false,
                    responsibilityerror: false,
                    typeError: false,
                    rowError: false
                  };
                })
              );

              return missingApproverRows;
            };

            const missingRows = await buildMissingApproverRows();

            const updatedApprowData: ForwardTo[] = [...EditApprowData, ...missingRows];

            setForwardToArr(updatedApprowData);
            setForwardToArrEdit(updatedApprowData);
          }

          if (ApprowData.length == 0 && sharewithuser?.length > 0) {
            // Remove Requester from the sharewithuser list
            const filteredUsers = sharewithuser.filter(user => user.value !== setBannerById[0].RequesterNameId);

            // Create EditApprowData based on filtered users
            // const EditApprowDataPre: ForwardTo[] = filteredUsers && filteredUsers.length > 0 ? filteredUsers.map((user: any, index: number) =>
            //    ({
            //   id: 0, // You can use user.value if a unique ID is needed
            //   leveltype: "One", // or whatever default logic you prefer
            //   role: 0, // Role is blank as per requirement
            //   level: index + 1, // Default level
            //   Responsibility: "Preparer",
            //   IsSignatureRequired: true,
            //   IsPreparedBy: true,
            //   approvers: [
            //     {
            //       value: user.value,
            //       label: user.label
            //     }
            //   ]
            // })) : [
            //   {
            //     id: 0,
            //     role: 0,
            //     level: 1,
            //     leveltype: "One",
            //     Responsibility: "Signer",
            //     IsSignatureRequired: true,
            //     IsPreparedBy: false,
            //     approvers: [],
            //     roleError: false,
            //     approverError: false,
            //     responsibilityerror: false,
            //     typeError: false,
            //     rowError: false
            //   }
            // ];
            const buildEditApprowDataPre = async (): Promise<ForwardTo[]> => {
              if (!filteredUsers || filteredUsers.length === 0) {
                return [
                  {
                    id: 0,
                    role: 0,
                    level: 1,
                    leveltype: "One",
                    Responsibility: "Signer",
                    IsSignatureRequired: true,
                    IsPreparedBy: false,
                    approvers: [],
                    roleError: false,
                    approverError: false,
                    responsibilityerror: false,
                    typeError: false,
                    rowError: false
                  }
                ];
              }

              // Get web URL once before mapping
              const web = await sp.web.select("Url")();
              console.log("Web Absolute URL:", web.Url);

              const rows: ForwardTo[] = await Promise.all(
                filteredUsers.map(async (user: any, index: number) => {
                  const loginName = `i:0#.f|membership|${user.UserEmail}`;
                  const encodedLogin = encodeURIComponent(loginName);
                  const url = `${web.Url}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodedLogin}'`;

                  let designation = "";

                  try {
                    const res = await fetch(url, {
                      method: "GET",
                      headers: {
                        "Accept": "application/json;odata=verbose"
                      }
                    });
                    const data = await res.json();
                    const props = data.d.UserProfileProperties.results;
                    const title = props.find((p: any) => p.Key === "Title")?.Value;
                    // You can also get department or other props if needed here
                    designation = title || "";
                    console.log("Job Title:", title);
                  } catch (err) {
                    console.error("Error getting user profile:", err);
                  }

                  const role = setRolesValue.find((role: any) => role.label === designation)?.value || 0;

                  return {
                    id: 0,
                    leveltype: "One",
                    role: role,
                    level: index + 1,
                    Responsibility: "Preparer",
                    IsSignatureRequired: true,
                    IsPreparedBy: true,
                    approvers: [
                      {
                        value: user.value,
                        label: user.label
                      }
                    ],
                    roleError: false,
                    approverError: false,
                    responsibilityerror: false,
                    typeError: false,
                    rowError: false
                  };
                })
              );

              return rows;

            };
            const EditApprowDataPre = await buildEditApprowDataPre();
            setForwardToArr(EditApprowDataPre);
            setForwardToArrEdit(EditApprowDataPre);
          }
          if (ApprowData.length > 0 && ApprowData.length > sharewithuser.length) {
            const EditApprowData1 = ApprowData.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id, // Assuming role comes from ApproverRole
              level: item.Level, // Default to 1 if missing
              Responsibility: item.Responsibility || "",
              IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
              IsPreparedBy: item.Responsibility == "Preparer" ? true : false,
              approvers: item.Approvers?.map((approver: any) => ({
                value: approver.Id,
                label: approver.Title,

              }))
            }));
            //setdisabledforwardarr(true);
            setForwardToArr(EditApprowData1);
            setForwardToArrEdit(EditApprowData1);
          }
          // MainListID
        }

        if (ProcessItemId && ProcessItemId.CurrentUserRole !== "OES" && ProcessItemId.IsInitiator == "No") {
          const ApprowData1: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_ChangeDocument, setBannerById[0].DocumentCode)

          if (ApprowData1.length > 0) {

            const EditApprowData1 = ApprowData1.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id, // Assuming role comes from ApproverRole
              level: item.Level, // Default to 1 if missing
              Responsibility: item.Responsibility || "",
              IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
              IsPreparedBy: item.Responsibility == "Preparer" ? true : false,
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
        if (ProcessItemId && ProcessItemId.CurrentUserRole == "Initiator" && ProcessItemId.IsInitiator == "Yes" && ProcessItemId.IsRework == "Yes") {
          const ApprowData1: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_ChangeDocument, setBannerById[0].DocumentCode)

          if (ApprowData1.length > 0) {

            const EditApprowData1 = ApprowData1.map((item: any) => ({
              id: item.ID,
              leveltype: item.LevelType,
              role: item.ApproverRole?.Id, // Assuming role comes from ApproverRole
              level: item.Level, // Default to 1 if missing
              Responsibility: item.Responsibility || "",
              IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
              IsPreparedBy: false,
              approvers: item.Approvers?.map((approver: any) => ({
                value: approver.Id,
                label: approver.Title,

              }))
            }));
            //setdisabledforwardarr(true);
            setForwardToArr(EditApprowData1);
            setForwardToArrEdit(EditApprowData1);

          }

          // MainListID
        }
        setissueNo(setBannerById[0].IssueNumber);
        setserialNo(setBannerById[0].SerialNumber);
        setFormData(prevData => ({
          ...prevData,
          FileName: setBannerById[0].FileName,
          IssueNumber: setBannerById[0].IssueNumber,
          ReferenceNumber: setBannerById[0].ReferenceNumber,
          RevisionNumber: setBannerById[0].RevisionNumber,
          ChangeRequestID: setBannerById[0].ID,
          RequestDate: setBannerById[0].RequestDate,
          IssueDate: setBannerById[0].IssueDate,
          LocationId: setBannerById[0].LocationId,
          filename: setBannerById[0].FileName,
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
          PreviousAttachmentID: setBannerById[0].PreviousAttachmentID,
          AttachmentJson: setBannerById[0].AttachmentJson,
          Status: setBannerById[0].Status,
          RequesterName: setBannerById[0].Title,
          Remark: setBannerById[0].Remarks,
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

        if (((setBannerById[0].PreviousAttachmentID != null || setBannerById[0].PreviousAttachmentID != undefined
          || setBannerById[0].PreviousAttachmentID != "") && setBannerById[0].PreviousAttachmentID > 0) || setBannerById[0].AttachmentId.length > 0) {
          if (setBannerById[0].AttachmentId.length > 0) {

            let arrn = await getDocumentLinkByIDarr(sp, setBannerById[0].AttachmentId[0], setBannerById[0].ID);
            //let arraynew: any[];
            //arraynew.push(arrn)
            console.log("arrrrrrnty", arrn);

            // if (setBannerById[0].Status != "Rework") {
            setAttachmentarr(arrn);
            // }
          }
          if ((setBannerById[0].PreviousAttachmentID != null || setBannerById[0].PreviousAttachmentID != undefined
            || setBannerById[0].PreviousAttachmentID != "") && setBannerById[0].PreviousAttachmentID > 0) {
            setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].PreviousAttachmentID, setBannerById[0].ID));
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
    // setLoading(false);
    //#endregion


  };
  const onSelectDocCode = async (selectedList: any) => {
    debugger
    setshowpreviousattachment(true);
    let IsdoccodeAllowed = await CheckifDocumentisApproved(sp, selectedList.DocumentCode);
    if (!IsdoccodeAllowed) {
      // Show popup and return early if not approved
      await Swal.fire({
        //icon: 'warning',
        //title: 'Document not approved',
        text: `This document code "${selectedList.DocumentCode}" is already in progress, so you cannot raise a request against it.`,
        confirmButtonText: 'OK'
      });
      return; // Stop further execution
    }
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
        filename: selectedList.filename,
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
        AttachmentJson: selectedList.AttachmentJson,
        PreviousAttachmentID: selectedList.AttachmentId.length > 0 ? selectedList.AttachmentId[0] : "", // Assuming AttachmentId is an array and we want the first element
        // Format as YYYY-MM-DD
      }));
      setSelectedOption(selectedList);
      const rowData: any[] = await getItemByIDChangeRequest(sp, Number(selectedList.ID)) //baseUrl
      const initialRows = rowData.map((item: any) => ({
        id: item.Id,
        description: item.ChangeDescription,
        reason: item.ReasonforChange,
      }));
      setLoading(false);
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
        let arrn = await getDocumentLinkByIDarr(sp, selectedList.AttachmentId[0], selectedList.ID);
        //let arraynew: any[];
        //arraynew.push(arrn)
        //console.log("arrrrrrn56", arrn);
        //setAttachmentarr(arrn);
        setDocumentLink(await getDocumentLinkByID(sp, selectedList.AttachmentId[0], selectedList.ID))
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
    setshowpreviousattachment(false);
    setselectedCheckboxIds([]);
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
    if (selectedList?.requestcode != "New") {
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
      RequestTypeId: selectedList?.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionReq(selectedList);  // Set the selected users
  };
  const onSelectAmend = (selectedList: any) => {
    console.log(selectedList, "selectedListamenddd");
    setFormData(prevData => ({
      ...prevData,
      AmendmentTypeId: selectedList?.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionAmend(selectedList);  // Set the selected users
  };
  const onSelectDepart = (selectedList: any) => {
    console.log(selectedList, "selectedListadepartttt");
    setFormData(prevData => ({
      ...prevData,
      DepartmentId: selectedList?.value
      // Format as YYYY-MM-DD
    }));
    let doccodearrew: any;
    if (selectedList) {
      doccodearrew = doccoderows.filter((x: any) => x.DepartmentId == selectedList?.value)
    }

    setRows(doccodearrew);

    setSelectedOptionDepart(selectedList);  // Set the selected users
  };
  const onSelectTemplatetype = (selectedList: any) => {
    console.log(selectedList, "selectedListadtemplatetype");
    setFormData(prevData => ({
      ...prevData,
      TemplateTypeId: selectedList?.value
      // Format as YYYY-MM-DD
    }));


    setSelectedOptionTemplate(selectedList);  // Set the selected users
  };
  const onSelectClassification = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      ClassificationId: selectedList?.value
      // Format as YYYY-MM-DD
    }));
    setSelectedOptionClassification(selectedList);  // Set the selected users
  };

  const onSelectLocation = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      LocationId: selectedList?.value
      // Format as YYYY-MM-DD
    }));
    setselectedOptionLoc(selectedList);  // Set the selected users
  };
  const onSelectsharewith = (selectedOptions: any) => {
    // const newSelections = selectedOptions || [];
    // const allOptions = [...sharewithusers, ...newSelections];

    // const uniqueOptions = allOptions.filter(
    //     (option, index, self) =>
    //         index === self.findIndex((o) => o.value === option.value)
    // );

    // setSharewithusers(uniqueOptions);
    const uniqueOptions = (selectedOptions || []).filter(
      (option: any, index: any, self: any) =>
        index === self.findIndex((o: any) => o.value === option.value)
    );
    setSharewithusers(uniqueOptions);

  };
  const onSelectCustodian = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      CustodianId: selectedList?.value
      // Format as YYYY-MM-DD
    }));
    setselectedOptionCusto(selectedList);  // Set the selected users
  };

  const onSelectDocumentType = (selectedList: any) => {
    console.log(selectedList, "selectedListclasss");
    setFormData(prevData => ({
      ...prevData,
      DocumentTypeId: selectedList?.value
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
  const handleChangeResp = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
    event.preventDefault();
    const updatedArr = forwardToArr.map(row =>
      // row.level === lvl ? { ...row, Responsibility: event.target.value} : row

      row.level === lvl ? { ...row, Responsibility: event.target.value, IsSignatureRequired: true } : row
    );
    //   setApprovalType(event.target.value);
    setForwardToArr(updatedArr);
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
    if (event.target.value == "Select") {
      return
    } else {
      const updatedArr1 = forwardToArr.map(row =>
        row.level === lvl ? { ...row, leveltype: event.target.value } : row
      );
      setForwardToArr(updatedArr1);
    }
    //setUserRoles(UserRoles.filter((x: any) => x.label !== event.target.value))
  };

  // const handleAddRow = () => {
  //     const newRow: ForwardTo = { id: Date.now(), role: 0, level: "", approvers: [] };
  //     setForwardToArr([...forwardToArr, newRow]);
  // };
  const handleAddRow = () => {
    setForwardToArr((prev) => [
      ...prev,
      { id: 0, role: 0, level: prev.length + 1, approvers: [], leveltype: "One", Responsibility: "Signer", IsSignatureRequired: true, IsPreparedBy: false, }
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
    const path1 = window.location.href;
    if (path1.includes("/view/") || path1.includes("/approve/") || path1.includes("/view/")) {
      setLoading(true);
    } else { setLoading(false); }


    ApiCallFunc().then((x) => {
      setLoading(false);
    });

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
  const cancelModalAction = (refresh?: boolean,) => {
    debugger
    setredirecturl(window.location.href);
    //setShowfileNew(false);
    setShowModalTemplateDoc(false);
    Showfile = false;
  }
  const renderIframeUrl = (fileUrl: string): string => {
    const isOfficeFile = /\.(docx?|xlsx?|pptx?)$/i.test(fileUrl);

    if (isOfficeFile) {
      return `https://edcadae.sharepoint.com/sites/ededms/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=embedview`;
    }

    if (/\.pdf$/i.test(fileUrl)) {
      return `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(fileUrl)}`;
    }

    // fallback
    return fileUrl;
  };

  const OpenFile = (obj: any, sts: string) => {
    debugger

    const fileUrl = `${Tenant_URL}${obj?.FileRef != "" ? obj.FileRef : obj.fileUrl}`;
    if (sts == "Open") {
      Showfile = true;
    }
    let redirecturl = obj?.FileRef != "" ? obj.FileRef : obj.fileUrl;
    //const viewerUrl = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=view`
    //const viewerUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(redirecturl)}`
    //setredirecturl(viewerUrl);
    //const fileUrl = `${Tenant_URL}${obj.FileRef}`;

    console.log("ttrtrtrtt", obj);
    const tenantUrl = "https://edcadae.sharepoint.com";
    const docurl = "/sites/ededms/ImportantDocuments/SomeFolder/Example.pdf";

    // Extract parent folder from docurl
    const parent = docurl.substring(0, docurl.lastIndexOf("/") + 1);

    // Construct the final redirect URL
    const redirectURLpdf = `${tenantUrl}/sites/ededms/ImportantDocuments/Forms/AllItems.aspx?id=${encodeURIComponent(docurl)}&parent=${encodeURIComponent(parent)}&p=true&ga=1`;

    console.log(redirectURLpdf);

    //const fileUrl = `${Tenant_URL}${obj?.FileRef != "" ? obj.FileRef : obj.fileUrl}`;
    if (sts == "Open") {
      if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
        const viewerUrlppt = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview`;
        //const viewerUrldoc = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=view&wdHideGridlines=true&wdHideToolbar=true&embedded=true`;
        //const viewerUrlxlsx = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview&wdAllowInteractivity=false&embedded=true`;

        setredirecturl(viewerUrlppt);


        //window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=view`);

      } else {
        setredirecturl(fileUrl);
        //window.open(fileUrl, "_blank"); // Open PDF and other files normally
      }

    } else if (sts == "Download") {
      // const link = document.createElement("a");
      // link.href = fileUrl;
      // link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
      // document.body.appendChild(link);
      // link.click();
      // document.body.removeChild(link);
      // const serverRelativeUrl = obj.FileRef; // e.g. "/sites/test/Shared Documents/sample.docx"
      // const fileUrls = `${SITE_URL}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(serverRelativeUrl)}')/$value`;

      // const link = document.createElement("a");
      // link.href = fileUrls;
      // link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
      // document.body.appendChild(link);
      // link.click();
      // document.body.removeChild(link);
      const serverRelativeUrl = obj?.FileRef || ""; // Ensure serverRelativeUrl is defined
      const fileUrls = `${SITE_URL}${serverRelativeUrl}`;
      downloadFile(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));
      //directDownload(fileUrls);
      //downloadFileNew(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));
      //downloadWithPnP(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));
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
  const downloadFile = async (serverRelativeUrl: string | number | boolean, fileName: string) => {

    try {

      const response = await fetch(

        `${SITE_URL}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(serverRelativeUrl)}')/$value`,

        {

          method: "GET",

          headers: {

            "Accept": "application/octet-stream"

          }

        }

      );

      if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

      const blob = await response.blob();

      const url = window.URL.createObjectURL(blob);

      const link = document.createElement("a");

      link.href = url;

      link.download = fileName || "download";

      document.body.appendChild(link);

      link.click();

      document.body.removeChild(link);

      window.URL.revokeObjectURL(url);

    } catch (err) {

      console.error("File download failed:", err);

    }

  };
  const directDownload = (filePath: any) => {
    const link = document.createElement("a");
    link.href = `${filePath}?download=1`;
    link.target = "_blank";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
  const downloadFileNew = async (serverRelativeUrl: string | number | boolean, fileName: string) => {
    const endpoint = `${SITE_URL}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(serverRelativeUrl)}')/$value`;

    const response = await fetch(endpoint, {
      method: "GET",
      headers: {
        "Accept": "application/octet-stream"
      },
      credentials: "include"
    });

    if (!response.ok) throw new Error(`Error fetching file: ${response.statusText}`);

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  };
  const downloadWithPnP = async (serverRelativeUrl: any, fileName: string) => {
    const fileItem = await (sp.web as any).getFileByServerRelativeUrl(serverRelativeUrl).getItem("ID", "AuthorId", "Modified");
    console.log(fileItem, 'fileItem');
    const file = await (sp.web as any).getFileByServerRelativeUrl(serverRelativeUrl).getBuffer();
    const blob = new Blob([file]);
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.click();
    window.URL.revokeObjectURL(url);
  };
  // Usage example:



  const OpenFileTemplate = (obj: any, sts: string) => {
    debugger
    setShowModalTemplateDoc(true);
    if (sts == "Open") {
      Showfile = true;
    }
    console.log("ttrtrtrtt", obj)
    const fileUrl = `${Tenant_URL}${obj?.FileRef != "" ? obj.FileRef : obj.fileUrl}`;
    if (sts == "Open") {
      if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
        const viewerUrl = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview`;

        //window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=view`);
        setredirecturl(viewerUrl);
      } else {
        setredirecturl(fileUrl);
        //window.open(fileUrl, "_blank"); // Open PDF and other files normally
      }

    } else if (sts == "Download") {
      // const serverRelativeUrl = obj?.FileRef || ""; // Ensure serverRelativeUrl is defined
      // downloadFile(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));
      const serverRelativeUrl = obj?.FileRef || ""; // Ensure serverRelativeUrl is defined
      const fileUrls = `${SITE_URL}${serverRelativeUrl}`;
      downloadFile(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));
      // directDownload(fileUrls);
      //downloadFileNew(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));
      //downloadWithPnP(serverRelativeUrl, obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name));

      // const serverRelativeUrl = obj.FileRef; // e.g. "/sites/test/Shared Documents/sample.docx"
      // const fileUrls = `${SITE_URL}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(serverRelativeUrl)}')/$value`;

      // const link = document.createElement("a");
      // link.href = fileUrls;
      // link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
      // document.body.appendChild(link);
      // link.click();
      // document.body.removeChild(link);
      // const link = document.createElement("a");
      // link.href = fileUrl;
      // link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
      // document.body.appendChild(link);
      // link.click();
      // document.body.removeChild(link);

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

  const cleanFileNameSave = async (filename: string) => {
    // Match a 14-digit datetime suffix before the file extension
    const datetimePattern = /_\d{14}(?=\.[^.]+$)/;

    if (datetimePattern.test(filename)) {
      return filename.replace(datetimePattern, '');
    }

    return filename;
  }

  const ApprovalTypeOptions = [
    { value: 'One', label: 'Anyone' },
    { value: 'All', label: 'Everyone' }

  ];
  const addCancelReason = () => {
    setcancellReason([...cancellReason, { id: 0, description: "", reason: "" }]);
  };
  const updatedigisignnew = async () => {
    let items = await updateDigitalsign("ChangeRequestList", sp, DigitalsignID, editItemID);
    if (items) {
      sethidedigisign(true);
      const isRecordExist = await getdigitalsignaturerequestbyIDYes("ChangeRequestList", sp, Number(editItemID));
      if (isRecordExist == "Yes" || isRecordExist == "NoRecord") {
        setshowdigisign(false);
      } else {
        setshowdigisign(true);
      }

    }
  }

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

  const validateForm = async (fmode: FormSubmissionMode) => {
    debugger
    const { RequesterName, RequesterDesignation, RequestDate, DocumentCode, IssueNumber, RevisionNumber, ReferenceNumber } = formData;
    // const { description } = richTextValues;
    let valid = true;
    let valid1 = true;
    // let validateOverview:boolean = false;
    // let validatetitlelength = false;
    // let validateTitle = false;
    setsharewitherr(false);
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
    setfilenameerr(false);
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
      if (selectedOptionReq && selectedOptionReq.requestcode != "New" && !selectedOption) {
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
      if (!sharewithusers || sharewithusers.length === 0) {
        setsharewitherr(true);
        valid = false;
      }
      if (cancellReason.length > 0) {
        let valid = true;
        const errors = cancellReason.map((row) => {
          const descEmpty = !row.description || row.description.trim() === "";
          const reasonEmpty = !row.reason || row.reason.trim() === "";

          if (descEmpty || reasonEmpty) {
            valid = false;
          }

          return {
            descriptionError: descEmpty,
            reasonError: reasonEmpty,
          };
        });

        setRowErrors(errors); // Update state with per-row error flags
        valid1 = valid;
      }
      if (formData.filename == "") {
        setfilenameerr(true);
        valid = false;
      }

      // if (cancellReason.length > 0) {
      //   let descriptionError = false;
      //   let reasonError = false;

      //   cancellReason.forEach((row: any) => {
      //     if (row.description === null || row.reason === null) {
      //       if (row.description === null) {
      //         descriptionError = true;
      //       }
      //       if (row.reason === null) {
      //         reasonError = true;
      //       }
      //     } else {
      //       if (row.description != null && row.description.trim() === "") {
      //         descriptionError = true;
      //       }
      //       if (row.reason != null && row.reason.trim() === "") {
      //         reasonError = true;
      //       }
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
      if (cancellReason.length == 0) {
        setchangedescriptionerr(true);
        setchangereasonerr(true);
        valid1 = false;
      }
      if (selectedCheckboxIds.length == 0) {
        setchangerequesttypeerr(true);
        valid = false;
      }
      // if (Attachmentarr.length == 0 && DocumentLink == null) {
      //   setattachmenterr(true);
      //   valid = false;
      // }
      if (Attachmentarr.length == 0) {
        setattachmenterr(true);
        valid = false;
      }
      // return true;

      setValidSubmit(valid);
      setValidCancelReason(valid1);
    }
    else {
      setsharewitherr(false);
      if (!RequesterName) {
        valid = false;
      }
      if (!selectedOptionReq) {
        setrequesttypeerr(true);
        valid = false;
      }
      if (selectedOptionReq && selectedOptionReq.requestcode != "New" && !selectedOption) {
        setdocumentcodeerr(true);
        valid = false;
      }
      if (!SelectedOptionDepart) {
        setdepartmenterr(true);
        valid = false;
      }
      if (formData.filename == "") {
        setfilenameerr(true);
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

  const updateFilename = async (filename: string, docCode: string, newIssueNo: string, newRevNo: string) => {
    if (!filename.includes(docCode)) return filename; // Return as-is if docCode not found

    const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
    const match = filename.match(pattern);

    if (!match) return filename; // If it doesn't match expected structure, return original

    const [, code, , , rest] = match;

    return `${code}-${newIssueNo}-${newRevNo}-${rest}`;
  }


  const formatDateTime = async (date: Date) => {
    const pad = (n: number) => n.toString().padStart(2, '0');

    const dd = pad(date.getDate());
    const MM = pad(date.getMonth() + 1); // Month is 0-based
    const yyyy = date.getFullYear();
    const hh = pad(date.getHours());
    const mm = pad(date.getMinutes());
    const ss = pad(date.getSeconds());

    return `${dd}${MM}${yyyy}${hh}${mm}${ss}`;
  };


  const handleFormSubmit = async () => {
    debugger
    scrollToTop();
    let url = window.location.href.split('/sites/')[0];
    console.log("topp submit", editItemID, cancellReason, DocumentLink, Attachmentarr);
    if (await validateForm(FormSubmissionMode.SUBMIT)) {
      debugger
      let changerequestdata: any = [];
      const selectedTemplate = SelectedOptionTemplate?.label;
      if (selectedTemplate != "ChangeRequest") {
        changerequestdata = await getchangerequesttemp(sp);
      }

      let serialnumber = await getDocumentCodeselected(sp, selectedOptionLoc.locationId, selectedOptionCusto.custodianId, selectedOptionDoctype.documentTypeId)
      let issueno = "";
      let serialno = "";
      let revisionno = "";
      if (selectedOptionReq.requestcode == "New") {
        if (serialnumber.length > 0) {
          issueno = formData.Status == "Rework" ? formData.IssueNumber : serialnumber[0].IssueNo;
          serialno = formData.Status == "Rework" ? formData.SerialNumber : (Number(serialnumber[0].SerialNo) + 1).toString();
          revisionno = formData.Status == "Rework" ? formData.RevisionNumber : serialnumber[0].RevisionNo;
          setissueNo(formData.Status == "Rework" ? formData.IssueNumber : serialnumber[0].IssueNo);
          setserialNo(formData.Status == "Rework" ? formData.SerialNumber : serialno);
          setrevisionNo(formData.Status == "Rework" ? formData.RevisionNumber : serialnumber[0].RevisionNo);
        } else {
          issueno = issueNo == "" || issueNo == null ? "01" : issueNo;
          serialno = serialNo == "" || issueNo == null ? "01" : serialNo;
          revisionno = revisionNo == "" || issueNo == null ? "00" : revisionNo;
          setissueNo(issueNo == "" || issueNo == null ? "01" : issueNo);
          setserialNo(serialNo == "" || issueNo == null ? "01" : serialNo);
        }
      } else {
        issueno = formData.Status == "Rework" ? formData.IssueNumber : (Number(selectedOption?.IssueNumber) + 1).toString();
        serialno = formData.Status == "Rework" ? formData.SerialNumber : selectedOption?.SerialNumber;
        revisionno = formData.Status == "Rework" ? formData.RevisionNumber : (Number(selectedOption?.RevisionNumber) + 1).toString();
        setissueNo(issueno);
        setserialNo(serialno);
        setrevisionNo(revisionno);
      }
      let ApprovedChanedoc: any = [];
      ApprovedChanedoc = selectedOptionReq?.requestcode == "Edit" && await getDocumentCodeselectedApproved(sp, formData.DocumentCode, formData.LocationId, formData.CustodianId, formData.DocumentTypeId);
      let finalissuedateNew: any;
      if (ApprovedChanedoc.length > 0) {
        finalissuedateNew = new Date(ApprovedChanedoc[0].IssueDate).toISOString();
      } else {
        finalissuedateNew = new Date().toISOString();
      }
      let doccode = selectedOptionReq?.requestcode == "New" && formData.Status != "Rework" ? await generateDocCode(serialno) : selectedOption?.DocumentCode;
      let referencecode = await generateReferenceCode(serialno, issueno);

      let finalrevisiondate: any;
      let finalissuedate: any;
      if (changerequestdata && changerequestdata.length > 0) {
        finalrevisiondate = changerequestdata[0]?.RevisionDate == null || changerequestdata[0]?.RevisionDate == undefined ? undefined : new Date(changerequestdata[0]?.RevisionDate).toISOString();
        finalissuedate = changerequestdata[0]?.IssueDate == null || changerequestdata[0]?.IssueDate == undefined ? undefined : new Date(changerequestdata[0]?.IssueDate).toISOString();
        console.log("vbvbvb", changerequestdata[0]?.RevisionDate);
      }

      console.log("doccode doccode", doccode, referencecode);
      const now = new Date();
      const dateTimeSuffix = await formatDateTime(now);
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
            const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/ChangeRequestDocs');
            let docCode = selectedOptionReq.requestcode == "New" ? doccode : selectedOption?.DocumentCode;
            let filenamenew: any;
            let newfileName: any;
            debugger
            if (Attachmentarr.length > 0) {
              if (Attachmentarr[0]?.files?.length > 0) {
                for (const file of Attachmentarr[0].files) {
                  //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                  //const newFileName = await getNewFileName(file.name);
                  newfileName = docCode + "-" + issueno + "-" + revisionno + "-" + file.name;

                  if (file.name.includes(docCode)) {

                    const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match = file.name.match(pattern);

                    if (!match) return file.name; // If it doesn't match expected structure, return original

                    const [, code, , , rest] = match;

                    filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                  }

                  let newfileNameNew = file.name.includes(docCode) ? filenamenew : newfileName;
                  newfileNameNew = await cleanFileNameSave(newfileNameNew);
                  DocumentName = newfileNameNew;
                  const finalFileName = `${newfileNameNew.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNew.substring(newfileNameNew.lastIndexOf('.'))}`;
                  const fileAddResult = await folder.files.addChunked(finalFileName, file);
                  const fileNew = fileAddResult.file;
                  let filenamenew1: any;
                  //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                  const newfileNameNewN1 = docCode + "-" + issueno + "-" + revisionno + "-" + fileAddResult.data.Name;
                  if (fileAddResult.data.Name.includes(docCode)) {
                    const pattern1 = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match1 = fileAddResult.data.Name.match(pattern1);

                    if (!match1) return fileAddResult.data.Name; // If it doesn't match expected structure, return original

                    const [, code1, , , rest1] = match1;

                    filenamenew1 = `${code1}-${issueno}-${revisionno}-${rest1}`;
                  }


                  let newfileNameNewN = fileAddResult.data.Name.includes(docCode) ? filenamenew1 : newfileNameNewN1;
                  newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                  const finalFileNameN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                  const documentName = newfileNameNewN;
                  bannerImageArray = fileAddResult;
                  // Get the item ID for the uploaded file
                  const currentItemId = await fileNew.getItem<{ Id: number }>();
                  const itemId = currentItemId.Id;
                  await currentItemId.update({
                    FileName: finalFileNameN, // Assuming FileName is the internal name of the column
                    DocumentCode: selectedOptionReq?.requestcode == "New" ? doccode : selectedOption?.DocumentCode,
                  });

                  // Save the document ID for the attachment field in ChangeRequestList
                  attachmentIds.push(itemId);
                }
              } else if (Attachmentarr[0].ID > 0) {

                const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
                debugger
                let filenamenew: any;
                const newfileNameNew1 = docCode + "-" + issueno + "-" + revisionno + "-" + item.File.Name;
                if (item.File.Name.includes(docCode)) {
                  const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                  const match = item.File.Name.match(pattern);

                  if (!match) return item.File.Name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                }


                let newfileNameNewN = item.File.Name.includes(docCode) ? filenamenew : newfileNameNew1;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const oldFilePath = item?.File?.ServerRelativeUrl;
                const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
                //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);

                const finalFileNameN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const newFilePath = `${folderPath}/${finalFileNameN}`;

                // 2. Use moveByPath to rename the file
                await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
                // 2. Move (rename) file
                // await item.update({
                //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
                //   DocumentCode: docCode
                // });
                const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).update({
                  FileName: finalFileNameN, // Assuming FileName is the internal name of the column
                  DocumentCode: docCode
                })
                attachmentIds.push(Attachmentarr[0]?.ID);
              }
            } else {
              if (DocumentLink && DocumentLink?.files?.length > 0) {
                for (const file of DocumentLink.files) {
                  //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                  //const newFileName = await getNewFileName(file.name);
                  let filenamenew: any;
                  const newfileName = docCode + "-" + issueno + "-" + revisionno + "-" + file.name;
                  if (file.name.includes(docCode)) {
                    const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match = file.name.match(pattern);

                    if (!match) return file.name; // If it doesn't match expected structure, return original

                    const [, code, , , rest] = match;

                    filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                  }
                  let newfileNameNew = file.name.includes(docCode) ? filenamenew : newfileName;
                  newfileNameNew = await cleanFileNameSave(newfileNameNew);
                  const finalFileNameNew = `${newfileNameNew.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNew.substring(newfileNameNew.lastIndexOf('.'))}`;
                  DocumentName = finalFileNameNew;
                  const fileAddResult = await folder.files.addChunked(finalFileNameNew, file);
                  const fileNew = fileAddResult.file;
                  let filenamenew1: any;
                  //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                  if (fileAddResult.data.Name.includes(docCode)) {
                    const pattern1 = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match1 = fileAddResult.data.Name.match(pattern1);

                    if (!match1) return fileAddResult.data.Name; // If it doesn't match expected structure, return original

                    const [, code1, , , rest1] = match1;

                    filenamenew1 = `${code1}-${issueno}-${revisionno}-${rest1}`;
                  }
                  const newfileNameNewN1 = docCode + "-" + issueno + "-" + revisionno + "-" + fileAddResult.data.Name;


                  let newfileNameNewN = fileAddResult.data.Name.includes(docCode) ? filenamenew1 : newfileNameNewN1;
                  newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                  const finalFileNameNewN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                  const documentName = finalFileNameNewN;
                  bannerImageArray = fileAddResult;
                  // Get the item ID for the uploaded file
                  const currentItemId = await fileNew.getItem<{ Id: number }>();
                  const itemId = currentItemId.Id;
                  await currentItemId.update({
                    FileName: documentName, // Assuming FileName is the internal name of the column
                    DocumentCode: selectedOptionReq?.requestcode == "New" ? doccode : selectedOption?.DocumentCode,
                  });

                  // Save the document ID for the attachment field in ChangeRequestList
                  attachmentIds.push(itemId);
                }
              } else if (DocumentLink && DocumentLink.ID > 0) {

                const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
                debugger
                let filenamenew: any;
                const newfileNameNew1 = docCode + "-" + issueno + "-" + revisionno + "-" + item.File.Name;
                if (item.File.Name.includes(docCode)) {
                  const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                  const match = item.File.Name.match(pattern);

                  if (!match) return item.File.Name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                }


                let newfileNameNewN = item.File.Name.includes(docCode) ? filenamenew : newfileNameNew1;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const oldFilePath = item?.File?.ServerRelativeUrl;
                const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
                //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
                const finalFileNameNewN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const newFilePath = `${folderPath}/${finalFileNameNewN}`;

                // 2. Use moveByPath to rename the file
                await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
                // 2. Move (rename) file
                // await item.update({
                //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
                //   DocumentCode: docCode
                // });
                const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).update({
                  FileName: finalFileNameNewN, // Assuming FileName is the internal name of the column
                  DocumentCode: docCode
                })
              }
              attachmentIds.push(DocumentLink?.ID);
            }
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            const sharewithIds: any[] = [];
            sharewithusers.forEach((user: any) => {
              if (user?.value) {
                sharewithIds.push(user.value);
              }
            });
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            let arr = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById,
              //IssueDate: new Date(formData.IssueDate).toISOString(),
              LocationId: formData.LocationId,
              FileName: formData.filename,
              CustodianId: formData.CustodianId,
              SerialNumber: formData.Status == "Rework" ? formData.SerialNumber : Number(serialno),
              IssueNumber: formData.Status == "Rework" ? formData.IssueNumber : Number(issueno),
              RevisionNumber: formData.Status == "Rework" ? formData.RevisionNumber : Number(revisionno),
              //RevisionDate: formData.RevisionDate,
              IssueDate: finalissuedateNew,
              DocumentCode: formData.Status == "Rework" ? formData.DocumentCode : docCode,
              ReferenceNumber: formData.Status == "Rework" ? formData.ReferenceNumber : referencecode,
              RequestTypeId: formData.RequestTypeId,
              AmendmentTypeId: formData.AmendmentTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              ClassificationId: formData.ClassificationId,
              SubmiitedDate: new Date(formData.RequestDate).toISOString(),
              SubmitStatus: "Yes",
              Status: "Pending",
              IsRework: "No",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              DocumentName: attachmentIds.length != 0 ? DocumentName : formData.DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              CDocumentCode: selectedTemplate != "Change Request" && changerequestdata && changerequestdata?.length > 0 ? changerequestdata[0]?.DocumentCode : docCode,
              CIssueNumber: selectedTemplate != "Change Request" && changerequestdata && changerequestdata?.length > 0 ? Number(changerequestdata[0]?.IssueNo) : Number(issueno),
              CRevisionNumber: selectedTemplate != "Change Request" && changerequestdata && changerequestdata?.length > 0 ? Number(changerequestdata[0]?.RevisionNo) : Number(revisionno),
              CIssueDate: selectedTemplate != "Change Request" ? finalissuedate : finalissuedateNew,
              CRevisionDate: selectedTemplate != "Change Request" ? finalrevisiondate : undefined,
              //AttachmentId: selectedOptionReq.label == "New" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "New" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso,
              PreviousAttachmentID: ""
            }
            let arrework = {
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById,
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              IssueDate: finalissuedateNew,
              RequestTypeId: formData.RequestTypeId,
              AmendmentTypeId: formData.AmendmentTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              ClassificationId: formData.ClassificationId,
              SubmiitedDate: new Date(formData.RequestDate).toISOString(),
              SubmitStatus: "Yes",
              Status: "Pending",
              IsRework: "No",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              DocumentName: attachmentIds.length != 0 ? DocumentName : formData.DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso,
              PreviousAttachmentID: ""
            }
            const postResult = await updateItemChangeRequestList(formData.Status == "Rework" ? arrework : arr, sp, editItemID);
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
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
              }
            });
            //sessionStorage.removeItem("ChangeRequestId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
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
            const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/ChangeRequestDocs');
            let docCode = selectedOptionReq?.requestcode == "New" ? doccode : selectedOption?.DocumentCode;
            debugger
            if (Attachmentarr.length > 0) {
              if (Attachmentarr[0]?.files?.length > 0) {
                for (const file of Attachmentarr[0].files) {
                  let filenamenew: any;
                  //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                  const newfileNameNewX = docCode + "-" + issueno + "-" + revisionno + "-" + file.name;
                  if (file.name.includes(docCode)) {
                    const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match = file.name.match(pattern);

                    if (!match) return file.name; // If it doesn't match expected structure, return original

                    const [, code, , , rest] = match;

                    filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                  }

                  let newfileNameNew = file.name.includes(docCode) ? filenamenew : newfileNameNewX;

                  newfileNameNew = await cleanFileNameSave(newfileNameNew);
                  const finalFileNameNew = `${newfileNameNew.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNew.substring(newfileNameNew.lastIndexOf('.'))}`;
                  DocumentName = finalFileNameNew;
                  const fileAddResult = await folder.files.addChunked(finalFileNameNew, file);
                  //DocumentName = newfileNameNew;
                  //const fileAddResult = await folder.files.addChunked(newfileNameNew, file);
                  const fileNew = fileAddResult.file;
                  let filenamenew1: any
                  //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                  const newfileNameNewP = docCode + "-" + issueno + "-" + revisionno + "-" + fileAddResult.data.Name;
                  if (fileAddResult.data.Name.includes(docCode)) {
                    const pattern1 = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match1 = fileAddResult.data.Name.match(pattern1);

                    if (!match1) return fileAddResult.data.Name; // If it doesn't match expected structure, return original

                    const [, code1, , , rest1] = match1;

                    filenamenew1 = `${code1}-${issueno}-${revisionno}-${rest1}`;

                  }


                  let newfileNameNewN = fileAddResult.data.Name.includes(docCode) ? filenamenew1 : newfileNameNewP;
                  newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                  const finalFileNameNewN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                  const documentName = finalFileNameNewN;
                  //const documentName = fileAddResult.data.Name;
                  bannerImageArray = fileAddResult;
                  // Get the item ID for the uploaded file
                  const currentItemId = await fileNew.getItem<{ Id: number }>();
                  const itemId = currentItemId.Id;
                  await currentItemId.update({
                    FileName: documentName, // Assuming FileName is the internal name of the column
                    DocumentCode: selectedOptionReq.requestcode == "New" ? doccode : selectedOption?.DocumentCode,
                  });

                  // Save the document ID for the attachment field in ChangeRequestList
                  attachmentIds.push(itemId);
                }
              } else if (Attachmentarr[0].ID > 0) {

                const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
                debugger
                let filenamenew: any;
                const newfileNameNew1 = docCode + "-" + issueno + "-" + revisionno + "-" + item.File.Name;
                if (item.File.Name.includes(docCode)) {
                  const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                  const match = item.File.Name.match(pattern);

                  if (!match) return item.File.Name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                }


                let newfileNameNewN = item.File.Name.includes(docCode) ? filenamenew : newfileNameNew1;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const oldFilePath = item?.File?.ServerRelativeUrl;
                const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
                //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
                const finalFileNameNewN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const newFilePath = `${folderPath}/${finalFileNameNewN}`;

                // 2. Use moveByPath to rename the file
                await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
                // 2. Move (rename) file
                // await item.update({
                //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
                //   DocumentCode: docCode
                // });
                const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).update({
                  FileName: finalFileNameNewN, // Assuming FileName is the internal name of the column
                  DocumentCode: docCode
                })
                attachmentIds.push(Attachmentarr[0]?.ID);
              }
            } else {

              if (DocumentLink && DocumentLink?.files?.length > 0) {
                let filenamenew: any;
                for (const file of DocumentLink.files) {
                  //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                  const newfileNameNewX = docCode + "-" + issueno + "-" + revisionno + "-" + file.name;
                  if (file.name.includes(docCode)) {
                    const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match = file.name.match(pattern);

                    if (!match) return file.name; // If it doesn't match expected structure, return original

                    const [, code, , , rest] = match;

                    filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                  }
                  let newfileNameNew = file.name.includes(docCode) ? filenamenew : newfileNameNewX;
                  newfileNameNew = await cleanFileNameSave(newfileNameNew);
                  const finalFileNameNewN = `${newfileNameNew.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNew.substring(newfileNameNew.lastIndexOf('.'))}`;
                  DocumentName = finalFileNameNewN;
                  const fileAddResult = await folder.files.addChunked(finalFileNameNewN, file);
                  const fileNew = fileAddResult.file;
                  let filenamenew1: any;
                  //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                  const newfileNameNewP = docCode + "-" + issueno + "-" + revisionno + "-" + fileAddResult.data.Name;
                  if (fileAddResult.data.Name.includes(docCode)) {
                    const pattern1 = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                    const match1 = fileAddResult.data.Name.match(pattern1);

                    if (!match1) return fileAddResult.data.Name; // If it doesn't match expected structure, return original

                    const [, code1, , , rest1] = match1;

                    filenamenew1 = `${code1}-${issueno}-${revisionno}-${rest1}`;
                  }



                  let newfileNameNewN = fileAddResult.data.Name.includes(docCode) ? filenamenew1 : newfileNameNewP;
                  newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                  const finalFileNameNewNO = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                  const documentName = finalFileNameNewNO;
                  //const documentName = fileAddResult.data.Name;
                  bannerImageArray = fileAddResult;
                  // Get the item ID for the uploaded file
                  const currentItemId = await fileNew.getItem<{ Id: number }>();
                  const itemId = currentItemId.Id;
                  await currentItemId.update({
                    FileName: documentName, // Assuming FileName is the internal name of the column
                    DocumentCode: selectedOptionReq.requestcode == "New" ? doccode : selectedOption?.DocumentCode,
                  });

                  // Save the document ID for the attachment field in ChangeRequestList
                  attachmentIds.push(itemId);
                }
              } else if (DocumentLink && DocumentLink.ID > 0) {

                const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
                debugger
                let filenamenew: any;
                const newfileNameNew1 = docCode + "-" + issueno + "-" + revisionno + "-" + item.File.Name;
                if (item.File.Name.includes(docCode)) {
                  const pattern = new RegExp(`^(${docCode})-(\\d+)-(\\d+)-(.*)$`);
                  const match = item.File.Name.match(pattern);

                  if (!match) return item.File.Name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${issueno}-${revisionno}-${rest}`;
                }


                let newfileNameNewN = item.File.Name.includes(docCode) ? filenamenew : newfileNameNew1;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const oldFilePath = item?.File?.ServerRelativeUrl;
                const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
                //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
                const finalFileNameNewNON = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const newFilePath = `${folderPath}/${finalFileNameNewNON}`;

                // 2. Use moveByPath to rename the file
                await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
                // 2. Move (rename) file
                // await item.update({
                //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
                //   DocumentCode: docCode
                // });
                const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).update({
                  FileName: finalFileNameNewNON, // Assuming FileName is the internal name of the column
                  DocumentCode: docCode
                })
                attachmentIds.push(DocumentLink?.ID);
              }
            }
            const sharewithIds: any[] = [];
            sharewithusers.forEach((user: any) => {
              if (user?.value) {
                sharewithIds.push(user.value);
              }
            });
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            const postPayload = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              RequestDate: formData.RequestDate != "" ? new Date(formData.RequestDate).toISOString() : new Date().toISOString(),
              DepartmentId: formData.DepartmentId,
              TemplateTypeId: formData.TemplateTypeId,
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById,
              //IssueDate: formData.IssueDate,
              IssueDate: finalissuedateNew,
              LocationId: formData.LocationId,
              FileName: formData.filename,
              CustodianId: formData.CustodianId,
              SerialNumber: formData.Status == "Rework" ? formData.SerialNumber : Number(serialno),
              IssueNumber: formData.Status == "Rework" ? formData.IssueNumber : Number(issueno),
              RevisionNumber: formData.Status == "Rework" ? formData.RevisionNumber : Number(revisionno),
              //RevisionNumber: selectedOption?.RevisionNumber,
              //RevisionDate: new Date().toISOString(),
              DocumentCode: formData.Status == "Rework" ? formData.DocumentCode : docCode,
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
              CDocumentCode: selectedTemplate != "Change Request" && changerequestdata && changerequestdata?.length > 0 ? changerequestdata[0]?.DocumentCode : docCode,
              CIssueNumber: selectedTemplate != "Change Request" && changerequestdata && changerequestdata?.length > 0 ? Number(changerequestdata[0]?.IssueNo) : Number(issueno),
              CRevisionNumber: selectedTemplate != "Change Request" && changerequestdata && changerequestdata?.length > 0 ? Number(changerequestdata[0]?.RevisionNo) : Number(revisionno),
              CIssueDate: selectedTemplate != "Change Request" ? finalissuedate : finalissuedateNew,
              CRevisionDate: selectedTemplate != "Change Request" ? finalrevisiondate : undefined,
              //AttachmentId: selectedOptionReq.label == "New" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "New" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso,
              PreviousAttachmentID: ""
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
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
              }
            });
            // Swal.fire('Submitted successfully.', '', 'success');
            // // sessionStorage.removeItem("bannerId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
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

  const onChange = (name: string, value: string) => {

    debugger

    setFormData((prevData) => ({

      ...prevData,

      Remark: value,

    }));

  };
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
    const now = new Date();
    const dateTimeSuffix = await formatDateTime(now);
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
            const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/ChangeRequestDocs');
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                //const newFileName = await getNewFileName(file.name);
                DocumentName = file.name;
                let filenamedraft = await cleanFileNameSave(file.name);
                const finalFileNameNewN = `${filenamedraft.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${filenamedraft.substring(filenamedraft.lastIndexOf('.'))}`;
                const fileAddResult = await folder.files.addChunked(finalFileNameNewN, file);
                const fileNew = fileAddResult.file;
                //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                const documentName = fileAddResult.data.Name;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: finalFileNameNewN, // Assuming FileName is the internal name of the column
                  NFinalDocumentCodePrinting: "No"
                });

                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            }
            const sharewithIds: any[] = [];
            sharewithusers.forEach((user: any) => {
              if (user?.value) {
                sharewithIds.push(user.value);
              }
            });
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            let arr = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId || undefined,
              TemplateTypeId: formData.TemplateTypeId || undefined,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              LocationId: formData.LocationId || undefined,
              CustodianId: formData.CustodianId || undefined,
              SerialNumber: selectedOptionReq.requestcode == "Edit" && selectedOption ? Number(selectedOption?.SerialNumber) : null,
              IssueNumber: selectedOptionReq.requestcode == "Edit" && selectedOption ? Number(selectedOption?.IssueNumber) : null,
              RevisionNumber: selectedOptionReq.requestcode == "Edit" && selectedOption ? Number(selectedOption?.RevisionNumber) : null,
              DocumentCode: selectedOptionReq.requestcode == "New" ? "" : selectedOption && selectedOption?.DocumentCode,
              ReferenceNumber: selectedOptionReq.requestcode == "Edit" ? test : "",
              AmendmentTypeId: formData.AmendmentTypeId || undefined,
              RequestTypeId: formData.RequestTypeId || undefined,
              ClassificationId: formData.ClassificationId || undefined,
              FileName: formData.filename,
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById || undefined,
              //ChangeRequestTypeId: formData.ChangeRequestTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              //SubmiitedDate: selectedOption?.SubmiitedDate,
              SubmiitedDate: new Date(formData.SubmiitedDate).toISOString(),
              SubmitStatus: "No",
              Status: formData.Status == "Rework" ? "Rework" : "Save as draft",
              DocumentName: DocumentName,
              DocumentTypeId: formData.DocumentTypeId || undefined,
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "No",
              CurrentUserRole: "OES",
              //AttachmentId: selectedOptionReq.label == "New" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "New" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentarr.length > 0 ? Attachmentidsss : [],
              AttachmentJson: Attachmentarr.length > 0 ? AttachmentJso : "",
              PreviousAttachmentID: selectedOptionReq.requestcode == "New" ? "" : formData.PreviousAttachmentID + ""

            }
            let descriptionError = false;
            let reasonError = false;
            console.log("postPayloaddrafttedit", arr, editItemID, cancellReason);
            const postResult = await updateItemChangeRequestList(arr, sp, editItemID);
            const postId = postResult?.data?.ID;
            // for (const row of cancellReason) {
            //   if (row.description.trim() === "") {
            //     descriptionError = true;
            //   }
            //   if (row.reason.trim() === "") {
            //     reasonError = true;
            //   }
            //   const postPayload2 = {
            //     ChangeRequestIDId: editItemID, // Assuming "Title" column exists
            //     ChangeDescription: row.description,
            //     ReasonforChange: row.reason,
            //   }
            //   if (!descriptionError || !reasonError) {
            //     if (!row.id) {
            //       const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
            //       const postId2 = postResult2?.data?.ID;
            //       // debugger
            //       if (!postId2) {
            //         console.error("Post creation failed.");
            //         return;
            //       }

            //     }
            //     else if (row.id > 0) {
            //       const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
            //       const postId2 = postResult2?.data?.ID;
            //     }
            //   }


            // }
            for (const row of cancellReason) {
              const isDescriptionEmpty = !row.description || row.description.trim() === "";
              const isReasonEmpty = !row.reason || row.reason.trim() === "";

              // Skip the row if BOTH description and reason are empty
              if (isDescriptionEmpty && isReasonEmpty) {
                continue;
              }

              const postPayload2 = {
                ChangeRequestIDId: editItemID,
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              };

              try {
                if (!row.id) {
                  const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                  const postId2 = postResult2?.data?.ID;
                  if (!postId2) {
                    console.error("Post creation failed.");
                    return;
                  }
                } else if (row.id > 0) {
                  const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                  const postId2 = postResult2?.data?.ID;
                }
              } catch (error) {
                console.error("Error saving data:", error);
                return;
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
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
              }
            });
            // Swal.fire('Saved successfully.', '', 'success');
            // sessionStorage.removeItem("ChangeRequestId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
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
            const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/ChangeRequestDocs');
            debugger
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                //const newFileName = await getNewFileName(file.name);
                DocumentName = file.name;
                let filenamedraft = await cleanFileNameSave(file.name);
                const finalFileNameNewN = `${filenamedraft.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${filenamedraft.substring(filenamedraft.lastIndexOf('.'))}`;
                const fileAddResult = await folder.files.addChunked(finalFileNameNewN, file);
                const fileNew = fileAddResult.file;
                //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                const documentName = fileAddResult.data.Name;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: finalFileNameNewN, // Assuming FileName is the internal name of the column
                  NFinalDocumentCodePrinting: "No"
                });
                console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            }
            const sharewithIds: any[] = [];
            sharewithusers.forEach((user: any) => {
              if (user?.value) {
                sharewithIds.push(user.value);
              }
            });
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            const postPayload = {
              Title: formData.RequesterName,
              RequesterNameId: formData.RequesterNameId,
              RequesterDesignation: formData.RequesterDesignation,
              DepartmentId: formData.DepartmentId || undefined,
              TemplateTypeId: formData.TemplateTypeId || undefined,
              RequestDate: new Date(formData.RequestDate).toISOString(),
              //IssueDate: new Date().toISOString(),
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById || undefined,
              LocationId: formData.LocationId || undefined,
              FileName: formData.filename,
              CustodianId: formData.CustodianId || undefined,
              SerialNumber: selectedOptionReq.requestcode == "Edit" && selectedOption ? Number(selectedOption?.SerialNumber) : null,
              IssueNumber: selectedOptionReq.requestcode == "Edit" && selectedOption ? Number(selectedOption?.IssueNumber) : null,
              RevisionNumber: selectedOptionReq.requestcode == "Edit" && selectedOption ? Number(selectedOption?.RevisionNumber) : null,
              DocumentCode: selectedOptionReq.requestcode == "New" ? "" : selectedOption && selectedOption?.DocumentCode,
              ReferenceNumber: selectedOptionReq.requestcode == "Edit" ? test : "",
              RequestTypeId: formData.RequestTypeId || undefined,
              AmendmentTypeId: formData.AmendmentTypeId || undefined,
              ClassificationId: formData.ClassificationId || undefined,
              ChangeRequestTypeId: selectedCheckboxIds,
              SubmiitedDate: new Date().toISOString(),
              SubmitStatus: "No",
              Status: "Save as draft",
              DocumentName: DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              //AttachmentId: selectedOptionReq.label == "New" ? Attachmentidsss : selectedOption?.AttachmentId,
              //AttachmentJson: selectedOptionReq.label == "New" ? AttachmentJso : selectedOption?.AttachmentJson,
              AttachmentId: Attachmentarr.length > 0 ? Attachmentidsss : [],
              AttachmentJson: Attachmentarr.length > 0 ? AttachmentJso : "",
              PreviousAttachmentID: selectedOptionReq.requestcode == "New" ? "" : formData.PreviousAttachmentID + ""

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

            // for (const row of cancellReason) {
            //   if (row.description.trim() === "") {
            //     descriptionError = true;
            //   }
            //   if (row.reason.trim() === "") {
            //     reasonError = true;
            //   }
            //   const postPayload2 = {
            //     ChangeRequestIDId: postId, // Assuming "Title" column exists
            //     ChangeDescription: row.description,
            //     ReasonforChange: row.reason,
            //   }
            //   if (!descriptionError || !reasonError) {
            //     const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
            //     const postId2 = postResult2?.data?.ID;
            //     // debugger
            //     if (!postId2) {
            //       console.error("Post creation failed.");
            //       return;
            //     }
            //   }
            // }
            for (const row of cancellReason) {
              const isDescriptionEmpty = !row.description || row.description.trim() === "";
              const isReasonEmpty = !row.reason || row.reason.trim() === "";

              // Skip the row if BOTH description and reason are empty
              if (isDescriptionEmpty && isReasonEmpty) {
                continue;
              }

              const postPayload2 = {
                ChangeRequestIDId: postId,
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              };

              try {
                if (!row.id) {
                  const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                  const postId2 = postResult2?.data?.ID;
                  if (!postId2) {
                    console.error("Post creation failed.");
                    return;
                  }
                } else if (row.id > 0) {
                  const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                  const postId2 = postResult2?.data?.ID;
                }
              } catch (error) {
                console.error("Error saving data:", error);
                return;
              }
            }



            setLoading(false);
            sessionStorage.removeItem("ChangeRequestId")
            Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
              }
            });
            // Swal.fire('Saved successfully.', '', 'success');
            // // sessionStorage.removeItem("bannerId")
            // setTimeout(() => {
            //   //window.location.reload();
            //   window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
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
  const ForwardApproval = async (status: string) => {
    debugger
    scrollToTop();
    const IsactionTaken = await CheckIfAlreadyactionTaken(sp, editID.Id, CONTENTTYPE_ChangeDocument);
    setValidRemark(true);
    let url = window.location.href.split('/sites/')[0];
    console.log("topp draf fprt", editItemID, cancellReason, url);
    let valid = true;
    let actionMessage = "";
    let successMessage = "";
    if (status === 'Rework' || status === 'Rejected') {
      setMandatRemark(true)
    } else {
      setMandatRemark(false)
    }
    if ((status === 'Rework' || status === 'Rejected') && (formData.Remark === "" || formData.Remark == null)) {
      setValidRemark(false);
      Swal.fire('Please fill the mandatory fields', '', 'warning');
      return;
    }
    switch (status) {
      case "Forward":
        actionMessage = "Have you verified that the approval hierarchy (Preparer > Reviewer > Endorser > Signer/Approver) has been followed before forwarding this request?";
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

      // debugger
      // const isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 && (row.leveltype == "One" || row.leveltype == "All"));

      // if (!isValid) {
      //   // alert("Each row must have a role selected and at least Anyone approver.");
      //   valid = false;
      // }
      // setValidforward(valid)
      // if (!valid) {
      //   Swal.fire('Please fill all the mandatory fields.');
      //   return;
      // }
      let isFormValid = true;
      if (forwardToArr.length === 0) {
        // alert("At least Anyone row is required.");
        isFormValid = false;
      }
      const updatedRows = forwardToArr.map((row) => {
        const roleError = row.role === 0;
        const approverError = row.approvers.length === 0;
        const typeError = !(row.leveltype === "One" || row.leveltype === "All");
        const responsibilityerror = row.Responsibility.trim() === "" || row.Responsibility == null;
        const isRowValid = !roleError && !approverError && !typeError && !responsibilityerror;

        if (!isRowValid) {
          isFormValid = false;
        }

        return {
          ...row,
          roleError,
          approverError,
          typeError,
          responsibilityerror,
          rowError: roleError && approverError && typeError && responsibilityerror, // mark full-row error only if fully empty
        };
      });

      setForwardToArr(updatedRows);
      setValidforward(isFormValid);

      if (!isFormValid) {
        Swal.fire("Please fill all the mandatory fields in every row.");
        return;
      }
      if (!IsactionTaken) {
        Swal.fire("Action has already been taken on this record.");
        return;
      }
      if (valid) {
        Swal.fire({
          title: actionMessage,
          showConfirmButton: true,
          showCancelButton: true,
          confirmButtonText: "Yes",
          cancelButtonText: "No",
          icon: 'warning',
          customClass: {
            title: status == "Forward" ? 'swal-title-large' : ''
          }
        }
        ).then(async (result) => {
          console.log(result)
          if (result.isConfirmed) {
            setLoading(true);
            debugger
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))

            let arr = {
              ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toISOString(),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: "Approved",
              Remark: formData.Remark,


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
                //ContentTitle: formData.ReferenceNumber,
                ContentTitle: formData.FileName,
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
                IsApprovalGenerated: "No",
                Responsibility: item.Responsibility || "",
                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",
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
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx`;
              }
            });

            // }
          }

        })
      }



    }
    else {
      if (!IsactionTaken) {
        Swal.fire("Action has already been taken on this record.");
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
              ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toISOString(),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: status,
              Remark: formData.Remark,


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
              if (status != "Rework") {
                let arr2 = {
                  Title: currentUser.Title,
                  //ContentTitle: selectedOption?.ReferenceNumber,
                  ContentTitle: selectedOption?.FileName || formData.FileName,
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
                  IsApprovalGenerated: "No",
                  Responsibility: item.Responsibility || "",
                  IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",
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
            let arr2 = {
              SubmiitedDate: new Date().toLocaleDateString("en-CA"),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: status === "Rework" ? "Rework" : "Rejected",
              IsRework: status === "Rework" ? "Yes" : "No",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "No",
              SubmitStatus: "No",
              //Remarks:formData.Remark
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
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/MyApprovals.aspx`;
              }
            });
            // }
          }
        })
      }
    }
  }
  const handlePaste = (e: React.ClipboardEvent<HTMLInputElement>) => {
    const pastedText = e.clipboardData.getData('text');
    if (!/^[a-zA-Z0-9-]*$/.test(pastedText)) {
      e.preventDefault();
    }
  };
  const handleKeyDowntext = (e: React.KeyboardEvent<HTMLInputElement>) => {
    const allowedKeys = [
      'Backspace',
      'ArrowLeft',
      'ArrowRight',
      'Delete',
      'Tab',
    ];

    const isLetter = /^[a-zA-Z]$/.test(e.key);   // a-z or A-Z
    const isNumber = /^[0-9]$/.test(e.key);      // 0-9
    const isHyphen = e.key === '-';
    const isSpace = e.key === ' ';
    if (isLetter || isNumber || isHyphen || isSpace || allowedKeys.includes(e.key)) {
      return; // ✅ allowed input
    } else {
      e.preventDefault(); // ❌ block everything else
    }
  };


  const onChangefilename = (name: string, value: string) => {
    debugger
    setFormData((prevData) => ({
      ...prevData,
      [name]: value,
    }));
  };

  const ForwardInitiatorApproval = async (status: string) => {
    let url = window.location.href.split('/sites/')[0];
    // let valid = true;
    let actionMessage = "";
    let successMessage = "";
    const now = new Date();
    console.log("forwardToArrEdit", forwardToArrEdit, forwardToArr);
    const dateTimeSuffix = await formatDateTime(now);
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
    const docCode1 = formData.DocumentCode;
    if (status === 'Approved') {
      setMandatRemark(true)
    } else {
      setMandatRemark(false)
    }
    const IsactionTaken = await CheckIfAlreadyactionTaken(sp, editID.Id, CONTENTTYPE_ChangeDocument);
    if (!IsactionTaken) {
      Swal.fire("Action has already been taken on this record.");
      return;
    }
    if (status == "Approved") {
      if (await validateForm(FormSubmissionMode.SUBMIT)) {

        if (status === 'Approved' && (formData.Remark === "" || formData.Remark == null)) {
          setValidRemark(false);
          Swal.fire('Please fill the mandatory fields', '', 'warning');
          return;
        }
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
            debugger
            let bannerImageArray: any = {};
            let DocumentName: string = "";
            let attachmentIds = [];
            const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/ChangeRequestDocs');
            // if (Attachmentarr.length > 0) {
            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                //const newFileName = await getNewFileName(file.name);

                const newfileNameNew = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + file.name;
                let filenamenew: any;
                DocumentName = newfileNameNew;
                if (file.name.includes(docCode1)) {
                  const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
                  const match = file.name.match(pattern);

                  if (!match) return file.name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
                }
                let filenew = file.name.includes(docCode1) ? filenamenew : newfileNameNew;
                filenew = await cleanFileNameSave(filenew);
                const finalFileNameNewin = `${filenew.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${filenew.substring(filenew.lastIndexOf('.'))}`;
                const fileAddResult = await folder.files.addChunked(finalFileNameNewin, file);
                const fileNew = fileAddResult.file;
                //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                const newfileNameNewN1 = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + file.name;
                //const newfileNameNew1 = docCode + "-" + issueno + "-" + revisionno + "-" + file.name;

                let newfileNameNewN = file.name.includes(docCode1) ? filenamenew : newfileNameNewN1;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const finalFileNameNewin1 = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const documentName = finalFileNameNewin1;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: documentName, // Assuming FileName is the internal name of the column
                  //DocumentCode: selectedOptionReq?.requestcode == "New" ? doccode : selectedOption?.DocumentCode,
                });

                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            } else
              if (Attachmentarr.length > 0 && Attachmentarr[0].ID > 0) {

                const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
                debugger
                let filenamenew: any;
                const newfileNameNewo = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + item.File.Name;
                if (item.File.Name.includes(docCode1)) {
                  const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
                  const match = item.File.Name.match(pattern);

                  if (!match) return item.File.Name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
                }

                let newfileNameNewN = item.File.Name.includes(docCode1) ? filenamenew : newfileNameNewo;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const oldFilePath = item?.File?.ServerRelativeUrl;
                const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
                //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
                const finalFileNameNewiIN = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const newFilePath = `${folderPath}/${finalFileNameNewiIN}`;

                // 2. Use moveByPath to rename the file
                await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
                // 2. Move (rename) file
                // await item.update({
                //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
                //   DocumentCode: docCode
                // });
                const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).update({
                  FileName: finalFileNameNewiIN, // Assuming FileName is the internal name of the column
                  DocumentCode: docCode1
                })
              }

            // else {
            //   if (DocumentLink && DocumentLink?.files?.length > 0) {
            //     for (const file of DocumentLink.files) {
            //       //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
            //       //const newFileName = await getNewFileName(file.name);
            //       let filenamenew: any;
            //       const newfileName = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + file.name;
            //       if (file.name.includes(docCode1)) {
            //         const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
            //         const match = file.name.match(pattern);

            //         if (!match) return file.name; // If it doesn't match expected structure, return original

            //         const [, code, , , rest] = match;

            //         filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
            //       }
            //       const newfileNameNew = file.name.includes(docCode1) ? filenamenew : newfileName;

            //       DocumentName = newfileNameNew;
            //       const fileAddResult = await folder.files.addChunked(newfileNameNew, file);
            //       const fileNew = fileAddResult.file;
            //       let filenamenew1: any;
            //       //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
            //       if (fileAddResult.data.Name.includes(docCode1)) {
            //         const pattern1 = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
            //         const match1 = fileAddResult.data.Name.match(pattern1);

            //         if (!match1) return fileAddResult.data.Name; // If it doesn't match expected structure, return original

            //         const [, code1, , , rest1] = match1;

            //         filenamenew1 = `${code1}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest1}`;
            //       }
            //       const newfileNameNewN1 = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + fileAddResult.data.Name;


            //       const newfileNameNewN = fileAddResult.data.Name.includes(docCode1) ? filenamenew1 : newfileNameNewN1;
            //       const documentName = newfileNameNewN;
            //       bannerImageArray = fileAddResult;
            //       // Get the item ID for the uploaded file
            //       const currentItemId = await fileNew.getItem<{ Id: number }>();
            //       const itemId = currentItemId.Id;
            //       await currentItemId.update({
            //         FileName: documentName, // Assuming FileName is the internal name of the column
            //         DocumentCode: selectedOptionReq?.requestcode == "New" ? docCode1 : selectedOption?.DocumentCode,
            //       });

            //       // Save the document ID for the attachment field in ChangeRequestList
            //       attachmentIds.push(itemId);
            //     }
            //   } else if (DocumentLink && DocumentLink.ID > 0) {

            //     const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
            //     debugger
            //     let filenamenew: any;
            //     const newfileNameNew1 = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + item.File.Name;
            //     if (item.File.Name.includes(docCode1)) {
            //       const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
            //       const match = item.File.Name.match(pattern);

            //       if (!match) return item.File.Name; // If it doesn't match expected structure, return original

            //       const [, code, , , rest] = match;

            //       filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
            //     }


            //     const newfileNameNewN = item.File.Name.includes(docCode1) ? filenamenew : newfileNameNew1;
            //     const oldFilePath = item?.File?.ServerRelativeUrl;
            //     const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
            //     //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
            //     const newFilePath = `${folderPath}/${newfileNameNewN}`;

            //     // 2. Use moveByPath to rename the file
            //     await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
            //     // 2. Move (rename) file
            //     // await item.update({
            //     //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
            //     //   DocumentCode: docCode
            //     // });
            //     const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).update({
            //       FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
            //       DocumentCode: docCode1
            //     })
            //   }
            //   attachmentIds.push(DocumentLink?.ID);
            // }
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            // //////////////Update Process Approval List when Submitted
            let arr = {
              ActionTakenById: currentUser.Id,
              ActionTakenOn: new Date().toISOString(),
              // ActionTakenRoleId: formData.RequesterDesignation,
              Status: "Approved",
              Remark: formData.Remark,

            }
            const postResult = await updateApprovalItem(arr, sp, editID.Id);
            const postId = postResult?.data?.ID;
            const sharewithIds: any[] = [];
            sharewithusers.forEach((user: any) => {
              if (user?.value) {
                sharewithIds.push(user.value);
              }
            });
            // //////////////Update Document cancellation List when Submitted
            let arr3 = {
              //Title: formData.RequesterName,
              //RequesterNameId: formData.RequesterNameId,
              //RequesterDesignation: formData.RequesterDesignation,
              //DepartmentId: formData.DepartmentId,
              Remarks: "",
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById || undefined,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: formData.RequestDate,
              //IssueDate: formData.IssueDate,
              LocationId: formData?.LocationId,
              FileName: formData.filename,
              CustodianId: formData?.CustodianId,
              //SerialNumber: formData?.SerialNumber,
              //IssueNumber: formData?.IssueNumber,
              //RevisionNumber: formData?.RevisionNumber,
              //RevisionDate: selectedOption?.RevisionDate,
              //DocumentCode: formData?.DocumentCode,
              //ReferenceNumber: formData?.ReferenceNumber,
              //PreparedById: sharewithIds,
              AmendmentTypeId: formData?.AmendmentTypeId,
              RequestTypeId: formData?.RequestTypeId,
              ClassificationId: formData?.ClassificationId,
              ChangeRequestTypeId: selectedCheckboxIds,
              SubmiitedDate: formData?.SubmiitedDate,
              SubmitStatus: "Yes",
              Status: "Pending",
              // DocumentName: "",
              // IsRework: false,
              // DigitalSignStatus: false,
              //ChangeRequestIDId: formData.ChangeRequestID,
              DocumentTypeId: formData?.DocumentTypeId,
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso


            }
            let arrework = {
              //Title: formData.RequesterName,
              //RequesterNameId: formData.RequesterNameId,
              //RequesterDesignation: formData.RequesterDesignation,
              //RequestDate: new Date(formData.RequestDate).toISOString(),
              //IssueDate: finalissuedateNew,
              Remarks: "",
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById || undefined,
              RequestTypeId: formData.RequestTypeId,
              AmendmentTypeId: formData.AmendmentTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              ClassificationId: formData.ClassificationId,
              SubmiitedDate: new Date(formData.RequestDate).toISOString(),
              //PreparedById: sharewithIds,
              SubmitStatus: "Yes",
              Status: "Pending",
              IsRework: "No",
              OESSubmitStatus: "No",
              InitiatorSubmitStatus: "Yes",
              CurrentUserRole: "OES",
              DocumentName: attachmentIds.length != 0 ? DocumentName : formData.DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso,
              PreviousAttachmentID: ""
            }
            const postResult3 = await updateItemChangeRequestList(formData.Status == "Rework" ? arrework : arr3, sp, editItemID);
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
            // const toDeletefor = forwardToArrEdit.filter(
            //   (itemEdit) => !forwardToArr.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
            // );
            if (sharewithIds.length > 0) {
              const preparerToDelete = forwardToArrEdit.filter(item => {
                return (
                  item.Responsibility === "Preparer" &&
                  item.approvers.length === 1 &&
                  !sharewithIds.includes(item.approvers[0].value)
                );
              });
              for (const item of preparerToDelete) {
                try {
                  await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(item.id).delete();
                  // console.log(`Deleted item with ID: ${item.ID}`);
                } catch (error) {
                  console.error(`Error deleting item with ID: ${item.id}`, error);
                }
              }
            }

            // Delete each item from SharePoint

            setLoading(false);
            Swal.fire(successMessage, '', 'success').then(async (result) => {
              if (result.isConfirmed) {
                sessionStorage.removeItem("ChangeRequestId")
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
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
            debugger
            // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
            // let arr = {
            //   ActionTakenById: currentUser.Id,
            //   ActionTakenOn: new Date().toLocaleDateString("en-CA"),
            //   // ActionTakenRoleId: formData.RequesterDesignation,
            //   Status: status,
            //   // Remark: remark,

            // }
            // const postResult = await updateApprovalItem(arr, sp, editID.Id);
            // const postId = postResult?.data?.ID;
            let bannerImageArray: any = {};
            let DocumentName: string = "";
            let attachmentIds = [];
            const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/ChangeRequestDocs');

            if (Attachmentarr.length > 0 && Attachmentarr[0]?.files?.length > 0) {
              for (const file of Attachmentarr[0].files) {
                //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                //const newFileName = await getNewFileName(file.name);

                const newfileNameNew = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + file.name;
                let filenamenew: any;
                DocumentName = newfileNameNew;
                if (file.name.includes(docCode1)) {
                  const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
                  const match = file.name.match(pattern);

                  if (!match) return file.name; // If it doesn't match expected structure, return original

                  const [, code, , , rest] = match;

                  filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
                }

                let filenew1 = file.name.includes(docCode1) ? filenamenew : newfileNameNew;
                filenew1 = await cleanFileNameSave(filenew1);
                const finalFileNameNewiIN = `${filenew1.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${filenew1.substring(filenew1.lastIndexOf('.'))}`;
                const fileAddResult = await folder.files.addChunked(finalFileNameNewiIN, file);
                const fileNew = fileAddResult.file;
                //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
                const newfileNameNewN1 = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + file.name;
                //const newfileNameNew1 = docCode + "-" + issueno + "-" + revisionno + "-" + file.name;

                let newfileNameNewN = file.name.includes(docCode1) ? filenamenew : newfileNameNewN1;
                newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
                const finalFileNameNewIn1 = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
                const documentName = finalFileNameNewIn1;
                bannerImageArray = fileAddResult;
                // Get the item ID for the uploaded file
                const currentItemId = await fileNew.getItem<{ Id: number }>();
                const itemId = currentItemId.Id;
                await currentItemId.update({
                  FileName: documentName, // Assuming FileName is the internal name of the column
                  //DocumentCode: selectedOptionReq?.requestcode == "New" ? doccode : selectedOption?.DocumentCode,
                });

                // Save the document ID for the attachment field in ChangeRequestList
                attachmentIds.push(itemId);
              }
            } else if (Attachmentarr.length > 0 && Attachmentarr[0].ID > 0) {

              const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
              debugger
              let filenamenew: any;
              const newfileNameNewo = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + item.File.Name;
              if (item.File.Name.includes(docCode1)) {
                const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
                const match = item.File.Name.match(pattern);

                if (!match) return item.File.Name; // If it doesn't match expected structure, return original

                const [, code, , , rest] = match;

                filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
              }

              let newfileNameNewN = item.File.Name.includes(docCode1) ? filenamenew : newfileNameNewo;
              newfileNameNewN = await cleanFileNameSave(newfileNameNewN);
              const oldFilePath = item?.File?.ServerRelativeUrl;
              const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
              //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
              const finalFileNameNewINo = `${newfileNameNewN.replace(/\.[^/.]+$/, "")}_${dateTimeSuffix}${newfileNameNewN.substring(newfileNameNewN.lastIndexOf('.'))}`;
              const newFilePath = `${folderPath}/${finalFileNameNewINo}`;

              // 2. Use moveByPath to rename the file
              await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
              // 2. Move (rename) file
              // await item.update({
              //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
              //   DocumentCode: docCode
              // });
              const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(Attachmentarr[0].ID).update({
                FileName: finalFileNameNewINo, // Assuming FileName is the internal name of the column
                DocumentCode: docCode1
              })
            }

            // else {
            //   if (DocumentLink && DocumentLink?.files?.length > 0) {
            //     for (const file of DocumentLink.files) {
            //       //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
            //       //const newFileName = await getNewFileName(file.name);
            //       let filenamenew: any;
            //       const newfileName = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + file.name;
            //       if (file.name.includes(docCode1)) {
            //         const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
            //         const match = file.name.match(pattern);

            //         if (!match) return file.name; // If it doesn't match expected structure, return original

            //         const [, code, , , rest] = match;

            //         filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
            //       }
            //       const newfileNameNew = file.name.includes(docCode1) ? filenamenew : newfileName;

            //       DocumentName = newfileNameNew;
            //       const fileAddResult = await folder.files.addChunked(newfileNameNew, file);
            //       const fileNew = fileAddResult.file;
            //       let filenamenew1: any;
            //       //const newFileNameN = await getNewFileName(fileAddResult.data.Name);
            //       if (fileAddResult.data.Name.includes(docCode1)) {
            //         const pattern1 = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
            //         const match1 = fileAddResult.data.Name.match(pattern1);

            //         if (!match1) return fileAddResult.data.Name; // If it doesn't match expected structure, return original

            //         const [, code1, , , rest1] = match1;

            //         filenamenew1 = `${code1}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest1}`;
            //       }
            //       const newfileNameNewN1 = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + fileAddResult.data.Name;


            //       const newfileNameNewN = fileAddResult.data.Name.includes(docCode1) ? filenamenew1 : newfileNameNewN1;
            //       const documentName = newfileNameNewN;
            //       bannerImageArray = fileAddResult;
            //       // Get the item ID for the uploaded file
            //       const currentItemId = await fileNew.getItem<{ Id: number }>();
            //       const itemId = currentItemId.Id;
            //       await currentItemId.update({
            //         FileName: documentName, // Assuming FileName is the internal name of the column
            //         DocumentCode: selectedOptionReq?.requestcode == "New" ? docCode1 : selectedOption?.DocumentCode,
            //       });

            //       // Save the document ID for the attachment field in ChangeRequestList
            //       attachmentIds.push(itemId);
            //     }
            //   } else if (DocumentLink && DocumentLink.ID > 0) {

            //     const item = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).select('File/ServerRelativeUrl', 'File/Name').expand('File')();
            //     debugger
            //     let filenamenew: any;
            //     const newfileNameNew1 = docCode1 + "-" + formData.IssueNumber + "-" + formData.RevisionNumber + "-" + item.File.Name;
            //     if (item.File.Name.includes(docCode1)) {
            //       const pattern = new RegExp(`^(${docCode1})-(\\d+)-(\\d+)-(.*)$`);
            //       const match = item.File.Name.match(pattern);

            //       if (!match) return item.File.Name; // If it doesn't match expected structure, return original

            //       const [, code, , , rest] = match;

            //       filenamenew = `${code}-${formData.IssueNumber}-${formData.RevisionNumber}-${rest}`;
            //     }


            //     const newfileNameNewN = item.File.Name.includes(docCode1) ? filenamenew : newfileNameNew1;
            //     const oldFilePath = item?.File?.ServerRelativeUrl;
            //     const folderPath = oldFilePath.substring(0, oldFilePath.lastIndexOf('/'));
            //     //const oldFilePathN = folderPath +"/"+ encodeURI(item.File.Name);
            //     const newFilePath = `${folderPath}/${newfileNameNewN}`;

            //     // 2. Use moveByPath to rename the file
            //     await sp.web.getFileByServerRelativePath(oldFilePath).moveByPath(newFilePath, true, false);
            //     // 2. Move (rename) file
            //     // await item.update({
            //     //   FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
            //     //   DocumentCode: docCode
            //     // });
            //     const itemnew = await sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(DocumentLink.ID).update({
            //       FileName: newfileNameNewN, // Assuming FileName is the internal name of the column
            //       DocumentCode: docCode1
            //     })
            //   }
            //   attachmentIds.push(DocumentLink?.ID);
            // }
            const sharewithIds: any[] = [];
            sharewithusers.forEach((user: any) => {
              if (user?.value) {
                sharewithIds.push(user.value);
              }
            });
            let Attachmentidsss = attachmentIds.length != 0 ? attachmentIds : formData.AttachmentId;
            let AttachmentJso = attachmentIds.length != 0 ? JSON.stringify(bannerImageArray) : formData.AttachmentJson;
            // //////////////Update Document cancellation List when Submitted
            let arr3 = {
              // Title: formData.RequesterName,
              // RequesterNameId: formData.RequesterNameId,
              // RequesterDesignation: formData.RequesterDesignation,
              // DepartmentId: formData.DepartmentId,
              Remarks: formData.Remark,
              TemplateTypeId: formData.TemplateTypeId,
              RequestDate: formData.RequestDate,
              IssueDate: formData.IssueDate,
              //LocationId: selectedOption?.LocationId,
              FileName: formData.filename,
              //CustodianId: selectedOption?.CustodianId,
              // SerialNumber: selectedOption?.SerialNumber,
              // IssueNumber: selectedOption?.IssueNumber,
              // RevisionNumber: selectedOption?.RevisionNumber,
              // //RevisionDate: selectedOption?.RevisionDate,
              // DocumentCode: selectedOption?.value,
              // ReferenceNumber: selectedOption?.ReferenceNumber,
              AmendmentTypeId: formData?.AmendmentTypeId,
              RequestTypeId: formData?.RequestTypeId,
              ClassificationId: formData?.ClassificationId,
              ChangeRequestTypeId: selectedCheckboxIds,
              SubmiitedDate: formData?.SubmiitedDate,
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById || undefined,
              SubmitStatus: "No",
              //Status: "Rework",
              // DocumentName: "",
              // IsRework: false,
              // DigitalSignStatus: false,
              //ChangeRequestIDId: formData.ChangeRequestID,
              DocumentTypeId: formData?.DocumentTypeId,
              //OESSubmitStatus: "No",
              //InitiatorSubmitStatus: "No",
              //CurrentUserRole: "Initiator",
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso


            }
            let arrework = {
              //Title: formData.RequesterName,
              //RequesterNameId: formData.RequesterNameId,
              //RequesterDesignation: formData.RequesterDesignation,
              //RequestDate: new Date(formData.RequestDate).toISOString(),
              //IssueDate: finalissuedateNew,
              Remarks: formData.Remark,
              PreparedById: sharewithIds.length > 0 ? sharewithIds : formData.PreparedById || undefined,
              RequestTypeId: formData.RequestTypeId,
              AmendmentTypeId: formData.AmendmentTypeId,
              ChangeRequestTypeId: selectedCheckboxIds,
              ClassificationId: formData.ClassificationId,
              //SubmiitedDate: new Date(formData.RequestDate).toISOString(),
              //SubmitStatus: "Yes",
              //Status: "Rework",
              //IsRework: "Yes",
              // OESSubmitStatus: "No",
              // InitiatorSubmitStatus: "Yes",
              // CurrentUserRole: "OES",
              DocumentName: attachmentIds.length != 0 ? DocumentName : formData.DocumentName,
              DocumentTypeId: formData.DocumentTypeId,
              AttachmentId: Attachmentidsss,
              AttachmentJson: AttachmentJso,
              PreviousAttachmentID: ""
            }
            const postResult3 = await updateItemChangeRequestList(formData.Status == "Rework" ? arrework : arr3, sp, editItemID);
            //const postId3 = postResult?.data?.ID;


            // //////////////Update Document cancellation Reason List when Submitted

            // for (const row of cancellReason) {

            //   const postPayload2 = {
            //     ChangeRequestIDId: formData.ChangeRequestID, // Assuming "Title" column exists
            //     ChangeDescription: row.description,
            //     ReasonforChange: row.reason,
            //   }

            //   if (!row.id) {

            //     const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
            //     const postId2 = postResult2?.data?.ID;
            //     // debugger
            //     if (!postId2) {
            //       console.error("Post creation failed.");
            //       return;
            //     }

            //   }
            //   else if (row.id > 0) {
            //     const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
            //     const postId2 = postResult2?.data?.ID;
            //   }

            // }
            for (const row of cancellReason) {
              const isDescriptionEmpty = !row.description || row.description.trim() === "";
              const isReasonEmpty = !row.reason || row.reason.trim() === "";

              // Skip the row if BOTH description and reason are empty
              if (isDescriptionEmpty && isReasonEmpty) {
                continue;
              }

              const postPayload2 = {
                ChangeRequestIDId: formData.ChangeRequestID,
                ChangeDescription: row.description,
                ReasonforChange: row.reason,
              };

              try {
                if (!row.id) {
                  const postResult2 = await addItemChangeRequestReasonlist(postPayload2, sp);
                  const postId2 = postResult2?.data?.ID;
                  if (!postId2) {
                    console.error("Post creation failed.");
                    return;
                  }
                } else if (row.id > 0) {
                  const postResult2 = await updateItemChangeRequestReasonList(postPayload2, sp, row.id);
                  const postId2 = postResult2?.data?.ID;
                }
              } catch (error) {
                console.error("Error saving data:", error);
                return;
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
                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
              }
            });

            // }
          }

        })
      }

    }


  }
  const handleKeyDown = (e: React.KeyboardEvent<HTMLSelectElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault(); // 🛑 Prevents page reload
    }
  };
  // const handleKeyDowntext = (e: React.KeyboardEvent<HTMLInputElement>) => {
  //   if (e.key === 'Enter') {
  //     e.preventDefault();
  //   }
  // };
  // const handleKeyDowntextarea = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
  //   if (e.key === 'Enter' && !e.shiftKey) {
  //     e.preventDefault(); // Prevent form submission or unwanted behavior
  //     // Optional: do something when Enter is pressed (like submit or blur)
  //   }
  // };
  const onFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>,
    libraryName: string,
    docLib: string
  ) => {
    event.preventDefault();


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
            //icon: "error",
            title: "Invalid File Type",
            text: "Only images and document files are allowed.",
          });
          return;
        }
        const fileExtension = files[0].name.split('.').pop().toLowerCase();
        if (fileExtension === 'doc') {
          Swal.fire({
            // icon: "error",
            title: "Invalid File Type",
            text: "Please convert the .doc file to .docx format and reattach the updated file.",
          });
          return;        // Stop further execution
        }
        const result = await Swal.fire({
          title: 'Please ensure that the document alignment and formatting are correct, or refer to the change request notes before attaching and submitting.',
          showConfirmButton: true,
          showCancelButton: true,
          confirmButtonText: "Yes",
          cancelButtonText: "No",
          //icon: 'warning',
          customClass: {
            title: 'swal-small-title'
            //icon: 'swal-small-icon'
          }
        });

        // If user cancels, exit early
        if (!result.isConfirmed) {
          const input = event.target as HTMLInputElement;
          const files = Array.from(input.files || []);
          input.value = '';

          return;
        }
        setAttachmentarr([]);
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
          FileLeafRef: files[0].name,
        };
        uloadBannerImageFiles.push(arr);
        setAttachmentarr(uloadBannerImageFiles);

      } else {
        Swal.fire("upload a document")
      }
    }
  };
  const generatePreviewUrl = async (serverRelativeUrl: string) => {
    // Encode the file name and construct the preview URL
    const encodedFilePath = encodeURIComponent(serverRelativeUrl);

    // Example: 
    // serverRelativeUrl = "/sites/ededms/test/DocumentLibraryInsideTest/Book.xlsx"
    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;

    // const previewUrl = `${siteUrl}/sites/ededms/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
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
        const fileUrl = `${fileObj.serverUrl?.trim()}${fileObj.serverRelativeUrl?.trim()}`;
        setPreviewUrl(fileUrl); // Set the preview URL
        setIsModalOpen(true);   // Open the modal
      } else {
        //alert("Invalid file object. Cannot generate preview URL.");
      }
    }
  };
  React.useEffect(() => {
    initializeCheckboxes();
  }, [selectedOptionReq]);


  const initializeCheckboxes = async () => {
    const allCheckboxes = await fetchChangeRequestTypes(); // Full list with IDs and names
    const filteredCheckboxes = await getChangeRequestTypeMaster(sp, selectedOptionReq?.label); // Only names match

    const filteredNames = filteredCheckboxes.map((item) =>
      item.name.trim().toLowerCase()
    );

    const enabledIds = allCheckboxes
      .filter(
        (checkbox) =>
          filteredNames.includes(checkbox.name.trim().toLowerCase()) ||
          checkbox.name.trim().toLowerCase() === "others" // Always include "Others"
      )
      .map((checkbox) => checkbox.id);

    setEnabledCheckboxIds(enabledIds);
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
        <div key={checkbox.id} className="form-check mb-3" title={checkbox.name}>
          {/* <input
            //onKeyDown={handleKeyDowntext}
            type="checkbox"
            className={`form-check-input ${(!ValidSubmit && changerequesttypeerr) ? "border-on-error" : ""}`}
            id={`checkbox-${checkbox.id}`}
            disabled={(InputDisabled && formData?.Status !== "Rework") || IsRecorddisabled}
            checked={selectedCheckboxIds.indexOf(checkbox.id) !== -1}
            onChange={() => handleCheckboxChange(checkbox.id)}
          /> */}
          <input
            type="checkbox"
            className={`form-check-input ${!ValidSubmit && changerequesttypeerr ? "border-on-error" : ""
              }`}
            id={`checkbox-${checkbox.id}`}
            disabled={
              // ✅ Disable if not in enabled list AND not "Others"
              !enabledCheckboxIds.includes(checkbox.id) ||
              (InputDisabled && formData?.Status !== "Rework") ||
              IsRecorddisabled
            }
            checked={selectedCheckboxIds.includes(checkbox.id)}
            onChange={() => {
              // ✅ Prevent selection if checkbox is disabled
              if (!enabledCheckboxIds.includes(checkbox.id) && checkbox.name.trim().toLowerCase() !== "others") {
                return;
              }
              handleCheckboxChange(checkbox.id);
            }} />
          <label className="form-check-label" htmlFor={`checkbox-${checkbox.id}`}>
            {checkbox.name}
          </label>
        </div>
      </div>
      // <div className="col-lg-3">
      //   <div key={checkbox.id} className="form-check mb-3" title={checkbox.name}>
      //     <input
      //       type="checkbox"
      //       // className="form-check-input"
      //       className={`form-check-input ${(!ValidSubmit && changerequesttypeerr) ? "border-on-error" : ""}`}
      //       id={`checkbox-${checkbox.id}`}
      //       // disabled={this.state.isReadonly} // Make the checkbox readonly if the condition is true
      //       // Use indexOf instead of includes
      //       disabled={InputDisabled && formData?.Status != "Rework"}
      //       checked={selectedCheckboxIds.indexOf(checkbox.id) !== -1}
      //       onChange={() => handleCheckboxChange(checkbox.id)}
      //     />
      //     <label className="form-check-label" htmlFor={`checkbox-${checkbox.id}`}>
      //       {checkbox.name}
      //     </label>
      //   </div>
      // </div>
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
          <div className="" >
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
                              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                <h4 style={{ textAlign: 'left', marginBottom: 0 }} className="text-dark font-16 fw-bold">
                                  Requested By
                                </h4>
                                {changereqNotes && changereqNotes.length > 0 &&
                                  <a
                                    style={{ fontSize: '0.875rem', marginLeft: '10px', color: '#333' }}
                                    onClick={() => setShowModalNotes(true)}
                                    title="View Notes"
                                  >
                                    <FontAwesomeIcon icon={faStickyNote} />
                                  </a>}
                              </div>
                              {/* <h4 style={{ textAlign: 'left' }} className="text-dark font-16 fw-bold mb-3">
                                Requested By
                              </h4>
                              <div>
                                <div>

                                  <a style={{ fontSize: '0.875rem' }} onClick={() => setShowModalNotes(true)}>
                                    <FontAwesomeIcon icon={faStickyNote} />
                                  </a>


                                </div>
                              </div> */}
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

                                {TemplateDoc && TemplateDoc.length > 0 && (
                                  <span
                                    onClick={() => OpenFileTemplate(TemplateDoc[0], "Open")}
                                    style={{ color: "blue", cursor: "pointer" }}
                                  >
                                    <div className="btn btn-primary p-2" title='Preview document'>
                                      <img
                                        style={{ cursor: 'pointer', height: '24px' }}
                                        className='mt-0'
                                        src={showdigisign ? require("../assets/signicon.png") : require("../assets/noun-download-5006210.png")}
                                        alt="Download Icon"
                                      />
                                    </div>
                                  </span>
                                )}
                              </div>
                            </div>

                            {/* <p className="sub-header">
                                                        Lorem ipsum dolor sit amet, consectetur adipisicing elit. Aspernatur, itaque.
                                                    </p> */}

                            <div className="row">
                              <div className="col-lg-4">


                                <div className="mb-3">
                                  <label htmlFor="RequesterName" className="form-label">Name:</label>
                                  <input title={formData.RequesterName} type="text" id="Name" name="RequesterName" className="form-control" value={formData.RequesterName} disabled={true} />
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={Departopt}
                                      value={SelectedOptionDepart}
                                      name="Department"
                                      className={`${(!ValidDraft && departmenterr) ? "border-on-error" : ""} ${(!ValidSubmit && departmenterr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectDepart(selectedOption)}
                                      placeholder="Search Department"
                                      isDisabled={InputDisabled || formData?.Status == "Rework" || IsRecorddisabled || !IsDepartmentEditable}
                                    />
                                  </div>


                                </div>
                              </div>
                              <div className="col-lg-4">


                                <div className="mb-3">
                                  <label htmlFor="RequesterDesignation" className="form-label">Designation:</label>
                                  <input title={formData.RequesterDesignation} type="text" id="RequesterDesignation" name="RequesterDesignation" className="form-control" value={formData.RequesterDesignation} disabled={true} />
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
                                     // maxDate={new Date()}
                                      //minDate={new Date()}
                                      disabled={InputDisabled && formData?.Status !== "Rework" || IsRecorddisabled}
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={ReqType}
                                      value={selectedOptionReq}
                                      name="Request Type"
                                      className={`${(!ValidDraft && requesttypeerr) ? "border-on-error" : ""} ${(!ValidSubmit && requesttypeerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectReq(selectedOption)}
                                      placeholder="Search Request Type"
                                      isDisabled={InputDisabled || formData?.Status == "Rework" || IsRecorddisabled}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">
                                {console.log("selectedOptionReqselectedOptionReq", selectedOptionReq)}
                                <div className="mb-3">
                                  <label htmlFor="DocumentCode" className="form-label">Document Code:
                                    {selectedOptionReq?.requestcode != "New" && <span className="text-danger1">*</span>}

                                  </label>
                                  {editItemID > 0 ?
                                    <input title={formData.DocumentCode} type="text" id="RequesterDesignation" name="RequesterDesignation" className="form-control" value={formData.DocumentCode} disabled={true} />
                                    :
                                    <div
                                      title={selectedOption?.label || "Select a document code"}
                                      style={{ width: "100%" }}
                                    >
                                      <Select
                                        //onKeyDown={handleKeyDown}

                                        options={rows}
                                        value={selectedOption}
                                        name="DocumentCode"
                                        isClearable={true}
                                        //isOptionDisabled={() => selectedOptionReq.label == "New"}
                                        isSearchable={true}
                                        className={`${(selectedOptionReq?.requestcode != "New" && !ValidDraft && documentcodeerr) ? "border-on-error" : ""} ${(selectedOptionReq?.label != "New" && !ValidSubmit && documentcodeerr) ? "border-on-error" : ""}`}
                                        //className={`${(selectedOptionReq?.label != "New" && !ValidSubmit && documentcodeerr) ? "border-on-error" : ""}`}
                                        onChange={(selectedOption: any) => onSelectDocCode(selectedOption)}
                                        placeholder={selectedOptionReq == null || (selectedOptionReq != null && selectedOptionReq?.requestcode == "New")
                                          || InputDisabled ? "" : "Search Document Code"}
                                        isDisabled={selectedOptionReq == null || (selectedOptionReq != null && selectedOptionReq?.requestcode == "New")
                                          || InputDisabled || IsRecorddisabled
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
                                  <input title={formData.IssueNumber} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" value={formData.IssueNumber} />
                                </div>
                              </div>

                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="example-email" className="form-label">Revision No:</label>
                                  <input title={formData.RevisionNumber} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" value={formData.RevisionNumber} />
                                </div>
                              </div>
                              <div className="col-lg-4">

                                <div className="mb-3">
                                  <label htmlFor="example-email" className="form-label">Reference No:</label>
                                  <input title={formData.ReferenceNumber} disabled type="text" id="example-email" name="example-email" className="form-control" placeholder="" value={formData.ReferenceNumber} />
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={DocumentTypeOpt}
                                      value={selectedOptionDoctype}
                                      name="Document Type"
                                      className={` ${(!ValidSubmit && documenttypeerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectDocumentType(selectedOption)}
                                      placeholder="Search Document Type"
                                      isDisabled={InputDisabled || IsRecorddisabled || (selectedOptionReq != null && selectedOptionReq?.requestcode != "New") || formData?.Status == "Rework"}
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={LocationOpt}
                                      value={selectedOptionLoc}
                                      name="Location"
                                      className={` ${(!ValidSubmit && locationerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectLocation(selectedOption)}
                                      placeholder="Search Location"
                                      isDisabled={InputDisabled || IsRecorddisabled || (selectedOptionReq != null && selectedOptionReq?.requestcode != "New") || formData?.Status == "Rework"}
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={Custodianopt}
                                      value={selectedOptionCusto}
                                      name="Custodian"
                                      className={` ${(!ValidSubmit && custodianerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectCustodian(selectedOption)}
                                      placeholder="Search Custodian"
                                      isDisabled={InputDisabled || IsRecorddisabled || (selectedOptionReq != null && selectedOptionReq?.requestcode != "New") || formData?.Status == "Rework"}
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={Amendtype}
                                      value={selectedOptionAmend}
                                      name="Amendment Type"
                                      className={`${(!ValidSubmit && amendmenterr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectAmend(selectedOption)}
                                      placeholder="Search" isDisabled={(InputDisabled && formData?.Status != "Rework") || IsRecorddisabled}
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={Classificationopt}
                                      value={selectedOptionClass}
                                      name="Classification"
                                      className={`${(!ValidSubmit && classificationerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectClassification(selectedOption)}
                                      placeholder="Search Classification" isDisabled={(InputDisabled && formData?.Status != "Rework") || IsRecorddisabled}
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
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={TemplateTypeopt}
                                      value={SelectedOptionTemplate}
                                      name="Template Type"
                                      className={`${(!ValidSubmit && templatetypeerr) ? "border-on-error" : ""}`}
                                      onChange={(selectedOption: any) => onSelectTemplatetype(selectedOption)}
                                      placeholder="Search Template type"
                                      isDisabled={InputDisabled || !enableTemplatetype || IsRecorddisabled}
                                    />
                                  </div>

                                </div>
                              </div>
                              <div className="col-lg-4">
                                <div className="mb-3">
                                  <label htmlFor="RequesterName" className="form-label">Document Name:<span className="text-danger1">*</span></label>
                                  {/* <input type="text" id="Name" name="department" className="form-control" value={formData.Department} disabled={true} />
   */}
                                  <div
                                    title={formData.filename || "Enter file name"}
                                    style={{ width: "100%" }}
                                  >
                                    <input type="text"
                                      id="example-email"
                                      name="example-email"
                                      className={`form-control ${(!ValidDraft && filenameerr) ? "border-on-error" : ""} ${(!ValidSubmit && filenameerr) ? "border-on-error" : ""}`}
                                      //className={`form-control ${(!ValidDRecomm) ? "border-on-error" : ""}`}
                                      onChange={(e) => onChangefilename("filename", e.target.value)}
                                      //disabled={InputDisabled || IsRecorddisabled || (selectedOptionReq != null && selectedOptionReq?.requestcode != "New") || formData?.Status == "Rework"}
                                      disabled={InputDisabled || formData?.Status == "Rework"}
                                      onPaste={(e) => { handlePaste(e) }}
                                      onKeyDown={(e: React.KeyboardEvent<HTMLInputElement>) => handleKeyDowntext(e)}
                                      placeholder="Document name"
                                      value={formData.filename} />
                                  </div>

                                </div>
                              </div>
                              {console.log("FormItemIdFormItemIdFormItemId", FormItemId, modeValue, selectedOption, Attachmentarr, DocumentLink, formData?.filename)}
                              {/* //modeValue != "view" || modeValue == "edit" || modeValue != "approve"  && */}
                              {(FormItemId == null || (FormItemId != null && modeValue == "edit") || (modeValue == "view" || modeValue == "approve")
                                || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                <div className="col-lg-4">

                                  <div className="mb-3">
                                    <div className='d-flex justify-content-between'>
                                      <div>
                                        <label htmlFor="bannerImage" className="form-label">
                                          Attachment <a style={{ fontSize: '0.875rem' }} onClick={() => setShowModaltemp(true)}>
                                            <FontAwesomeIcon title='Download Template' icon={faDownload} />
                                          </a><span className="text-danger1">*</span>
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
                                      //onKeyDown={handleKeyDowntext}
                                      type="file"
                                      id="attachment"
                                      name="attachment"
                                      disabled={(InputDisabled && formData?.Status != "Rework") || IsRecorddisabled}
                                      //disabled={handleSectionState('requestedBySection')}
                                      accept=".jpg,.jpeg,.png,.gif,.bmp,.svg,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx"
                                      className={`form-control ${(!ValidSubmit && attachmenterr) ? "border-on-error" : ""}`}
                                      onChange={(e) => onFileChange(e, "bannerimg", "Document")}
                                    />

                                  </div>
                                </div>
                              }
                              {console.log("ghghghghghghgh", showpreviousattachment, "jjjj", (showpreviousattachment && (((modeValue == "view" || modeValue == "approve" ||
                                (selectedOptionReq?.requestcode != "New" && selectedOption)) ||
                                (modeValue == "edit" && formData?.Status == "Save as draft")) && DocumentLink && Attachmentarr.length == 0)))}
                              {(showpreviousattachment && selectedOptionReq?.requestcode != "New" && (((modeValue == "view" || modeValue == "approve" ||
                                (selectedOptionReq?.requestcode != "New" && selectedOption)) ||
                                (modeValue == "edit" && formData?.Status == "Save as draft")) &&
                                ((DocumentLink && Attachmentarr.length == 0) || formData?.PreviousAttachmentID != null))) &&

                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label htmlFor="DocumentCode" className="form-label">Previous Document:</label>
                                    {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}

                                    <div className="text-dark mt-0"> <span >
                                      <a onClick={() => setShowModal(true)} ><FontAwesomeIcon icon={faPaperclip} />{DocumentLink && DocumentLink.length == 1 ? "1 file Attached" : "2 files Attached"}</a>

                                    </span>
                                    </div>

                                  </div>
                                </div>
                              }
                              <div className="col-lg-4">
                                <div className="mb-3">
                                  <label htmlFor="revisionNo">Prepared By<span className="text-danger1"> *</span></label>
                                  <div title={sharewithusers.map(user => user.label).join(', ')}>
                                    <Select
                                      //onKeyDown={handleKeyDown}
                                      isClearable={true}
                                      options={rows1}
                                      isMulti
                                      value={sharewithusers}
                                      name="share with"
                                      className={`newse ${(!ValidSubmit && sharewitherr) ? "border-on-error" : ""}`}
                                      // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                      onChange={(selectedOptions: any) => onSelectsharewith(selectedOptions)}
                                      placeholder="Enter Prepared By"
                                      isDisabled={(InputDisabled && formData?.Status != "Rework") || IsRecorddisabled}
                                    />
                                  </div>
                                </div>
                              </div>
                            </div>

                          </div>
                        </div>


                        <div className="card mt-2">
                          <div className="card-body">
                            <div className='row'>
                              <div className='col-sm-12'>
                                <h3 className="text-dark font-16 fw-bold mb-3">Change Request Type<span className="text-danger1">*</span></h3>
                                {/* <label className="form-label text-muted font-16">Change Request Type</label> */}
                                <p style={{ fontSize: '11px', color: '#6c757d', marginTop: '-12px', marginBottom: '8px' }}>
                                  In case of "New Documented Information" choose the Addition in Preface, Chapters, Annexures, others and<br />
                                  In case of "Change in Existing Documented Information" choose the Revision in Preface, Chapters, Annexures, others
                                </p>
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
                              {console.log("formData?.Status ", InputDisabled, IsRecorddisabled, formData)}
                              <div className='col-sm-4'>
                                <div style={{ textAlign: "right" }} className="mt-2 float-end text-right">
                                  {/* <i style={{ cursor: "pointer" }} onClick={addField}  className="fe-plus-circle  font-20 text-warning"></i> */}
                                  {/* <i style={{ cursor: "pointer" }} className="fe-plus-circle  font-20 text-warning"></i> */}
                                  {(modeValue === "" || modeValue === "edit" || InputDisabled != true || (modeValue == "approve" && formData?.Status == "Rework")) && !IsRecorddisabled &&
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
                              <table className="mtbalenewscrollnew4 table-centered table-nowrap table-borderless mb-0" id="tbl">
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: "40px", maxWidth: "40px" }}>S.No</th>
                                    <th>Change Description<span className="text-danger1">*</span></th>
                                    <th>Reason for Change<span className="text-danger1">*</span></th>
                                    {(modeValue === "" || modeValue === "edit" || InputDisabled != true || (modeValue == "approve" && formData?.Status == "Rework")) && !IsRecorddisabled &&
                                      <th style={{ minWidth: "30px", maxWidth: "30px" }}>Action</th>
                                    }
                                  </tr>

                                </thead>
                                <tbody >
                                  {console.log("cancellReasonn", cancellReason)}



                                  {cancellReason.map((row, index) => {
                                    const hasRowError = rowErrors[index]?.descriptionError || rowErrors[index]?.reasonError;

                                    return (
                                      <tr key={index} className={hasRowError ? "row-error" : ""}>
                                        <td className="text-center" style={{ minWidth: "30px", maxWidth: "30px" }}>
                                          <div className="indexdesign" style={{ marginLeft: "0px" }}>
                                            {index + 1}
                                          </div>
                                        </td>

                                        <td title={row.description}>
                                          <textarea
                                            //onKeyDown={handleKeyDowntextarea}
                                            id="simpleinput"
                                            disabled={(InputDisabled && formData?.Status !== "Rework") || IsRecorddisabled}
                                            value={row.description}
                                            className={`form-control mb-0 ${rowErrors[index]?.descriptionError ? "border-on-error" : ""}`}
                                            onChange={(e) => {
                                              const newRows = [...cancellReason];
                                              newRows[index].description = e.target.value;
                                              setcancellReason(newRows);
                                            }}
                                          />
                                        </td>

                                        <td title={row.reason}>
                                          <textarea
                                            //onKeyDown={handleKeyDowntextarea}
                                            id="simpleinput"
                                            disabled={(InputDisabled && formData?.Status !== "Rework") || IsRecorddisabled}
                                            className={`form-control mb-0 ${rowErrors[index]?.reasonError ? "border-on-error" : ""}`}
                                            value={row.reason}
                                            onChange={(e) => {
                                              const newRows = [...cancellReason];
                                              newRows[index].reason = e.target.value;
                                              setcancellReason(newRows);
                                            }}
                                          />
                                        </td>

                                        {(modeValue === "" || modeValue === "edit" || InputDisabled !== true || (modeValue === "approve" && formData?.Status === "Rework")) && !IsRecorddisabled && (
                                          <td style={{ minWidth: "30px", maxWidth: "30px", textAlign: "center" }}>
                                            <img src={require("../assets/del.png")} onClick={() => deleteLocalFile(index, cancellReason)} />
                                          </td>
                                        )}
                                      </tr>
                                    );
                                  })}


                                </tbody>
                              </table>


                            </div>


                          </div>


                        </div>
                        {(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.IsRework == "Yes" && editID.CurrentUserRole === "Initiator")
                          &&
                          <div className="card mt-2">
                            <div className="card-body">
                              <div className='row'>
                                <div className="col-lg-12">

                                  <div className="mb-0" >

                                    <label htmlFor="example-textarea" className="form-label text-dark font-14">Remarks: <span className="text-danger1"> *</span></label>

                                    <textarea style={{ height: '80px' }} className={`form-control ${(!ValidRemark) ? "border-on-error" : ""}`}
                                      id="example-textarea"
                                      rows={5}
                                      name="Remark"
                                      value={formData.Remark}

                                      onChange={(e) => setFormData({ ...formData, Remark: e.target.value })}>

                                    </textarea>

                                  </div>

                                </div>
                              </div>

                            </div>
                          </div>

                        }
                        {console.log("editiiiiifhifassignmentt", editID, modeValue, InputDisabled, ApprovalTypeOptions, MainEditItem,
                          (modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES"),
                          (MainEditItem !== null && MainEditItem.length != 0 && MainEditItem.Status != "Save as draft" && MainEditItem.Status != "Rework" && modeValue !== "view"))}
                        {/* {((modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES") ||
                          (MainEditItem !== null && MainEditItem.length != 0 && MainEditItem.Status != "Save as draft" && MainEditItem.Status != "Rework" && modeValue !== "view"))
                          && */}
                        {(modeValue === "approve" || modeValue === "view" || modeValue === "edit") && editID != null && (editID.Status === "Pending" || editID.Status === "Approved" || editID.Status === "Rejected") && editID.CurrentUserRole !== "Initiator" &&
                          <div className="card mt-2" style={{ marginBottom: '17px' }}>
                            <div className="card-body">
                              <div className='row'>
                                <div className='col-sm-8'>
                                  <h3 className="header-title text-dark font-16 fw-bold mb-3 ">Forward Approval To (Preparer {'>'} Reviewer {'>'} Endorser {'>'} Signer\Approver)</h3>
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
                                      <th style={{ minWidth: "35px", maxWidth: "35px" }}>S.No</th>
                                      <th style={{ borderBottomLeftRadius: "0px", minWidth: "60px", maxWidth: "60px" }}>Role<span className="text-danger1">*</span></th>
                                      <th style={{ borderBottomLeftRadius: "0px", minWidth: '86px', maxWidth: '86px' }}>Responsibility<span className="text-danger1"> *</span></th>
                                      <th style={{ minWidth: "46px", maxWidth: "46px" }} title="Use your electronic digital signature to sign this document digitally">E-Sign?</th>
                                      <th style={{ minWidth: '40px', maxWidth: '40px' }} >Level<span className="text-danger1">*</span></th>
                                      <th style={{ minWidth: "110px", maxWidth: "110px" }}  >Approver Name<span className="text-danger1">*</span></th>
                                      <th style={{ minWidth: "75px", maxWidth: "75px" }} >Approval Criteria<span className="text-danger1">*</span></th>
                                      <th style={{ minWidth: '45px', maxWidth: '45px' }}>Action</th>
                                    </tr>
                                  </thead>
                                  <tbody style={{ maxHeight: "8007px", overflow: 'inherit' }}>
                                    {console.log("forwardToArrforwardToArrforwardToArr", forwardToArr, UserRoles, ApprovalTypeOptions)}
                                    {forwardToArr.map((row, index) => (

                                      <tr key={index}> <td style={{ minWidth: "35px", maxWidth: "35px" }}>
                                        <div
                                          style={{ marginLeft: "0px", overflow: 'inherit' }}
                                          className="indexdesign"
                                        >
                                          {index + 1}</div></td>
                                        <td
                                          title={row.role != 0 && UserRoles.filter((role: any) => role.value == row.role)[0]?.label}
                                          style={{ overflow: 'inherit', minWidth: '60px', maxWidth: '60px' }}
                                          className="ng-binding">
                                          <select
                                            //onKeyDown={handleKeyDown}

                                            onChange={(e) => onSelectRole(e, row.level)}
                                            value={row.role}
                                            disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                            //</td>className={`form-select ${(!Validforward) ? "border-on-error" : ""}`}
                                            className={`form-select ${row.roleError ? "border-on-error" : ""}`}
                                          >
                                            <option value="" selected>Select Role</option>
                                            {UserRoles
                                            //.filter((role: any) =>
                                              //!forwardToArr.some((r) => r.role === role.value && r.level !== row.level) || role.value === row.role // Allow the current row's role
                                            //)
                                            .map((role: any, index: number) => (
                                              <option key={index} value={role.value}

                                              //disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                              >
                                                {role.label}</option>
                                            ))}
                                          </select>

                                        </td>
                                        <td style={{ overflow: 'inherit', minWidth: '86px', maxWidth: '86px' }}>
                                          <div
                                          //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                          >
                                            <select
                                              id="responsibleId"
                                              value={row.Responsibility}
                                              onChange={(e) => handleChangeResp(e, row.level)}
                                              className={`form-select ${row.responsibilityerror ? "border-on-error" : ""}`}
                                              disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES") || row.IsPreparedBy == true || row.Responsibility === "Preparer"}
                                              title={row.Responsibility ? row.Responsibility : "Select"}
                                            >
                                              <option value="">Select </option>
                                              <option value="Preparer" disabled>Preparer</option>
                                              <option value="Reviewer">Reviewer</option>
                                              <option value="Endorser">Endorser</option>
                                              <option value="Signer">Signer</option>


                                            </select>

                                          </div>
                                        </td>
                                        <td style={{ minWidth: "46px", maxWidth: "46px", overflow: 'inherit' }}>
                                          <div
                                          //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                          >
                                            <input
                                              type="checkbox"
                                              checked={row.IsSignatureRequired}
                                              disabled={row.Responsibility === "Signer" || row.Responsibility === "Endorser" || row.Responsibility === "Preparer" || row.Responsibility === "" || (!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES"))}
                                              style={{ marginLeft: '17px', width: "15px" }}

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
                                        <td style={{ minWidth: '40px', maxWidth: '40px', overflow: 'inherit' }}>Level {index + 1}</td>
                                        {console.log("row.approvers", row.approvers)}
                                        <td style={{ overflow: 'inherit', minWidth: '110px', maxWidth: '110px', }} title={row.approvers && row.approvers.map(x => x.label).join(',')}>

                                          <Select
                                            //onKeyDown={handleKeyDown}
                                            isClearable={true}
                                            options={rows1}
                                            isMulti
                                            value={row.approvers}
                                            name="Approvers"
                                            //className={` ${(!Validforward) ? "border-on-error" : ""}`}
                                            className={`react-select ${row.approverError ? "border-on-error" : ""}`}
                                            //className={`form-control ${(!ValidDraft) ? "border-on-error" : ""} ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                            // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                            onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                            placeholder="Enter Approver Name"
                                            isDisabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES") || row.IsPreparedBy == true || row.Responsibility === "Preparer"}
                                          />
                                        </td>
                                        <td
                                          //title={ApprovalTypeOptions.filter(x => x.value = row.leveltype)[0].label}
                                          style={{ overflow: 'inherit', minWidth: '75px', maxWidth: '75px', }} className="ng-binding">
                                          {/* <select
                                            //onKeyDown={handleKeyDown}
                                            //className={`form-select ${(!Validforward) ? "border-on-error" : ""}`}
                                            className={`form-select ${row.typeError ? "border-on-error" : ""}`}
                                            onChange={(e) => onSelectApprovalType(e, row.level)}
                                            value={row.leveltype}
                                            disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES")}
                                          >
                                            <option value="" selected>Select Type</option>
                                            {ApprovalTypeOptions.map((x: any, index: number) => (
                                              <option key={index} value={x.value}

                                              >{x.label}</option>
                                            ))}
                                          </select> */}
                                          <select id="approvalType" value={row.leveltype} onChange={(e) => onSelectApprovalType(e, row.level)}
                                            className={`form-select ${row.typeError ? "border-on-error" : ""}`}
                                            disabled={!(modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES") || row.IsPreparedBy == true || row.Responsibility === "Preparer"}
                                            title={row.leveltype === "One" ? "Anyone" : row.leveltype === "All" ? "Everyone" : "Select Approval Type"}>
                                            <option value="">Select </option>
                                            <option value="One">Anyone</option>
                                            <option value="All">Everyone</option>
                                          </select>
                                        </td>
                                        <td style={{ minWidth: '45px', maxWidth: '45px', overflow: 'inherit' }}>
                                          {/* <i className="fe-trash-2 text-danger"></i> */}
                                          {editID.CurrentUserRole === "OES" && row.Responsibility !== "Preparer" && !IsRecorddisabled ? <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} /> :
                                            <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' />}
                                          {/* <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} className='sidebariconsmall' /> */}

                                        </td>
                                      </tr>

                                    ))}

                                  </tbody>
                                </table>
                              </div>
                              {modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES" &&
                                <div className="col-lg-12">

                                  <div className="mb-0" >

                                    <label htmlFor="example-textarea" className="form-label text-dark font-14">Remarks: <span className="text-danger1"> *</span></label>

                                    <textarea style={{ height: '80px' }} className={`form-control ${(!ValidRemark) ? "border-on-error" : ""}`}
                                      id="example-textarea"
                                      rows={5}
                                      name="Remark"
                                      value={formData.Remark}

                                      onChange={(e) => setFormData({ ...formData, Remark: e.target.value })}>

                                    </textarea>

                                  </div>

                                </div>
                              }
                              {modeValue === "approve" && editID != null && editID.ApprovalType === "Assignment" && editID.Status === "Pending" && editID.CurrentUserRole === "OES" &&
                                <div className="row mt-3">
                                  <div className="col-12 text-center">
                                    {/* <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardApproval("Forward")} >
                                      <i className="fe-check-circle me-1"></i> Forward
                                    </button>                                  {/* <a href="#"> */}
                                    {/* <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={() => ForwardApproval("Rework")} >
                                      <i className="fe-corner-up-left me-1"></i> Rework
                                    </button>
                                    <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={() => ForwardApproval("Rejected")} >
                                      <i className="fe-x me-1"></i> Reject
                                    </button>
                                    <button type="button" className="btn btn-light waves-effect waves-light m-1" onClick={handleCancel}>
                                      <i className="fe-x me-1"></i> Cancel
                                    </button> */}
                                    <div
                                      role="button"
                                      tabIndex={0}
                                      className="btn btn-primary waves-effect waves-light m-1"
                                      onClick={() => ForwardApproval("Forward")}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter' || e.key === ' ') {
                                          e.preventDefault();
                                          ForwardApproval("Forward");
                                        }
                                      }}
                                    >
                                      <i className="fe-check-circle me-1"></i> Forward
                                    </div>

                                    <div
                                      role="button"
                                      tabIndex={0}
                                      className="btn btn-warning waves-effect waves-light m-1"
                                      onClick={() => ForwardApproval("Rework")}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter' || e.key === ' ') {
                                          e.preventDefault();
                                          ForwardApproval("Rework");
                                        }
                                      }}
                                    >
                                      <i className="fe-corner-up-left me-1"></i> Rework
                                    </div>

                                    <div
                                      role="button"
                                      tabIndex={0}
                                      className="btn btn-danger waves-effect waves-light m-1"
                                      onClick={() => ForwardApproval("Rejected")}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter' || e.key === ' ') {
                                          e.preventDefault();
                                          ForwardApproval("Rejected");
                                        }
                                      }}
                                    >
                                      <i className="fe-x me-1"></i> Reject
                                    </div>

                                    <div
                                      role="button"
                                      tabIndex={0}
                                      className="btn btn-light waves-effect waves-light m-1"
                                      onClick={handleCancel}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter' || e.key === ' ') {
                                          e.preventDefault();
                                          handleCancel();
                                        }
                                      }}
                                    >
                                      <i className="fe-x me-1"></i> Cancel
                                    </div>

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
                              DisableApproval={false} DisableCancel={false} maxlevel={maxlevelAllprocess}
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
                        {/* <a href="../sites/ededms/SitePages/EDCMAIN.aspx">       */}
                        {/* <button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}><i className="fe-x me-1"></i> Cancel</button>
                      {/* </a> */}
                        {/* </div>
                  </div>  */}
                        {/* (((InputDisabled != true && editItemID == null && MainEditItem == null) ||
                        (MainEditItem?.Status === "Save as draft" && editID == null &&
                        (modeValue === "" || modeValue === "edit"))) ||
                        (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) &&

                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") &&
                          (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) &&
                          (modeValue === "approve") && */}

                        <div className="row mt-3">
                          <div className="col-12 text-center">
                            {(((InputDisabled != true && editItemID == null && MainEditItem == null) || 
                            (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || 
                              (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && (editID?.IsInitiator == "Yes" || editID == null) &&
                              // <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                              //   <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                              //   Save As Draft</button>
                              <div
                                role="button"
                                tabIndex={0}
                                style={{ width: '145px' }}
                                className="btn btn-primary waves-effect waves-light m-1"
                                onClick={handleSaveAsDraft}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' || e.key === ' ') {
                                    e.preventDefault();
                                    handleSaveAsDraft();
                                  }
                                }}
                              >
                                <img
                                  src={require('../../../Assets/ExtraImage/checkcircle.svg')}
                                  style={{ width: '1rem' }}
                                  className="me-1"
                                  alt="Check"
                                />
                                Save As Draft
                              </div>

                            }

                            {(((InputDisabled != true && editItemID == null && MainEditItem == null) || 
                            (MainEditItem?.Status === "Save as draft" && editID == null && (modeValue === "" || modeValue === "edit"))) || 
                              (editID && editID != null && editID.ApprovalType !== "Approval" && editID.ApprovalType !== "Assignment")) && (editID?.IsInitiator == "Yes" || editID == null) &&
                              // <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                              //   <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                              //   Submit</button>
                              <div
                                role="button"
                                tabIndex={0}
                                style={{ width: '145px' }}
                                className="btn btn-primary waves-effect waves-light m-1"
                                onClick={handleFormSubmit}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' || e.key === ' ') {
                                    e.preventDefault();
                                    handleFormSubmit();
                                  }
                                }}
                              >
                                <img
                                  src={require('../../../Assets/ExtraImage/checkcircle.svg')}
                                  style={{ width: '1rem' }}
                                  className="me-1"
                                  alt="Check"
                                />
                                Submit
                              </div>

                            }
                            {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") &&
                              // <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</button>

                              <div
                                role="button"
                                tabIndex={0}
                                style={{ width: '145px' }}
                                className="btn btn-primary waves-effect waves-light m-1"
                                onClick={() => ForwardInitiatorApproval("Save as draft")}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' || e.key === ' ') {
                                    e.preventDefault();
                                    ForwardInitiatorApproval("Save as draft");
                                  }
                                }}
                              >
                                <img
                                  src={require('../../../Assets/ExtraImage/checkcircle.svg')}
                                  style={{ width: '1rem' }}
                                  className="me-1"
                                  alt="Check"
                                />
                                Save As Draft
                              </div>

                            }
                            {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID?.IsInitiator == "Yes")) && (modeValue === "approve") &&
                              // <button style={{ width: '145px' }} type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</button>
                              <div
                                role="button"
                                tabIndex={0}
                                style={{ width: '145px' }}
                                className="btn btn-primary waves-effect waves-light m-1"
                                onClick={() => ForwardInitiatorApproval("Approved")}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' || e.key === ' ') {
                                    e.preventDefault();
                                    ForwardInitiatorApproval("Approved");
                                  }
                                }}
                              >
                                <img
                                  src={require('../../../Assets/ExtraImage/checkcircle.svg')}
                                  style={{ width: '1rem' }}
                                  className="me-1"
                                  alt="Check"
                                />
                                Submit
                              </div>

                            }
                            {((modeValue === "" || modeValue === "edit" || modeValue === "view") ||
                              (InputDisabled && editID != null && modeValue === "approve" && editID.Status === "Approved") ||
                              (editID !== null && (editID?.IsInitiator == "Yes" || editID?.IsInitiator == "No"))) &&
                              // <button style={{ width: '145px' }} type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                              //   className='me-1' alt="x" /> Cancel</button>
                              <div
                                role="button"
                                tabIndex={0}
                                style={{ width: '145px' }}
                                className="btn cancel-btn waves-effect waves-light m-1"
                                onClick={handleCancel}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' || e.key === ' ') {
                                    e.preventDefault();
                                    handleCancel();
                                  }
                                }}
                              >
                                <img
                                  src={require('../../../Assets/ExtraImage/xIcon.svg')}
                                  style={{ width: '1rem' }}
                                  className="me-1"
                                  alt="x"
                                />
                                Cancel
                              </div>

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
                        <Modal show={showModaltemp} onHide={() => setShowModaltemp(false)} size={Showfile ? "xl" : "lg"} className='newmobmodal'>
                          <Modal.Header closeButton>
                            <Modal.Title> Download Template <br></br>
                              <p className='text-muted font-14 fw-400'>Please download the provided template to proceed with creating the document
                              </p>

                            </Modal.Title>
                            {/* {ImagepostArr1.length > 0 && showBannerModal && <Modal.Title>Media Images</Modal.Title>} */}
                          </Modal.Header>
                          <Modal.Body className="" id="style-5">
                            <>
                              {Showfile ?

                                <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                :
                                <table className="mtbalenew" style={{ height: '400px', overflowY: 'auto' }} >
                                  <thead style={{ background: '#eef6f7' }}>
                                    <tr>
                                      <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                      <th style={{ minWidth: '140px', maxWidth: '1400px', textAlign: 'left' }}>File Name</th>
                                      {/* {editForm && <th>File Link</th>} */}
                                      {/* <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Upload date</th> */}
                                      {/* {!InputDisabled && <th className='text-center'>Action</th>} */}
                                      <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Action</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {console.log("Templatelink doc link", Templatelink)}
                                    {Templatelink && Templatelink?.length > 0 && (
                                      Templatelink?.map((row: any, index: number) => {
                                        return (
                                          <tr key={index}>
                                            <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>

                                            <td style={{ minWidth: '140px', maxWidth: '140px', textAlign: 'left' }} title={row != null && `${cleanFileName(row?.FileLeafRef)}`}>{row != null && `${cleanFileName(row?.FileLeafRef)}`}</td>

                                            {/* <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center' title={row && moment(row?.Created).format("DD/MMM/YYYY")}>{row && moment(row?.Created).format("DD/MMM/YYYY")}</td> */}

                                            <td style={{ minWidth: '50px', maxWidth: '50px', textAlign: 'center' }}>

                                              <span title='Download template' onClick={() => OpenFile(row != null && row, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                <FontAwesomeIcon icon={faDownload} /></span>
                                            </td>

                                          </tr>
                                        )
                                      }))}
                                  </tbody>

                                </table>
                              }
                            </>
                          </Modal.Body>
                        </Modal>
                        <Modal show={showModal} onHide={() => setShowModal(false)} size={Showfile ? "xl" : "lg"} className='newmobmodal'>
                          <Modal.Header closeButton>
                            <Modal.Title> Attachment Details <br></br>
                              <p className='text-muted font-14 fw-400 mb-0'>Below are the attachment details for Change Request
                              </p>

                            </Modal.Title>
                            {/* {Previous Document} */}
                          </Modal.Header>
                          <Modal.Body className="" id="style-5">
                            <>
                              {Showfile ?

                                <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                :
                                <table className="mtbalenew" >
                                  <thead style={{ background: '#eef6f7' }}>
                                    <tr>
                                      <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                      <th style={{ minWidth: '220px', maxWidth: '220px' }}>File Name</th>

                                      <th style={{ minWidth: '80px', maxWidth: '80px' }} className='text-center'>Upload date</th>
                                      {/* {(DocumentLink?.NDocumentCodePrinting == "Yes" || DocumentLink?.NDocumentCodePrinting == "No" || DocumentLink?.NDocumentCodePrinting == "" || DocumentLink?.NDocumentCodePrinting == undefined) && <th > Action </th>} */}
                                      <th style={{ minWidth: '60px', maxWidth: '60px' }}> Action </th>
                                    </tr>
                                  </thead>

                                  <tbody>
                                    {console.log("Attachmentarrnmnm doc link", DocumentLink, DocumentLink?.NDocumentCodePrinting, DocumentLink != null)}
                                    {/* <tr >
                                      <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>1</td>

                                      <td title={DocumentLink != null && `${cleanFileName(DocumentLink?.FileLeafRef)}`}>{DocumentLink != null && `${cleanFileName(DocumentLink?.FileLeafRef)}`}</td>
                                      <td className='text-center' title={DocumentLink && moment(DocumentLink?.Created).format("DD/MMM/YYYY")}>{DocumentLink && moment(DocumentLink?.Created).format("DD/MMM/YYYY")}</td>
                                      {(DocumentLink?.NDocumentCodePrinting == "Yes" || DocumentLink?.NDocumentCodePrinting == "No" || DocumentLink?.NDocumentCodePrinting == "" || DocumentLink?.NDocumentCodePrinting == undefined) &&
                                        <td style={{ textAlign: 'center' }}>

                                          <span title='Preview file' onClick={() => OpenFile(DocumentLink != null && DocumentLink, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                            <FontAwesomeIcon icon={faEye} /></span>
                                          <span title='Download file' onClick={() => OpenFile(DocumentLink != null && DocumentLink, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                            <FontAwesomeIcon icon={faDownload} /></span>
                                        </td>
                                      }
                                      {/* <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
                                      <img src={require("../assets/del.png")} className='' onClick={() => deleteLocalFileAttachment(0, Attachmentarr)}></img>
                                    </td> */}
                                    {/* </tr> */}
                                    {DocumentLink && DocumentLink?.length > 0 && (
                                      DocumentLink?.map((row: any, index: number) => {
                                        return (
                                          <tr key={index}>
                                            {/* <tr > */}
                                            <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>

                                            <td style={{ minWidth: '220px', maxWidth: '220px' }} title={row != null && `${cleanFileName(row?.FileLeafRef)}`}>{row != null && `${cleanFileName(row?.FileLeafRef)}`}</td>
                                            <td style={{ minWidth: '80px', maxWidth: '80px' }} className='text-center' title={row && moment(row?.Created).format("DD/MMM/YYYY")}>{row && moment(row?.Created).format("DD/MMM/YYYY")}</td>
                                            {(row?.NDocumentCodePrinting == "Yes" || row?.NDocumentCodePrinting == "No" || row?.NDocumentCodePrinting == "" || row?.NDocumentCodePrinting == undefined) &&
                                              <td style={{ textAlign: 'center', minWidth: '60px', maxWidth: '60px' }}>

                                                <span title='Preview file' onClick={() => OpenFile(row != null && row, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                  <FontAwesomeIcon icon={faEye} /></span>
                                                <span title='Download file' onClick={() => OpenFile(row != null && row, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                  <FontAwesomeIcon icon={faDownload} /></span>
                                              </td>
                                            }

                                          </tr>
                                        )
                                      }))}
                                  </tbody>

                                </table>
                              }
                            </>
                          </Modal.Body>
                        </Modal>
                        <Modal show={ShowModalAtt} onHide={() => setShowModalAtt(false)} size={Showfile ? "xl" : "lg"} className='newmobmodal'>
                          <Modal.Header closeButton>
                            <Modal.Title> Attachment Details <br></br>
                              <p className='text-muted font-14 fw-400 mb-0'>Below are the attachment details for Change Request
                              </p>

                            </Modal.Title>
                            {/* {ImagepostArr1.length > 0 && showBannerModal && <Modal.Title>Media Images</Modal.Title>} */}
                          </Modal.Header>
                          <Modal.Body className="" id="style-5">
                            <>
                              {Showfile ?

                                <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                :
                                <table className="mtbalenew" >

                                  <thead>
                                    <tr>
                                      <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                      <th style={{ minWidth: '170px', maxWidth: '170px' }}>File Name</th>
                                      {/* {showButton && */}

                                      <th style={{ minWidth: '60px', maxWidth: '60px' }} className='text-center'>Upload date</th>
                                      {((((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                        || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                        (Attachmentarr[0]?.NDocumentCodePrinting == "Yes" || Attachmentarr[0]?.NDocumentCodePrinting == "No" || Attachmentarr[0]?.NDocumentCodePrinting == "" || Attachmentarr[0]?.NDocumentCodePrinting == undefined || Attachmentarr[0]?.FileRef.includes("/ChangeRequestAttachDigitalSignedDocs/")))
                                        || (modeValue == "edit" && formData?.Status == "Save as draft")
                                        || modeValue == null || modeValue == "") &&
                                        <th style={{ minWidth: '50px', maxWidth: '50px' }}> Action </th>
                                      }
                                      {/* {((modeValue == "edit" && formData?.Status == "Save as draft") || modeValue == null || modeValue == ""
                                      || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                      <th style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}> Action </th>
                                    } */}
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {console.log("Attachmentarrnmnm attach only", Attachmentarr, showviewdownload)}
                                    {Attachmentarr.length > 0 &&

                                      <tr >
                                        <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>1</td>
                                        {/* {Attachmentarr && (Attachmentarr[0]?.FileName.includes('_') ? Attachmentarr[0]?.FileName.split('_')[2] : Attachmentarr[0]?.FileName)} */}
                                        <td style={{ minWidth: '170px', maxWidth: '170px' }} className='text-center' title={Attachmentarr && (cleanFileName(Attachmentarr?.[0]?.FileName) || cleanFileName(Attachmentarr?.[0]?.FileLeafRef))}>
                                          {Attachmentarr && (cleanFileName(Attachmentarr?.[0]?.FileName) || cleanFileName(Attachmentarr?.[0]?.FileLeafRef))}</td>
                                        {/* {showButton && */}

                                        <td style={{ minWidth: '60px', maxWidth: '60px' }} className='text-center' title={Attachmentarr && moment(Attachmentarr[0]?.Created).format("DD/MMM/YYYY")}>{Attachmentarr && moment(Attachmentarr[0]?.Created).format("DD/MMM/YYYY")}</td>
                                        {(((((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                          || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                          (Attachmentarr[0]?.NDocumentCodePrinting == "Yes" || Attachmentarr[0]?.NDocumentCodePrinting == "No" || Attachmentarr[0]?.NDocumentCodePrinting == "" || Attachmentarr[0]?.FileRef.includes("/ChangeRequestAttachDigitalSignedDocs/")))
                                          || (modeValue == "edit" && formData?.Status == "Save as draft")) ||
                                          ((modeValue == "edit" && formData?.Status == "Save as draft") || modeValue == null || modeValue == ""
                                            || (modeValue == "approve" && formData?.Status == "Rework"))) &&
                                          <td style={{ minWidth: '50px', maxWidth: '50px', textAlign: 'center' }}>
                                            {/* <span onClick={() => OpenFile(Attachmentarr && Attachmentarr[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                            <FontAwesomeIcon icon={faEye} /></span> */}
                                            {((((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                              || (modeValue == "approve" && formData?.Status == "Rework")) && showviewdownload &&
                                              (Attachmentarr[0]?.NDocumentCodePrinting == "Yes" || Attachmentarr[0]?.NDocumentCodePrinting == "No" || Attachmentarr[0]?.NDocumentCodePrinting == "" || Attachmentarr[0]?.FileRef.includes("/ChangeRequestAttachDigitalSignedDocs/")))
                                              || (modeValue == "edit" && formData?.Status == "Save as draft")) && !filechanged &&
                                              <>
                                                <span title='Preview file' onClick={() => OpenFile(Attachmentarr && Attachmentarr[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                  <FontAwesomeIcon icon={faEye} /></span>
                                                <span title='Download file' onClick={() => OpenFile(Attachmentarr && Attachmentarr[0], "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                  <FontAwesomeIcon icon={faDownload} /></span>
                                              </>
                                            }
                                            {((modeValue == "edit" && formData?.Status == "Save as draft") || modeValue == null || modeValue == ""
                                              || (modeValue == "approve" && formData?.Status == "Rework")) &&

                                              <img src={require("../assets/del.png")} className='' onClick={() => deleteLocalFileAttachment(0, Attachmentarr)}></img>

                                            }
                                          </td>
                                        }

                                      </tr>
                                    }
                                  </tbody>
                                </table>
                              }
                            </>
                          </Modal.Body>
                        </Modal>
                        <Modal show={ShowModalNotes} onHide={() => setShowModalNotes(false)} size={"xl"} className='newmobmodal'>
                          <Modal.Header closeButton style={{ borderBottom: 'none' }} className="custom-modal-header">

                            <Modal.Title> Change Request Notes <br></br>
                              <p className='text-muted font-14 fw-400 mb-0'>Kindly refer to the notes below
                              </p>

                            </Modal.Title>
                          </Modal.Header>
                          <Modal.Body className="" id="style-5">
                            <>

                              {changereqNotes && changereqNotes.length > 0 && (
                                <div style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', marginBottom: '10px' }}>

                                  <div
                                    style={{ flex: 1 }}
                                    dangerouslySetInnerHTML={{
                                      __html: changereqNotes[0]?.Notes || '<p>No content available</p>'
                                    }}
                                  ></div>
                                </div>
                              )}

                            </>
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

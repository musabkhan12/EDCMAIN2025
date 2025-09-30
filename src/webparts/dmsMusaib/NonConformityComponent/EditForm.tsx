import * as React from 'react';
import styles from './AuditPlan.module.scss';
import type { IAuditPlanProps } from './IAuditPlanProps';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  IDatePickerStyles,
  PrimaryButton,
  DefaultButton,
  Label,
  TooltipHost
} from "@fluentui/react";
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/presets/all";
import { Checkbox } from '@fluentui/react';
import Swal from 'sweetalert2';
import Select from "react-select";
import moment from 'moment';
import { auditHistoryDelegationBgColor, auditHistoryDelegationTextColor } from "../../../Shared/Constants";
import CustomBreadcrumb from '../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
import { icon } from '@fortawesome/fontawesome-svg-core';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faDownload, faEye, faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Item } from '@pnp/sp/items';
import { redirect } from 'react-router-dom';
import { getMemoNumberAuditReport, getNCNumbers } from '../AnnualAuditReportComponent/AuditReportService';
import { Modal } from 'react-bootstrap';
import FileViewer from '../ChangerequestComponent/fileviewer';
import { CONTENTTYPE_NonComformity } from '../ChangerequestComponent/Constants';
let EditSubmitStatus: any;

let EditStatus: any;
let EditCurrentUserrole: any;
let EditDelegateToSubmitStatus: any;
let EditLastInitiatorSubmitStatus: any;
let Approvallistitemid = 0;
let ApproverEmail = "";
let isRequesterDelegated: boolean = false;
let isAssignedtoDelegated: boolean = false;
let isDelegatedToDelegated: boolean = false;
let isApproverDelegated: boolean = false;
let IsAnalyzedBy: boolean = false;
let showimsupdated: boolean = false;
let showfinalapproval: boolean = false;
let isdisablefinal: boolean = false;
let showcorrectionappicable: boolean = false;
let isdisableims: boolean = false;
let showimsoncefilled: boolean = false;
let showreworkremarks: boolean = false;
let CurrentuserEmail = "";
let RequesterEmail = "";
let setloading: boolean = false;
let optionsmemoNumbernewnc: any[] = [];
let editoptsmemoAllNC: any[] = [];
let editoptsmemoAllObs: any[] = [];
let optionsmemoNumbernewobs: any[] = [];
let forwardisdisabled: boolean = false;
let Approveclicked: boolean = false;
let Rejectclicked: boolean = false;
let Reworkclicked: boolean = false;
let resubmitclicked: boolean = false;
let editforwardrecord: boolean = false;
let Showfile: boolean = false;
let depart: any = "";
const datePickerErrorStyles: Partial<IDatePickerStyles> = {
  root: {
    border: "1px solid #ffcccb", // Apply red border
    backgroundColor: "#ffcccb",
    borderRadius: "4px",
  },
  textField: {
    selectors: {
      "input": {
        color: "#ffcccb", // Optional: Change text color
      },
    },
  },
};

export interface IEditState {
  // mainItemId?: any | null;
  Auditeeuser: any;
  DelegatetoUser: any;
  Analyzedbyuser: any;
  allusersoption: any;
  editoptionsmemoNC: any;
  editoptionsmemoObs: any;
  Loading: boolean;
  Approveremailnew: string;
  mainItemId?: any;
  edType?: string;
  approvalItemId?: string;
  editDepartmentOption: any[];
  editmemonumberOptions: any[];
  editmemonumberOptionsall: any[];
  edittypeoptions: any[];
  editncType: any;
  departmentselected: any;
  ApprovedAuditSelected: any;
  NCNumberselected: any[];
  editMemoNumber: string;
  editNCNumberOptions: any[];
  editNCNumber: string;
  editApprovedAuditReport: string | number;
  editDepartment: string | number;
  editfromdepartment: string | number;
  editNCNumberID: string | number;
  editdepartmentCode: string;
  editserialNo: number;
  notUpdateDepartmentCode: string;
  notUpdateSerialNo: number;
  editfromdepartmentselected: any;
  editNCRNo: string;
  editReferenceNumber: String;
  editIssueNo: any;
  Requester: any;
  editRevisionNo: any;
  editIssueDate: any;
  editRevisionDate: any;
  TemplateDoc: any[];
  isIMSUpdated: any;
  riskandopportunitiesUpdated: any;
  correctionApplicable: any;
  notEffective: any;
  effectiveClosed: any;
  NotEffective: any;
  EffectiveandProblemClosed: any;
  IsFinalapprover: boolean;
  editCriteria: string;
  editCloseOutStatus: string;
  showcategoryothers: boolean;
  showsubcategoryothers: boolean;
  showlocationothers: boolean;
  CategoryOthers: string;
  SubCategoryOthers: string;
  LocationOthers: string;
  editCategoryCheckOption: IDropdownOption[];
  editCategoryValueIsCheck: number[];
  editSubCategoryCheckOption: IDropdownOption[];
  editSubCategoryValueIsCheck: number[];
  editLocationCheckOption: IDropdownOption[];
  editLocationValueIsCheck: number[];
  editAssignTo: string;
  editAssignToEmail: string;
  editDelegateToEmail: string;
  editAssignToId: number | null;
  editProblemDescription: string;
  editDueDate: any;
  editPersonAssigned: string;
  editPersonAssignedId: number | null;
  editDate: any;
  editDeadlineCompletion: any;
  editCorrection: string;
  editRootCause: string;
  editCorrectiveAction: string;
  editDelegateTo: string;
  editDelegateToId: number | null;
  editAnalyzedBy: string;
  editAnalyzedById: number | null;
  editReviewedBy: string;
  editReviewedById: number | null;
  editCorrectiveActionImplementedOn: string;
  editSubmitedDate: any;
  editSubmitStatus: string;
  editStatus: string;
  editIsRework: string;
  editSerialNumber: string;
  editDocumentCode: string;
  editAttachmentPreArray: any[];
  editAttachmentJson: any[];
  editErrors: { [key: string]: string };
  editCurrentUserRole: string;
  editFirstInitiatorSubmitStatus: string;
  editFirstAssignedToSubmitStatus: string;
  editDelegateToSubmitStatus: string;
  editAnalyzedBySubmitStatus: string;
  editReviewedBySubmitStatus: string;
  editLastAssignedToSubmitStatus: string;
  editLastInitiatorSubmitStatus: string;
  disable: boolean;
  editProcessCurrentUserRole: string;
  editProcessStatus: string;
  editProcessListNameId: number | null;
  editProcessListName: string;
  editProcessActionTakenById: number | null;
  editProcessActionTakenBy: string;
  processListItemID: number | null;
  approvers: any[];
  optionsRole: IDropdownOption[];
  apprDelId: any[];
  indApp: number;
  mainListId: number;
  formNameId: number;
  reqRolId: number;
  ncItemId: number | null;
  remarks: string;
  reworkremarks: string;
  showDialog: boolean;
  showDialogauditee: boolean;
  copyFil: any[];
  copyFilauditee: any[];
  fileCount: number;
  exFiles: any[];
  ShowDeleteicon: boolean;
  ShowDeleteiconInitiator: boolean;
  fileDeleteId: any[];
  fileDeleteIdauditee: any[];
  files: FileList;
  filesauditee: FileList;
  fileCountauditee: number;
  exFilesauditee: any[];
  apprItems: any[];
  isDisabled: boolean;
  showApprove: boolean;
  showSubmit: boolean;
  showDraft: boolean;
  deptSectionDisable: boolean;
  forwarDisable: boolean;
  showForward: boolean;
  showReject: boolean;
  showDelegate: boolean;
  redirecturl: string;
  ShowModalTemplateDoc: boolean;
  ShowModalAtt: boolean;
  hidedigisign: boolean;
  showdigisign: boolean;
  DigitalsignID: any;
}
const optionsApp: IDropdownOption[] = [
  { key: 'One', text: 'Anyone' },
  { key: 'All', text: 'Everyone' }


]
const optionsResponsibility: IDropdownOption[] = [

  { key: 'Signer', text: 'Signer' },
  { key: 'Reviewer', text: 'Reviewer' }

]
export default class EditForm extends React.Component<IAuditPlanProps, IEditState> {

  constructor(props: IAuditPlanProps) {
    super(props);
    const selectedTextDiv = document.getElementById('selectedText');
    selectedTextDiv.style.display = 'none';
    // const breadcrumbElement = document.getElementById("breadcrumb");
    // breadcrumbElement.style.display = 'none';

    this.state = {
      // mainItemId: props.edItm || null,
      Auditeeuser: [],
      DelegatetoUser: [],
      Analyzedbyuser: [],
      allusersoption: [],
      editoptionsmemoNC: [],
      editoptionsmemoObs: [],
      Loading: false,
      Approveremailnew: "",
      mainItemId: '',
      edType: this.props.edType,
      approvalItemId: this.props.approvalItemId,
      editDepartmentOption: [],
      editmemonumberOptions: [],
      editmemonumberOptionsall: [],
      departmentselected: [],
      NCNumberselected: [],
      ApprovedAuditSelected: [],
      edittypeoptions: [],
      showsubcategoryothers: false,
      showlocationothers: false,
      showcategoryothers: false,

      CategoryOthers: "",
      SubCategoryOthers: "",
      LocationOthers: "",
      editncType: "",
      isIMSUpdated: "No",
      riskandopportunitiesUpdated: "No",
      correctionApplicable: "",
      notEffective: "",
      effectiveClosed: "",
      NotEffective: "",
      EffectiveandProblemClosed: "",
      IsFinalapprover: false,
      editMemoNumber: "",
      editNCNumberOptions: [],
      editNCNumber: "",
      editApprovedAuditReport: "",
      editNCNumberID: "",
      editDepartment: "",
      editfromdepartment: "",
      editdepartmentCode: "",
      editserialNo: 0,
      notUpdateDepartmentCode: "",
      notUpdateSerialNo: 0,
      editfromdepartmentselected: [],
      editCriteria: "",
      editNCRNo: "",
      Requester: null,
      editReferenceNumber: "",
      editIssueNo: "",
      editRevisionNo: "",
      editIssueDate: null,
      editRevisionDate: null,
      TemplateDoc: [],
      editCloseOutStatus: "",
      editCategoryCheckOption: [],
      editCategoryValueIsCheck: [],
      editSubCategoryCheckOption: [],
      editSubCategoryValueIsCheck: [],
      editLocationCheckOption: [],
      editLocationValueIsCheck: [],
      editAssignTo: "",
      editAssignToEmail: "",
      editDelegateToEmail: "",
      editAssignToId: null,
      editProblemDescription: "",
      editDueDate: null,
      editPersonAssigned: "",
      editPersonAssignedId: null,
      editDate: null,
      editDeadlineCompletion: null,
      editCorrection: "",
      editRootCause: "",
      editCorrectiveAction: "",
      editDelegateTo: "",
      editDelegateToId: null,
      editAnalyzedBy: "",
      editAnalyzedById: null,
      editReviewedBy: "",
      editReviewedById: null,
      editCorrectiveActionImplementedOn: "",
      editSubmitedDate: null,
      editSubmitStatus: "",
      editStatus: "",
      editIsRework: "No",
      editSerialNumber: "",
      editDocumentCode: "",
      editAttachmentPreArray: [],
      editAttachmentJson: [],
      editErrors: {},
      editCurrentUserRole: "",
      editFirstInitiatorSubmitStatus: "",
      editFirstAssignedToSubmitStatus: "",
      editDelegateToSubmitStatus: "",
      editAnalyzedBySubmitStatus: "",
      editReviewedBySubmitStatus: "",
      editLastAssignedToSubmitStatus: "",
      editLastInitiatorSubmitStatus: "",
      disable: true,
      editProcessCurrentUserRole: "",
      editProcessStatus: "",
      editProcessListNameId: null,
      editProcessListName: "",
      editProcessActionTakenById: null,
      editProcessActionTakenBy: "",
      processListItemID: null,
      approvers: [{
        Role: "", Level: "", Name: "", Type: "One", Index: 0, appEx: "", itemId: "", Responsibility: "Signer", IsSignatureRequired: true,
        responsibilityerror: false
      }],
      optionsRole: [],
      apprDelId: [],
      indApp: 0,
      mainListId: 0,
      formNameId: 0,
      reqRolId: 0,
      ncItemId: null,
      remarks: "",
      reworkremarks: "",
      showDialog: false,
      showDialogauditee: false,
      copyFil: [],
      fileCount: 0,
      exFiles: [],
      ShowDeleteicon: false,
      ShowDeleteiconInitiator: false,
      fileDeleteId: [],
      fileDeleteIdauditee: [],
      files: {} as FileList,
      copyFilauditee: [],
      filesauditee: {} as FileList,
      fileCountauditee: 0,
      exFilesauditee: [],
      apprItems: [],
      isDisabled: false,
      showApprove: false,
      showSubmit: false,
      showDraft: false,
      deptSectionDisable: false,
      forwarDisable: false,
      showForward: false,
      showReject: false,
      showDelegate: false,
      redirecturl: "",
      ShowModalTemplateDoc: false,
      ShowModalAtt: false,
      hidedigisign: false,
      showdigisign: false,
      DigitalsignID: null
    };
    this.addApprover = this.addApprover.bind(this);
    this.deleteItemApp = this.deleteItemApp.bind(this);
    this._getPeoplePickerItemsApp = this._getPeoplePickerItemsApp.bind(this);
    this.onRoleChange = this.onRoleChange.bind(this);
    this.onTypeChange = this.onTypeChange.bind(this);
    this._OpenModal = this._OpenModal.bind(this);
    this._CloseModal = this._CloseModal.bind(this);
    this.removeFiles = this.removeFiles.bind(this);
    this._OpenModalauditee = this._OpenModalauditee.bind(this);
    this._CloseModalauditee = this._CloseModalauditee.bind(this);
    this.removeFilesAuditee = this.removeFilesAuditee.bind(this);
    this.toBeDeleted = this.toBeDeleted.bind(this);
    this.toBeDeletedauditee = this.toBeDeletedauditee.bind(this);

    this.handleFileChange = this.handleFileChange.bind(this);
    this.updateData = this.updateData.bind(this);
  }
  //   async componentDidMount(){


  //     // Extracting the part after `#/`
  //     const url=window.location.href;

  //     const parts = url.split("#/")[1].split("/");

  //     const programName = decodeURIComponent(parts[0]); // "Non Confirmity"
  //     const editType = parts[1]; // "edit"
  //     const id = parts[2]; // "165"

  //     console.log("Program Name:", programName);
  //     console.log("Edit Type:", editType);
  //     console.log("ID:", id);
  //     // alert(`${programName},${editType},${id}`);

  //     // Set state
  //     // this.setState((prevState) => ({
  //     //   edItm: id ? id : prevState.edItm, // Update only if `id` exists
  //     //   edType: editType ? editType : prevState.edType // Update only if `editType` exists
  //     // }));
  //     if(id){
  //       this.setState({mainItemId:id})
  //     }
  //     if(editType){
  //       this.setState({edType:editType})
  //     }

  //     // alert(`after setting edtype and edItem ${this.state.edItm},${this.state.edType}`);
  //     if(this.state.edType=="approve"){
  //     //  alert("approve");
  //     const approvalItemId= parts[3];
  //     if(approvalItemId){
  //       this.setState({approvalItemId:approvalItemId})
  //     }


  //       this.setState({isDisabled:true})
  //       this.setState({showApprove:true});
  //       this.setState({showSubmit: false});

  //     }
  //     else if(this.state.edType =="edit"){
  // //alert("edit");

  // this.setState({isDisabled:false})
  // this.setState({showApprove:false});
  // this.setState({showSubmit: true});

  //     }
  //     else if(this.state.edType=="view"){

  //       this.setState({isDisabled:true})
  // this.setState({showApprove:false});
  // this.setState({showSubmit: false});

  //     }
  //     else{
  // //alert("new");

  //     }
  //   }
  // private handleFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
  //   this.setState({ copyFil: [], exFiles: [] })
  //   if (e.target.files) {
  //     _self.setState({ fileCount: e.target.files.length });
  //     _self.setState({ files: e.target.files });
  //     var allfiles: any[] = [];
  //     [].forEach.call(e.target.files, function (file: File) {
  //       allfiles.push(file);
  //     })
  //     _self.setState({ copyFil: allfiles });
  //   }
  // };
  // private handleFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
  //   if (e.target.files) {
  //     const fileArray = Array.from(e.target.files);
  //     _self.setState({
  //       copyFil: fileArray,
  //       fileCount: fileArray.length,
  //       files: e.target.files // raw FileList
  //     });
  //   }
  // }
  private handleFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
    if (e.target.files) {
      const newFilesArray = Array.from(e.target.files); // Convert FileList to Array
      const existingFiles = _self.state.copyFil || []; // previously uploaded

      const allFiles = [...existingFiles, ...newFilesArray];

      _self.setState({
        fileCount: allFiles.length,
        files: e.target.files, // optional: might not represent all files now
        copyFil: allFiles
      });
    }
  }
  // private handleAuditeeFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
  //   this.setState({ copyFilauditee: [], exFilesauditee: [] });
  //   if (e.target.files) {
  //     _self.setState({ fileCountauditee: e.target.files.length });
  //     _self.setState({ filesauditee: e.target.files });
  //     var allfiles: any[] = [];
  //     [].forEach.call(e.target.files, function (file: File) {
  //       allfiles.push(file);
  //     })
  //     _self.setState({ copyFilauditee: allfiles });
  //   }
  // };
  // private handleAuditeeFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
  //   if (e.target.files) {
  //     const filesArray = Array.from(e.target.files);
  //     _self.setState({
  //       fileCountauditee: filesArray.length,
  //       filesauditee: e.target.files,
  //       copyFilauditee: filesArray
  //     });
  //   }
  // }
  private handleAuditeeFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
    if (e.target.files) {
      const newFilesArray = Array.from(e.target.files); // Convert FileList to Array
      const existingFiles = _self.state.copyFilauditee || []; // previously uploaded

      const allFiles = [...existingFiles, ...newFilesArray];

      _self.setState({
        fileCountauditee: allFiles.length,
        filesauditee: e.target.files, // optional: might not represent all files now
        copyFilauditee: allFiles
      });
    }
  }
  private Breadcrumb = [
    {
      MainComponent: "My Request",
      MainComponentURl: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EDCMAIN.aspx`,
    },
    {
      ChildComponent: "Non Conformity" + " / " + "Observation",
      ChildComponentURl: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EDCMAIN.aspx#/NonConformity`,
    },
  ];
  private _OpenModal() {
    this.setState({
      showDialog: true
    });
  }
  private _CloseModal() {
    this.setState({
      showDialog: false
    });
  }
  private _OpenModalauditee() {
    this.setState({
      showDialogauditee: true
    });
  }
  private _CloseModalauditee() {
    this.setState({
      showDialogauditee: false
    });
  }
  private removeFiles(i: number) {
    var items = this.state.copyFil.filter(function (it, val) {
      if (val != i) {
        return it;
      }
    })
    this.setState({ copyFil: items, fileCount: items.length });
    if (items.length === 0) {
      const fileInput = document.getElementById("newfile") as HTMLInputElement;
      if (fileInput) {
        fileInput.value = "";
      }
    }
  }
  private clearFileInputIfEmpty() {
    const fileInput = document.getElementById("newfile") as HTMLInputElement;
    if (fileInput) fileInput.value = "";
  }
  private removeFilesAuditee(i: number) {
    const remainingFiles = this.state.copyFilauditee.filter((_, index) => index !== i);

    this.setState({
      copyFilauditee: remainingFiles,
      fileCountauditee: remainingFiles.length
    });

    if (remainingFiles.length === 0) {
      const fileInput = document.getElementById("newfile") as HTMLInputElement;
      if (fileInput) fileInput.value = "";
    }
    if (remainingFiles.length === 0) this.clearFileInputIfEmpty();
  }

  // private removeFilesAuditee(i: number) {
  //   var items = this.state.copyFilauditee.filter(function (it, val) {
  //     if (val != i) {
  //       return it;
  //     }
  //   })
  //   this.setState({ copyFilauditee: items, fileCountauditee: items.length });
  //   if (items.length === 0) {
  //     const fileInput = document.getElementById("newfile") as HTMLInputElement;
  //     if (fileInput) {
  //       fileInput.value = "";
  //     }
  //   }
  // }
  private toBeDeleted(i: number) {
    var itemId = this.state.exFiles.filter(function (it, val) {
      return val == i
    })

    var remFiles = this.state.exFiles.filter(function (it, val) {
      return val != i
    })
    //var exFiles:any[]= 
    this.state.fileDeleteId.push(itemId[0].Id)
    this.setState({ fileDeleteId: this.state.fileDeleteId });
    this.setState({ exFiles: remFiles, fileCount: remFiles.length });
    if (remFiles.length === 0) {
      const fileInput = document.getElementById("newfile") as HTMLInputElement;
      if (fileInput) {
        fileInput.value = "";
      }
    }
  }
  private toBeDeletedauditee(i: number) {
    const deletedItem = this.state.exFilesauditee[i];
    const remainingFiles = this.state.exFilesauditee.filter((_, index) => index !== i);
    const updatedDeleteIds = [...this.state.fileDeleteIdauditee, deletedItem.Id];

    this.setState({
      exFilesauditee: remainingFiles,
      fileDeleteIdauditee: updatedDeleteIds,
      fileCountauditee: remainingFiles.length
    });

    if (remainingFiles.length === 0) {
      const fileInput = document.getElementById("newfile") as HTMLInputElement;
      if (fileInput) fileInput.value = "";
    }
    if (remainingFiles.length === 0) this.clearFileInputIfEmpty();
  }

  // private toBeDeletedauditee(i: number) {
  //   var itemId = this.state.exFilesauditee.filter(function (it, val) {
  //     return val == i
  //   })

  //   var remFiles = this.state.exFilesauditee.filter(function (it, val) {
  //     return val != i
  //   })
  //   //var exFiles:any[]= 
  //   this.state.fileDeleteIdauditee.push(itemId[0].Id)
  //   this.setState({ fileDeleteIdauditee: this.state.fileDeleteIdauditee });
  //   this.setState({ exFilesauditee: remFiles, fileCountauditee: remFiles.length });
  //   if (remFiles.length === 0) {
  //     const fileInput = document.getElementById("newfile") as HTMLInputElement;
  //     if (fileInput) {
  //       fileInput.value = "";
  //     }
  //   }
  // }
  public changeDepartment = (item: any): void => {
    let selecteddepartment = this.state.editDepartmentOption.filter((x: any) => x.value == item.value);
    this.setState({ editDepartment: item.value, editdepartmentCode: item.data.departmentCode, departmentselected: selecteddepartment });
  };
  public async getUniqueBy(array: any[], key: string) {

    const seen: any[] = [];
    const result = [];

    for (let i = 0; i < array.length; i++) {
      const value = array[i][key];
      if (value != null && !seen.includes(value)) {
        seen.push(value);
        result.push(array[i]);
      }
    }

    return result;
  }
  public handleChangeCategoryOthers = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {

    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      CategoryOthers: value,
    }));
  };
  public handleChangeSubCategoryOthers = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {

    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      SubCategoryOthers: value,
    }));
  };
  public handleChangeLocationOthers = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {

    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      LocationOthers: value,
    }));
  };
  public changeMemoNumber = async (item: any): Promise<void> => {

    const sp = spfi().using(SPFx(this.props.context));
    let nctypenew: string = this.state.editncType == "NC" ? "NC Number" : "Observation NUmber";
    //let NCNumberoptionnew = await getNCNumbers(sp, item.text, nctypenew);
    let optionsNCNumber: any = [];
    const NCNumberoptionnew = item && await getNCNumbers(sp, item?.label, nctypenew);
    let existingrecords = item && await this.getNCdata(item?.label, this.state.editncType);
    if (Array.isArray(NCNumberoptionnew) && NCNumberoptionnew.length > 0) {
      // Safely extract existing NCNumbers, even if the array is empty
      const existingNCNumbersSet = new Set(
        (Array.isArray(existingrecords) ? existingrecords : []).map((rec: any) => rec.NCNumber)
      );

      // Filter out NCNumbers already in existingrecords
      optionsNCNumber = NCNumberoptionnew[0]
        .filter((entry: any) => !existingNCNumbersSet.has(entry.NCNumber))
        .map((entry: any) => ({
          value: entry.ID,
          label: entry.NCNumber,
          ncNo: entry.NCNumber,
          reportcode: entry.ReportCode,
          nctype: entry.NCType
        }));
    }
    debugger
    let optionsNCNumbernew: any[] = [];
    optionsNCNumbernew = optionsNCNumber.length > 0 && await this.getUniqueBy(optionsNCNumber, "ncNo");
    optionsNCNumbernew && optionsNCNumbernew.sort((a, b) => a.label.localeCompare(b.label));
    let approvedauditreportselected = this.state.editmemonumberOptions.filter((x: any) => x.value == item?.value);
    const selectedOption = this.state.editDepartmentOption.find(user => user?.value === approvedauditreportselected[0]?.department);
    this.setState({ editNCNumberOptions: optionsNCNumbernew })
    this.setState({
      editApprovedAuditReport: item?.value, editMemoNumber: item?.memoNumber, ApprovedAuditSelected: approvedauditreportselected,
      departmentselected: selectedOption, editDepartment: approvedauditreportselected[0]?.department

    });
    this.setState({ editNCNumber: "", editNCNumberID: "", NCNumberselected: [] });
  };
  public async getNCdata(reportcode: string, nctype: string) {

    const sp = spfi().using(SPFx(this.props.context));
    let arr: any[] = []
    let arrs = []
    let bannerimg = []
    const currentUser = await sp.web.currentUser();
    await sp.web.lists.getByTitle("NonConformityList").items
      .select("*")
      .expand("")
      .filter(`ApprovedAuditReportMemoNumber eq '${reportcode}' and NCType eq '${nctype}'`)
      .orderBy("Modified", false)
      ()
      .then((res: any) => {
        //console.log(res, 'Memonumbers from audit report');

        //arr.push(res)
        arr = res;
      })
      .catch((error: any) => {
        console.log("Error fetching data: ", error);
      });
    console.log(arr, 'arr bnbnbn');
    return arr;
  }
  public async getNCdatadraft(reportcode: string, nctype: string, itemid: number) {

    const sp = spfi().using(SPFx(this.props.context));
    let arr: any[] = []
    let arrs = []
    let bannerimg = []
    const currentUser = await sp.web.currentUser();
    await sp.web.lists.getByTitle("NonConformityList").items
      .select("*")
      .expand("")
      .filter(`ApprovedAuditReportMemoNumber eq '${reportcode}' and NCType eq '${nctype}' and ID ne ${itemid}`)
      .orderBy("Modified", false)
      ()
      .then((res: any) => {
        //console.log(res, 'Memonumbers from audit report');

        //arr.push(res)
        arr = res;
      })
      .catch((error: any) => {
        console.log("Error fetching data: ", error);
      });
    console.log(arr, 'arr bnbnbn');
    return arr;
  }
  public changeNCNumber = (item: any): void => {
    let ncnumberselected = this.state.editNCNumberOptions.filter((x: any) => x.value == item?.value);
    this.setState({ editNCNumber: item?.label, editNCNumberID: item?.value, NCNumberselected: ncnumberselected });
  };
  public handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      [name]: value,
      reworkremarks: value
    }));
  };
  public handleChangeReworkremarks = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      [name]: value,
    }));
  };

  public handleChangeobsdescription = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {

    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      editProblemDescription: value,
    }));
  };
  private _handlePeoplePickerChange = (field: keyof IEditState, idField: keyof IEditState) => (items: any[]) => {
    if (items.length > 0) {
      this.setState({
        [field]: items[0].text,
        [idField]: items[0].id,
      } as Pick<IEditState, keyof IEditState>);
    } else {
      this.setState({
        [field]: "",
        [idField]: null,
      } as unknown as Pick<IEditState, keyof IEditState>);
    }
  };
  private onSelectAuditee = (selectedOptions: any, field: keyof IEditState, idField: keyof IEditState) => {
    if (selectedOptions.length > 0) {
      this.setState({
        [field]: selectedOptions[0].label,
        [idField]: selectedOptions[0].value,
      } as Pick<IEditState, keyof IEditState>);
    } else {
      this.setState({
        [field]: "",
        [idField]: null,
      } as unknown as Pick<IEditState, keyof IEditState>);
    }
    const uniqueOptions = (selectedOptions || []).filter(
      (option: any, index: any, self: any) =>
        index === self.findIndex((o: any) => o.value === option.value)
    );
    this.setState({ Auditeeuser: uniqueOptions })

  };
  private onSelectDelegateto = (selectedOptions: any, field: keyof IEditState, idField: keyof IEditState) => {
    if (selectedOptions.length > 0) {
      this.setState({
        [field]: selectedOptions[0].label,
        [idField]: selectedOptions[0].value,
      } as Pick<IEditState, keyof IEditState>);
    } else {
      this.setState({
        [field]: "",
        [idField]: null,
      } as unknown as Pick<IEditState, keyof IEditState>);
    }
    const uniqueOptions = (selectedOptions || []).filter(
      (option: any, index: any, self: any) =>
        index === self.findIndex((o: any) => o.value === option.value)
    );
    this.setState({ DelegatetoUser: uniqueOptions })

  };
  private onSelectAnalyzedby = (selectedOptions: any, field: keyof IEditState, idField: keyof IEditState) => {
    if (selectedOptions.length > 0) {
      this.setState({
        [field]: selectedOptions[0].label,
        [idField]: selectedOptions[0].value,
      } as Pick<IEditState, keyof IEditState>);
    } else {
      this.setState({
        [field]: "",
        [idField]: null,
      } as unknown as Pick<IEditState, keyof IEditState>);
    }
    const uniqueOptions = (selectedOptions || []).filter(
      (option: any, index: any, self: any) =>
        index === self.findIndex((o: any) => o.value === option.value)
    );
    this.setState({ Analyzedbyuser: uniqueOptions })

  };
  private _handleCheckboxChange = (stateKey: keyof IEditState, itemKey: number, itemtext: String) =>
    (_ev: React.FormEvent<HTMLElement>, isChecked?: boolean) => {
      // if (stateKey == "editCategoryValueIsCheck" && itemtext == "Others") {
      //   this.setState({ showcategoryothers: true })
      // } else if (stateKey == "editSubCategoryValueIsCheck" && itemtext == "Others") {
      //   this.setState({ showsubcategoryothers: true })
      // } else if (stateKey == "editLocationValueIsCheck" && itemtext == "Others") {
      //   this.setState({ showlocationothers: true })
      // } 
      // this.setState((prevState) => {
      //   const updatedValues = isChecked
      //     ? [...(prevState[stateKey] as number[]), itemKey]
      //     : (prevState[stateKey] as number[]).filter((key) => key !== itemKey);
      //   return { [stateKey]: updatedValues } as unknown as Pick<IEditState, keyof IEditState>;
      // });
      const newState: Partial<IEditState> = {};

      if (stateKey === "editCategoryValueIsCheck" && itemtext === "Others") {
        newState.showcategoryothers = isChecked || false;
        if (!isChecked) newState.CategoryOthers = ""; // Clear text
      }

      if (stateKey === "editSubCategoryValueIsCheck" && itemtext === "Others") {
        newState.showsubcategoryothers = isChecked || false;
        if (!isChecked) newState.SubCategoryOthers = ""; // Clear text
      }

      if (stateKey === "editLocationValueIsCheck" && itemtext === "Others") {
        newState.showlocationothers = isChecked || false;
        if (!isChecked) newState.LocationOthers = ""; // Clear text
      }
      this.setState((prevState) => {
        const updatedValues = isChecked
          ? [...(prevState[stateKey] as number[]), itemKey]
          : (prevState[stateKey] as number[]).filter((key) => key !== itemKey);
        return {
          ...newState,
          [stateKey]: updatedValues,
        } as unknown as Pick<IEditState, keyof IEditState>;
      });
    };
  private getnctypeoptions = async () => {
    const _sp = spfi().using(SPFx(this.props.context));
    const field2 = await _sp.web.lists.getByTitle("NonConformityList").fields.getByInternalNameOrTitle("NCType")();
    const choices: string[] = field2["Choices"];

    // Convert string array to dropdown option objects
    const dropdownOptions = choices.map(choice => ({
      key: choice,
      text: choice
    }));

    this.setState({ edittypeoptions: dropdownOptions });
  };
  public async componentDidMount() {

    const _sp = spfi().using(SPFx(this.props.context));
    const currentUser = await _sp.web.currentUser();
    const users = await _sp.web.siteUsers();
    const people = users.filter(user => user.PrincipalType === PrincipalType.User);

    const Selectedoptions = people.map(item => ({
      value: item.Id,
      label: item.Title,
      UserName: item.Title,
      UserEmail: item.Email
    }));
    this.setState({ allusersoption: Selectedoptions })
    CurrentuserEmail = currentUser.Email;
    // Extracting the part after `#/`
    const url = window.location.href;
    // alert(
    //   url + "url"
    // )
    //let filenamee = "12_20250814131526501_Chapter Sample.doc";
    //this.cleanFileName(filenamee);
    const parts = url.split("#/")[1].split("/");

    const programName = decodeURIComponent(parts[0]); // "Non Confirmity"
    const editType = parts[1]; // "edit"
    const id = parts[2]; // "165"
    const sp = spfi().using(SPFx(this.props.context));
    console.log("Program Name:", programName);
    console.log("Edit Type:", editType);
    console.log("ID:", id);
    // alert("Program Name:"+ programName)
    // alert("Edit Type:"+ editType)
    // alert("ID:"+ id)
    this.setState({ Loading: true });
    setloading = true;
    await this.getnctypeoptions();
    await this.getDepartment();
    //await this.getAuditreport();
    await this.getDataRoles();
    await this.getMainListName();
    await this.getRequestorRole();
    await this.getFormName();
    if (id) {
      this.setState({ mainItemId: id }, async () => {
        if (editType === "view") {
          // this.setState({ isDisabled: true });
          // this.setState({ deptSectionDisable: true });
          // this.setState({ showDelegate: true });
          // this.setState({ forwarDisable: true });
          // this.setState({ showApprove: false });
          // this.setState({ showSubmit: false });
          // this.setState({ showDraft: false });
          // this.setState({ showForward: false });
          // this.setState({ showReject: false });
          this.setState({
            isDisabled: true,
            deptSectionDisable: true,
            showDelegate: true,
            forwarDisable: true,
            showApprove: false,
            showSubmit: false,
            showDraft: false,
            showForward: false,
            showReject: false,
          });
        }
        // if (editType === "approve") {
        //   const approvalItemId = parts[3];
        //   // console.log("approvalItemId",approvalItemId)
        //   // alert("approvalItemId"+ approvalItemId)
        //   // alert("approvalItemId"+ typeof(approvalItemId))
        //   if (approvalItemId) {
        //     Approvallistitemid = Number(approvalItemId);
        //     this.setState({ approvalItemId: approvalItemId })
        //     // alert('here is my state ' + this.state.approvalItemId)
        //   }
        //   // this.setState({ isDisabled: true });
        //   // this.setState({ deptSectionDisable: true });
        //   // this.setState({ showDelegate: true });
        //   // this.setState({ showApprove: true });
        //   // this.setState({ showSubmit: false });
        //   // this.setState({ showDraft: false });
        //   // this.setState({ showReject: false });
        //   // this.setState({
        //   //   isDisabled: true,
        //   //   deptSectionDisable: true,
        //   //   showDelegate: true,
        //   //   showApprove: true,
        //   //   showSubmit: false,
        //   //   showDraft: false,
        //   //   showReject: false,
        //   // });
        //   // if (EditLastInitiatorSubmitStatus == "No" && EditCurrentUserrole == "LastInitiator") {
        //   //   this.setState({ showApprove: false, showForward: true });
        //   //   //this.setState({ showForward: true });
        //   // }
        //   // if (EditLastInitiatorSubmitStatus == "Yes" && EditCurrentUserrole == "LastInitiator") {
        //   //   this.setState({ showApprove: false, forwarDisable: true, showForward: false });
        //   //   // this.setState({ forwarDisable: true });
        //   //   // this.setState({ showForward: false });
        //   //   // this.setState({ showReject: true });
        //   // }
        //   // if (EditLastInitiatorSubmitStatus == "Yes") {
        //   //   if (EditCurrentUserrole == "Approverrole") {
        //   //     this.setState({ showApprove: false, forwarDisable: true, showForward: false, showReject: true });
        //   //     // this.setState({ forwarDisable: true });
        //   //     // this.setState({ showForward: false });
        //   //     // this.setState({ showReject: true });
        //   //   }

        //   //   forwardisdisabled = true;
        //   // }

        //   // if (EditStatus == "Approved" || EditStatus == "Rejected") {
        //   //   this.setState({ showReject: false });
        //   // }
        // }
        // this.getAllapprovalitems(Number(id));
        // This will run AFTER the state update is completed
        // alert("Updated mainItemId: " + this.state.mainItemId);
        await this.getListData(Number(id)); // Fetch list data after updating state
        await this.getGeneratedTemplateDocNC(Number(id));
        // if (editType === "edit") {
        //   // this.setState({ showApprove: false });
        //   // this.setState({ showSubmit: true });
        //   // this.setState({ showForward: false });
        //   // this.setState({ showReject: false });
        //   // this.setState({ forwarDisable: false });
        //   this.setState({
        //     showApprove: false,
        //     showSubmit: true,
        //     showForward: false,
        //     showReject: false,
        //     forwarDisable: false,
        //   });
        //   if (EditSubmitStatus === 'No' || (EditStatus == "Rework" && EditCurrentUserrole == "FirstInitiator")) {
        //     this.setState({ showDraft: true, isDisabled: false });
        //     //this.setState({ isDisabled: false });
        //   }
        //   else {
        //     // this.setState({ deptSectionDisable: false });
        //     // this.setState({ showDelegate: false });
        //     // this.setState({ showDraft: false });
        //     // this.setState({ isDisabled: true });
        //     this.setState({
        //       deptSectionDisable: false,
        //       showDelegate: false,
        //       showDraft: false,
        //       isDisabled: true,
        //     });

        //   }
        //   if (EditDelegateToSubmitStatus == "No" && EditCurrentUserrole == "DelegateTo") {
        //     this.setState({ showDelegate: true });
        //   }
        // }


        // else if (editType === "view") {
        //   // this.setState({ isDisabled: true });
        //   // this.setState({ deptSectionDisable: true });
        //   // this.setState({ showDelegate: true });
        //   // this.setState({ forwarDisable: true });
        //   // this.setState({ showApprove: false });
        //   // this.setState({ showSubmit: false });
        //   // this.setState({ showDraft: false });
        //   // this.setState({ showForward: false });
        //   // this.setState({ showReject: false });
        //   this.setState({
        //     isDisabled: true,
        //     deptSectionDisable: true,
        //     showDelegate: true,
        //     forwarDisable: true,
        //     showApprove: false,
        //     showSubmit: false,
        //     showDraft: false,
        //     showForward: false,
        //     showReject: false,
        //   });
        // }
        // else

      });
    }
    if (editType) {
      this.setState({ edType: editType })
    }

    // await this.getListData();

    this.setState({ Loading: false });
    setloading = false;
    await this.getAuditreport();

    // setTimeout(() => {
    //   this.setState({ Loading: false });
    //   setloading = false;
    // }, 7000); // 5000ms = 5 seconds
  }
  public async getGeneratedTemplateDocNC(itemId: number) {

    const _sp = spfi().using(SPFx(this.props.context));
    let results: any = [];
    // for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("NonConformityDigitalSignedDocs").items
      .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
      .then(async (res) => {
        console.log(res, 'tem let arrs=[]');
        if (res.length > 0) {
          results = res;
          this.setState({ TemplateDoc: res })
        } else {
          await _sp.web.lists.getByTitle("NonConformityGeneratedTemplateDoc").items
            .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
            .then((res) => {
              console.log(res, 'tem let arrs=[]');
              results = res;
              this.setState({ TemplateDoc: res })
            })
        }

      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });

    // }
    console.log(results, 'results');
    return results;
  }
  public async CheckIfAlreadyactionTaken(id: number, processName: string) {

    const _sp = spfi().using(SPFx(this.props.context));
    try {
      const currentUser = await _sp.web.currentUser();

      const item = await _sp.web.lists
        .getByTitle("ProcessApprovalList")
        .items
        .getById(id)
        .select("Id", "ActionTakenById", "ActionTakenOn", "AssignedTo/Id", "ProcessName")
        .expand("AssignedTo")();

      const isUnprocessed = (!item.ActionTakenById || item.ActionTakenById == null) && (!item.ActionTakenOn || item.ActionTakenOn == null);

      // Optional: further check if it's assigned to current user and matches processName


      if (isUnprocessed) {
        return true;
      }

      return false;
    } catch (error) {
      console.error("Error in CheckIfAlreadyactionTaken:", error);
      return false;
    }
  };
  public async getapprovalbyID(id: number, processName: string) {

    const _sp = spfi().using(SPFx(this.props.context));
    let arr: any[] = []
    let arrs = []
    let bannerimg = []
    let val = "Yes"
    let Sts = "Save as draft";
    let sts = "Pending"
    //const currentUser = await _sp.web.currentUser();
    const today = new Date().toISOString();
    const currentUser = await _sp.web.currentUser();
    // await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    //   .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title,AssignedTo/EMail").expand("Author,RequesterName,AssignedTo")()
    //   .then((res) => {
    //     console.log(res, 'ghghghghgh let arrs=[]');
    //     if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName) {
    //       arr.push(res);
    //     }
    //   })
    await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
      .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title,AssignedTo/EMail").expand("Author,RequesterName,AssignedTo")()
      .then(async (res) => {
        console.log(res, 'ghghghgh let arrs=[]');

        //arr = res;

        await _sp.web.lists.getByTitle("ARGDelegateList").items
          .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
          .expand("Author,DelegateName,ActingFor")
          .filter(`DelegateName/ID eq '${res.AssignedTo.Id}' and ActingFor/ID eq '${currentUser.Id}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
          .orderBy("Created", false).top(5000)()
          .then(async (result) => {
            // if (result.length > 0) {

            //   if (res && (res.AssignedTo.Id == currentUser.Id || res.AssignedTo.Id == result[0].ActingForId) && res.ProcessName === processName && (res?.Status == "Pending" || res?.Status === "Save as draft")
            //     // && res.Level === 0
            //     ) {
            //     // arr = res;
            //     arr.push(res);
            //   }

            // }
            if (result.length > 0) {
              const isAssignedToUserOrActingFor = result.some(r =>
                res && (
                  res.AssignedTo.Id === currentUser.Id ||
                  res.AssignedTo.Id === r.DelegateNameId
                )
              );

              if (
                isAssignedToUserOrActingFor &&
                res.ProcessName === processName &&
                (res.Status === "Pending" || res.Status === "Save as draft")
                // && res.Level === 0
              ) {
                arr.push(res);
              }
            }
            else {

              if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft")
                //&& res.Level === 0
              ) {
                // arr = res;
                arr.push(res);
              }


            }


          })
          .catch((error) => {
            console.log("Error fetching data: ", error);
          });

      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    console.log(arr, 'arr approval of current user');
    return arr;
  }
  public async getListData(idNumber: any) {
    let url = window.location.href;
    let parts = url.split("#/")[1].split("/");
    let editType = parts[1]; // "edit"
    let id = parts[2];
    let currentlevel: any;
    let finallevel: any;
    this.setState({ Loading: true });
    setloading = true;
    const sp = spfi().using(SPFx(this.props.context));
    const currentUser = await sp.web.currentUser();
    // let memoopt = await this.getAuditreport().then(async (x) => {
    //const auditData = await this.getAuditreport();
    //console.log("Audit Report Result", auditData);
    const memoItems = await getMemoNumberAuditReport(sp);
    // console.log("Fetched memoItems:", memoItems);

    if (!Array.isArray(memoItems) || memoItems.length === 0) {
      console.warn("No memo items returned or not an array");
      return;
    }

    // Flatten if memoItems is nested: [ [ items ] ]
    const flatMemoItems = Array.isArray(memoItems[0]) ? memoItems[0] : memoItems;
    const optionsallnc = flatMemoItems.map((item: any) => ({
      value: item.ID,
      label: item.ReportCode,
      itemId: item.ID,
      reportCode: item.ReportCode,
      ncNo: item.NCNumber,
      department: item.DepartmentAuditedId
    }));
    console.log("nmnngfhjagfhjdagfjadfdhjmnm", flatMemoItems, this.state.editmemonumberOptions, optionsmemoNumbernewnc, editoptsmemoAllNC, editoptsmemoAllObs);
    //let departopt = await this.getDepartment();
    const deptItems = await sp.web.lists.getByTitle("ProcessDepartmentMasterList").items();
    const optionsdept = deptItems.map((item: {
      DepartmentCode: any; Title: string; Id: number, ADDepartmentName: string
    }) => ({
      value: item.Id,
      label: item.Title,
      adDepartmentName: item.ADDepartmentName,
      data: { departmentCode: item.DepartmentCode },
    }));

    try {
      const Items: any = await sp.web.lists.getByTitle("NonConformityList").items.getById(idNumber)
        .select("*, Category/Id, Category/Title, SubCategory/Id, Location/Id, Location/Title, SubCategory/Title, AssignedTo/Id, AssignedTo/Title,AssignedTo/EMail,DelegateTo/EMail, DelegateTo/Id, DelegateTo/Title, AnalyzedBy/Id, AnalyzedBy/Title, ReviewedBy/Id, ReviewedBy/Title, PersonAssigned/Id, PersonAssigned/Title,Author/Id,Author/Title,Author/EMail")
        .expand("Category, SubCategory, Location, AssignedTo, DelegateTo, AnalyzedBy, ReviewedBy, PersonAssigned,Author")();
      console.log("Itemsedit", Items);
      EditSubmitStatus = Items.SubmitStatus;

      EditStatus = Items.Status;
      EditCurrentUserrole = Items.CurrentUserRole;
      EditDelegateToSubmitStatus = Items.DelegateToSubmitStatus;
      EditLastInitiatorSubmitStatus = Items.LastInitiatorSubmitStatus;
      let editmemonumberOptionsselect = Items?.NCType == "NC" ? this.state.editoptionsmemoNC : this.state.editoptionsmemoObs;
      let approvedauditreportselected = optionsallnc && optionsallnc.filter((x: any) => Number(x.value) == Number(Items.ApprovedAuditReportId));
      let selecteddepartment = optionsdept.filter((x: any) => x.value == Items.DepartmentId);
      let fromselecteddepartment = optionsdept.filter((x: any) => x.value == Items.FromDepartmentId);
      const showCategoryOthers = Items.Category?.some((cat: any) => cat.Title === "Others") || false;
      const showSubCategoryOthers = Items.SubCategory?.some((sub: any) => sub.Title === "Others") || false;
      const showLocationOthers = Items.Location?.some((loc: any) => loc.Title === "Others") || false;
      //let auditreport=this.state.editmemonumberOptions.length > 0 && Items.ApprovedAuditReportId && this.state.editmemonumberOptions.filter((item: any) => item?.value == Items.ApprovedAuditReportId)
      this.setState({
        ApprovedAuditSelected: approvedauditreportselected,
        departmentselected: selecteddepartment,
        editfromdepartmentselected: fromselecteddepartment,
        ncItemId: Items.Id,
        editDepartment: Items.DepartmentId,
        editfromdepartment: Items.FromDepartmentId,
        editncType: Items.NCType,
        isIMSUpdated: Items.IMSUpdated,
        riskandopportunitiesUpdated: Items.RiskOpportunitiesUpdated,
        correctionApplicable: Items.Correctionapplicable == "Yes" ? true : false,
        notEffective: Items.NotEffective == "Yes" ? true : false,
        effectiveClosed: Items.EffectiveandProblemClosed == "Yes" ? true : false,
        editMemoNumber: Items.ApprovedAuditReportMemoNumber,
        editNCNumber: Items.NCNumber,
        editApprovedAuditReport: Items.ApprovedAuditReportId,
        editNCNumberID: Items.NCNumberID,
        editCriteria: Items.Criteria,
        editNCRNo: Items.NCRNo,
        editReferenceNumber: Items.ReferenceNumber,
        editDocumentCode: Items.DocumentCode,
        editRevisionDate: Items.RevisionDate,
        editIssueDate: Items.IssueDate,
        editRevisionNo: Items.RevisionNumber,
        editIssueNo: Items.IssueNumber,
        CategoryOthers: Items.CategoryOthers,
        LocationOthers: Items.LocationOthers,
        SubCategoryOthers: Items.SubCategoryOthers,
        editCloseOutStatus: Items.CloseOutStatus,
        editCategoryValueIsCheck: Items.Category ? Items.Category.map((cat: any) => cat.Id) : [],
        editSubCategoryValueIsCheck: Items.SubCategory ? Items.SubCategory.map((sub: any) => sub.Id) : [],
        editLocationValueIsCheck: Items.Location ? Items.Location.map((loc: any) => loc.Id) : [],
        editAssignToId: Items.AssignedTo ? Items.AssignedTo.Id : null,
        editAssignTo: Items.AssignedTo ? Items.AssignedTo.Title : null,
        editDelegateToEmail: Items.DelegateTo ? Items.DelegateTo.EMail : null,
        editAssignToEmail: Items.AssignedTo ? Items.AssignedTo.EMail : null,
        editProblemDescription: Items.ProblemDescription,
        editDueDate: Items.DueDate ? new Date(Items.DueDate) : null,
        editPersonAssignedId: Items.AssignedTo ? Items.AssignedTo.Id : null,
        editPersonAssigned: Items.AssignedTo ? Items.AssignedTo.Title : null,
        editDate: Items.Date ? new Date(Items.Date) : null,
        editDeadlineCompletion: Items.DueDate ? new Date(Items.DueDate) : null,
        editCorrection: Items.Correctionproblem,
        editRootCause: Items.RootCause,
        editCorrectiveAction: Items.CorrectiveAction,
        editDelegateToId: Items.DelegateTo ? Items.DelegateTo.Id : null,
        editDelegateTo: Items.DelegateTo ? Items.DelegateTo.Title : null,
        editAnalyzedById: Items.AnalyzedBy ? Items.AnalyzedBy.Id : null,
        editAnalyzedBy: Items.AnalyzedBy ? Items.AnalyzedBy.Title : null,
        editReviewedById: Items.ReviewedBy ? Items.ReviewedBy.Id : null,
        editReviewedBy: Items.ReviewedBy ? Items.ReviewedBy.Title : null,
        editCorrectiveActionImplementedOn: Items.CorrectiveActionImplementedOn,
        editSubmitStatus: Items.SubmitStatus,
        editCurrentUserRole: Items.CurrentUserRole,
        editFirstInitiatorSubmitStatus: Items.FirstInitiatorSubmitStatus,
        editFirstAssignedToSubmitStatus: Items.FirstAssignedToSubmitStatus,
        editDelegateToSubmitStatus: Items.DelegateToSubmitStatus,
        editAnalyzedBySubmitStatus: Items.AnalyzedBySubmitStatus,
        editReviewedBySubmitStatus: Items.ReviewedBySubmitStatus,
        editLastAssignedToSubmitStatus: Items.LastAssignedToSubmitStatus,
        editLastInitiatorSubmitStatus: Items.LastInitiatorSubmitStatus,
        editStatus: Items.Status,
        editAttachmentPreArray: [],
        editAttachmentJson: [],
        notUpdateDepartmentCode: Items.NCRNo,
        notUpdateSerialNo: Items.SerialNumber,
        Requester: Items.Author,
        //remarks: Items.FinalRemarks,
        reworkremarks: Items.ReworkRemarks,
        showcategoryothers: showCategoryOthers,
        showlocationothers: showLocationOthers,
        showsubcategoryothers: showSubCategoryOthers
        //}, () => {

        // this.setState({ showcategoryothers: showCategoryOthers, showlocationothers: showLocationOthers, showsubcategoryothers: showSubCategoryOthers });
      });
      let delegatetouser: any[] = [];

      if (Items.DelegateTo) {
        const approver = Items.DelegateTo; // single object
        delegatetouser = [{
          value: approver.Id,
          label: approver.Title,
          UserName: approver.Title,
          UserEmail: approver.EMail
        }];

      }

      let Assignedtouser: any[] = [];

      if (Items.AssignedTo) {
        const approver = Items.AssignedTo; // single object
        Assignedtouser = [{
          value: approver.Id,
          label: approver.Title,
          UserName: approver.Title,
          UserEmail: approver.EMail
        }];
      }
      let AnalyzedByuser: any[] = [];

      if (Items.AnalyzedBy) {
        const approver = Items.AnalyzedBy; // single object
        AnalyzedByuser = [{
          value: approver.Id,
          label: approver.Title,
          UserName: approver.Title,
          UserEmail: approver.EMail
        }];
      }
      this.setState({ DelegatetoUser: delegatetouser, Auditeeuser: Assignedtouser, Analyzedbyuser: AnalyzedByuser })
      if (editType == "edit") {
        this.setState({
          showApprove: false,
          showSubmit: true,
          showForward: false,
          showReject: false,
          forwarDisable: false,
        });
        if (Items.SubmitStatus === 'No' || (Items.Status == "Rework" && Items.CurrentUserRole == "FirstInitiator")) {
          this.setState({ showDraft: true, isDisabled: false });
        }
        else {

          this.setState({
            deptSectionDisable: false,
            showDelegate: false,
            showDraft: false,
            isDisabled: true,
          });
        }
        if (Items.DelegateToSubmitStatus == "No" && Items.CurrentUserRole == "DelegateTo") {
          this.setState({ showDelegate: true });
        }
        // if (currentApprover.CurrentUserRole !== "LastInitiator") {
        showreworkremarks = true;
        showimsupdated = true;
        isdisableims = false;
        // }
      }
      if (editType === "approve") {
        const approvalItemId = parts[3];
        if (approvalItemId) {
          Approvallistitemid = Number(approvalItemId);
          this.setState({ approvalItemId: approvalItemId })
          // alert('here is my state ' + this.state.approvalItemId)
        }

        this.setState({
          isDisabled: true,
          deptSectionDisable: true,
          showDelegate: true,
          showApprove: true,
          showSubmit: false,
          showDraft: false,
          showReject: false,
        });
      }
      // if (Items?.NCType) {
      //   this.setState({
      //     editncType: Items?.NCType,
      //     editmemonumberOptions: Items?.NCType == "NC" ? auditData.optionsmemoNumbernewnc : auditData.optionsmemoNumbernewobs
      //   });
      // }
      var cnt: any = 0;
      var appItems: any[] = [];
      const apprItems = await sp.web.lists
        .getByTitle("ProcessApprovalList")
        .items.select(
          "*",
          "AssignedTo/Title,AssignedTo/Id,AssignedTo/EMail,ActionTakenRole,ActionTakenRole/Role,RequesterName/Title,ActionTakenBy/Title"
        )
        .expand("AssignedTo,ActionTakenRole,RequesterName,ActionTakenBy")
        .filter(
          "ListItemId eq '" +
          idNumber +
          "' and ProcessName eq 'Non Conformity'"
        )
        .orderBy("Id", true)();
      if (apprItems && apprItems.length > 0) {
        apprItems.forEach(async function (itm: any) {
          //Audit Report
          var objToAdd: any = {};
          objToAdd["Level"] = itm.Level;
          objToAdd["AssignedTo"] = itm.AssignedTo?.Title;
          objToAdd["AssignedToEmail"] = itm.AssignedTo?.EMail;
          objToAdd["RequesterName"] = itm.RequesterName?.Title;
          objToAdd["ActionTakenRole"] = itm.ActionTakenRoleId == null ? itm.CurrentUserRole : itm.ActionTakenRole?.Role;
          {
            /* Divyansh Changes */
          }
          if (itm.RequestedDate == "" || itm.RequestedDate == null) {
            objToAdd["RequestedDate"] = "";
          } else {
            objToAdd["RequestedDate"] = itm.RequestedDate;
          }
          if (itm.ActionTakenById != null) {
            objToAdd["ActionTakenBy"] = itm.ActionTakenBy.Title;
          }
          else {
            objToAdd["ActionTakenBy"] = "";
          }

          objToAdd["ActionTakenOn"] = itm.ActionTakenOn;
          objToAdd["Remarks"] = itm.Remark;
          objToAdd["Status"] = itm.Status;
          objToAdd["Index"] = itm.Level;

          appItems.push(objToAdd);
          cnt = cnt + 1;
        });
        this.setState({ apprItems: appItems });
      }
      const newItem1 = await this.getdigitalsignaturerequestbyID("NonConformityList", sp, Number(idNumber));
      const isRecordExist = await this.getdigitalsignaturerequestbyIDYes("NonConformityList", sp, Number(idNumber));
      console.log("newItem1newItem1", newItem1, isRecordExist);

      if (newItem1.length > 0) {
        this.setState({ DigitalsignID: newItem1[0].ID })
      }
      if (isRecordExist == "Yes" || isRecordExist == "NoRecord") {
        this.setState({ showdigisign: false });
      } else {
        this.setState({ showdigisign: true });
      }
      // setTimeout(() => {
      //   this.setState({ Loading: false });
      //   setloading = false;
      // }, 7000);
      debugger
      let optnew: any;
      let nctypenew: string = Items.NCType == "NC" ? "NC Number" : "Observation NUmber";
      let NCNumberoptionnew = await getNCNumbers(sp, Items.ApprovedAuditReportMemoNumber, nctypenew);
      if (NCNumberoptionnew.length > 0) {
        optnew = NCNumberoptionnew[0].map((item: any) => ({
          value: item.ID,
          label: item.NCNumber,
          ncNo: item.NCNumber,
          reportcode: item.ReportCode,
          nctype: item.NCType
        }));
      }
      let ncnumberselectednew = optnew.length > 0 && optnew.filter((x: any) => x.value == Items.NCNumberID);

      this.setState({
        //editNCNumberOptions: optionsNCNumbernew,
        NCNumberselected: ncnumberselectednew,

      })
      let existingrecords = Items && await this.getNCdatadraft(Items.ApprovedAuditReportMemoNumber, Items.NCType, Items.ID);
      let optionsNCNumber: any = [];
      if (Array.isArray(NCNumberoptionnew) && NCNumberoptionnew.length > 0) {
        // Safely extract existing NCNumbers, even if the array is empty
        const existingNCNumbersSet = new Set(
          (Array.isArray(existingrecords) ? existingrecords : []).map((rec: any) => rec.NCNumber)
        );

        // Filter out NCNumbers already in existingrecords
        optionsNCNumber = NCNumberoptionnew[0]
          .filter((entry: any) => !existingNCNumbersSet.has(entry.NCNumber))
          .map((entry: any) => ({
            value: entry.ID,
            label: entry.NCNumber,
            ncNo: entry.NCNumber,
            reportcode: entry.ReportCode,
            nctype: entry.NCType
          }));
      }

      // if (NCNumberoptionnew.length > 0) {
      //   optionsNCNumber = NCNumberoptionnew[0].map((item: any) => ({
      //     value: item.ID,
      //     label: item.NCNumber,
      //     ncNo: item.NCNumber,
      //     reportcode: item.ReportCode,
      //     nctype: item.NCType
      //   }));
      // }
      let optionsNCNumbernew: any[] = [];

      // console.log("nmnnmnm", this.state.editmemonumberOptions, optionsmemoNumbernewnc, editoptsmemoAllNC, editoptsmemoAllObs);
      optionsNCNumbernew = optionsNCNumber.length > 0 && await this.getUniqueBy(optionsNCNumber, "ncNo");

      let ncnumberselected = optionsNCNumbernew.filter((x: any) => x.value == Items.NCNumberID);

      this.setState({
        editNCNumberOptions: optionsNCNumbernew,
        NCNumberselected: ncnumberselected,

      })
      //this.setState({ editApprovedAuditReport: item.key, editMemoNumber: item.memoNumber });
      RequesterEmail = Items.Author.EMail;
      isRequesterDelegated = await this.isUserDelegatedFor(Items.Author?.EMail, currentUser.Id);
      isAssignedtoDelegated = await this.isUserDelegatedFor(Items.AssignedTo && Items.AssignedTo?.EMail, currentUser.Id);
      isDelegatedToDelegated = await this.isUserDelegatedFor(Items.DelegateTo && Items.DelegateTo?.EMail, currentUser.Id);




      EditSubmitStatus = Items.SubmitStatus;

      EditStatus = Items.Status;
      EditCurrentUserrole = Items.CurrentUserRole;
      EditDelegateToSubmitStatus = Items.DelegateToSubmitStatus;
      EditLastInitiatorSubmitStatus = Items.LastInitiatorSubmitStatus;

      if (editType === "approve") {
        if (Items.LastInitiatorSubmitStatus == "No" && Items.CurrentUserRole == "LastInitiator") {
          this.setState({ showApprove: false, showForward: true });
          //this.setState({ showForward: true });
        }
        if (Items.LastInitiatorSubmitStatus == "Yes" && Items.CurrentUserRole == "LastInitiator") {
          this.setState({ showApprove: false, forwarDisable: true, showForward: false });
        }
        if (Items.LastInitiatorSubmitStatus == "Yes") {
          if (EditCurrentUserrole == "Approverrole") {
            this.setState({ showApprove: false, forwarDisable: true, showForward: false, showReject: true });
          }
          forwardisdisabled = true;
        }

        if (EditStatus == "Approved" || EditStatus == "Rejected") {
          this.setState({ showReject: false });
        }
        let approvalItemIdnew = parts[3];
        let Approverdata = await this.getapprovalbyID(Number(approvalItemIdnew), "Non Conformity");
        console.log("Approverdata", Approverdata, "Approverdata0", Approverdata && Approverdata[0], Approvallistitemid, approvalItemIdnew);
        const today = new Date().toISOString();

        let isCurrentuserdelagated: boolean = false;
        if (Approverdata.length > 0) {
          const currentApprover = Approverdata[0];
          let delegateduser = await sp.web.lists.getByTitle("ARGDelegateList").items
            .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
            .expand("Author,DelegateName,ActingFor")
            .filter(`DelegateName/EMail eq '${currentApprover.AssignedTo?.EMail}' and ActingFor/ID eq '${currentUser.Id}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
            .orderBy("Created", false).top(5000)()
            .then(async (result) => {
              console.log("resultresultresult", result);
              if (result.length > 0) {
                isCurrentuserdelagated = true;
              }
            });
          const isAnalyzedRole = currentApprover.CurrentUserRole === "AnalyzedBy";
          const isFirstAssigned = currentApprover.CurrentUserRole === "FirstAssignedTo" || currentApprover.CurrentUserRole === "DelegateTo";
          const isCurrentUser = (currentApprover.CurrentUserRole === "FirstAssignedTo" && (currentApprover.AssignedTo?.EMail === CurrentuserEmail || isCurrentuserdelagated)) ||
            (currentApprover.CurrentUserRole === "DelegateTo" && (currentApprover.DelegateTo?.EMail === CurrentuserEmail || isCurrentuserdelagated)) ||
            (currentApprover.CurrentUserRole === "AnalyzedBy" && (currentApprover.AssignedTo?.EMail === CurrentuserEmail || isCurrentuserdelagated));

          const currentstatus = currentApprover.Status == "Approved";
          const finalstatus = Items.Status == "Approved";
          currentlevel = currentApprover.Level;
          finallevel = currentApprover.Maxlevel;
          ApproverEmail = currentApprover.AssignedTo?.EMail;
          isApproverDelegated = await this.isUserDelegatedFor(currentApprover.AssignedTo?.EMail, currentUser.Id);

          if (this.state.editLastInitiatorSubmitStatus == "No" && this.state.editCurrentUserRole == "LastInitiator" && (currentApprover.AssignedTo?.EMail === CurrentuserEmail || isCurrentuserdelagated)) {
            forwardisdisabled = false
          } else {
            forwardisdisabled = true;
          }
          // if (isFirstAssigned && isCurrentUser) {
          //   IsAnalyzedBy = true;
          //   showimsupdated = true;
          //   isdisableims = false;
          // }
          if (currentApprover.CurrentUserRole !== "LastInitiator") {
            showreworkremarks = true;
          }
          // This will override previous value only if role is "FirstAssignedTo"
          if (Items.IMSUpdated != "" && Items.IMSUpdated != null && Items.RiskOpportunitiesUpdated != "" && Items.RiskOpportunitiesUpdated != null && !(isFirstAssigned && isCurrentUser)) {
            showimsupdated = true;
            isdisableims = true;
          }
          if (Items.Status != "Approved" && Items.CurrentUserRole == "LastInitiator") {
            showcorrectionappicable = true;
            isdisablefinal = false;
            //this.setState({ IsFinalapprover: isFinalApprover });
          } else {
            isdisablefinal = true;
          }
          if (Items.Status == "Approved") {
            showcorrectionappicable = true;
            isdisablefinal = true;
          }

          // showfinalapproval = currentstatus;
          // isdisablefinal = finalstatus;
        }

      }
      if (editType == "view") {
        this.setState({
          isDisabled: true,
          deptSectionDisable: true,
          showDelegate: true,
          forwarDisable: true,
          showApprove: false,
          showSubmit: false,
          showDraft: false,
          showForward: false,
          showReject: false,
        });
        if (Items.IMSUpdated != "" && Items.IMSUpdated != null && Items.RiskOpportunitiesUpdated != "" && Items.RiskOpportunitiesUpdated != null) {
          showimsupdated = true;
          isdisableims = true;
        }

        if (Items.Status == "Approved") {
          showcorrectionappicable = true;
          isdisablefinal = true;
        }
      }

      //  console.log("apprItems111", apprItems);
      if (apprItems && apprItems.length > 0) {
        if (Items.SubmitStatus == "Yes" && (Items.CurrentUserRole == "FirstAssignedTo" || Items.CurrentUserRole == "DelegateTo")) {
          //this.setState({ approvalItemId: apprItems[0].ID })
          if (apprItems.length == 1) {
            Approvallistitemid = apprItems[0].ID
          } else if (apprItems.length > 1) {
            let initatoritem = apprItems.filter((x) => (x.CurrentUserRole == "FirstAssignedTo" || x.CurrentUserRole == "DelegateTo") && x.Status == "Pending")
            if (initatoritem.length > 0) {
              Approvallistitemid = initatoritem[0].ID
            }
          }

        }
        if ((Items.SubmitStatus == "Yes" || Items.SubmitStatus == "No") && Items.Status == "Rework" && Items.CurrentUserRole == "FirstInitiator" && apprItems.length > 1) {
          //this.setState({ approvalItemId: apprItems[0].ID })
          let initatoritem = apprItems.filter((x) => x.CurrentUserRole == "FirstInitiator" && x.Status == "Pending")
          if (initatoritem.length > 0) {
            Approvallistitemid = initatoritem[0].ID
          }

        }
        apprItems.forEach(async function (itm: any) {
          //Audit Report
          var objToAdd: any = {};
          objToAdd["Level"] = itm.Level;
          objToAdd["AssignedTo"] = itm.AssignedTo?.Title;
          objToAdd["AssignedToEmail"] = itm.AssignedTo?.EMail;
          objToAdd["RequesterName"] = itm.RequesterName?.Title;
          objToAdd["ActionTakenRole"] = itm.ActionTakenRoleId == null ? itm.CurrentUserRole : itm.ActionTakenRole?.Role;
          {
            /* Divyansh Changes */
          }
          if (itm.RequestedDate == "" || itm.RequestedDate == null) {
            objToAdd["RequestedDate"] = "";
          } else {
            objToAdd["RequestedDate"] = itm.RequestedDate;
          }
          if (itm.ActionTakenById != null) {
            objToAdd["ActionTakenBy"] = itm.ActionTakenBy.Title;
          }
          else {
            objToAdd["ActionTakenBy"] = "";
          }

          objToAdd["ActionTakenOn"] = itm.ActionTakenOn;
          objToAdd["Remarks"] = itm.Remark;
          objToAdd["Status"] = itm.Status;
          objToAdd["Index"] = itm.Level;

          appItems.push(objToAdd);
          cnt = cnt + 1;
        });
        this.setState({ apprItems: appItems });
      }
      //Get Files
      const upFiles = await sp.web.lists.getByTitle("NonConformityDocs").items.select("*", "File/Name,FileLeafRef,FileRef,EncodedAbsUrl").expand("File")
        .filter("NonConformityId eq '" + idNumber + "'")();
      if (upFiles.length > 0) {
        var obJFiles: any[] = [];
        let fCount: number = 0;
        upFiles.forEach(async function (item: any) {
          obJFiles.push({
            "Name": item?.File.Name,
            "type": "old",
            "Id": item.Id,
            "FileRef": item.FileRef,
            "FileLeafRef": item.FileLeafRef,
            "Uploaded": item.Modified,
            "Path": item.EncodedAbsUrl
          })
        })
        fCount = upFiles.length;
        this.setState({ exFiles: obJFiles, fileCount: fCount });
      }
      const upFilesauditee = await sp.web.lists.getByTitle("AuditeeAttachmentDocs").items.select("*", "File/Name,FileLeafRef,FileRef,EncodedAbsUrl").expand("File")
        .filter("NonConformityId eq '" + idNumber + "'")();
      if (upFilesauditee.length > 0) {
        var obJFiles: any[] = [];
        let fCount: number = 0;
        upFilesauditee.forEach(async function (item: any) {
          obJFiles.push({
            "Name": item?.File.Name,
            "type": "old",
            "Id": item.Id,
            "FileRef": item.FileRef,
            "FileLeafRef": item.FileLeafRef,
            "Uploaded": item.Modified,
            "Path": item.EncodedAbsUrl
          })
        })
        fCount = upFilesauditee.length;
        this.setState({ exFilesauditee: obJFiles, fileCountauditee: fCount });
      }
      if (Items.SubmitStatus == "No" || (Items.Status == "Rework" && Items.CurrentUserRole == "FirstAssignedTo") ||
        (Items.CurrentUserRole == "FirstAssignedTo" && (Items.FirstAssignedToSubmitStatus == "No" || Items.FirstAssignedToSubmitStatus == undefined)) ||
        (Items.CurrentUserRole == "DelegateTo" && (Items.delegateToSubmitStatus == "No" || Items.delegateToSubmitStatus == undefined))) {
        this.setState({ ShowDeleteicon: true })
      }
      if (Items.SubmitStatus == "No" || Items.FirstInitiatorSubmitStatus == "No"
      ) {
        this.setState({ ShowDeleteiconInitiator: true })
      }
      //AllProcessApproval Table data

      const approvalItems = await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.select("*", "Approvers/Name", "Approvers/Title", "Approvers/EMail", "Approvers/ID").expand("Approvers")
        .filter("MainListID eq '" + idNumber + "'and ProcessName eq 'Non Conformity'")
        .orderBy("Level", true)();
      var allApp: any[] = [];
      var cnt: any = 0;
      const sorted = [...approvalItems].sort((a, b) => b.Level - a.Level);



      if (approvalItems.length > 0) {
        approvalItems.forEach(async function (itm: any) {
          var objToAdd: any = {};
          let approve: string[] = [];
          itm.Approvers.forEach(function (it: any) {
            approve.push(it.Name.split('|membership|')[1].toString());
            // ids.push(it.Id);
          })
          objToAdd["Role"] = itm.ApproverRoleId;
          objToAdd["Type"] = itm.LevelType;
          objToAdd["Name"] = itm.ApproversId;
          objToAdd["AppName"] = itm.Approvers;
          objToAdd["itemId"] = itm.Id;
          objToAdd["Index"] = itm.Level;
          objToAdd["appEx"] = approve;
          objToAdd["Responsibility"] = itm.Responsibility || "";
          objToAdd["IsSignatureRequired"] = itm.IsSignatureRequired == "Yes" ? true : false;
          allApp.push(objToAdd);
          cnt = cnt + 1;
        });
        editforwardrecord = true;
        //const sorted = [...allApp].sort((a, b) => a.index - b.index);
        //this.setState({ approvers: sorted });
        this.setState({ approvers: allApp });
        const approvers = sorted[0]?.Approvers || [];
        const currentUserEmail = this.props.context.pageContext.user.email;

        const finalApprover = approvers.map((user: any) => ({
          id: user.ID,
          email: user.EMail
        }));
        // const isFinalApprover = finalApprover.some(
        //   (approver: any) => approver.email?.toLowerCase() === currentUserEmail?.toLowerCase() || isCurrentuserdelagated
        // ) && Items.CurrentUserRole == "Approverrole";
        const today = new Date().toISOString();
        let isFinalApprover = false;

        for (const approver of finalApprover) {
          const approverEmail = approver.email?.toLowerCase();
          const currentUserEmailLower = currentUserEmail?.toLowerCase();

          // Direct match
          if (approverEmail === currentUserEmailLower && Items.CurrentUserRole === "Approverrole") {
            isFinalApprover = true;
            break;
          }

          // Delegation check
          const delegateResults = await sp.web.lists.getByTitle("ARGDelegateList").items
            .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
            .expand("Author,DelegateName,ActingFor")
            .filter(`DelegateName/EMail eq '${approver.email}' and ActingFor/ID eq ${currentUser.Id} and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
            .top(1)(); // only need existence

          if (delegateResults.length > 0 && Items.CurrentUserRole === "Approverrole") {
            isFinalApprover = true;
            break;
          }
        }

        // Final check: must be final approver and have role
        //const isUserFinalApproverWithRole = isFinalApprover && Items.CurrentUserRole === "Approverrole";

        //console.log("finallevel", finallevel, currentlevel);

        if (editType === "approve") {
          if (Items.Status != "Approved" && isFinalApprover && finallevel == currentlevel && Items.CurrentUserRole == "Approverrole") {
            //   showcorrectionappicable = true;
            //   isdisablefinal = false;
            this.setState({ IsFinalapprover: isFinalApprover });
          }
          if (Items.Status != "Approved" && Items.CurrentUserRole == "Approverrole") {
            showcorrectionappicable = true;
            isdisablefinal = true;
            this.setState({ IsFinalapprover: isFinalApprover });
          }
        }

      }
      setTimeout(() => {
        this.setState({ Loading: false });
        setloading = false;
      }, 2000);
      // this.setState({ Loading: false });
      // setloading = false;
      //Get latest last rec the NC list   
      const latestItem = await sp.web.lists
        .getByTitle("NonConformityList").items.select("Id", "SerialNumber", "Created", "SubmitStatus")
        .orderBy("SerialNumber", false).top(1)();
      if (latestItem.length > 0) {
        var serialNo = latestItem[0].SerialNumber + 1;
        if (serialNo < 999)
          serialNo = ("0000" + serialNo).slice(-3);
        this.setState({ editserialNo: serialNo })
      }
      else {
        var serialNo: any = "001";
        this.setState({ editserialNo: serialNo })
      }

    } catch (e) {
      console.error(e);
    }


  }
  private async getDataRoles() {
    const sp = spfi().using(SPFx(this.props.context));
    const listItems = await sp.web.lists.getByTitle("ApproverRoleMaster").items();

    let dropdownItems: IDropdownOption[] = [];
    listItems.map(item => {
      dropdownItems.push({
        key: item.Id,
        text: item.Role
      })
    });
    this.setState({ optionsRole: dropdownItems });
  }

  private async getMainListName() {
    const sp = spfi().using(SPFx(this.props.context));
    const listItems = await sp.web.lists.getByTitle("ListNameMaster").items.filter("ListName eq 'NonConformityList'")();
    this.setState({ mainListId: listItems[0].Id })

  }
  private async getFormName() {
    const sp = spfi().using(SPFx(this.props.context));
    const listItems = await sp.web.lists.getByTitle("FormNameMaster").items.filter("FormName eq 'Non Conformity'")();
    this.setState({ formNameId: listItems[0].Id })

  }
  private async getRequestorRole() {
    const sp = spfi().using(SPFx(this.props.context));
    const listItems = await sp.web.lists.getByTitle("RequesterRoleMaster").items.filter("Role eq 'Initiator'")();
    this.setState({ reqRolId: listItems[0].Id })
  }
  public cancelRequest(redirectto: string) {
    const { context } = this.props;
    if (redirectto == "myapproval") {
      window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
    } else {
      window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
    }
    // Swal.fire({
    //   title: 'Do you want to cancel this request?',
    //   showCancelButton: true,
    //   confirmButtonText: 'Yes',
    //   cancelButtonText: 'No'
    // }).then(function (val) {
    //   if (val.isConfirmed) {
    //     Swal.fire({
    //       title: "Cancelled Successfully.",
    //       icon: "success"
    //     }).then(() => {
    //       if (redirectto == "myapproval") {
    //         window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
    //       } else {
    //         window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
    //       }

    //       // window.location.reload();
    //     });
    //   }
    // });
  }
  // public async getAuditreport() {
  //   const sp = spfi().using(SPFx(this.props.context));
  //   
  //   try {
  //     const memoItems = await getMemoNumberAuditReport(sp);
  //     let optionsNCNumber: any = [];
  //     let optionsObservationNumber: any = [];
  //     if (memoItems.length > 0) {
  //       const filteredItemsNC = memoItems[0].filter((item: any) => item.FailureofIntentNonconformity === "Yes");
  //       const filteredItemsObs = memoItems[0].filter((item: any) => item.Observations === "Yes");
  //       // if (filteredItemsNC.length > 0) {
  //       //   optionsNCNumber = filteredItemsNC.map((item: any) => ({
  //       //     value: item.ID,
  //       //     label: item.ReportCode,
  //       //     itemId: item.ID,
  //       //     reportCode: item.ReportCode,
  //       //     ncNo: item.NCNumber,
  //       //     department: item.DepartmentAuditedId
  //       //   }));
  //       // }
  //       // if (filteredItemsObs.length > 0) {
  //       //   optionsObservationNumber = filteredItemsObs.map((item: any) => ({
  //       //     value: item.ID,
  //       //     label: item.ReportCode,
  //       //     itemId: item.ID,
  //       //     reportCode: item.ReportCode,
  //       //     ncNo: item.NCNumber,
  //       //     department: item.DepartmentAuditedId
  //       //   }));
  //       // }
  //       const groupItemsByReportCode =async (items: any[]) => {
  //         const result: { [reportCode: string]: Set<string> } = {};
  //         items.forEach((item) => {
  //           const reportCode = item.ReportCode?.trim();
  //           const number = item.NCNumber?.toString().trim();

  //           // Only proceed if both reportCode and NCNumber are non-empty
  //           if (reportCode && number && number !== "") {
  //             if (!result[reportCode]) {
  //               result[reportCode] = new Set();
  //             }
  //             result[reportCode].add(number);
  //           }
  //         });
  //         return result;
  //       };

  //       const reportCodeToExpectedNCs =await groupItemsByReportCode(filteredItemsNC);
  //       const cleanedFilteredItemsObs = filteredItemsObs.filter((item: any) => {
  //         const nc = item.NCNumber?.toString().trim();
  //         return nc && nc !== "";
  //       });
  //       const reportCodeToExpectedObs =await groupItemsByReportCode(cleanedFilteredItemsObs);
  //       // Process NC report codes
  //       await Promise.all(Object.keys(reportCodeToExpectedNCs).map(async (reportCode) => {
  //         const ncData = await this.getNCdata(reportCode);
  //         const existingNumbers = new Set(ncData.map((item: any) => item.NCNumber));
  //         const expectedNumbers = reportCodeToExpectedNCs[reportCode];

  //         const allCreated = Array.from(expectedNumbers).every((num) => existingNumbers.has(num));

  //         if (!allCreated) {
  //           const exampleItem = filteredItemsNC.find((item: any) => item.ReportCode === reportCode);
  //           if (exampleItem) {
  //             optionsmemoNumbernewnc.push({
  //               value: exampleItem.ID,
  //               label: exampleItem.ReportCode,
  //               itemId: exampleItem.ID,
  //               reportCode: exampleItem.ReportCode,
  //               ncNo: exampleItem.NCNumber,
  //               department: exampleItem.DepartmentAuditedId
  //             });
  //           }
  //         }
  //       }));

  //       // Process Observation report codes
  //       await Promise.all(Object.keys(reportCodeToExpectedObs).map(async (reportCode) => {
  //         const obsData = await this.getNCdata(reportCode); // Assuming same list for Observations
  //         const existingNumbers = new Set(obsData.map((item: any) => item.NCNumber));
  //         const expectedNumbers = reportCodeToExpectedObs[reportCode];

  //         const allCreated = Array.from(expectedNumbers).every((num) => existingNumbers.has(num));

  //         if (!allCreated) {
  //           const exampleItem = filteredItemsObs.find((item: any) => item.ReportCode === reportCode);
  //           if (exampleItem) {
  //             optionsmemoNumbernewobs.push({
  //               value: exampleItem.ID,
  //               label: exampleItem.ReportCode,
  //               itemId: exampleItem.ID,
  //               reportCode: exampleItem.ReportCode,
  //               ncNo: exampleItem.NCNumber,
  //               department: exampleItem.DepartmentAuditedId
  //             });
  //           }
  //         }
  //       }));

  //       // Sort options
  //       optionsmemoNumbernewobs = await this.getUniqueBy(optionsmemoNumbernewobs, "reportCode");
  //       optionsmemoNumbernewnc = await this.getUniqueBy(optionsmemoNumbernewnc, "reportCode");
  //       optionsmemoNumbernewnc.sort((a, b) => a.label.localeCompare(b.label));
  //       optionsmemoNumbernewobs.sort((a, b) => a.label.localeCompare(b.label));
  //       const editoptsmemoAll1 = this.state.editncType === "NC" ? filteredItemsNC : filteredItemsObs;

  //       const optionsallnc = filteredItemsNC.map((item: any) => ({
  //         value: item.ID,
  //         label: item.ReportCode,
  //         itemId: item.ID,
  //         reportCode: item.ReportCode,
  //         ncNo: item.NCNumber,
  //         department: item.DepartmentAuditedId
  //       }));
  //       const optionsallobs = filteredItemsObs.map((item: any) => ({
  //         value: item.ID,
  //         label: item.ReportCode,
  //         itemId: item.ID,
  //         reportCode: item.ReportCode,
  //         ncNo: item.NCNumber,
  //         department: item.DepartmentAuditedId
  //       }));
  //       editoptsmemoAllNC = optionsallnc;
  //       editoptsmemoAllObs = optionsallobs;
  //       this.setState({
  //         editmemonumberOptions: this.state.editncType == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs,
  //         editmemonumberOptionsall: this.state.editncType == "NC" ? optionsNCNumber : optionsObservationNumber
  //       });
  //     }

  //     //this.state.ncType
  //     //let optionsmemoNumbernewnc: any[] = [];
  //     // optionsmemoNumbernewnc = await this.getUniqueBy(optionsNCNumber, "reportCode");
  //     // optionsmemoNumbernewnc = [...optionsmemoNumbernewnc].sort((a, b) =>
  //     //   a.label.localeCompare(b.label)
  //     // );

  //     // //let optionsmemoNumbernewobs: any[] = [];
  //     // optionsmemoNumbernewobs = await this.getUniqueBy(optionsObservationNumber, "reportCode");
  //     // optionsmemoNumbernewobs = [...optionsmemoNumbernewobs].sort((a, b) =>
  //     //   a.label.localeCompare(b.label)
  //     // );

  //   } catch (e) {
  //     console.error(e);
  //   }
  // }
  public async getAuditreport(): Promise<{
    optionsmemoNumbernewnc: any[],
    optionsmemoNumbernewobs: any[],
    editoptsmemoAllNC: any[],
    editoptsmemoAllObs: any[]
  }> {
    const sp = spfi().using(SPFx(this.props.context));

    // Reset global or class-level variables (if needed)
    let optionsmemoNumbernewncloc: any[] = [];
    let optionsmemoNumbernewobsloc: any[] = [];

    let optionsNCNumberloc: any[] = [];
    let optionsObservationNumberloc: any[] = [];

    try {
      const memoItems = await getMemoNumberAuditReport(sp);
      // console.log("Fetched memoItems:", memoItems);

      if (!Array.isArray(memoItems) || memoItems.length === 0) {
        console.warn("No memo items returned or not an array");
        return;
      }

      // Flatten if memoItems is nested: [ [ items ] ]
      const flatMemoItems = Array.isArray(memoItems[0]) ? memoItems[0] : memoItems;

      const filteredItemsNC = flatMemoItems.filter((item: any) => (item.FailureofIntentNonconformity === "Yes" || item.FailureofImplementation == "Yes" || item.FailureofEffectiveness == "Yes"));
      const filteredItemsObs = flatMemoItems.filter((item: any) => item.Observations === "Yes");

      const groupItemsByReportCode = async (items: any[]) => {
        const result: { [reportCode: string]: Set<string> } = {};
        items.forEach((item) => {
          const reportCode = item.ReportCode?.trim();
          const number = item.NCNumber?.toString().trim();

          if (reportCode && number && number !== "") {
            if (!result[reportCode]) {
              result[reportCode] = new Set();
            }
            result[reportCode].add(number);
          }
        });
        return result;
      };

      const reportCodeToExpectedNCs = await groupItemsByReportCode(filteredItemsNC);

      const cleanedFilteredItemsObs = filteredItemsObs.filter((item: any) => {
        const nc = item.NCNumber?.toString().trim();
        return nc && nc !== "";
      });

      const reportCodeToExpectedObs = await groupItemsByReportCode(cleanedFilteredItemsObs);

      // Process NC report codes
      await Promise.all(Object.keys(reportCodeToExpectedNCs).map(async (reportCode) => {
        try {
          const ncData = await this.getNCdata(reportCode, this.state.editncType);
          const existingNumbers = new Set(ncData.map((item: any) => item.NCNumber));
          const expectedNumbers = reportCodeToExpectedNCs[reportCode];

          const allCreated = Array.from(expectedNumbers).every((num) => existingNumbers.has(num));

          if (!allCreated) {
            const exampleItem = filteredItemsNC.find((item: any) => item.ReportCode === reportCode);
            if (exampleItem) {
              optionsmemoNumbernewncloc.push({
                value: exampleItem.ID,
                label: exampleItem.ReportCode,
                itemId: exampleItem.ID,
                reportCode: exampleItem.ReportCode,
                ncNo: exampleItem.NCNumber,
                department: exampleItem.DepartmentAuditedId
              });
            }
          }
        } catch (error) {
          console.error(`Error fetching NC data for reportCode ${reportCode}:`, error);
        }
      }));

      // Process Observation report codes
      await Promise.all(Object.keys(reportCodeToExpectedObs).map(async (reportCode) => {
        try {
          const obsData = await this.getNCdata(reportCode, this.state.editncType);
          const existingNumbers = new Set(obsData.map((item: any) => item.NCNumber));
          const expectedNumbers = reportCodeToExpectedObs[reportCode];

          const allCreated = Array.from(expectedNumbers).every((num) => existingNumbers.has(num));

          if (!allCreated) {
            const exampleItem = filteredItemsObs.find((item: any) => item.ReportCode === reportCode);
            if (exampleItem) {
              optionsmemoNumbernewobsloc.push({
                value: exampleItem.ID,
                label: exampleItem.ReportCode,
                itemId: exampleItem.ID,
                reportCode: exampleItem.ReportCode,
                ncNo: exampleItem.NCNumber,
                department: exampleItem.DepartmentAuditedId
              });
            }
          }
        } catch (error) {
          console.error(`Error fetching Obs data for reportCode ${reportCode}:`, error);
        }
      }));

      // Sort and deduplicate
      optionsmemoNumbernewnc = await this.getUniqueBy(optionsmemoNumbernewncloc, "reportCode");
      optionsmemoNumbernewobs = await this.getUniqueBy(optionsmemoNumbernewobsloc, "reportCode");

      optionsmemoNumbernewnc.sort((a, b) => a.label.localeCompare(b.label));
      optionsmemoNumbernewobs.sort((a, b) => a.label.localeCompare(b.label));

      const optionsallnc = flatMemoItems.map((item: any) => ({
        value: item.ID,
        label: item.ReportCode,
        itemId: item.ID,
        reportCode: item.ReportCode,
        ncNo: item.NCNumber,
        department: item.DepartmentAuditedId
      }));

      const optionsallobs = flatMemoItems.map((item: any) => ({
        value: item.ID,
        label: item.ReportCode,
        itemId: item.ID,
        reportCode: item.ReportCode,
        ncNo: item.NCNumber,
        department: item.DepartmentAuditedId
      }));

      editoptsmemoAllNC = optionsallnc;
      editoptsmemoAllObs = optionsallobs;

      this.setState({
        editmemonumberOptions: this.state.editncType === "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs, editoptionsmemoNC: optionsallnc,
        editoptionsmemoObs: optionsallobs
        //editmemonumberOptionsall: this.state.editncType === "NC" ? optionsNCNumber : optionsObservationNumber
      });

      return {
        optionsmemoNumbernewnc,
        optionsmemoNumbernewobs,
        editoptsmemoAllNC,
        editoptsmemoAllObs
      };
    } catch (e) {
      console.error(e);
      return {
        optionsmemoNumbernewnc: [],
        optionsmemoNumbernewobs: [],
        editoptsmemoAllNC: [],
        editoptsmemoAllObs: []
      };
    } finally {
      // Always stop the loader
      // this.setState({ Loading: false });
    }
  }
  private async getAuditReportNCNumbersAll() {

    const sp = spfi().using(SPFx(this.props.context));
    //let nctype = NCType == "NC" ? "NC Number" : "Observation Number";
    try {
      //const filter = `ReportCode eq '${reportCode}' and NCType eq '${nctype}'`;

      const items = await sp.web.lists
        .getByTitle("AuditReportNCNumber")
        .items
        //.filter(filter)
        .select("ID", "ReportCode", "NCType", "NCNumber", "Description", "AnnualAuditReportListId").top(5000)
        ();

      return items;
    } catch (error) {
      //console.error(`Error fetching AuditReportNCNumbers for ${reportCode} (${NCType}):`, error);
      return [];
    }
  }
  // public async getAuditreport() {
  //   const sp = spfi().using(SPFx(this.props.context));
  //   
  //   try {
  //     const memoItems = await getMemoNumberAuditReport(sp);
  //     let optionsmemoNumber: any = [];
  //     if (memoItems.length > 0) {
  //       optionsmemoNumber = memoItems[0].map((item: any) => ({
  //         key: item.ID,
  //         text: item.MemoNumber,
  //         itemId: item.ID,
  //         memoNumber: item.MemoNumber,
  //         ncNo: item.NCNumber
  //       }));
  //     }
  //     let optionsmemoNumbernew: any[] = [];
  //     optionsmemoNumbernew = await this.getUniqueBy(optionsmemoNumber, "memoNumber");
  //     this.setState({ editmemonumberOptions: optionsmemoNumbernew, editmemonumberOptionsall: optionsmemoNumber });
  //   } catch (e) {
  //     console.error(e);
  //   }
  // }
  public async getDepartment() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const deptItems = await sp.web.lists.getByTitle("ProcessDepartmentMasterList").items();
      const options = deptItems.map((item: {
        DepartmentCode: any; Title: string; Id: number, ADDepartmentName: string
      }) => ({
        value: item.Id,
        label: item.Title,
        adDepartmentName: item.ADDepartmentName,
        data: { departmentCode: item.DepartmentCode },
      }));
      this.setState({ editDepartmentOption: options });
      const userProfile = await sp.profiles.myProperties();
      const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
      let currentuserdepartment = UserDept;
      const selectedOption = options.find(user => user?.adDepartmentName === currentuserdepartment);
      // Find the selected department's departmentCode and set it
      const selectedDeptArray = options.filter(opt => opt.value === this.state.editDepartment);
      if (selectedDeptArray.length > 0) {
        this.setState({ editdepartmentCode: selectedDeptArray[0].data.departmentCode });
      }

      await this.getCategory();
      return options;
    } catch (e) {
      console.error(e);
    }
  }

  public async getCategory() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const categoryItems = await sp.web.lists.getByTitle("NCCategoryMasterList").items();
      const options = categoryItems.map((item: { Title: string; Id: number }) => ({
        key: item.Id,
        text: item.Title,
      }));
      this.setState({ editCategoryCheckOption: options });
      await this.getSubCategory();
    } catch (e) {
      console.error(e);
    }
  }

  public async getSubCategory() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const subCategoryItems = await sp.web.lists.getByTitle("NCSubCategoryMasterList").items();
      const options = subCategoryItems.map((item: { Title: string; Id: number }) => ({
        key: item.Id,
        text: item.Title,
      }));
      this.setState({ editSubCategoryCheckOption: options });
      await this.getLocation();
    } catch (e) {
      console.error(e);
    }
  }

  public async getLocation() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const locationItems = await sp.web.lists.getByTitle("NCLocationMasterList").items();
      const options = locationItems.map((item: { Title: string; Id: number }) => ({
        key: item.Id,
        text: item.Title,
      }));
      this.setState({ editLocationCheckOption: options });
    } catch (e) {
      console.error(e);
    }
  }

  //< ------- Start Submit Validation -------->
  public handleSubmit = (formsubmode: string) => {
    Reworkclicked = false;
    //resubmitclicked = false;
    if (this.state.editCurrentUserRole == "FirstInitiator" && this.state.editStatus == "Rework" && this.state.editFirstInitiatorSubmitStatus == "No") {
      resubmitclicked = true
    }

    if (resubmitclicked ? (this.validateReworkRemark() && this.validateFormSubmit()) : this.validateFormSubmit()) {
      this._updateSubmitData(formsubmode);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public validateFormSubmit = (): boolean => {
    let editErrors: { [key: string]: string } = {};
    if (this.state.editSubmitStatus == "No") {
      if (!this.state.editncType) editErrors.editncType = "editncType is required";
      //if (this.state.departmentselected.length == 0) editErrors.editDepartment = "Department is required";
      if (this.state.ApprovedAuditSelected.length == 0) editErrors.editApprovedAuditReport = "Report code is required";
      if (this.state.NCNumberselected.length == 0) editErrors.editNCNumber = "NCR number is required";
      if (!this.state.editCriteria) editErrors.editCriteria = "Criteria is required";
      //if (!this.state.editCloseOutStatus) editErrors.editCloseOutStatus = "Close Out Status is required";
      if (this.state.editCategoryValueIsCheck.length == 0) {
        document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.add(styles.errCh);
        });
        editErrors.editcategory = "Category is required";
        Swal.fire({ title: "Please select at least one category!" });
      }
      else {
        document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.remove(styles.errCh);
        });
      }
      if (this.state.editSubCategoryValueIsCheck.length == 0) {
        document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.add(styles.errCh);
        });
        editErrors.editsubcategory = "Sub-Category is required";
        Swal.fire({ title: "Please select at least one Sub category!" });
      }
      else {
        document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.remove(styles.errCh);
        });
      }
      if (this.state.editLocationValueIsCheck.length == 0) {
        document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.add(styles.errCh);
        });
        editErrors.editlocation = "location is required";
        Swal.fire({ title: "Please select at least one location!" });
      }
      else {
        document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.remove(styles.errCh);
        });
      }
      if (!this.state.editAssignTo) {
        editErrors.editAssignTo = "AssignTo is required";
        document.querySelectorAll("#AssigntoPeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.add(styles.errCh);
        });
      } else {
        document.querySelectorAll("#AssigntoPeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.remove(styles.errCh);
        });
        document.querySelectorAll("#AssigntoPeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.add(styles.peoplepickerstyleAuditte);
        });
      }
      //if (!this.state.editAssignTo) editErrors.editAssignTo = "AssignTo is required";
      if (!this.state.editDueDate) editErrors.editDueDate = "dueDate is required";
      // if (!this.state.fileCount && this.state.exFiles.length == 0) {
      //   editErrors.Attchments = "Attachments are required";
      // }
      // if (!this.state.fileCount && this.state.exFiles.length == 0) {
      //   editErrors.Attchments = "Attachments are required";
      //   //isValid = false;
      //   document.querySelectorAll("#newfile").forEach((el) => {
      //     el.classList.remove(styles.errCh);
      //   });

      // } else {
      //   document.querySelectorAll("#newfile").forEach((el) => {
      //     el.classList.remove(styles.errCh);
      //   });
      // }
      if (!this.state.editProblemDescription) editErrors.editProblemDescription = "Problem Description is required";
    }
    else if ((this.state.editCurrentUserRole == "FirstAssignedTo" && this.state.editDelegateToId == null) || (this.state.editCurrentUserRole == "DelegateTo" && this.state.editDelegateToId != null)) {
      if (!this.state.editPersonAssigned) editErrors.editPersonAssigned = "PersonAssigned is required";
      if (!this.state.editDate) editErrors.editDate = "Date is required";
      if (!this.state.editDeadlineCompletion) editErrors.editDeadlineCompletion = "Deadline for completion is required";
      if (!this.state.editCorrection) editErrors.editCorrection = "Correction is required";
      if (!this.state.editRootCause) editErrors.editRootCause = "Root Cause is required";
      if (!this.state.editCorrectiveAction) editErrors.editCorrectiveAction = "Corrective Action is required";

      if (!this.state.editAnalyzedBy) {
        editErrors.editAnalyzedBy = "Analyzed By is required";
        document.querySelectorAll("#analyzedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.add(styles.errCh);
        });
      } else {
        document.querySelectorAll("#analyzedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.remove(styles.errCh);
        });
        document.querySelectorAll("#delegatetopeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.remove(styles.errCh);
        });

      }
      // if (!this.state.editReviewedBy) {
      //   editErrors.editReviewedBy = "Reviewed By is required";
      //   document.querySelectorAll("#reviewedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
      //     el.classList.add(styles.errCh);
      //   });
      // } else {
      //   document.querySelectorAll("#reviewedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
      //     el.classList.remove(styles.errCh);
      //   });
      // }
      //if (!this.state.editCorrectiveActionImplementedOn) editErrors.editCorrectiveActionImplementedOn = "Corrective Action Implemented on is required";
    }
    else {
      console.log('No field require for any conditions');
    }
    this.setState({ editErrors });
    return Object.keys(editErrors).length === 0;
  };
  //< ------- End Submit Validation -------->

  //< ------- Start Forward Validation -------->
  public handleForward = (formsubmode: string) => {
    console.log("aghghghghghghgh", this.state.approvers)
    if (this.validateFormForward()) {
      this.forwardRequest(formsubmode);

    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public validateFormForward = (): boolean => {
    const editErrors: { [key: string]: string } = {};
    let hasError = false;

    this.state.approvers.forEach((row, index) => {
      const isNameEmpty = !row.Name || row.Name.length === 0;
      const isRoleEmpty = !row.Role || row.Role.toString().trim() === '';
      const isResponsibilityEmpty = !row.Responsibility || row.Responsibility.trim() === '' || row.Responsibility == "Select";

      const peoplePickerElements = document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`);

      if (isNameEmpty) {
        editErrors[`approvers[${index}].Name`] = 'Approver is required';
        hasError = true;

        peoplePickerElements.forEach((el) => {
          el.classList.add(styles.errCh);
          //el.classList.remove(styles.peoplepickerstyleAuditteerrch);
        });
      } else {
        peoplePickerElements.forEach((el) => {
          el.classList.remove(styles.errCh);
          //el.classList.add(styles.peoplepickerstyleAuditteerrch);
        });
      }

      if (isRoleEmpty) {
        editErrors[`approvers[${index}].Role`] = 'Role is required';
        hasError = true;
      }
      if (isResponsibilityEmpty) {
        editErrors[`approvers[${index}].Responsibility`] = 'Responsibility is required';
        hasError = true;
      }
    });
    //let editErrors: { [key: string]: string } = {};
    if (!this.state.editCorrectiveActionImplementedOn) {
      editErrors.editCorrectiveActionImplementedOn = "Corrective Action ImplementedOn is required";
      hasError = true;
    }
    this.setState({ editErrors });
    //return Object.keys(editErrors).length === 0;
    // this.setState({ editErrors });

    return !hasError;
  };

  // public validateFormForward = (): boolean => {
  //   const editErrors: { [key: string]: string } = {};
  //   let hasError = false;

  //   this.state.approvers.forEach((row, index) => {
  //     const isNameEmpty = !row.Name || row.Name.length === 0;
  //     const isRoleEmpty = !row.Role || row.Role.toString().trim() === '';

  //     if (isNameEmpty) {
  //       editErrors[`approvers[${index}].Name`] = 'Approver is required';
  //       hasError = true;

  //       document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`).forEach((el) => {
  //         el.classList.add(styles.errCh);
  //       });

  //     }
  //     else {
  //       document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`).forEach((el) => {
  //         el.classList.remove(styles.errCh);
  //       });
  //       document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`).forEach((el) => {
  //         el.classList.add(styles.peoplepickerstyleAuditteerrch);
  //       });

  //     }

  //     if (isRoleEmpty) {
  //       editErrors[`approvers[${index}].Role`] = 'Role is required';
  //       hasError = true;
  //     }
  //   });

  //   this.setState({ editErrors });

  //   return !hasError;
  // };


  // public validateFormForward = (): boolean => {
  //   let editErrors: { [key: string]: string } = {};
  //   const isValidApp = this.state.approvers.every(row => row.Name.length == 0);
  //   const isValidrole = this.state.approvers.every(row => row.Role !== 0);
  //   if (this.state.approvers.length === 0 || !isValidApp) {
  //     editErrors.approvers = "Approvers required";
  //     document.querySelectorAll("#approverpeoplepicker .ms-BasePicker-text").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   } else {
  //     document.querySelectorAll("#approverpeoplepicker .ms-BasePicker-text").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }
  //   if (this.state.approvers.length === 1) {
  //     if (!this.state.approvers[0].Role || !isValidrole) {
  //       editErrors.approvers = "Approver role is required";
  //     }
  //   }

  //   this.setState({ editErrors });
  //   return Object.keys(editErrors).length === 0;
  // };
  //< ------- End Forward Validation -------->

  //< ------- Start Draft Validation -------->
  public handleDraft = (formsubmode: string) => {
    resubmitclicked = false;

    if (this.validateFormDraft() || (this.state.editCurrentUserRole === "FirstAssignedTo" || this.state.editCurrentUserRole === "DelegateTo")) {
      this._updateSubmitData(formsubmode);

    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public validateFormDraft = (): boolean => {
    let editErrors: { [key: string]: string } = {};

    // if (this.state.departmentselected.length == 0) {
    //   editErrors.editDepartment = "Department is required";
    //   this.setState({ editErrors });
    //   Swal.fire('Please select a department.');
    //   return false;
    // }
    if (this.state.ApprovedAuditSelected.length == 0) {
      editErrors.editApprovedAuditReport = "Memo number is required";
      this.setState({ editErrors });
      Swal.fire('Please select a Approved memo number.');
      return false;
    }
    if (this.state.NCNumberselected.length == 0) {
      editErrors.editApprovedAuditReport = "Memo number is required";
      this.setState({ editErrors });
      Swal.fire('Please select a NC/Observation number.');
      return false;
    }
    this.setState({ editErrors });
    return true;
  };
  public _approveRequest = (formsubmode: string) => {
    Approveclicked = true;
    Reworkclicked = false;
    if (this.state.editncType == "Observation") {
      this.approveRequest(formsubmode);
    } else {
      //if (this.validateFormRemark()) {
      this.approveRequest(formsubmode);

      // }
      // else {
      //   Swal.fire('Please fill the mandatory fields.');
      // }
    }

  };
  public _rejectRequest = (formsubmode: string) => {
    Rejectclicked = true;
    Reworkclicked = false;
    if (this.validateFormRemark()) {
      this.rejectRequest(formsubmode);

    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public _reworkRequest = (formsubmode: string) => {
    Reworkclicked = true;
    document.querySelectorAll("#delegatetopeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.remove(styles.errCh);
    });
    document.querySelectorAll("#reviewedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.remove(styles.errCh);
    });
    document.querySelectorAll("#analyzedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.remove(styles.errCh);
    });

    if (this.validateFormRemark()) {
      this.reworkRequest(formsubmode);

    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public validateFormRemark = (): boolean => {
    let editErrors: { [key: string]: string } = {};
    if (!this.state.remarks) editErrors.remarks = "Remarks is required";
    this.setState({ editErrors });
    return Object.keys(editErrors).length === 0;
  };
  public validateReworkRemark = (): boolean => {
    let editErrors: { [key: string]: string } = {};
    if (!this.state.reworkremarks && !this.state.remarks) editErrors.remarks = "Remarks is required";
    this.setState({ editErrors });
    return Object.keys(editErrors).length === 0;
  };

  //< ------- End Draft Validation -------->

  private onRoleChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Role = item.key;
    this.setState({ approvers: this.state.approvers });
  }

  private onTypeChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Type = item.key;
    this.setState({ approvers: this.state.approvers });
  }
  private handleChangeResp(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Responsibility = item.key;
    this.state.approvers[i].IsSignatureRequired = item.key == "Signer" ? true : false;
    this.setState({ approvers: this.state.approvers });
  }
  // private handleChangeResp(item: any, i: number) {
  //   this.state.approvers[i].Responsibility = item.key;
  //   this.setState({ approvers: this.state.approvers });
  // }
  // private _getPeoplePickerItemsApp(items: any[], i: number) {

  //   var arr: any[];
  //   arr = [];
  //   items.forEach(function (it) {
  //     arr.push(it.id);
  //   })
  //   this.state.approvers.filter(function (it) {
  //     if (it.Index == (editforwardrecord ? i + 1 : i)) {
  //       it.Name = arr
  //     }
  //   });
  //   this.setState({ approvers: this.state.approvers });
  // }
  private _getPeoplePickerItemsApp(items: any[], i: number) {
    const selectedIds = items.map(it => it.id);
    const selectedNames = items.map(it => it.text); // For display in tooltip or validation

    const updatedApprovers = [...this.state.approvers];

    if (updatedApprovers[i]) {
      updatedApprovers[i] = {
        ...updatedApprovers[i],
        Name: selectedIds,
        appEx: selectedNames
      };

      this.setState({ approvers: updatedApprovers });
    }
  }

  private addApprover() {

    var itm = this.state.indApp + 1;
    this.setState({ indApp: itm });
    this.state.approvers.push({
      Role: "", Level: "", Name: "", Index: itm, itemId: "", Responsibility: "Signer", IsSignatureRequired: true,
      responsibilityerror: false
    });
    this.setState({ approvers: this.state.approvers });
  }

  private deleteItemApp(i: number) {
    //i = editforwardrecord ? i + 1 : i;
    var items = this.state.approvers.filter(function (it, val) {
      return val != i
    });

    this.setState({ approvers: items });

    if (this.state.approvers[i].itemId != "") {
      var delIDs = [];
      delIDs.push(this.state.approvers[i].itemId)
      this.setState({ apprDelId: delIDs });
    }
    var items = this.state.approvers.filter(function (it, val) {
      return val != i
    });
    this.setState({ approvers: items });

  }

  public async updateData(_editsubmitStatus: string, firstInitiatorSubmitStatus: string, firstAssignedToSubmitStatus: string, delegateToSubmitStatus: string, analyzedBySubmitStatus: string, reviewedBySubmitStatus: string, lastAssignedToSubmitStatus: string, lastInitiatorSubmitStatus: string, currentUserRole: string, reworkById: any, serialNumber: number, documentCode: string, ncrnumber: string, FirstAssignedToSavedStatus: string, FirstDelegateToSavedStatus: string) {

    const sp = spfi().using(SPFx(this.props.context));
    let observationStatus: string = "";
    let observationStatus1: string = "";
    let closeoutstatuss: string = "";
    let test1 = this.state.correctionApplicable ? "Yes" : "No";
    let test2 = this.state.notEffective ? "Yes" : "No";
    let test3 = this.state.effectiveClosed ? "Yes" : "No";
    let test4 = this.state.isIMSUpdated == "Yes" ? "Yes" : "No";
    let test5 = this.state.riskandopportunitiesUpdated == "Yes" ? "Yes" : "No";
    let ncStatus = (_editsubmitStatus == "Rework" || (this.state.editStatus == "Rework" && _editsubmitStatus == "draft")) && this.state.editCurrentUserRole != "FirstInitiator" || (this.state.editStatus == "Rework" && _editsubmitStatus == "draft" && this.state.editCurrentUserRole == "FirstInitiator")
      //(_editsubmitStatus == "submit" && this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator")
      ? "Rework" : "Pending";

    if (this.state.editncType == "Observation" && this.state.editDelegateToId != null && this.state.editCurrentUserRole == "LastInitiator" && _editsubmitStatus != "Rework") {
      observationStatus1 = "Approved";
      closeoutstatuss = "Completed"
    } else
      if (this.state.editncType == "Observation" && this.state.editDelegateToId == null && this.state.editCurrentUserRole == "LastInitiator" && _editsubmitStatus != "Rework") {
        observationStatus1 = "Approved";
        closeoutstatuss = "Completed"
      } else {
        observationStatus1 = "Pending";
        closeoutstatuss = this.state.editCloseOutStatus
      };
    observationStatus = _editsubmitStatus == "Rework" || (this.state.editStatus == "Rework" && _editsubmitStatus == "draft") ? "Rework" : observationStatus1;
    // let submitstatus: string = "";
    // submitstatus = ((_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo")
    //   || (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "DelegateTo")
    //   || (_editsubmitStatus == "draft")) ?
    //   "No" : "Yes";
    //_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo"
    await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId).update({
      NCNumber: this.state.editNCNumber,
      NCNumberID: this.state.editNCNumberID,
      ApprovedAuditReportId: this.state.editApprovedAuditReport || null,
      ApprovedAuditReportMemoNumber: this.state.editMemoNumber,
      DepartmentId: this.state.editDepartment || null,
      FromDepartmentId: this.state.editfromdepartment,
      Criteria: this.state.editCriteria,
      CloseOutStatus: this.state.editncType == "Observation" ? closeoutstatuss : this.state.editCloseOutStatus,
      //NCRNo: this.state.editNCRNo,
      ReferenceNumber: this.state.editReferenceNumber,
      RevisionNumber: this.state.editRevisionNo !== "" ? Number(this.state.editRevisionNo) : null,
      IssueNumber: this.state.editIssueNo !== "" ? Number(this.state.editIssueNo) : null,
      IssueDate: this.state.editIssueDate,
      RevisionDate: this.state.editRevisionDate,
      CategoryId: this.state.editCategoryValueIsCheck,
      SubCategoryId: this.state.editSubCategoryValueIsCheck,
      LocationId: this.state.editLocationValueIsCheck,
      // AssignedToId: _editsubmitStatus == "submit" && this.state.editCurrentUserRole == null ? this.props.currentUserID : this.state.editAssignToId || null,
      AssignedToId: this.state.editAssignToId || null,
      DueDate: this.state.editDueDate,
      ProblemDescription: this.state.editProblemDescription,
      PersonAssignedId: this.state.editPersonAssignedId || null,
      Date: this.state.editDate,
      Deadlineforcompletion: this.state.editDeadlineCompletion,
      Correctionproblem: this.state.editCorrection,
      RootCause: this.state.editRootCause,
      CorrectiveAction: this.state.editCorrectiveAction,
      DelegateToId: this.state.editDelegateToId || null,
      AnalyzedById: this.state.editAnalyzedById || null,
      ReviewedById: this.state.editReviewedById || null,
      CorrectiveActionImplementedOn: this.state.editCorrectiveActionImplementedOn,
      SubmiitedDate: new Date(),
      AuditeeSubmittedDate: this.state.editCurrentUserRole == "FirstAssignedTo" || this.state.editCurrentUserRole == "DelegateTo" ? new Date() : null,
      // SubmitStatus: _editsubmitStatus == "draft" || this.state.editStatus == "Rework" || (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo") ? "No" : "Yes",
      SubmitStatus: _editsubmitStatus == "draft" && (currentUserRole != "FirstAssignedTo" && currentUserRole != "DelegateTo") ? "No" : "Yes",
      SubmiitedById: this.props.currentUserID || null,
      CurrentUserRole: currentUserRole,
      FirstInitiatorSubmitStatus: firstInitiatorSubmitStatus,
      FirstAssignedToSubmitStatus: firstAssignedToSubmitStatus,
      DelegateToSubmitStatus: delegateToSubmitStatus,
      FirstAssignedToSavedStatus: FirstAssignedToSavedStatus,
      FirstDelegateToSavedStatus: FirstDelegateToSavedStatus,
      AnalyzedBySubmitStatus: analyzedBySubmitStatus,
      ReviewedBySubmitStatus: reviewedBySubmitStatus,
      LastAssignedToSubmitStatus: lastAssignedToSubmitStatus,
      LastInitiatorSubmitStatus: lastInitiatorSubmitStatus,
      Status: this.state.editncType == "Observation" ? observationStatus : ncStatus,
      //IsRework: _editsubmitStatus == "Rework" || _editsubmitStatus == "Reject" ? "Yes" : "No",
      IsRework: _editsubmitStatus == "Rework" || this.state.editStatus == "Rework" ? "Yes" : "No",
      ReworkById: reworkById,
      SerialNumber: serialNumber,
      NCRNo: ncrnumber,
      DocumentCode: this.state.editDocumentCode,
      IMSUpdated: this.state.editCurrentUserRole == "FirstAssignedTo" || this.state.editCurrentUserRole == "DelegateTo" ? test4 : this.state.isIMSUpdated,
      RiskOpportunitiesUpdated: this.state.editCurrentUserRole == "FirstAssignedTo" || this.state.editCurrentUserRole == "DelegateTo" ? test5 : this.state.riskandopportunitiesUpdated,
      LocationOthers: this.state.LocationOthers,
      SubCategoryOthers: this.state.SubCategoryOthers,
      CategoryOthers: this.state.CategoryOthers,
      ReworkRemarks: _editsubmitStatus == "draft" && (currentUserRole != "FirstAssignedTo" && currentUserRole != "DelegateTo") ? this.state.remarks : "",
      Correctionapplicable: this.state.editCurrentUserRole == "LastInitiator" || this.state.editCurrentUserRole == "Approverrole" ? test1 : "",
      NotEffective: this.state.editCurrentUserRole == "LastInitiator" || this.state.editCurrentUserRole == "Approverrole" ? test2 : "",
      EffectiveandProblemClosed: this.state.editCurrentUserRole == "LastInitiator" || this.state.editCurrentUserRole == "Approverrole" ? test3 : ""
    });
  }

  public getNewFileName = async (originalFileName: string): Promise<string> => {
    const sp = spfi().using(SPFx(this.props.context));
    const currentUser = await sp.web.currentUser();
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
    //const fileExtension = originalFileName.split('.').pop();
    const fileNameWithoutExtension = originalFileName.split('.').slice(0, -1).join('.');
    return `${userId}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
  };
  //Update Function
  private _updateSubmitData = async (_editsubmitStatus: string) => {

    let mText = "";
    let cText = "";
    //Start Flow condition
    let currentUserRole = "";
    let firstInitiatorSubmitStatus = "";
    let FirstAssignedToSavedStatus = "";
    let FirstDelegateToSavedStatus = "";
    let firstAssignedToSubmitStatus = "";
    let delegateToSubmitStatus = "";
    let analyzedBySubmitStatus = "";
    let reviewedBySubmitStatus = "";
    let lastAssignedToSubmitStatus = "";
    let lastInitiatorSubmitStatus = "";
    let reworkById: any;
    let serialNumber: number;
    let ncrnumber = "";
    let documentCode = "";
    const IsactionTaken = Number(Approvallistitemid) > 0 ?
      await this.CheckIfAlreadyactionTaken(Number(Approvallistitemid), CONTENTTYPE_NonComformity) : true;
    if (_editsubmitStatus == "draft" && this.state.editCurrentUserRole == null) {
      firstInitiatorSubmitStatus = "No"
      firstAssignedToSubmitStatus = "No";
      FirstAssignedToSavedStatus = "No";
      FirstDelegateToSavedStatus = "No";
      delegateToSubmitStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "";
      reworkById = null;
      serialNumber = 0;
      ncrnumber = "";
      documentCode = "";
    }
    else if (_editsubmitStatus == "submit" && this.state.editCurrentUserRole == null) {
      firstInitiatorSubmitStatus = "Yes";
      firstAssignedToSubmitStatus = "No";
      FirstAssignedToSavedStatus = "No";
      FirstDelegateToSavedStatus = "No";
      delegateToSubmitStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "FirstAssignedTo";
      reworkById = null;
      serialNumber = this.state.editserialNo;
      ncrnumber = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
      //documentCode = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
    }
    else if (_editsubmitStatus == "submit" && this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator") {
      firstInitiatorSubmitStatus = "Yes";
      firstAssignedToSubmitStatus = "No";
      FirstAssignedToSavedStatus = "No";
      FirstDelegateToSavedStatus = "No";
      delegateToSubmitStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "FirstAssignedTo";
      reworkById = null;
      serialNumber = this.state.editserialNo;
      ncrnumber = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
      //documentCode = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
    }
    else if (this.state.editStatus == "Rework" && _editsubmitStatus == "draft" && this.state.editCurrentUserRole == "FirstInitiator") {
      firstInitiatorSubmitStatus = "No"
      firstAssignedToSubmitStatus = "No";
      FirstAssignedToSavedStatus = "No";
      FirstDelegateToSavedStatus = "No";
      delegateToSubmitStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "FirstInitiator"; //LastInitiator
      //reworkById = this.state.editAssignToId || null;
      //serialNumber = this.state.notUpdateSerialNo;
      //ncrnumber = this.state.notUpdateDepartmentCode;
    }
    else if (this.state.editStatus != "Rework" && this.state.editCurrentUserRole == "FirstInitiator") {
      firstInitiatorSubmitStatus = "Yes";
      firstAssignedToSubmitStatus = "No";
      FirstAssignedToSavedStatus = "No";
      FirstDelegateToSavedStatus = "No";
      delegateToSubmitStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "FirstAssignedTo";
      reworkById = null;
      serialNumber = this.state.editserialNo;
      ncrnumber = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
      //documentCode = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
    }
    else {
      console.log("No status");
    }

    //<-------- Start Case1 --------->
    if (this.state.editDelegateToId == null) {
      if (_editsubmitStatus == "submit" && this.state.editCurrentUserRole == "FirstAssignedTo") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        FirstAssignedToSavedStatus = "No";
        FirstDelegateToSavedStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "AnalyzedBy";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
        //documentCode = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case1 --------->
    //<-------- Start Case2 --------->
    else {
      if (_editsubmitStatus == "submit" && this.state.editCurrentUserRole == "FirstAssignedTo") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        FirstAssignedToSavedStatus = "No";
        FirstDelegateToSavedStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "DelegateTo";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "submit" && this.state.editCurrentUserRole == "DelegateTo") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        FirstAssignedToSavedStatus = "No";
        FirstDelegateToSavedStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "AnalyzedBy";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
    }
    if (_editsubmitStatus == "draft" && this.state.editCurrentUserRole == "FirstAssignedTo") {
      firstInitiatorSubmitStatus = "Yes"
      firstAssignedToSubmitStatus = "No";
      delegateToSubmitStatus = "No";
      FirstAssignedToSavedStatus = "Yes";
      FirstDelegateToSavedStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "FirstAssignedTo";
      reworkById = null;
      // serialNumber = this.state.notUpdateSerialNo;
      //ncrnumber = this.state.notUpdateDepartmentCode;
      //documentCode = this.state.notUpdateDepartmentCode;
    }
    if (_editsubmitStatus == "draft" && this.state.editCurrentUserRole == "DelegateTo") {
      firstInitiatorSubmitStatus = "Yes"
      firstAssignedToSubmitStatus = "Yes";
      FirstAssignedToSavedStatus = "No";
      FirstDelegateToSavedStatus = "Yes";
      delegateToSubmitStatus = "No";
      analyzedBySubmitStatus = "No";
      reviewedBySubmitStatus = "No";
      lastAssignedToSubmitStatus = "No";
      lastInitiatorSubmitStatus = "No";
      currentUserRole = "DelegateTo";
      reworkById = null;
      // serialNumber = this.state.notUpdateSerialNo;
      //ncrnumber = this.state.notUpdateDepartmentCode;
      //documentCode = this.state.notUpdateDepartmentCode;
    }
    //<-------- End Case2 --------->

    if (_editsubmitStatus == "submit") {
      mText = "submit";
      cText = "Submitted";
    }
    else {
      mText = "save";
      cText = "Saved";
    }
    const sp = spfi().using(SPFx(this.props.context));
    const itemid = this.state.mainItemId;
    const { editProblemDescription } = this.state;
    const { currentUserID, approvalItemId, context } = this.props;
    const { updateData } = this;
    if (!IsactionTaken) {
      Swal.fire("Action has already been taken on this record.");
      return;
    }
    Swal.fire({
      title: "Do you want to " + mText + " this request?",
      showCancelButton: true,
      confirmButtonText: "Yes",
      cancelButtonText: "No"
    }).then(async (result) => {
      if (result.isConfirmed) {

        setloading = true;
        console.log("ApprovallistitemidApprovallistitemid", Approvallistitemid, approvalItemId, _editsubmitStatus)
        this.setState({ Loading: true });
        await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber, FirstAssignedToSavedStatus, FirstDelegateToSavedStatus);
        if (_editsubmitStatus == "submit") {
          if (currentUserRole == "AnalyzedBy" || currentUserRole == "DelegateTo") {
            // alert("Approval Item ID" + approvalItemId)
            let uploadedIds: number[] = [];

            // 1. Delete removed files
            if (this.state.fileDeleteIdauditee.length > 0) {
              this.state.fileDeleteIdauditee.forEach(function (ids) {
                sp.web.lists.getByTitle("AuditeeAttachmentDocs").items.getById(ids).delete();
              });
            }

            // 2. Upload new files
            if (this.state.copyFilauditee.length > 0) {
              const uploadPromises = this.state.copyFilauditee.map(async (file) => {


                const currentUser = await sp.web.currentUser();
                const userId = currentUser.Id; // Or however you get the current user ID
                const date = new Date();
                const fileExtension = file.name.split('.').pop();

                const components = [
                  date.getFullYear(),
                  (date.getMonth() + 1).toString().padStart(2, '0'),
                  date.getDate().toString().padStart(2, '0'),
                  date.getHours().toString().padStart(2, '0'),
                  date.getMinutes().toString().padStart(2, '0'),
                  date.getSeconds().toString().padStart(2, '0'),
                  date.getMilliseconds().toString().padStart(3, '0')
                ];
                //var fileNamePath = encodeURI(file.name);
                let newfileNameNew: any = await this.cleanFileNameSave(file.name);
                const fileNameWithoutExtension = newfileNameNew.split('.').slice(0, -1).join('.');
                var fileNamePath = `${userId}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
                //return `${userId}_${components.join('')}_${file.name}`;
                //const fileNamePath = encodeURI(file.name);
                //const fileNamePath = await this.getNewFileName(file.name);
                return sp.web.getFolderByServerRelativePath("AuditeeAttachmentDocs").files
                  .addUsingPath(fileNamePath, file, { Overwrite: true })
                  .then((response) => {
                    return response.file.getItem().then((fileItem: any) => {
                      uploadedIds.push(fileItem.Id); // Store ID of uploaded file
                      return fileItem.update({
                        NonConformityId: itemid // Optional: link document to current item
                      });
                    });
                  });
              });

              // 3. Once all uploads are done, update multi-lookup
              Promise.all(uploadPromises).then(() => {
                if (uploadedIds.length > 0) {
                  const lookupValues = uploadedIds.map((id) => ({ Id: id }));

                  sp.web.lists.getByTitle("NonConformityList").items.getById(itemid).update({
                    AuditeeAttachmentId: { results: lookupValues } // This field must be a multi-lookup pointing to AuditeeAttachmentDocs
                  });
                }
              });
            }
            if (approvalItemId == null || approvalItemId == undefined || approvalItemId == "") {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: (this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator") ||
                  (this.state.editStatus == "Pending" && this.state.editCurrentUserRole == "FirstAssignedTo" && this.state.reworkremarks !== "") ? this.state.reworkremarks : this.state.remarks,
              });
            } else {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: (this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator") ||
                  (this.state.editStatus == "Pending" && this.state.editCurrentUserRole == "FirstAssignedTo" && this.state.reworkremarks !== "") ? this.state.reworkremarks : this.state.remarks,
              });
            }

          }
          else {
            if (approvalItemId == null || approvalItemId == undefined || approvalItemId == "") {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: (this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator") ||
                  (this.state.editStatus == "Pending" && this.state.editCurrentUserRole == "FirstAssignedTo" && this.state.reworkremarks !== "") ? this.state.reworkremarks : this.state.remarks,
              });
            } else {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: (this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator") ||
                  (this.state.editStatus == "Pending" && this.state.editCurrentUserRole == "FirstAssignedTo" && this.state.reworkremarks !== "") ? this.state.reworkremarks : this.state.remarks,
              });
            }
            if (this.state.fileDeleteId.length > 0) {
              this.state.fileDeleteId.forEach(function (ids) {
                sp.web.lists.getByTitle("NonConformityDocs").items.getById(ids).delete();
              });
            }
            if (this.state.copyFil.length > 0) {
              let uploadedIdsatt: number[] = [];
              const uploadPromises = this.state.copyFil.map(async (file) => {

                //const sp = spfi().using(SPFx(this.props.context));
                const currentUser = await sp.web.currentUser();
                const userId = currentUser.Id; // Or however you get the current user ID
                const date = new Date();
                const fileExtension = file.name.split('.').pop();

                const components = [
                  date.getFullYear(),
                  (date.getMonth() + 1).toString().padStart(2, '0'),
                  date.getDate().toString().padStart(2, '0'),
                  date.getHours().toString().padStart(2, '0'),
                  date.getMinutes().toString().padStart(2, '0'),
                  date.getSeconds().toString().padStart(2, '0'),
                  date.getMilliseconds().toString().padStart(3, '0')
                ];

                //var fileNamePath = encodeURI(file.name);
                let newfileNameNew: any = await this.cleanFileNameSave(file.name);
                const fileNameWithoutExtension = newfileNameNew.split('.').slice(0, -1).join('.');
                var fileNamePath = `${userId}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
                //var fileNamePath = `${userId}_${components.join('')}_${newfileNameNew}`;
                //return `${userId}_${components.join('')}_${file.name}`;
                //var fileNamePath = encodeURI(file.name);
                // var fileNamePath = await this.getNewFileName(file.name);
                sp.web.getFolderByServerRelativePath("NonConformityDocs").files
                  .addUsingPath(fileNamePath, file, { Overwrite: true })
                  .then(async (response) => {
                    return response.file.getItem().then((fileItem: any) => {
                      uploadedIdsatt.push(fileItem.Id); // Store ID of uploaded file
                      return fileItem.update({
                        NonConformityId: itemid // Optional: link document to current item
                      });
                    });

                  });
              })
              Promise.all(uploadPromises).then(() => {
                if (uploadedIdsatt.length > 0) {
                  const lookupValues = uploadedIdsatt.map((id) => ({ Id: id }));

                  // sp.web.lists.getByTitle("NonConformityList").items.getById(itemid).update({
                  //   AuditeeAttachmentId: { results: lookupValues } // This field must be a multi-lookup pointing to AuditeeAttachmentDocs
                  // });
                }
              });
            }

          }
        }
        else if (_editsubmitStatus == "draft") {
          if (currentUserRole == "FirstAssignedTo" || currentUserRole == "DelegateTo") {
            // alert("Approval Item ID" + approvalItemId)
            let uploadedIds: number[] = [];

            // 1. Delete removed files
            if (this.state.fileDeleteIdauditee.length > 0) {
              this.state.fileDeleteIdauditee.forEach(function (ids) {
                sp.web.lists.getByTitle("AuditeeAttachmentDocs").items.getById(ids).delete();
              });
            }

            // 2. Upload new files
            if (this.state.copyFilauditee.length > 0) {
              const uploadPromises = this.state.copyFilauditee.map(async (file) => {


                const currentUser = await sp.web.currentUser();
                const userId = currentUser.Id; // Or however you get the current user ID
                const date = new Date();
                const fileExtension = file.name.split('.').pop();

                const components = [
                  date.getFullYear(),
                  (date.getMonth() + 1).toString().padStart(2, '0'),
                  date.getDate().toString().padStart(2, '0'),
                  date.getHours().toString().padStart(2, '0'),
                  date.getMinutes().toString().padStart(2, '0'),
                  date.getSeconds().toString().padStart(2, '0'),
                  date.getMilliseconds().toString().padStart(3, '0')
                ];
                //var fileNamePath = encodeURI(file.name);
                let newfileNameNew: any = await this.cleanFileNameSave(file.name);
                const fileNameWithoutExtension = newfileNameNew.split('.').slice(0, -1).join('.');
                var fileNamePath = `${userId}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
                //var fileNamePath = `${userId}_${components.join('')}_${newfileNameNew}`;
                //return `${userId}_${components.join('')}_${file.name}`;
                //const fileNamePath = encodeURI(file.name);
                //const fileNamePath = await this.getNewFileName(file.name);
                return sp.web.getFolderByServerRelativePath("AuditeeAttachmentDocs").files
                  .addUsingPath(fileNamePath, file, { Overwrite: true })
                  .then((response) => {
                    return response.file.getItem().then((fileItem: any) => {
                      uploadedIds.push(fileItem.Id); // Store ID of uploaded file
                      return fileItem.update({
                        NonConformityId: itemid // Optional: link document to current item
                      });
                    });
                  });
              });

              // 3. Once all uploads are done, update multi-lookup
              Promise.all(uploadPromises).then(() => {
                if (uploadedIds.length > 0) {
                  const lookupValues = uploadedIds.map((id) => ({ Id: id }));

                  sp.web.lists.getByTitle("NonConformityList").items.getById(itemid).update({
                    AuditeeAttachmentId: { results: lookupValues } // This field must be a multi-lookup pointing to AuditeeAttachmentDocs
                  });
                }
              });
            }
          } else {
            if (this.state.fileDeleteId.length > 0) {
              this.state.fileDeleteId.forEach(function (ids) {
                sp.web.lists.getByTitle("NonConformityDocs").items.getById(ids).delete();
              });
            }
            if (this.state.copyFil.length > 0) {
              let uploadedIdsatt: number[] = [];
              const uploadPromises = this.state.copyFil.map(async (file) => {

                //const sp = spfi().using(SPFx(this.props.context));
                const currentUser = await sp.web.currentUser();
                const userId = currentUser.Id; // Or however you get the current user ID
                const date = new Date();
                const fileExtension = file.name.split('.').pop();

                const components = [
                  date.getFullYear(),
                  (date.getMonth() + 1).toString().padStart(2, '0'),
                  date.getDate().toString().padStart(2, '0'),
                  date.getHours().toString().padStart(2, '0'),
                  date.getMinutes().toString().padStart(2, '0'),
                  date.getSeconds().toString().padStart(2, '0'),
                  date.getMilliseconds().toString().padStart(3, '0')
                ];

                //var fileNamePath = encodeURI(file.name);
                let newfileNameNew: any = await this.cleanFileNameSave(file.name);
                const fileNameWithoutExtension = newfileNameNew.split('.').slice(0, -1).join('.');
                var fileNamePath = `${userId}_${fileNameWithoutExtension}_${components.join('')}.${fileExtension}`;
                //var fileNamePath = `${userId}_${components.join('')}_${newfileNameNew}`;
                //return `${userId}_${components.join('')}_${file.name}`;
                //var fileNamePath = encodeURI(file.name);
                // var fileNamePath = await this.getNewFileName(file.name);
                sp.web.getFolderByServerRelativePath("NonConformityDocs").files
                  .addUsingPath(fileNamePath, file, { Overwrite: true })
                  .then(async (response) => {
                    return response.file.getItem().then((fileItem: any) => {
                      uploadedIdsatt.push(fileItem.Id); // Store ID of uploaded file
                      return fileItem.update({
                        NonConformityId: itemid // Optional: link document to current item
                      });
                    });

                  });
              })
              Promise.all(uploadPromises).then(() => {
                if (uploadedIdsatt.length > 0) {
                  const lookupValues = uploadedIdsatt.map((id) => ({ Id: id }));

                  // sp.web.lists.getByTitle("NonConformityList").items.getById(itemid).update({
                  //   AuditeeAttachmentId: { results: lookupValues } // This field must be a multi-lookup pointing to AuditeeAttachmentDocs
                  // });
                }
              });
            }
          }
        }
        setloading = false;
        setTimeout(() => {
          this.setState({ Loading: false }, () => {
            Swal.fire({
              title: cText + " Successfully.",
              icon: "success"
            }).then(() => {
              const baseUrl = context.pageContext.web.absoluteUrl;
              const redirectUrl =
                this.state.editCurrentUserRole === "DelegateTo" || this.state.editCurrentUserRole === "FirstAssignedTo"
                  ? `${baseUrl}/SitePages/MyApprovals.aspx`
                  : `${baseUrl}/SitePages/EDCMAIN.aspx`;

              window.location.href = redirectUrl;
            });
          });
        }, 2000); // Delay of 5000 milliseconds = 5 seconds


      }
    })
      .catch(error => {
        console.error("Error while saving:", error);
      });
  }
  //Forward Call
  public forwardRequest = async (_editsubmitStatus: string) => {

    const IsactionTaken = Number(Approvallistitemid) > 0 ?
      await this.CheckIfAlreadyactionTaken(Number(Approvallistitemid), CONTENTTYPE_NonComformity) : true;

    var _self = this;
    //Start Flow condition
    let currentUserRole = "";
    let FirstAssignedToSavedStatus = "No";
    let FirstDelegateToSavedStatus = "No";
    let firstInitiatorSubmitStatus = "";
    let firstAssignedToSubmitStatus = "";
    let delegateToSubmitStatus = "";
    let analyzedBySubmitStatus = "";
    let reviewedBySubmitStatus = "";
    let lastAssignedToSubmitStatus = "";
    let lastInitiatorSubmitStatus = "";
    let reworkById: any;
    let serialNumber: number;
    let ncrnumber = "";
    let documentCode = "";

    //<-------- Start Case1 --------->
    if (this.state.editDelegateToId == null) {
      if (_editsubmitStatus == "Forward" && this.state.editCurrentUserRole == "LastInitiator") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "Yes";
        currentUserRole = "Approverrole";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case1 --------->
    //<-------- Start Case2 --------->
    else {
      if (_editsubmitStatus == "Forward" && this.state.editCurrentUserRole == "LastInitiator") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "Yes";
        lastInitiatorSubmitStatus = "Yes";
        currentUserRole = "Approverrole";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case2 --------->
    if (!IsactionTaken) {
      Swal.fire("Action has already been taken on this record.");
      return;
    }
    const { updateData } = this;
    const { remarks, approvers, apprDelId } = this.state;
    const { currentUserID, approvalItemId, context } = this.props;
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: 'Do you want to forward this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    })
      //.then(async function (val) {
      .then(async (val) => {
        if (val.isConfirmed) {

          setloading = true;
          this.setState({ Loading: true });
          await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber, FirstAssignedToSavedStatus, FirstDelegateToSavedStatus);
          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
            Status: "Approved",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
          });
          //Forward Button
          console.log("this.state.....", this.state.approvers);
          let depart: any = this.state.editDepartmentOption.filter((item: any) => item.value == this.state.editDepartment)[0]?.label;

          if (approvers.length > 0) {
            var maxLength = approvers.length;
            approvers.forEach(function (it: any, val: any) {
              if (it.itemId == 0) {
                sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
                  Title: "Title",
                  MainListNameId: _self.state.mainListId,
                  ApproverRoleId: it.Role,
                  Level: val + 1,
                  LevelType: it.Type || "One",
                  SubmitStatus: "Yes",
                  Maxlevel: maxLength,
                  ContentTitle: _self.state.editProblemDescription,
                  RequestId: _self.state.editNCNumber + " / " + _self.state.editncType + " / " + depart,
                  RequesterNameId: _self.props.currentUserID,
                  RequestedDate: new Date(),
                  ProcessName: "Non Conformity",
                  FormNameId: _self.state.formNameId,
                  MainListID: _self.state.ncItemId,
                  RequesterRoleId: _self.state.reqRolId,
                  ApproversId: it.Name,
                  ApprovalType: "Approval",
                  Responsibility: it.Responsibility || "",
                  IsSignatureRequired: it.IsSignatureRequired ? "Yes" : "No",
                }).catch(function (ex) {
                  console.log(ex.errorMessage);
                })
              }
              else {
                sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(it.itemId).update({
                  Title: "Title",
                  MainListNameId: _self.state.mainListId,
                  ApproverRoleId: it.Role,
                  Level: val + 1,
                  LevelType: it.Type || "One",
                  SubmitStatus: "Yes",
                  Maxlevel: maxLength,
                  ContentTitle: _self.state.editProblemDescription,
                  RequestId: _self.state.editNCNumber + " / " + _self.state.editncType + " / " + depart,
                  RequesterNameId: _self.props.currentUserID,
                  RequestedDate: new Date(),
                  ProcessName: "Non Conformity",
                  FormNameId: _self.state.formNameId,
                  MainListID: _self.state.ncItemId,
                  RequesterRoleId: _self.state.reqRolId,
                  ApproversId: it.Name,
                  ApprovalType: "Approval",
                  Responsibility: it.Responsibility || "",
                  IsSignatureRequired: it.IsSignatureRequired ? "Yes" : "No",
                }).catch(function (ex) {
                  console.log(ex.errorMessage);
                })
              }
            });
          }
          if (apprDelId.length > 0) {
            apprDelId.forEach(function (ids: number) {
              spfi(this._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(ids).delete();
            });
          }
          //End Forward Button
          setloading = false;
          setTimeout(() => {
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Forwarded Successfully",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });
          }, 2000); // Delay of 5000 milliseconds = 5 seconds
        }
      });

  }
  //Approver call
  public approveRequest = async (_editsubmitStatus: string) => {
    //Start Flow condition

    let currentUserRole = "";
    let firstInitiatorSubmitStatus = "";
    let FirstAssignedToSavedStatus = "No";
    let FirstDelegateToSavedStatus = "No";
    let firstAssignedToSubmitStatus = "";
    let delegateToSubmitStatus = "";
    let analyzedBySubmitStatus = "";
    let reviewedBySubmitStatus = "";
    let lastAssignedToSubmitStatus = "";
    let lastInitiatorSubmitStatus = "";
    let reworkById: any;
    let serialNumber: number;
    let ncrnumber = "";
    let documentCode = "";
    const IsactionTaken = Number(Approvallistitemid) > 0 ?
      await this.CheckIfAlreadyactionTaken(Number(Approvallistitemid), CONTENTTYPE_NonComformity) : true;

    //<-------- Start Case1 --------->
    if (this.state.editDelegateToId == null) {
      // if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
      //   firstInitiatorSubmitStatus = "Yes"
      //   firstAssignedToSubmitStatus = "Yes";
      //   delegateToSubmitStatus = "No";
      //   analyzedBySubmitStatus = "Yes";
      //   reviewedBySubmitStatus = "No";
      //   lastAssignedToSubmitStatus = "No";
      //   lastInitiatorSubmitStatus = "No";
      //   currentUserRole = "ReviewedBy";
      //   reworkById = null;
      //   serialNumber = this.state.notUpdateSerialNo;
      //   ncrnumber = this.state.notUpdateDepartmentCode;
      // }
      //else
      if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "LastInitiator";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case1 --------->
    //<-------- Start Case2 --------->
    else {
      // if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
      //   firstInitiatorSubmitStatus = "Yes"
      //   firstAssignedToSubmitStatus = "Yes";
      //   delegateToSubmitStatus = "Yes";
      //   analyzedBySubmitStatus = "Yes";
      //   reviewedBySubmitStatus = "No";
      //   lastAssignedToSubmitStatus = "No";
      //   lastInitiatorSubmitStatus = "No";
      //   currentUserRole = "ReviewedBy";
      //   reworkById = null;
      //   serialNumber = this.state.notUpdateSerialNo;
      //   ncrnumber = this.state.notUpdateDepartmentCode;
      // }
      // else 
      // if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
      //   firstInitiatorSubmitStatus = "Yes"
      //   firstAssignedToSubmitStatus = "Yes";
      //   delegateToSubmitStatus = "Yes";
      //   analyzedBySubmitStatus = "Yes";
      //   reviewedBySubmitStatus = "Yes";
      //   lastAssignedToSubmitStatus = "No";
      //   lastInitiatorSubmitStatus = "No";
      //   currentUserRole = "LastAssignedTo";
      //   reworkById = null;
      //   serialNumber = this.state.notUpdateSerialNo;
      //   ncrnumber = this.state.notUpdateDepartmentCode;
      // }
      // else 
      if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "Yes";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "LastInitiator";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else
        if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "LastInitiator") {
          firstInitiatorSubmitStatus = "Yes"
          firstAssignedToSubmitStatus = "Yes";
          delegateToSubmitStatus = "Yes";
          analyzedBySubmitStatus = "Yes";
          reviewedBySubmitStatus = "Yes";
          lastAssignedToSubmitStatus = "Yes";
          lastInitiatorSubmitStatus = "Yes";
          currentUserRole = "";
          reworkById = null;
          serialNumber = this.state.notUpdateSerialNo;
          ncrnumber = this.state.notUpdateDepartmentCode;
        }
    }
    //<-------- End Case2 --------->
    if (!IsactionTaken) {
      Swal.fire("Action has already been taken on this record.");
      return;
    }
    const { updateData } = this;
    const { remarks, editLastInitiatorSubmitStatus } = this.state;
    const { currentUserID, approvalItemId, context } = this.props;
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: 'Do you want to approve this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    })
      //.then(async function (val) {
      .then(async (val) => {
        if (val.isConfirmed) {

          setloading = true;
          this.setState({ Loading: true });
          if (editLastInitiatorSubmitStatus != "Yes") {
            await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber, FirstAssignedToSavedStatus, FirstDelegateToSavedStatus);
          }
          const sp = spfi().using(SPFx(this.props.context));
          if (this.state.IsFinalapprover) {
            let test1 = this.state.correctionApplicable ? "Yes" : "No";
            let test2 = this.state.notEffective ? "Yes" : "No";
            let test3 = this.state.effectiveClosed ? "Yes" : "No";
            await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId).update({
              // Correctionapplicable: this.state.IsFinalapprover
              //   ? test1
              //   : "",
              CloseOutStatus: this.state.IsFinalapprover ? "Completed" : "Open",
              // NotEffective: this.state.IsFinalapprover
              //   ? test2
              //   : "",
              // EffectiveandProblemClosed: this.state.IsFinalapprover
              //   ? test3
              //   : "",
              FinalRemarks: this.state.IsFinalapprover
                ? this.state.remarks : "",
            })
          }

          // alert("Approved" + approvalItemId + typeof(approvalItemId));
          // alert("Approved" + Approvallistitemid + typeof(Approvallistitemid));

          if (approvalItemId == undefined || approvalItemId == null || approvalItemId == "") {
            // sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(approvalItemId)).update({
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
              Status: "Approved",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            setTimeout(() => {
              this.setState({ Loading: false }, () => {
                Swal.fire({
                  title: "Approved Successfully",
                  icon: "success"
                }).then(() => {
                  window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
                });
              });
            }, 2000); // Delay of 5000 milliseconds = 5 seconds

          } else {
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({

              Status: "Approved",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            setTimeout(() => {
              this.setState({ Loading: false }, () => {
                Swal.fire({
                  title: "Approved Successfully",
                  icon: "success"
                }).then(() => {
                  window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
                });
              });
            }, 2000); // Delay of 5000 milliseconds = 5 seconds
          }

        }
      });

  }
  //Reject call
  private rejectRequest = async (_editsubmitStatus: string) => {
    const { remarks } = this.state
    const { currentUserID, approvalItemId, context } = this.props
    const sp = spfi().using(SPFx(this.props.context));
    const IsactionTaken = Number(Approvallistitemid) > 0 ?
      await this.CheckIfAlreadyactionTaken(Number(Approvallistitemid), CONTENTTYPE_NonComformity) : true;
    if (!IsactionTaken) {
      Swal.fire("Action has already been taken on this record.");
      return;
    }
    Swal.fire({
      title: 'Do you want to reject this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    })
      .then(async (val) => {
        //.then(function (val) {
        if (val.isConfirmed) {
          setloading = true;
          this.setState({ Loading: true });
          if (approvalItemId == undefined || approvalItemId == null || approvalItemId == "") {
            await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId).update({
              CloseOutStatus: "Completed",
            })
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
              Status: "Rejected",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            setTimeout(() => {
              this.setState({ Loading: false }, () => {
                Swal.fire({
                  title: "Rejected Successfully",
                  icon: "success"
                }).then(() => {
                  window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
                });
              });
            }, 1000); // Delay of 5000 milliseconds = 5 seconds
          } else {
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
              Status: "Rejected",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            setTimeout(() => {
              this.setState({ Loading: false }, () => {
                Swal.fire({
                  title: "Rejected Successfully",
                  icon: "success"
                }).then(() => {
                  window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
                });
              });
            }, 1000); // Delay of 5000 milliseconds = 5 seconds

          }

        }
      });

  }
  private cancelModalAction = (refresh?: boolean,) => {

    this.setState({ redirecturl: window.location.href });
    //setShowfileNew(false);
    this.setState({ ShowModalTemplateDoc: false, redirecturl: "" });
    Showfile = false;
  }
  private OpenFileTemplate = (obj: any, sts: string) => {

    this.setState({ ShowModalTemplateDoc: true });
    if (sts == "Open") {
      Showfile = true;
    }
    let url = this.props.context.pageContext.web.absoluteUrl;
    let tenanturl = url.match(/^https:\/\/[^\/]+/)[0];
    console.log("ttrtrtrtt", obj)
    const fileUrl = `${tenanturl}${obj.FileRef}`;
    if (sts == "Open") {
      if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
        const viewerUrl = `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview`;
        this.setState({ redirecturl: viewerUrl })
      } else {
        this.setState({ redirecturl: fileUrl })

      }

    } else if (sts == "Download") {
      const link = document.createElement("a");
      link.href = fileUrl;
      link.setAttribute("download", obj?.FileLeafRef != "" ? this.cleanFileName(obj.FileLeafRef) : this.cleanFileName(obj.name)); // Suggests a filename for download
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

    }

  }
  private OpenFile = (obj: any, sts: string) => {

    this.setState({ ShowModalAtt: true });
    if (sts == "Open") {
      Showfile = true;
    }
    let url = this.props.context.pageContext.web.absoluteUrl;
    let tenanturl = url.match(/^https:\/\/[^\/]+/)[0];
    console.log("obbbj", obj)
    const fileUrl = `${tenanturl}${obj.FileRef}`;

    if (sts == "Open") {
      if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
        const viewerUrlppt = `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview`;
        this.setState({ redirecturl: viewerUrlppt })
      } else {
        this.setState({ redirecturl: fileUrl })

      }
    } else if (sts == "Download") {
      const link = document.createElement("a");
      link.href = fileUrl;
      link.setAttribute("download", obj?.FileLeafRef != "" ? this.cleanFileName(obj.FileLeafRef) : this.cleanFileName(obj.name)); // Suggests a filename for download
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  }
  private cleanFileName = (filename: string) => {
    debugger
    // Match a 14-digit datetime suffix before the file extension
    // const datetimePattern = /_\d{14}(?=\.[^.]+$)/;
    // const prefixPattern = /^\d+_\d{17}_/;
    // if (prefixPattern.test(filename)) {
    //   return filename.replace(prefixPattern, '');
    // }
    const match = filename.match(/^\d+_(.*?)_\d{17}\.[^.]+$/);
    if (match) {
      return `${match[1]}.${filename.split('.').pop()}`;
    }
    console.log("ghghggh", filename)
    return filename;
  }

  private cleanFileNameSave = async (filename: string) => {
    // Match a 14-digit datetime suffix before the file extension
    const datetimePattern = /_\d{14}(?=\.[^.]+$)/;

    if (datetimePattern.test(filename)) {
      return filename.replace(datetimePattern, '');
    }
    console.log("ghghggh", filename)
    return filename;
  }
  //Rework Call
  private reworkRequest = async (_editsubmitStatus: string) => {
    //Start Flow condition

    const IsactionTaken = Number(Approvallistitemid) > 0 ?
      await this.CheckIfAlreadyactionTaken(Number(Approvallistitemid), CONTENTTYPE_NonComformity) : true;
    let currentUserRole = "";
    let FirstAssignedToSavedStatus = "No";
    let FirstDelegateToSavedStatus = "No";
    let firstInitiatorSubmitStatus = "";
    let firstAssignedToSubmitStatus = "";
    let delegateToSubmitStatus = "";
    let analyzedBySubmitStatus = "";
    let reviewedBySubmitStatus = "";
    let lastAssignedToSubmitStatus = "";
    let lastInitiatorSubmitStatus = "";
    let reworkById: any;
    let serialNumber: number;
    let ncrnumber = "";
    let documentCode = "";
    const { currentUserID, approvalItemId, context } = this.props;
    //<-------- Start Case1 --------->
    if (this.state.editDelegateToId == null) {
      if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstAssignedTo";
        reworkById = this.state.editAnalyzedById || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      } else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo") {
        firstInitiatorSubmitStatus = "No"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstInitiator"; //LastInitiator
        reworkById = this.state.editAssignToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      // else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "ReviewedBy") {
      //   firstInitiatorSubmitStatus = "Yes"
      //   firstAssignedToSubmitStatus = "No";
      //   delegateToSubmitStatus = "No";
      //   analyzedBySubmitStatus = "No";
      //   reviewedBySubmitStatus = "No";
      //   lastAssignedToSubmitStatus = "No";
      //   lastInitiatorSubmitStatus = "No";
      //   currentUserRole = "FirstAssignedTo";
      //   reworkById = this.state.editReviewedById || null;
      //   serialNumber = this.state.notUpdateSerialNo;
      //   ncrnumber = this.state.notUpdateDepartmentCode;
      // }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "LastInitiator") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstAssignedTo";
        reworkById = this.state.Requester.Id || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "Approverrole") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "Yes";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "LastInitiator";
        reworkById = currentUserID || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case1 --------->
    //<-------- Start Case2 --------->
    else {
      if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstAssignedTo",// "DelegateTo";
          reworkById = this.state.editAnalyzedById || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      } else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo") {
        firstInitiatorSubmitStatus = "No"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstInitiator"; //LastInitiator
        reworkById = this.state.editDelegateToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "DelegateTo") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstAssignedTo",//"FirstInitiator"; //LastInitiator
          reworkById = this.state.editDelegateToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      // else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "ReviewedBy") {
      //   firstInitiatorSubmitStatus = "Yes"
      //   firstAssignedToSubmitStatus = "Yes";
      //   delegateToSubmitStatus = "No";
      //   analyzedBySubmitStatus = "No";
      //   reviewedBySubmitStatus = "No";
      //   lastAssignedToSubmitStatus = "No";
      //   lastInitiatorSubmitStatus = "No";
      //   currentUserRole = "DelegateTo";
      //   reworkById = this.state.editReviewedById || null;
      //   serialNumber = this.state.notUpdateSerialNo;
      //   ncrnumber = this.state.notUpdateDepartmentCode;
      // }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "LastAssignedTo") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstInitiator";
        reworkById = this.state.editDelegateToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "LastInitiator") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstAssignedTo";
        reworkById = this.state.Requester.Id || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "Approverrole") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "Yes";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "LastInitiator";
        reworkById = currentUserID || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case2 --------->

    if (!IsactionTaken) {
      Swal.fire("Action has already been taken on this record.");
      return;
    }
    const { remarks, reworkremarks } = this.state;

    const sp = spfi().using(SPFx(this.props.context));
    const { updateData } = this;
    Swal.fire({
      title: 'Do you want to rework this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    })
      //.then(async function (val) {
      .then(async (val) => {
        if (val.isConfirmed) {
          // return;
          setloading = true;
          this.setState({ Loading: true });
          await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber, FirstAssignedToSavedStatus, FirstDelegateToSavedStatus);
          if (approvalItemId == undefined || approvalItemId == null || approvalItemId == "") {

            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
              Status: "Rework",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
              ReworkRemarks: reworkremarks,
              IsRework: "Yes"
            });
            setloading = false;
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Sent for Rework.",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });

          } else {

            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
              Status: "Rework",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
              ReworkRemarks: reworkremarks,
              IsRework: "Yes"
            });

            setloading = false;
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Send for Rework.",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });

          }


        }
      });
  }
  private onChangenctype = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({
        editncType: option.key as string,
        editmemonumberOptions: option?.text == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs,
        editNCNumber: "", editNCNumberID: "", NCNumberselected: [], ApprovedAuditSelected: []
      });
    }
  };
  private handleChangecheckbox = (value: any, field: string) => {
    if (field == "IMSUpdated") {
      this.setState((prevState) => ({
        ...prevState,
        isIMSUpdated: value
      }))
    } else {
      this.setState((prevState) => ({
        ...prevState,
        riskandopportunitiesUpdated: value
      }))
    }

  };
  private handleChangecheckboxsingle = (field: string, event: any) => {
    if (field == "correctionApplicable") {
      this.setState({ correctionApplicable: event.target.checked });
    }
    if (field == "notEffective") {
      this.setState({ notEffective: event.target.checked });
    }
    if (field == "effectiveClosed") {
      this.setState({ effectiveClosed: event.target.checked });
    }

  }

  private async isUserDelegatedFor(actingForEmail: string, currentUserId: number): Promise<boolean> {
    const today = new Date().toISOString();
    const _sp = spfi().using(SPFx(this.props.context));
    try {
      const delegatedUsers = await _sp.web.lists.getByTitle("ARGDelegateList").items
        .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
        .expand("Author,DelegateName,ActingFor")
        .filter(`DelegateName/EMail eq '${actingForEmail}' and ActingFor/ID eq ${currentUserId} and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
        .orderBy("Created", false)
        .top(1)();

      return delegatedUsers.length > 0;
    } catch (error) {
      console.error("Error checking delegation:", error);
      return false;
    }
  }
  private getdigitalsignaturerequestbyID = async (listname: any, _sp: any, id: any) => {
    let arr: any = []
    try {
      console.log("iddddd", id);
      const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
        .filter(`ListName eq '${listname}' and ListItemID eq ${id} and DocSignedStatus eq 'No'`)
        .top(100)
        ()
        .then((res: any) => {
          console.log(res, ' let arrs=[]');

          arr = res
          // arr = res;
        })
      // Perform any necessary actions after successful addition
    } catch (error) {
      console.log('Error adding item:', error);
      // Handle errors appropriately
      arr = null
    }
    return arr;
  };
  private getdigitalsignaturerequestbyIDYes = async (listname: any, _sp: any, id: any) => {
    debugger
    let arr = [];
    let Norecrodsexist = "No";
    try {
      console.log("iddddd", id);
      const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
        .filter(`ListName eq '${listname}' and ListItemID eq ${id}`)
        .top(100)
        ()
        .then((res: any) => {
          console.log(res, ' let arrs=[]');
          if (res.length > 0) {

            for (let i = 0; i < res.length; i++) {
              if (res[i].DocSignedStatus === "No" || res[i].DestinationIDUpdated === "No") {
                Norecrodsexist = "Yes";
                break; // Exit the loop early since we found a match
              }
            }

            arr = res
          } else {
            Norecrodsexist = "NoRecord";
          }

          // arr = res;
        })
      // Perform any necessary actions after successful addition
    } catch (error) {
      console.log('Error adding item:', error);
      // Handle errors appropriately
      arr = null
    }
    console.log("NorecrodsexistNorecrodsexist", Norecrodsexist);
    return Norecrodsexist;
  };
  private updateDigitalsign = async (listname: any, _sp: any, id: any, formitemid: any) => {
    let resultArr = []
    try {
      console.log("iddddd", id);
      // const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
      //   .filter(`ListName eq '${listname}' and ListItemID eq ${id}`)
      //   .top(1)
      //   ().then(async (res) => {
      let newItem;
      const postPayload2 = {
        DocSignedStatus: "Yes"
      }
      const newItem1 = await this.getdigitalsignaturerequestbyID("NonConformityList", _sp, Number(formitemid))
        .then(async (res: any) => {
          console.log("NC digi doc", res);
          for (var i = 0; i < res.length; i++) {
            newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items.getById(res[i].ID).update(postPayload2);
            console.log('Item added successfully:', newItem);

          }
        })
      resultArr = newItem
      // Perform any necessary actions after successful addition
    } catch (error) {
      console.log('Error adding item:', error);
      // Handle errors appropriately
      resultArr = null
    }
    return resultArr;
  };
  private updatedigisignnew = async () => {
    const sp = spfi().using(SPFx(this.props.context));
    let items = await this.updateDigitalsign("NonConformityList", sp, this.state.DigitalsignID, this.state.mainItemId);
    if (items) {
      this.setState({ hidedigisign: true });
      const isRecordExist = await this.getdigitalsignaturerequestbyIDYes("NonConformityList", sp, Number(this.state.mainItemId));
      debugger
      console.log("isRecordExistisRecordExist", isRecordExist)
      if (isRecordExist == "Yes" || isRecordExist == "NoRecord") {
        this.setState({ showdigisign: false });

      } else {
        this.setState({ showdigisign: true });

      }

    }
  }
  public render(): React.ReactElement<IAuditPlanProps> {
    // Group 1: Action visibility conditions
    const isApproveVisible = this.state.showApprove === true;
    const isRejectVisible = this.state.showReject === true;
    const isReworkVisible = showreworkremarks;
    const isIMSVisible = showimsupdated && !isdisableims;
    const isCorrectionVisible = showcorrectionappicable && !isdisablefinal;
    const isNotFirstInitiator = this.state.editCurrentUserRole !== "FirstInitiator";

    // Group 2: Delegation/submission status
    const isReviewedBySubmit = this.state.editReviewedBySubmitStatus === "Yes" && this.state.editDelegateToId == null;
    const isLastAssignedSubmit = this.state.editLastAssignedToSubmitStatus === "Yes" && this.state.editDelegateToId != null;

    // Group 3: FirstInitiator in Rework status
    const isFirstInitiatorRework =
      this.state.editCurrentUserRole === "FirstInitiator" &&
      this.state.editStatus === "Rework" &&
      this.state.editFirstInitiatorSubmitStatus === "No";

    // Group 4: Loading state
    const isNotLoading = !this.state.Loading || !setloading;

    // Group 5: Edit type or override condition
    const isEditableType = this.state.edType !== "view";
    const isIMSOverride = showimsupdated && !isdisableims;
    const isCorrectionOverride = showcorrectionappicable && !isdisablefinal;

    console.log("isApproveVisible:", isApproveVisible);
    console.log("isRejectVisible:", isRejectVisible);
    console.log("isReworkVisible:", isReworkVisible);
    console.log("isIMSVisible:", isIMSVisible);
    console.log("isCorrectionVisible:", isCorrectionVisible);
    console.log("isNotFirstInitiator:", isNotFirstInitiator);

    console.log("isReviewedBySubmit:", isReviewedBySubmit);
    console.log("isLastAssignedSubmit:", isLastAssignedSubmit);

    console.log("isFirstInitiatorRework:", isFirstInitiatorRework);

    console.log("isNotLoading:", isNotLoading);

    console.log("isEditableType:", isEditableType);
    console.log("isIMSOverride:", isIMSOverride);
    console.log("isCorrectionOverride:", isCorrectionOverride);

    // Final condition
    const finalCondition =
      (
        (
          (isApproveVisible || isRejectVisible || isReworkVisible || isIMSVisible || isCorrectionVisible)
          && isNotFirstInitiator
        ) ||
        isReviewedBySubmit ||
        isLastAssignedSubmit ||
        isFirstInitiatorRework
      )
      && isNotLoading
      && (isEditableType || isIMSOverride || isCorrectionOverride);

    console.log("FINAL CONDITION:", finalCondition);

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
    if (this.state.editCurrentUserRole != "FirstInitiator") {
      document.querySelectorAll("#AssigntoPeoplepicker .ms-BasePicker-text").forEach((el) => {
        el.classList.add(styles.peoplepickerstyleAuditte);
      });
    }
    document.querySelectorAll("#AssigntoPeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.add(styles.peoplepickerstyle);
    });
    // document.querySelectorAll("#approverpeoplepicker .ms-BasePicker-text").forEach((el) => {
    //   el.classList.add(styles.peoplepickerstyle);
    // });
    document.querySelectorAll("#reviewedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.add(styles.peoplepickerstyle);
    });
    document.querySelectorAll("#analyzedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.add(styles.peoplepickerstyle);
    });
    if (this.state.editDelegateToId == null) {
      document.querySelectorAll("#delegatetopeoplepicker .ms-BasePicker-text").forEach((el) => {
        el.classList.add(styles.peoplepickerstyle);
      });
    } else {
      document.querySelectorAll("#delegatetopeoplepicker .ms-BasePicker-text").forEach((el) => {
        el.classList.add(styles.peoplepickerstyleAuditte);
      });
    }

    const selectedTextDiv = document.getElementById('selectedText');
    if (selectedTextDiv) {
      selectedTextDiv.style.display = 'none';
    }
    var approval =
      //let approvers = [...this.state.approvers].sort((a, b) => a.index - b.index);

      this.state.approvers.map((item: any, i: number) => {
        console.log(this.state.approvers, "this.state.approvers", i)
        return (
          <tr key={i} className='tblCls'>
            {/* <td>
            <TextField value={(i + 1).toString()} disabled={true} className={styles.width} ></TextField>
          </td> */}
            <td style={{ minWidth: "40px", maxWidth: "40px" }}>
              <div
                style={{ marginLeft: "0px", overflow: 'inherit' }}
                className="indexdesign"
              >
                {i + 1}</div>
            </td>
            <td title={this.state.optionsRole.find(opt => opt.key === this.state.approvers[i].Role)?.text || "Select role"}
              style={{ overflow: 'inherit', minWidth: "110px", maxWidth: "110px" }} className="ng-binding">
              <Dropdown disabled={this.state.forwarDisable || forwardisdisabled} placeholder="Select options"

                selectedKey={this.state.approvers[i].Role}
                options={this.state.optionsRole
                  // .filter((opt) =>
                  //   !this.state.approvers.some((app, index) => index !== i && app.Role === opt.key)
                  // )
                }
                onChange={(e, itm: IDropdownOption) => this.onRoleChange(e, itm, i)}
                className={this.state.editErrors?.[`approvers[${i}].Role`] ? 'dropdown-error' : ''}
              // className={this.state.editErrors?.approvers ? 'dropdown-error' : ''}
              />
            </td>
            <td title={optionsResponsibility.find(opt => opt.key === this.state.approvers[i].Responsibility)?.text || "Select"}
              style={{ overflow: 'inherit', minWidth: "85px", maxWidth: "85px" }} className="ng-binding">
              <Dropdown
                disabled={this.state.forwarDisable || forwardisdisabled}
                placeholder="Select options"
                selectedKey={this.state.approvers[i].Responsibility || 'Signer'}
                options={optionsResponsibility} onChange={(e, itm: IDropdownOption) => this.handleChangeResp(e, itm, i)}
                className={this.state.editErrors?.Responsibility ? 'dropdown-error' : ''}

              />
            </td>

            <td style={{ minWidth: "46px", maxWidth: "46px", overflow: 'inherit' }}>
              <div
              //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
              >
                <input
                  type="checkbox"
                  checked={this.state.approvers[i].IsSignatureRequired}
                  disabled={this.state.approvers[i].Responsibility === "Signer" || this.state.approvers[i].Responsibility === "" || (this.state.forwarDisable || forwardisdisabled)}
                  style={{ marginLeft: '17px', width: "15px" }}

                  onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                    const isChecked = e.target.checked;
                    this.state.approvers[i].IsSignatureRequired = isChecked;
                    this.setState({ approvers: this.state.approvers });

                  }}
                />

              </div>
            </td>
            <td title={`Level ${(i + 1).toString()}` || "Level "} style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>
              <TextField value={`Level ${(i + 1).toString()}`} disabled={true}></TextField>
            </td>
            <td title={this.state.approvers[i].AppName && this.state.approvers[i].AppName.length > 0
              ? this.state.approvers[i].AppName.map((user: any) => user.Title).join(', ')
              : "Select approver"}
              style={{ overflow: 'inherit', minWidth: '100px', maxWidth: '100px' }} id={`approverpeoplepicker-${i}`}>

              <PeoplePicker
                context={peoplePickerContext}
                personSelectionLimit={5}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                disabled={this.state.forwarDisable || forwardisdisabled}
                ensureUser={true}
                defaultSelectedUsers={this.state.approvers[i].appEx ? this.state.approvers[i].appEx : []}
                onChange={(e) => this._getPeoplePickerItemsApp(e, i)}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                styles={{
                  root: {
                    backgroundColor: this.state.editErrors?.[`approvers[${i}].name`] ? '#ffe6e6' : 'white'
                  }
                }}
              />
            </td>
            <td title={optionsApp.find(opt => opt.key === this.state.approvers[i].Type)?.text || "AnyOne"}
              style={{ overflow: 'inherit', minWidth: '100px', maxWidth: '100px' }} className="ng-binding">
              <Dropdown disabled={this.state.forwarDisable || forwardisdisabled} placeholder="Select options"
                selectedKey={this.state.approvers[i].Type || 'One'}
                options={optionsApp} onChange={(e, itm: IDropdownOption) => this.onTypeChange(e, itm, i)}
                className={this.state.editErrors?.approvers ? 'dropdown-error' : ''}

              />
            </td>
            <td style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>
              {/* <i className="fe-trash-2 text-danger"></i> */}
              {
                // ((this.state.editReviewedBySubmitStatus == "Yes" && this.state.editDelegateToId == null) ||
                //   (this.state.editLastAssignedToSubmitStatus == "Yes" && this.state.editDelegateToId != null))
                !this.state.forwarDisable && !forwardisdisabled
                  ?
                  <img src={require("../assets/del.png")} onClick={(e) => this.deleteItemApp(i)} />
                  :
                  <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' />}
              {/* <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} className='sidebariconsmall' /> */}

            </td>
            {/* <td><button type="button" onClick={(e) => this.deleteItemApp(i)}>Delete</button></td> */}
          </tr>
        )
      })
    var fileData = this.state.copyFil.map((item: any, i: number) => {

      return (
        <tr style={{ display: 'table', width: '100%' }}>
          <td style={{ minWidth: '50px', maxWidth: '50px' }} className="text-center">{i + 1}</td>
          <td>
            {decodeURIComponent(item.name)}
          </td>
          {/* <td></td> */}
          <td>{new Date().toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric"
          }).replace(/ /g, "/")}</td>
          <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
            <img src={require("../assets/del.png")} className='' onClick={() => this.removeFiles(i)}></img>
          </td>
          {/* <td>
            <button type="button" onClick={(e) => this.removeFiles(i)}>Delete</button>
          </td> */}
        </tr>
      )
    });
    var upFiles = this.state.exFiles.map((item: any, i: number) => {
      console.log("iteeeeee", item.Name, item.Name || (item.Name.includes('_') ? item.Name.split('_')[2] : item.Name))
      return (
        <tr >
          <td style={{ minWidth: '50px', maxWidth: '50px' }} className="text-center">
            {i + 1}
          </td>
          {/* <td title={decodeURIComponent(item.Name)}>
            {decodeURIComponent(item.Name)}
          </td> */}
          <td style={{ minWidth: '150px', maxWidth: '150px', textAlign: 'center' }} title={(item.Name.includes('_') ? item.Name.split('_')[2] : item.Name)}>
            {(item.Name.includes('_') ? item.Name.split('_')[2] : item.Name)}</td>

          <td style={{ minWidth: '100px', maxWidth: '100px', textAlign: 'center' }} title={moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")} className="text-center">
            {moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")}
          </td>
          <td style={{ minWidth: '60px', maxWidth: '60px', textAlign: 'center' }} >
            <span onClick={() => this.OpenFile(item && item, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon title='Preview file' icon={faEye} /></span>
            <span onClick={() => this.OpenFile(item && item, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon title='Download file' icon={faDownload} /></span>
            {this.state.ShowDeleteicon &&
              <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeleted(i)}></img>
            }
            {/* {<a href={item.Path} target="_blank">Link</a>} */}
          </td>
          {/* {this.state.ShowDeleteicon &&
            <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
              <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeleted(i)}></img>
            </td>
          } */}
        </tr>
      )
    });
    { console.log("this.state.fileDataAuditee", this.state.copyFilauditee) }
    var fileDataAuditee = this.state.copyFilauditee.map((item: any, i: number) => {
      return (
        <tr >
          <td style={{ minWidth: '60px', maxWidth: '60px' }}>{i + 1}</td>
          <td style={{ minWidth: '150px', maxWidth: '150px' }}>
            {item.name}
          </td>
          {/* <td>NA</td> */}
          <td style={{ minWidth: '100px', maxWidth: '100px' }} className="text-center">{new Date().toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric"
          }).replace(/ /g, "/")}</td>
          <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
            <img src={require("../assets/del.png")} className='' onClick={() => this.removeFilesAuditee(i)}></img>
          </td>
          {/* <td>
            <button type="button" onClick={(e) => this.removeFiles(i)}>Delete</button>
          </td> */}
        </tr>
      )
    });
    { console.log("this.state.exFilesauditee", this.state.exFilesauditee) }
    var upFilesauditee = this.state.exFilesauditee.map((item: any, i: number) => {
      return (
        <tr >
          <td style={{ minWidth: '60px', maxWidth: '60px' }} >
            {i + 1}
          </td>
          {/* <td title={decodeURIComponent(item.Name)}>
            {decodeURIComponent(item.Name)}
          </td> */}
          <td title={(item.Name.includes('_') ? item.Name.split('_')[2] : item.Name)}>
            {(item.Name.includes('_') ? item.Name.split('_')[2] : item.Name)}</td>

          <td title={moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")} style={{ minWidth: '100px' }} className="text-center">
            {moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")}
          </td>
          <td style={{ textAlign: 'center' }}>
            <span onClick={() => this.OpenFile(item && item, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon title='Preview file' icon={faEye} /></span>
            <span onClick={() => this.OpenFile(item && item, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon title='Download file' icon={faDownload} /></span>
            {this.state.ShowDeleteicon &&
              <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeletedauditee(i)}></img>
            }
            {/* {<a href={item.Path} target="_blank">Link</a>} */}
          </td>
          {/* {this.state.ShowDeleteicon &&
            // <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
            <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeletedauditee(i)}></img>
            // </td>
          } */}
        </tr>
      )
    });
    { console.log("apprItemsapprItems", this.state.apprItems) }
    var auditHistory = this.state.apprItems.map((item: any, i: number) => {
      const total = this.state.apprItems.length;
      return (
        <tr>
          <td style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
            {i + 1}
          </td>
          <td title={`Level ${i + 1}`} style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
            {/* {item.Level} */}
            Level {i + 1}  {/* Reverse Level */}
          </td>
          <td title={item.AssignedTo} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.AssignedTo}
          </td>
          <td title={item.ActionTakenRole == "LastInitiator" ? "Initiator" :
            item.ActionTakenRole == "FirstAssignedTo" ? "Auditee" :
              item.ActionTakenRole == "DelegateTo" ? "Delegate" :
                item.ActionTakenRole == "AnalyzedBy" ? "Analyzed By" :
                  item.ActionTakenRole == "FirstInitiator" ? "Initiator" :
                    item.ActionTakenRole} style={{ minWidth: "90px", maxWidth: "90px" }}>
            {item.ActionTakenRole == "LastInitiator" ? "Initiator" :
              item.ActionTakenRole == "FirstAssignedTo" ? "Auditee" :
                item.ActionTakenRole == "DelegateTo" ? "Delegate" :
                  item.ActionTakenRole == "AnalyzedBy" ? "Analyzed By" :
                    item.ActionTakenRole == "FirstInitiator" ? "Initiator" :
                      item.ActionTakenRole} {/* Divyansh Changes */}
          </td>
          <td title={item.RequesterName} style={{ minWidth: "90px", maxWidth: "90px" }}>
            {item.RequesterName}
          </td>
          <td title=
            {`${new Intl.DateTimeFormat('en-GB', {
              day: '2-digit',
              month: 'short',
              year: 'numeric'
            }).format(new Date(new Date(item?.RequestedDate).getTime() - 4 * 60 * 60 * 1000)).replace(/ /g, "/")} ${new Date(new Date(item?.RequestedDate).getTime() - 4 * 60 * 60 * 1000).toLocaleTimeString('en-GB', {
              hour: '2-digit',
              minute: '2-digit',
              hour12: false
            })}`}
            // {moment(new Date(item.RequestedDate)).format("DD/MMM/YYYY HH:mm")}
            style={{ minWidth: '90px', maxWidth: '90px' }}>
            {/* {item.RequestedDate ? moment(new Date(item.RequestedDate)).format("DD/MMM/YYYY HH:mm") : ""} */}
            {`${new Intl.DateTimeFormat('en-GB', {
              day: '2-digit',
              month: 'short',
              year: 'numeric'
            }).format(new Date(new Date(item?.RequestedDate).getTime() - 4 * 60 * 60 * 1000)).replace(/ /g, "/")} ${new Date(new Date(item?.RequestedDate).getTime() - 4 * 60 * 60 * 1000).toLocaleTimeString('en-GB', {
              hour: '2-digit',
              minute: '2-digit',
              hour12: false
            })}`}
          </td>
          <td title={item.ActionTakenBy} style={{
            minWidth: '90px', maxWidth: '90px',
            backgroundColor: item.AssignedTo && item?.ActionTakenBy && item.AssignedTo !== item.ActionTakenBy && item.Status !== "Auto Approved"
              ? auditHistoryDelegationBgColor
              : undefined,
            color: item.AssignedTo && item?.ActionTakenBy && item.AssignedTo !== item.ActionTakenBy && item.Status !== "Auto Approved"
              ? auditHistoryDelegationTextColor
              : undefined
          }}>
            {item.ActionTakenBy}
          </td>
          <td title=
            {item?.ActionTakenOn == null || item?.ActionTakenOn == undefined || item?.ActionTakenOn == "" ? "" :
              (`${new Intl.DateTimeFormat('en-GB', {
                day: '2-digit',
                month: 'short',
                year: 'numeric'
              }).format(new Date(new Date(item?.ActionTakenOn).getTime() - 0 * 60 * 60 * 1000)).replace(/ /g, "/")} ${new Date(new Date(item?.ActionTakenOn).getTime() - 0 * 60 * 60 * 1000).toLocaleTimeString('en-GB', {
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
              })}`)
            }
            // {moment(new Date(item.ActionTakenOn)).format("DD/MMM/YYYY HH:mm")} 
            style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item?.ActionTakenOn == null || item?.ActionTakenOn == undefined || item?.ActionTakenOn == "" ? "" :
              (`${new Intl.DateTimeFormat('en-GB', {
                day: '2-digit',
                month: 'short',
                year: 'numeric'
              }).format(new Date(new Date(item?.ActionTakenOn).getTime() - 0 * 60 * 60 * 1000)).replace(/ /g, "/")} ${new Date(new Date(item?.ActionTakenOn).getTime() - 0 * 60 * 60 * 1000).toLocaleTimeString('en-GB', {
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
              })}`)
            }
            {/* {item.ActionTakenOn ? moment(new Date(item.ActionTakenOn)).format("DD/MMM/YYYY HH:mm") : ""} */}
          </td>
          <td title={item.Remarks} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.Remarks}
          </td>
          <td title={item.Status} style={{ minWidth: '70px', maxWidth: '70px' }}>
            {item.Status}
          </td>
        </tr>
      )

    });
    return (
      <section style={{ padding: '10px 0px' }}>
        <div className={styles.welcome}>
          <div className="row">
            <div className="col-lg-6 newbread">
              <CustomBreadcrumb Breadcrumb={this.Breadcrumb} />
            </div>

          </div>
          {console.log("this.state.Loading || setloading", this.state.Loading, setloading)}
          {this.state.Loading && setloading ?

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

            <section className='card card-body' >
              <fieldset>
                <form>
                  {/* Start save as draft */}

                  <div className="previewIcon">
                    <h4 style={{ textAlign: 'left' }} className="text-dark font-16 fw-bold mb-3">Non Conformity / Observation Details</h4>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0px' }}>
                      {(this.state.editStatus === "Approved" || this.state.editStatus === "Rejected") && !this.state.hidedigisign && this.state.DigitalsignID != null && (
                        <span
                          onClick={() => this.updatedigisignnew()}
                          style={{ cursor: "pointer" }}
                        >
                          <div className="" title='Sync digital signed document from Signing Hub'>
                            <img
                              style={{ cursor: 'pointer' }}
                              className='mt-0'
                              src={require("../assets/digisign.png")}
                              alt="Signature Icon"
                            />
                          </div>
                        </span>
                      )}
                      {this.state.TemplateDoc && this.state.TemplateDoc.length > 0 && (
                        <span
                          onClick={() => this.OpenFileTemplate(this.state.TemplateDoc[0], "Open")}
                          style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                        >
                          <div className="btn btn-primary">
                            <img style={{ cursor: 'pointer' }} className='mt-0'
                              src={this.state.showdigisign ? require("../assets/signicon.png") : require("../assets/noun-download-5006210.png")}
                            ></img></div>
                        </span>
                      )}
                    </div>
                  </div>

                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">

                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editncType || ""}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <Dropdown
                          required
                          placeholder="Category"
                          label="Category:"
                          disabled={this.state.isDisabled}
                          options={this.state.edittypeoptions}
                          selectedKey={this.state.editncType}
                          onChange={this.onChangenctype}
                          className={this.state.editErrors?.editnctype ? 'dropdown-error' : ''}
                        />
                      </TooltipHost>
                    </div>
                    {console.log("this.state.ApprovedAuditSelected", this.state.ApprovedAuditSelected)}
                    <div className="form-group col-md-4 mb-3">
                      <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }}>Approved Report Code:<span className="text-danger1">*</span>
                      </label>
                      <TooltipHost
                        content={this.state.editmemonumberOptions.length > 0 && this.state.editApprovedAuditReport && this.state.editmemonumberOptions.filter((item: any) => item?.value == this.state.editApprovedAuditReport)[0]?.label || ""}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >

                        {/* {this.state.ApprovedAuditSelected.length ==0 ?
                          <TextField
                            value={this.state.editMemoNumber + ""} disabled={true} /> : */}
                        <Select
                          options={this.state.editmemonumberOptions}
                          value={this.state.ApprovedAuditSelected}
                          isDisabled={this.state.isDisabled}
                          name="Auditreportcode"
                          isClearable={true}
                          isSearchable={true}
                          className={this.state.editErrors?.editApprovedAuditReport ? 'border-on-error' : ''}
                          onChange={(selectedOption: any) => this.changeMemoNumber(selectedOption)}
                          placeholder={"Approved Report Code"}

                        />
                        {/* } */}
                        {/* <Dropdown
                          disabled={this.state.isDisabled}
                          required
                          placeholder="Approved Report Code"
                          label="Approved Report Code:"
                          options={this.state.editmemonumberOptions}
                          defaultSelectedKey={this.state.editApprovedAuditReport}
                          selectedKey={this.state.editApprovedAuditReport}
                          onChange={this.changeMemoNumber}
                          className={this.state.editErrors?.editApprovedAuditReport ? 'dropdown-error' : ''}
                        /> */}
                      </TooltipHost>
                    </div>
                    {console.log("editNCNumberOptions", this.state.editNCNumberOptions, this.state.editNCNumberID)}
                    <div className="form-group col-md-4 mb-3">
                      <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }}>NC/Observation Number:<span className="text-danger1">*</span>
                      </label>
                      <TooltipHost
                        content={this.state.editNCNumberOptions.length > 0 && this.state.editNCNumberOptions.filter((item: any) => item?.value == this.state.editNCNumberID)[0]?.label || ""}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >

                        <Select
                          options={this.state.editNCNumberOptions}
                          value={this.state.NCNumberselected}
                          isDisabled={this.state.isDisabled}
                          name="ncnumber"
                          isClearable={true}
                          isSearchable={true}
                          //disabled={this.state.isDisabled}
                          className={this.state.editErrors?.editNCNumber ? 'border-on-error' : ''}
                          onChange={(selectedOption: any) => this.changeNCNumber(selectedOption)}
                          placeholder={"NC/Observation Number"}

                        />

                        {/* <Dropdown
                          disabled={this.state.isDisabled}
                          required
                    placeholder="NC/Observation Number"
                          label="NC/Observation Number:"
                          options={this.state.editNCNumberOptions}
                          defaultSelectedKey={this.state.editNCNumberID}
                          selectedKey={this.state.editNCNumberID}
                          onChange={this.changeNCNumber}
                          className={this.state.editErrors?.editNCNumber ? 'dropdown-error' : ''}
                        /> */}
                      </TooltipHost>
                    </div>
                    {/* <div className="form-group col-md-4">
                    <TooltipHost
                      content={this.state.editNCRNo}
                      calloutProps={{ gapSpace: 0 }}
                      styles={{ root: { display: 'inline-block', width: '100%' } }}
                    >
                      <TextField label="NCR No:" name='editncrNo' required value={this.state.editNCRNo} disabled={true} onChange={this.handleChange}

                      />
                    </TooltipHost>
                  </div> */}

                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 mb-3">
                      <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }} >From Department:<span className="text-danger1">*</span>
                      </label>
                      <TooltipHost
                        content={this.state.editDepartmentOption.filter((item: any) => item.value == this.state.editfromdepartment)[0]?.label || "Select Department"}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <Select
                          options={this.state.editDepartmentOption}
                          value={this.state.editfromdepartmentselected}
                          name="Department"
                          isClearable={true}
                          isSearchable={true}
                          isDisabled
                          disabled={this.state.isDisabled}
                          className={this.state.editErrors?.department ? 'border-on-error' : ''}
                          onChange={(selectedOption: any) => this.changeDepartment(selectedOption)}
                          placeholder={"Department"}

                        />
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editDocumentCode}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Document Code:" name='editDocumentCode' required value={this.state.editDocumentCode} disabled={true} onChange={this.handleChange}

                        />
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editIssueNo}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Issue Number:" name='editIssueNo' required value={this.state.editIssueNo + ""} disabled={true} onChange={this.handleChange}

                        />
                      </TooltipHost>
                    </div>

                  </div>

                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editRevisionNo == 0 || this.state.editRevisionNo == "0" ? "0" : this.state.editRevisionNo}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Revision Number:" name='editRevisionNo' required value={this.state.editRevisionNo + ""} disabled={true} onChange={this.handleChange}

                        /></TooltipHost>
                    </div>

                    <div className="form-group col-md-4 mb-3">
                      <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }}>Department:<span className="text-danger1">*</span>
                      </label>
                      <TooltipHost
                        content={this.state.editDepartmentOption.filter((item: any) => item?.value == this.state.editDepartment)[0]?.label || ""}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <Select
                          options={this.state.editDepartmentOption}
                          value={this.state.departmentselected}
                          // isDisabled={this.state.isDisabled}
                          name="Department"
                          isDisabled
                          //disabled={this.state.isDisabled}
                          isClearable={true}
                          isSearchable={true}
                          className={this.state.editErrors?.editdepartment ? 'border-on-error' : ''}
                          onChange={(selectedOption: any) => this.changeDepartment(selectedOption)}
                          placeholder={"Department"}

                        />
                        {/* <Dropdown
                          required
                          disabled={this.state.isDisabled}
                          label="Department:"
                          options={this.state.editDepartmentOption}
                          defaultSelectedKey={this.state.editDepartment}
                          selectedKey={this.state.editDepartment}
                          onChange={this.changeDepartment}
                          className={this.state.editErrors?.editdepartment ? 'dropdown-error' : ''}

                        /> */}
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editCriteria}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Criteria:" disabled={this.state.isDisabled} name='editCriteria'
                          required value={this.state.editCriteria} onChange={this.handleChange}
                          className={this.state.editErrors?.editCriteria ? 'textfield-error' : ''}
                        // styles={{
                        //   fieldGroup: {
                        //     backgroundColor: this.state.editErrors.editCriteria ? "#ffcccb" : "white",
                        //   }
                        // }}
                        />
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editCloseOutStatus}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Close Out Status:" disabled={true} name='editCloseOutStatus' required value={this.state.editCloseOutStatus} onChange={this.handleChange}
                          className={this.state.editErrors?.editCloseOutStatus ? 'textfield-error' : ''}
                        // styles={{
                        //   fieldGroup: {
                        //     backgroundColor: this.state.editErrors.editCloseOutStatus ? "#ffcccb" : "white",
                        //   }
                        // }}
                        />
                      </TooltipHost>
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4" id="categoryCheckbox">
                      <label>Category: <span className={styles.textdanger}>*</span></label>
                      {this.state.editCategoryCheckOption.map((item: any) => {
                        return (
                          <div style={{ margin: "2px", padding: "3px" }}>
                            <Checkbox label={item.text} disabled={this.state.isDisabled} checked={this.state.editCategoryValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("editCategoryValueIsCheck", item.key as number, item.text)} />
                          </div>
                        )
                      }
                      )}
                      {this.state.showcategoryothers &&
                        <div style={{ margin: "2px", padding: "3px" }}>
                          <TooltipHost
                            content={this.state.CategoryOthers || ""}
                            calloutProps={{ gapSpace: 0 }}
                            styles={{ root: { display: 'inline-block', width: '100%' } }}
                          >
                            <TextField disabled={this.state.isDisabled} label="Category Others:" name='CategoryOthers' value={this.state.CategoryOthers} onChange={this.handleChangeCategoryOthers}
                            //className={this.state.errors?.categoryothers ? 'textfield-error' : ''}
                            /></TooltipHost>

                        </div>

                      }
                    </div>
                    <div className="form-group col-md-4" id="SubCategoryCheckbox">
                      <label>Sub Category: <span className={styles.textdanger}>*</span></label>
                      {this.state.editSubCategoryCheckOption.map((item: any) => {
                        return (
                          <div style={{ margin: "2px", padding: "3px" }}>
                            <Checkbox label={item.text} disabled={this.state.isDisabled} checked={this.state.editSubCategoryValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("editSubCategoryValueIsCheck", item.key as number, item.text)} />
                          </div>
                        )
                      }
                      )}
                      {this.state.showsubcategoryothers &&
                        <div style={{ margin: "2px", padding: "3px" }}>
                          <TooltipHost
                            content={this.state.SubCategoryOthers || ""}
                            calloutProps={{ gapSpace: 0 }}
                            styles={{ root: { display: 'inline-block', width: '100%' } }}
                          >
                            <TextField disabled={this.state.isDisabled} label="SubCategory Others:" name='SubCategoryOthers' value={this.state.SubCategoryOthers} onChange={this.handleChangeSubCategoryOthers}
                            //className={this.state.errors?.subCategoryothers ? 'textfield-error' : ''} 
                            /></TooltipHost>

                        </div>
                      }
                    </div>
                    <div className="form-group col-md-4" id="locationCheckbox">
                      <label>Location: <span className={styles.textdanger}>*</span></label>
                      {this.state.editLocationCheckOption.map((item: any) => {
                        return (
                          <div style={{ margin: "2px", padding: "3px" }}>
                            <Checkbox label={item.text} disabled={this.state.isDisabled} checked={this.state.editLocationValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("editLocationValueIsCheck", item.key as number, item.text)}
                            />
                          </div>
                        )
                      }
                      )}
                      {this.state.showlocationothers &&
                        <div style={{ margin: "2px", padding: "3px" }}>
                          <TooltipHost
                            content={this.state.LocationOthers || ""}
                            calloutProps={{ gapSpace: 0 }}
                            styles={{ root: { display: 'inline-block', width: '100%' } }}
                          >
                            <TextField disabled={this.state.isDisabled} label="Location Others:" name='LocationOthers' value={this.state.LocationOthers} onChange={this.handleChangeLocationOthers}
                            //className={this.state.errors?.locationthers ? 'textfield-error' : ''} 
                            /></TooltipHost>

                        </div>
                      }
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 " id='AssigntoPeoplepicker'>
                      {/* <TooltipHost
                        content={this.state.editAssignTo}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <PeoplePicker
                          context={peoplePickerContext}
                          titleText="Auditee:"
                          personSelectionLimit={1}
                          required={true}
                          groupName={""} // Leave this blank in case you want to filter from all users
                          showtooltip={true}
                          disabled={this.state.isDisabled}
                          ensureUser={true}
                          defaultSelectedUsers={this.state.editAssignTo ? [this.state.editAssignTo] : []}
                          onChange={this._handlePeoplePickerChange("editAssignTo", "editAssignToId")}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          styles={{
                            root: {
                              backgroundColor: this.state.editErrors.editAssignTo ? "#ffcccb" : "white",
                            }
                          }}
                        />
                      </TooltipHost> */}
                      <label htmlFor="revisionNo">Auditee<span className="text-danger1"> *</span></label>
                      <div title={this.state.Auditeeuser.map((user: any) => user.label).join(', ')}>
                        <Select
                          //onKeyDown={handleKeyDown}
                          isClearable={true}
                          options={this.state.allusersoption}
                          isMulti
                          value={this.state.Auditeeuser}
                          name="Auditee"
                          //className={`newse ${(!ValidSubmit && sharewitherr) ? "border-on-error" : ""}`}
                          // onChange={(selectedOption: any) => onSelect(selectedOption)}
                          onChange={(selectedOptions: any) => this.onSelectAuditee(selectedOptions, "editAssignTo", "editAssignToId")}
                          placeholder="Enter Prepared By"
                          isDisabled={this.state.isDisabled}
                        /></div>
                    </div>

                    <div className="form-group col-md-4">
                      <Label>
                        Due Date <span className={styles.textdanger}>*</span>
                      </Label>
                      <div
                      // title={
                      //   this.state.editDueDate
                      //     ? moment(new Date(this.state.editDueDate)).format('DD/MMM/YYYY')
                      //     : "Select a request date"
                      // }
                      >
                        <TooltipHost
                          content={moment(new Date(this.state.editDueDate)).format('DD/MMM/YYYY') || "Select a request date"}
                          calloutProps={{ gapSpace: 0 }}
                          styles={{ root: { display: 'inline-block', width: '100%' } }}
                        >
                          <DatePicker
                            disabled={this.state.isDisabled}
                            formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                            placeholder="Select a Due Date"
                            value={this.state.editDueDate}
                            onSelectDate={(date: Date) => this.setState({ editDueDate: date })}
                            //styles={this.state.editErrors.editDueDate ? datePickerErrorStyles : {}}
                            className={this.state.editErrors?.editDueDate ? 'textfield-error' : ''}
                          />
                        </TooltipHost>
                      </div>
                    </div>
                    <div style={{ position: 'relative' }} className="col-lg-4 mt-1">
                      <label htmlFor="Attchments" style={{ marginRight: "10px" }}>Attachments </label>

                      <input disabled={this.state.isDisabled} className="form-control" type="file" name="myFile" onChange={(e) => this.handleFileChange(e, this)} id="newfile" multiple
                        style={{
                          backgroundColor: this.state.editErrors.Attchments ? "#ffe6e6" : "white",
                          borderColor: this.state.editErrors.Attchments ? '1px red' : '1px solid #dee2e6'
                        }} />
                      {(this.state.copyFil.length + this.state.exFiles.length) > 0 && (
                        <span style={{ fontSize: '0.875rem' }} onClick={this._OpenModal} className='newpo'>
                          <FontAwesomeIcon icon={faPaperclip} />{" "}
                          {(this.state.copyFil.length + this.state.exFiles.length)}{" "}
                          {(this.state.copyFil.length + this.state.exFiles.length) === 1 ? "file" : "files"} Attached
                        </span>
                      )}
                      {this.state.showDialog && <div id="myModal" className={styles.modal}>
                        <div className={`${styles.modalcontent} ${Showfile ? styles.wideModal : ''}`}>
                          <span className={styles.close} onClick={() => this._CloseModal()}>&times;</span>
                          <h4 className="font-16 text-dark fw-bold mb-1">Attachment Details</h4>
                          <p className="text-muted font-14 mb-3 fw-400">Below are the attachment details for Non Conformity / Observation</p>

                          {/*} <table className='mtbalenew'>
                            <thead>
                              <tr>
                                <th style={{ minWidth: '50px', maxWidth: '50px' }} >S.No.</th>
                                <th style={{ minWidth: '150px', maxWidth: '150px' }}>File Name</th>

                                <th style={{ minWidth: '100px', maxWidth: '100px' }} className="text-center">Upload Date</th>
                                {/* {this.state.exFiles.length > 0 && <th className="text-center">File Link</th>} */}
                          {/*} {(this.state.ShowDeleteicon || this.state.copyFil.length > 0 || this.state.exFiles.length > 0) &&
                                  <th className="text-center" style={{ minWidth: "60px", maxWidth: "60px" }}>Action</th>
                                }
                              </tr>
                            </thead>
                            {upFiles}{fileData}
                          </table>*/}
                          {Showfile ?

                            <FileViewer showfile={Showfile} docurl={this.state.redirecturl} cancelAction={this.cancelModalAction} />

                            :
                            <table className="mtbalenew">
                              <thead>
                                <tr>
                                  <th style={{ minWidth: '50px', maxWidth: '50px' }} className="text-center">S.No.</th>
                                  <th style={{ minWidth: '150px', maxWidth: '150px' }}>File Name</th>
                                  <th style={{ minWidth: '80px', maxWidth: '80px' }} className="text-center">Upload Date</th>
                                  {(this.state.ShowDeleteiconInitiator || this.state.copyFil.length > 0 || this.state.exFiles.length > 0) && (
                                    <th className="text-center" style={{ minWidth: '60px', maxWidth: '60px' }}>Action</th>
                                  )}
                                </tr>
                              </thead>
                              <tbody>
                                {[...this.state.exFiles, ...this.state.copyFil].map((item: any, index: number) => {
                                  const isExistingFile = index < this.state.exFiles.length;
                                  const serial = index + 1;

                                  // const fileName = isExistingFile
                                  //   ? (item.Name.includes('_') ? item.Name.split('_')[2] : item.Name)
                                  //   : decodeURIComponent(item.name);
                                  const fileName = isExistingFile
                                    ? (() => {
                                      const match = item.Name.match(/^\d+_(.*?)_\d{17}\.[^.]+$/);
                                      return match ? `${match[1]}.${item.Name.split('.').pop()}` : item.Name;
                                    })()
                                    : decodeURIComponent(item.name);
                                  const uploadDate = isExistingFile
                                    ? moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")
                                    : new Date().toLocaleDateString("en-GB", {
                                      day: "2-digit",
                                      month: "short",
                                      year: "numeric"
                                    }).replace(/ /g, "/");

                                  return (
                                    <tr key={index} style={{ display: 'table', width: '100%' }}>
                                      <td style={{ minWidth: '50px', maxWidth: '50px' }} className="text-center">{serial}</td>
                                      <td style={{ minWidth: '150px', maxWidth: '150px' }} title={fileName}>{fileName}</td>
                                      <td style={{ minWidth: '80px', maxWidth: '80px' }} className="text-center">{uploadDate}</td>
                                      <td style={{ minWidth: '60px', maxWidth: '60px' }} className="text-center">
                                        {isExistingFile ? (
                                          <>
                                            <span
                                              onClick={() => this.OpenFile(item, "Open")}
                                              style={{ color: "blue", cursor: "pointer", margin: "0 5px" }}
                                            >
                                              <FontAwesomeIcon title="Preview file" icon={faEye} />
                                            </span>
                                            <span
                                              onClick={() => this.OpenFile(item, "Download")}
                                              style={{ color: "blue", cursor: "pointer", margin: "0 5px" }}
                                            >
                                              <FontAwesomeIcon title="Download file" icon={faDownload} />
                                            </span>
                                            {this.state.ShowDeleteiconInitiator && (
                                              <img
                                                src={require("../assets/del.png")}
                                                onClick={() => this.toBeDeleted(index)}
                                                title="Delete file"
                                                style={{ cursor: "pointer", marginLeft: "5px" }}
                                              />
                                            )}
                                          </>
                                        ) : (
                                          <img
                                            src={require("../assets/del.png")}
                                            onClick={() => this.removeFiles(index - this.state.exFiles.length)}
                                            title="Remove new file"
                                            style={{ cursor: "pointer" }}
                                          />
                                        )}
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          }

                        </div>
                      </div>}
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row mb-3'>
                    <div className="form-group col-md-12">
                      <TooltipHost
                        content={this.state.editProblemDescription}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="NC/Observation Description:"
                          required
                          name='problemDescription'
                          value={this.state.editProblemDescription}
                          multiline rows={5}
                          onChange={this.handleChangeobsdescription}
                          disabled={this.state.isDisabled}
                          //errorMessage={this.state.errors.problemDescription}
                          className={this.state.editErrors?.editProblemDescription ? 'textfield-error' : ''}
                        ></TextField>

                      </TooltipHost>
                    </div>
                  </div>
                  {/* End Save as draft */}
                </form>
              </fieldset>
            </section>
          }
          {this.state.editSubmitStatus == "Yes" && (!this.state.Loading || !setloading) && this.state.editCurrentUserRole != "FirstInitiator" ? //this.state.editFirstInitiatorSubmitStatus == "Yes" || this.state.editFirstAssignedToSubmitStatus == "Yes" ?
            <section className='card card-body mt-2' >
              <fieldset>
                <form>
                  {/* Section 2 */}
                  <div className="row">
                    <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-3'>To be filled by Auditee/Department Head</h3></div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4" id='AssigntoPeoplepicker'>
                      <TooltipHost
                        content={this.state.editAssignTo}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <PeoplePicker
                          context={peoplePickerContext}
                          //disabled={this.state.showDelegate}
                          disabled={true}
                          titleText="Auditee:"
                          personSelectionLimit={1}
                          required={true}
                          onChange={this._handlePeoplePickerChange("editPersonAssigned", "editPersonAssignedId")}
                          defaultSelectedUsers={[this.state.editPersonAssigned]}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          ensureUser={true}
                          styles={{
                            root: {
                              backgroundColor: this.state.editErrors.editPersonAssigned ? "#ffe6e6" : "white",
                            }
                          }}
                        />
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4">
                      <Label>
                        Date <span className={styles.textdanger}>*</span>
                      </Label>
                      <TooltipHost
                        content={this.state?.editDate && !isNaN(Date.parse(this.state?.editDate)) ? new Date(this.state?.editDate).toLocaleString() : null}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <DatePicker
                          disabled={this.state.deptSectionDisable}
                          formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                          placeholder="Select a Date"
                          value={this.state.editDate}
                          onSelectDate={(date: Date) => this.setState({ editDate: date })}
                          className={this.state.editErrors?.editDate ? 'textfield-error' : ''}
                        //styles={this.state.editErrors.editDate ? datePickerErrorStyles : {}}
                        /></TooltipHost>
                    </div>
                    <div className="form-group col-md-4">
                      <Label>
                        Deadline for completion <span className={styles.textdanger}>*</span>
                      </Label>
                      <TooltipHost
                        content={this.state?.editDeadlineCompletion && !isNaN(Date.parse(this.state?.editDeadlineCompletion)) ? moment(new Date(this.state.editDueDate)).format('DD/MMM/YYYY') : null}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <DatePicker
                          formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                          disabled={this.state.deptSectionDisable}
                          placeholder="Select a Deadline"
                          value={this.state.editDeadlineCompletion}
                          onSelectDate={(date: Date) => this.setState({ editDeadlineCompletion: date })}
                          className={this.state.editErrors?.editDeadlineCompletion ? 'textfield-error' : ''}
                        //styles={this.state.editErrors.editDeadlineCompletion ? datePickerErrorStyles : {}}
                        /></TooltipHost>
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-12">
                      <TooltipHost
                        content={this.state.editCorrection}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Correction(Immediate Steps to stop the problem):" required name='editCorrection'
                          value={this.state.editCorrection}
                          multiline rows={3}
                          disabled={this.state.deptSectionDisable}
                          onChange={this.handleChange}
                          className={this.state.editErrors?.editCorrection ? 'textfield-error' : ''}
                        // styles={{
                        //   fieldGroup: {
                        //     backgroundColor: this.state.editErrors.editCorrection ? "#ffcccb" : "white", // Red tint for errors
                        //   }
                        // }}
                        />
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-12">
                      <TooltipHost
                        content={this.state.editRootCause}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Root Cause:" required name='editRootCause'
                          value={this.state.editRootCause} multiline rows={3}
                          onChange={this.handleChange}
                          disabled={this.state.deptSectionDisable}
                          className={this.state.editErrors?.editRootCause ? 'textfield-error' : ''}
                        // styles={{
                        //   fieldGroup: {
                        //     backgroundColor: this.state.editErrors.editRootCause ? "#ffcccb" : "white", // Red tint for errors
                        //   }
                        // }}
                        />
                      </TooltipHost>
                    </div>
                  </div>

                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 mb-3" id='delegatetopeoplepicker'>
                      {/* <TooltipHost
                        content={this.state.editDelegateTo}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <PeoplePicker
                          disabled={this.state.showDelegate}
                          context={peoplePickerContext}
                          titleText="Delegate To:"
                          personSelectionLimit={1}
                          required={false}
                          onChange={this._handlePeoplePickerChange("editDelegateTo", "editDelegateToId")}
                          defaultSelectedUsers={[this.state.editDelegateTo]}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          ensureUser={true}
                        />
                      </TooltipHost> */}
                      <label htmlFor="revisionNo">Delegate To:</label>
                      <div title={this.state.DelegatetoUser.map((user: any) => user.label).join(', ')}>
                        <Select
                          //onKeyDown={handleKeyDown}
                          isClearable={true}
                          options={this.state.allusersoption}
                          
                          value={this.state.DelegatetoUser}
                          name="Delegate To"
                          //className={`newse ${(!ValidSubmit && sharewitherr) ? "border-on-error" : ""}`}
                          // onChange={(selectedOption: any) => onSelect(selectedOption)}
                          onChange={(selectedOptions: any) => this.onSelectDelegateto(selectedOptions, "editDelegateTo", "editDelegateToId")}
                          placeholder="Delegate To"
                          isDisabled={this.state.showDelegate}
                        /></div>
                    </div>
                    <div className="form-group col-md-4" id='analyzedBypeoplepicker'>
                      {/* <TooltipHost
                        content={this.state.editAnalyzedBy}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <PeoplePicker
                          disabled={this.state.deptSectionDisable}
                          context={peoplePickerContext}
                          titleText="Analyzed and Reviewed By:"
                          personSelectionLimit={1}
                          required={true}
                          onChange={this._handlePeoplePickerChange("editAnalyzedBy", "editAnalyzedById")}
                          defaultSelectedUsers={[this.state.editAnalyzedBy]}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          ensureUser={true}
                          styles={{
                            root: {
                              backgroundColor: this.state.editErrors.editAnalyzedBy ? "#ffe6e6" : "white",
                            }
                          }}
                        />
                      </TooltipHost> */}
                      <label htmlFor="revisionNo">Analyzed By:</label>
                      <div title={this.state.Analyzedbyuser.map((user: any) => user.label).join(', ')}>
                        <Select
                          //onKeyDown={handleKeyDown}
                          isClearable={true}
                          options={this.state.allusersoption}
                          
                          value={this.state.Analyzedbyuser}
                          name="Analyzed by"
                          //className={`newse ${(!ValidSubmit && sharewitherr) ? "border-on-error" : ""}`}
                          // onChange={(selectedOption: any) => onSelect(selectedOption)}
                          onChange={(selectedOptions: any) => this.onSelectAnalyzedby(selectedOptions, "editAnalyzedBy", "editAnalyzedById")}
                          placeholder="Analyzed By"
                          isDisabled={this.state.showDelegate}
                        /></div>
                    </div>
                    {/* <div className="form-group col-md-4" id='reviewedBypeoplepicker'>
                      <TooltipHost
                        content={this.state.editReviewedBy}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <PeoplePicker
                          context={peoplePickerContext}
                          disabled={this.state.deptSectionDisable}
                          titleText="Reviewed By:"
                          personSelectionLimit={1}
                          required={true}
                          onChange={this._handlePeoplePickerChange("editReviewedBy", "editReviewedById")}
                          defaultSelectedUsers={[this.state.editReviewedBy]}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          ensureUser={true}
                          styles={{
                            root: {
                              backgroundColor: this.state.editErrors.editReviewedBy ? "#ffe6e6" : "white",
                            }
                          }}
                        />
                      </TooltipHost>
                    </div> */}
                    <div style={{ position: 'relative' }} className="col-lg-4 mt-1">
                      <label htmlFor="Attchments" style={{ marginRight: "10px" }}>Auditee Attachments </label>

                      <input
                        disabled={this.state.deptSectionDisable}
                        className="form-control"
                        type="file" name="myFile" onChange={(e) => this.handleAuditeeFileChange(e, this)} id="newfile" multiple
                      //className={`form-control ${this.state.errors?.Attachments} ? 'textfield-error' : ''`}
                      />
                      {/* {this.state.fileCountauditee > 0 ?
                        (<span style={{ fontSize: '0.875rem' }} onClick={this._OpenModalauditee} className='newpo'>
                          <FontAwesomeIcon icon={faPaperclip} /> {this.state.fileCountauditee} {this.state.fileCountauditee > 0 ? "files" : "file"} Attached
                        </span>) : ""
                      } */}
                      {(this.state.copyFilauditee.length + this.state.exFilesauditee.length) > 0 && (
                        <span onClick={this._OpenModalauditee} className='newpo'>
                          <FontAwesomeIcon icon={faPaperclip} />{" "}
                          {(this.state.copyFilauditee.length + this.state.exFilesauditee.length)}{" "}
                          {(this.state.copyFilauditee.length + this.state.exFilesauditee.length) === 1 ? "file" : "files"} Attached
                        </span>
                      )}
                      {this.state.showDialogauditee && (
                        <div id="myModal" className={styles.modal}>
                          <div className={`${styles.modalcontent} ${Showfile ? styles.wideModal : ''}`}>
                            {/* Close button */}
                            <span className={styles.close} onClick={() => this._CloseModalauditee()}>&times;</span>

                            {/* Modal title and subtitle */}
                            <h4 className="font-16 text-dark fw-bold mb-1">Attachment Details</h4>
                            <p className="text-muted font-14 mb-3 fw-400">Below are the attachment details for Non Conformity / Observation</p>

                            {/* Table */}
                            {/*} <table className={styles.mtbalenew}>
                              <thead>
                                <tr>
                                  <th style={{ minWidth: '60px', maxWidth: '60px' }}>S.No.</th>
                                  <th style={{ minWidth: '150px', maxWidth: '150px' }}>File Name</th>

                                  <th style={{ minWidth: '100px', maxWidth: '100px' }} className="text-center">Upload Date</th>
                                  {/* {this.state.exFilesauditee.length > 0 && <th className="text-center">File Link</th>} */}
                            {/*} {(this.state.ShowDeleteicon || this.state.copyFilauditee.length > 0 || this.state.exFilesauditee.length > 0) &&
                                    <th style={{ minWidth: '60px', maxWidth: '60px' }} className="text-center">Action</th>}
                                </tr>
                              </thead>
                              <tbody >
                                {upFilesauditee}
                                {fileDataAuditee}
                              </tbody>*/}
                            {/*</table>*/}
                            {Showfile ?
                              <FileViewer showfile={Showfile} docurl={this.state.redirecturl} cancelAction={this.cancelModalAction} />
                              :
                              <table className={styles.mtbalenew} style={{ width: '100%', borderCollapse: 'collapse' }}>
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: '60px', maxWidth: '60px' }}>S.No.</th>
                                    <th style={{ minWidth: '150px', maxWidth: '150px' }}>File Name</th>
                                    <th style={{ minWidth: '100px', maxWidth: '100px' }} className="text-center">Upload Date</th>
                                    {(this.state.ShowDeleteicon || this.state.copyFilauditee.length > 0 || this.state.exFilesauditee.length > 0) && (
                                      <th style={{ minWidth: '100px', maxWidth: '100px' }} className="text-center">Action</th>
                                    )}
                                  </tr>
                                </thead>
                                <tbody>
                                  {[...this.state.exFilesauditee, ...this.state.copyFilauditee].map((item: any, index: number) => {
                                    const isExisting = index < this.state.exFilesauditee.length;
                                    const serial = index + 1;
                                    console.log("isExisting", isExisting, this.state.ShowDeleteicon, index)
                                    // const fileName = isExisting
                                    //   ? (item.Name.includes('_') ? item.Name.split('_')[2] : item.Name)
                                    //   : decodeURIComponent(item.name);
                                    const fileName = isExisting
                                      ? (() => {
                                        const match = item.Name.match(/^\d+_(.*?)_\d{17}\.[^.]+$/);
                                        return match ? `${match[1]}.${item.Name.split('.').pop()}` : item.Name;
                                      })()
                                      : decodeURIComponent(item.name);
                                    const uploadDate = isExisting
                                      ? moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")
                                      : new Date().toLocaleDateString("en-GB", {
                                        day: "2-digit",
                                        month: "short",
                                        year: "numeric"
                                      }).replace(/ /g, "/");

                                    return (
                                      <tr key={index}>
                                        <td style={{ textAlign: 'center' }}>{serial}</td>
                                        <td title={fileName}>{fileName}</td>
                                        <td className="text-center">{uploadDate}</td>
                                        <td className="text-center">
                                          {isExisting ? (
                                            <>
                                              <span
                                                onClick={() => this.OpenFile(item, "Open")}
                                                style={{ color: "blue", cursor: "pointer", margin: "0 5px" }}
                                              >
                                                <FontAwesomeIcon title='Preview file' icon={faEye} />
                                              </span>
                                              <span
                                                onClick={() => this.OpenFile(item, "Download")}
                                                style={{ color: "blue", cursor: "pointer", margin: "0 5px" }}
                                              >
                                                <FontAwesomeIcon title='Download file' icon={faDownload} />
                                              </span>
                                              {this.state.ShowDeleteicon && (
                                                <img
                                                  src={require("../assets/del.png")}
                                                  onClick={() => this.toBeDeletedauditee(index)}
                                                  title="Delete existing file"
                                                  style={{ cursor: "pointer", marginLeft: "5px" }}
                                                />
                                              )}
                                            </>
                                          ) : (
                                            <img
                                              src={require("../assets/del.png")}
                                              onClick={() => this.removeFilesAuditee(index - this.state.exFilesauditee.length)}
                                              title="Remove new file"
                                              style={{ cursor: "pointer" }}
                                            />
                                          )}
                                        </td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            }
                          </div>
                        </div>
                      )}
                    </div>
                  </div>

                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-12">
                      <TooltipHost
                        content={this.state.editCorrectiveAction}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Corrective Action (Action to eliminate the root cause):"
                          required
                          name='editCorrectiveAction'
                          value={this.state.editCorrectiveAction} multiline rows={3}
                          onChange={this.handleChange}
                          disabled={this.state.deptSectionDisable}
                          className={this.state.editErrors?.editCorrectiveAction ? 'textfield-error' : ''}
                        // styles={{
                        //   fieldGroup: {
                        //     backgroundColor: this.state.editErrors.editCorrectiveAction ? "#ffcccb" : "white", // Red tint for errors
                        //   }
                        // }}
                        />
                      </TooltipHost>
                    </div>
                  </div>
                  {console.log("showimsupdatedshowimsupdated", showimsupdated, isdisableims)}
                  {(showimsupdated && isdisableims) &&
                    <div style={{ textAlign: 'left', display: 'flex', justifyContent: 'space-between', gap: '2rem', marginTop: '1rem' }}>

                      {/* IMS Updated */}
                      <div style={{ flex: 1 }}>
                        <div style={{ marginBottom: '0.3rem' }}>
                          <strong>IMS Updated</strong>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '2rem' }}>
                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.isIMSUpdated === 'Yes'}
                              onChange={() => this.handleChangecheckbox('Yes', "IMSUpdated")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            Yes
                          </label>

                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.isIMSUpdated === 'No'}
                              onChange={() => this.handleChangecheckbox('No', "IMSUpdated")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            No
                          </label>
                        </div>
                      </div>
                      {/* Risk & Opportunities Updated */}
                      <div style={{ flex: 1 }}>
                        <div style={{ marginBottom: '0.3rem' }}><strong>Risk & Opportunities Updated:</strong></div>

                        <div style={{ display: 'flex', alignItems: 'center', gap: '2rem' }}>
                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.riskandopportunitiesUpdated === 'Yes'}
                              onChange={() => this.handleChangecheckbox('Yes', "RiskOpportunities")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            Yes
                          </label>

                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.riskandopportunitiesUpdated === 'No'}
                              onChange={() => this.handleChangecheckbox('No', "RiskOpportunities")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            No
                          </label>
                        </div>
                      </div>

                    </div>
                  }
                  {(showimsupdated && !isdisableims) &&
                    <div style={{ textAlign: 'left', display: 'flex', justifyContent: 'space-between', gap: '2rem', marginTop: '1rem' }}>

                      {/* IMS Updated */}
                      <div style={{ flex: 1 }}>
                        <div style={{ marginBottom: '0.3rem' }}>
                          <strong>IMS Updated</strong>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '2rem' }}>
                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.isIMSUpdated === 'Yes'}
                              onChange={() => this.handleChangecheckbox('Yes', "IMSUpdated")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            Yes
                          </label>

                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.isIMSUpdated === 'No'}
                              onChange={() => this.handleChangecheckbox('No', "IMSUpdated")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            No
                          </label>
                        </div>
                      </div>
                      {/* Risk & Opportunities Updated */}
                      <div style={{ flex: 1 }}>
                        <div style={{ marginBottom: '0.3rem' }}><strong>Risk & Opportunities Updated:</strong></div>

                        <div style={{ display: 'flex', alignItems: 'center', gap: '2rem' }}>
                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.riskandopportunitiesUpdated === 'Yes'}
                              onChange={() => this.handleChangecheckbox('Yes', "RiskOpportunities")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            Yes
                          </label>

                          <label style={{ display: 'flex', alignItems: 'center' }}>
                            <input
                              disabled={isdisableims}
                              type="checkbox"
                              checked={this.state.riskandopportunitiesUpdated === 'No'}
                              onChange={() => this.handleChangecheckbox('No', "RiskOpportunities")}
                              style={{ marginRight: '0.4rem' }}
                            />
                            No
                          </label>
                        </div>
                      </div>

                    </div>
                  }

                  {(this.state.editCurrentUserRole == "LastInitiator" || this.state.editCurrentUserRole == "Approverrole") &&
                    this.state.editncType == "NC" &&
                    <><div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                      <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-0'>Non Conformity / Observation Close Out Details</h3></div>
                    </div><div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                        <div className="form-group col-md-12">
                          <TooltipHost
                            content={this.state.editCorrectiveActionImplementedOn}
                            calloutProps={{ gapSpace: 0 }}
                            styles={{ root: { display: 'inline-block', width: '100%' } }}
                          >
                            <TextField label="Corrective Action Implemented On:"
                              required
                              name='editCorrectiveActionImplementedOn'
                              value={this.state.editCorrectiveActionImplementedOn}
                              disabled={this.state.forwarDisable || forwardisdisabled}
                              multiline rows={3} onChange={this.handleChange}
                              className={this.state.editErrors?.editCorrectiveActionImplementedOn ? 'textfield-error' : ''} />
                          </TooltipHost>
                        </div>
                      </div></>
                  }
                  {console.log("isFinalApprover 5332", this.state.IsFinalapprover, IsAnalyzedBy, isdisablefinal, showcorrectionappicable)}
                  {showcorrectionappicable && this.state.editncType == "NC" &&
                    //!isdisablefinal &&
                    <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem', marginTop: '0.5rem' }}>
                      <label style={{ display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                        <input
                          disabled={isdisablefinal}
                          type="checkbox"
                          checked={this.state.correctionApplicable}
                          onChange={(e) =>
                            this.handleChangecheckboxsingle("correctionApplicable", e)
                          }
                          style={{ marginRight: '0.3rem' }}
                        />
                        Correction Applicable
                      </label>

                      <label style={{ display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                        <input
                          disabled={isdisablefinal}
                          type="checkbox"
                          checked={this.state.notEffective}
                          onChange={(e) =>
                            this.handleChangecheckboxsingle("notEffective", e)
                          }
                          style={{ marginRight: '0.3rem' }}
                        />
                        Not Effective
                      </label>

                      <label style={{ display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                        <input
                          disabled={isdisablefinal}
                          type="checkbox"
                          checked={this.state.effectiveClosed}
                          onChange={(e) =>
                            this.handleChangecheckboxsingle("effectiveClosed", e)
                          }
                          style={{ marginRight: '0.3rem' }}
                        />
                        Effective & Problem Closed
                      </label>
                    </div>

                  }
                  {/* Approve/Rework */}
                </form>
              </fieldset>
            </section> : null}
          {/* Approval Table */}
          {(((this.state.editReviewedBySubmitStatus == "Yes" && this.state.editDelegateToId == null && (!this.state.Loading || !setloading)) ||
            (this.state.editLastAssignedToSubmitStatus == "Yes" && this.state.editDelegateToId != null && (!this.state.Loading || !setloading))) && this.state.editncType == "NC") ?
            (<section className="card card-body mb-2">
              <fieldset disabled={this.state.forwarDisable}>
                <form>
                  <div className='row mt-2'>
                    <div className='col-sm-8'>
                      <h3 style={{ textAlign: 'left' }} className="header-title text-dark font-16 fw-bold mb-1 ">Forward Approval To</h3>
                      <label style={{ textAlign: 'left' }}>Define the approval hierarchy to ensure requests are routed to the appropriate approvers.
                      </label>
                    </div>
                    <div className='col-sm-4'>
                      <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right", paddingRight: "22px" }}>
                        {!this.state.forwarDisable && !forwardisdisabled &&
                          <img style={{ width: '30px', cursor: 'pointer' }} src={require("../assets/plus.png")}
                            onClick={this.addApprover} className='' />
                        }
                      </div>
                    </div>
                  </div>
                  {/* <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row'>
                  <div className="form-group col-md-12">
                  <h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-3'>Forward Detail</h3></div>
                </div> */}
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="">
                      {/* <button type="button" onClick={this.addApprover}>Add</button> */}

                      <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">

                        <table id="tblAppr" className='mtbalenew'>
                          <thead >
                            <tr>
                              <th style={{ minWidth: "40px", maxWidth: "40px" }}>S.No</th>
                              <th style={{ borderBottomLeftRadius: "0px", minWidth: "110px", maxWidth: "110px" }}>Role</th>
                              <th style={{ borderBottomLeftRadius: "0px", minWidth: '85px', maxWidth: '85px' }}>Responsibility<span className="text-danger1"> *</span></th>
                              <th style={{ minWidth: "46px", maxWidth: "46px" }} title="Use your electronic digital signature to sign this document digitally">E-Sign?</th>
                              <th style={{ minWidth: '70px', maxWidth: '70px' }} >Level</th>
                              <th style={{ minWidth: '100px', maxWidth: '100px' }}>Approver Name</th>
                              <th style={{ minWidth: '100px', maxWidth: '100px' }}>Approval Criteria</th>
                              <th style={{ minWidth: '70px', maxWidth: '70px' }}>Action</th>
                            </tr>
                          </thead>
                          <tbody>

                            {approval}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </form>
              </fieldset>
            </section>) : null}

          {/* {showreworkremarks && (!this.state.Loading || !setloading) &&
            <section style={{ justifyContent: 'left', textAlign: 'left' }} id="approvalSection" className='card card-body'>

              <TextField label="Remarks" required name="reworkremarks" value={this.state.reworkremarks} multiline rows={3} onChange={this.handleChangeReworkremarks}

                className={this.state.editErrors?.remarks ? 'textfield-error' : ''}// styles={{
              />
            </section>
          } */}
          {console.log("ghghghgh", this.state.editCurrentUserRole == "FirstInitiator", this.state.editStatus == "Rework", this.state.editFirstInitiatorSubmitStatus == "No",
            (!this.state.Loading || !setloading) && !this.state.showDraft, (this.state.edType !== "view" || (showimsupdated && !isdisableims) || showcorrectionappicable && !isdisablefinal)
          )}
          {(((this.state.showApprove === true || this.state.showReject === true || (showreworkremarks && this.state.editCurrentUserRole != null)
            //|| (showimsupdated && !isdisableims) 
            ||
            (showcorrectionappicable && !isdisablefinal)) && this.state.editCurrentUserRole != "FirstInitiator")
            ||
            ((this.state.editReviewedBySubmitStatus == "Yes" && this.state.editDelegateToId == null) ||
              (this.state.editLastAssignedToSubmitStatus == "Yes" && this.state.editDelegateToId != null)) ||
            (this.state.editCurrentUserRole == "FirstInitiator" && this.state.editStatus == "Rework" && this.state.editFirstInitiatorSubmitStatus == "No"))

            //&& (!this.state.showDraft) 
            && (this.state.edType !== "view"
              //|| (showimsupdated && !isdisableims) 
              || showcorrectionappicable && !isdisablefinal)
            && (!this.state.Loading || !setloading)
            ?
            <section style={{ justifyContent: 'left', textAlign: 'left' }} id="approvalSection" className='card card-body'>
              {this.state.edType !== "view" &&
                <TooltipHost
                  content={this.state.remarks}
                  calloutProps={{ gapSpace: 0 }}
                  styles={{ root: { display: 'inline-block', width: '100%' } }}
                >
                  <TextField label="Remarks"
                    //Approveclicked ||
                    required={(Rejectclicked || Reworkclicked || resubmitclicked || this.state.showApprove === true ||
                      this.state.showReject === true) ? true : false}
                    name="remarks"
                    value={this.state.editCurrentUserRole == "FirstInitiator" && this.state.editStatus == "Rework" && this.state.editFirstInitiatorSubmitStatus == "No" ? this.state.reworkremarks : this.state.remarks}
                    multiline rows={3}
                    onChange={this.handleChange}

                    className={this.state.editErrors?.remarks ? 'textfield-error' : ''}// styles={{
                  /></TooltipHost>
              }

              {console.log("isFinalApprover 5539", this.state.IsFinalapprover, IsAnalyzedBy, isdisablefinal, showcorrectionappicable)}
              {/* {showcorrectionappicable && !isdisablefinal &&
                <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem', marginTop: '0.5rem' }}>
                  <label style={{ display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                    <input
                      disabled={isdisablefinal}
                      type="checkbox"
                      checked={this.state.correctionApplicable}
                      onChange={(e) =>
                        this.handleChangecheckboxsingle("correctionApplicable", e)
                      }
                      style={{ marginRight: '0.3rem' }}
                    />
                    Correction Applicable
                  </label>

                  <label style={{ display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                    <input
                      disabled={isdisablefinal}
                      type="checkbox"
                      checked={this.state.notEffective}
                      onChange={(e) =>
                        this.handleChangecheckboxsingle("notEffective", e)
                      }
                      style={{ marginRight: '0.3rem' }}
                    />
                    Not Effective
                  </label>

                  <label style={{ display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                    <input
                      disabled={isdisablefinal}
                      type="checkbox"
                      checked={this.state.effectiveClosed}
                      onChange={(e) =>
                        this.handleChangecheckboxsingle("effectiveClosed", e)
                      }
                      style={{ marginRight: '0.3rem' }}
                    />
                    Effective & Problem Closed
                  </label>
                </div>

              } */}

            </section> : null
          }
          {/* Vishnu Changes  */}
          {console.log("bnbnbn", !this.state.Loading, !setloading, this.state.showDraft, RequesterEmail, CurrentuserEmail, isRequesterDelegated)}
          {this.state.showSubmit && (!this.state.Loading || !setloading) && ((this.state.editSubmitStatus == "No" && (RequesterEmail == CurrentuserEmail || isRequesterDelegated)) ||
            (this.state.editSubmitStatus == "Yes" && this.state.edType == "edit" && this.state.editStatus !== "Rework" && (this.state.editAssignToEmail == CurrentuserEmail || isAssignedtoDelegated) && this.state.editCurrentUserRole != "DelegateTo") ||
            (this.state.editSubmitStatus == "Yes" && this.state.edType == "edit" && this.state.editStatus !== "Rework" && (this.state.editDelegateToEmail == CurrentuserEmail || isDelegatedToDelegated) && this.state.editCurrentUserRole == "DelegateTo") ||
            (this.state.editStatus == "Rework" && (RequesterEmail == CurrentuserEmail || isRequesterDelegated) && this.state.editCurrentUserRole == "FirstInitiator") ||
            (this.state.edType == "edit" && this.state.editStatus == "Rework" && (this.state.editAssignToEmail == CurrentuserEmail || isAssignedtoDelegated) && this.state.editCurrentUserRole == "FirstAssignedTo") ||
            (this.state.edType == "edit" && this.state.editStatus == "Rework" && (this.state.editDelegateToEmail == CurrentuserEmail || isDelegatedToDelegated) && this.state.editCurrentUserRole == "DelegateTo")

          ) && (!this.state.Loading && !setloading) &&

            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }} className='newnbu'>

              {this.state.showDraft && (RequesterEmail == CurrentuserEmail || isRequesterDelegated) &&
                //this.state.editStatus != "Rework"&&

                // <PrimaryButton text="Save as Draft" onClick={() => this.handleDraft("draft")} />
                <div
                  role="button"
                  tabIndex={0}
                  //style={{ width: '145px' }}
                  className="btn btn-primary waves-effect waves-light m-1"
                  onClick={() => this.handleDraft("draft")}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      this.handleDraft("draft");
                    }
                  }}
                >
                  {/* <img
                    src={require('../../../Assets/ExtraImage/checkcircle.svg')}
                    style={{ width: '1rem' }}
                    className="me-1"
                    alt="Check"
                  /> */}
                  Save As Draft
                </div>

              }
              {(this.state.editCurrentUserRole === "FirstAssignedTo" || this.state.editCurrentUserRole === "DelegateTo") &&
                //this.state.editStatus != "Rework"&&

                // <PrimaryButton text="Save as Draft" onClick={() => this.handleDraft("draft")} />
                <div
                  role="button"
                  tabIndex={0}
                  //style={{ width: '145px' }}
                  className="btn btn-primary waves-effect waves-light m-1"
                  onClick={() => this.handleDraft("draft")}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      this.handleDraft("draft");
                    }
                  }}
                >
                  {/* <img
                    src={require('../../../Assets/ExtraImage/checkcircle.svg')}
                    style={{ width: '1rem' }}
                    className="me-1"
                    alt="Check"
                  /> */}
                  Save As Draft
                </div>

              }
              {/* <PrimaryButton text="Submit" onClick={() => this.handleSubmit("submit")} /> */}
              <div
                role="button"
                tabIndex={0}
                style={{ width: '105px' }}
                className="btn btn-primary waves-effect waves-light m-1"
                onClick={() => this.handleSubmit("submit")}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.handleSubmit("submit");
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
              {(this.state.editCurrentUserRole === "FirstAssignedTo" || this.state.editCurrentUserRole === "DelegateTo") &&
                // <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} />
                <a>
                  <div
                    role="button"
                    tabIndex={0}
                    style={{ width: '105px' }}
                    className="btn btn-warning waves-effect waves-light m-1"
                    onClick={(e) => this._reworkRequest("Rework")}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        this._reworkRequest("Rework");
                      }
                    }}
                  >
                    <i className="fe-corner-up-left me-1"></i> Rework
                  </div>
                </a>
              }
              {/* <DefaultButton text="Cancel" onClick={() => this.cancelRequest("Edcmain")} /> */}

              <div
                role="button"
                tabIndex={0}
                style={{ width: '105px' }}
                className="btn cancel-btn waves-effect waves-light m-1"
                onClick={() => this.cancelRequest("Edcmain")}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.cancelRequest("Edcmain");
                  }
                }}
              >
                <i className="fe-x me-1"></i> Cancel
              </div>


            </div>
          }
          {console.log("ApproverEmailApproverEmail", ApproverEmail, CurrentuserEmail)}
          {this.state.showApprove && (!this.state.Loading && !setloading) &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>
              {(ApproverEmail == CurrentuserEmail || isApproverDelegated) &&
                <>
                  {/* <PrimaryButton text="Approve" onClick={() => this._approveRequest("Approve")} />
                  <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} /> */}
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-success waves-effect waves-light m-1"
                      onClick={(e) => this._approveRequest("Approve")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this._approveRequest("Approve");
                        }
                      }}
                    >
                      <i className="fe-check-circle me-1"></i> Approve
                    </div>
                  </a>
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-warning waves-effect waves-light m-1"
                      onClick={(e) => this._reworkRequest("Rework")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this._reworkRequest("Rework");
                        }
                      }}
                    >
                      <i className="fe-corner-up-left me-1"></i> Rework
                    </div>
                  </a>
                </>
              }

              {/* <DefaultButton text="Cancel" onClick={() => this.cancelRequest("myapproval")} /> */}

              <div
                role="button"
                tabIndex={0}
                className="btn cancel-btn waves-effect waves-light m-1"
                onClick={() => this.cancelRequest("myapproval")}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.cancelRequest("myapproval");
                  }
                }}
              >
                <i className="fe-x me-1"></i> Cancel
              </div>



            </div>
          }
          {console.log("ApproverEmailforward", ApproverEmail, CurrentuserEmail)}
          {this.state.showForward && (!this.state.Loading && !setloading) &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>
              {(RequesterEmail == CurrentuserEmail || isRequesterDelegated) &&
                <>
                  {/* <PrimaryButton text="Forward" onClick={() => this.handleForward("Forward")} />
                  <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} /> */}
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-primary waves-effect waves-light m-1"
                      onClick={() => this.state.editncType == "Observation" ? this._approveRequest("Approve") : this.handleForward("Forward")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this.state.editncType == "Observation" ? this._approveRequest("Approve") : this.handleForward("Forward");
                        }
                      }}
                    >
                      <i className="fe-check-circle me-1"></i> {this.state.editncType == "Observation" ? "Approve" : "Forward"}
                    </div>
                  </a>
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-warning waves-effect waves-light m-1"
                      onClick={(e) => this._reworkRequest("Rework")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this._reworkRequest("Rework");
                        }
                      }}
                    >
                      <i className="fe-corner-up-left me-1"></i> Rework
                    </div>
                  </a>
                </>
              }
              {/* <a href='#/form'>
               {/* <DefaultButton onClick={() => this.cancelRequest("myapproval")}>Cancel</DefaultButton> 
               </a>*/}

              <div
                role="button"
                tabIndex={0}
                className="btn cancel-btn waves-effect waves-light m-1"
                onClick={() => this.cancelRequest("myapproval")}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.cancelRequest("myapproval");
                  }
                }}
              >
                <i className="fe-x me-1"></i> Cancel
              </div>


            </div>
          }
          {console.log("Approverreject", this.state.showReject, ApproverEmail, CurrentuserEmail)}
          {this.state.showReject && (!this.state.Loading && !setloading) &&

            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>
              {(ApproverEmail == CurrentuserEmail || isApproverDelegated) &&
                <>
                  {/* <PrimaryButton text="Approve" onClick={() => this._approveRequest("Approve")} />
                  <PrimaryButton text="Reject" onClick={() => this._rejectRequest("Reject")} />
                  <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} /> */}
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-success waves-effect waves-light m-1"
                      onClick={(e) => this._approveRequest("Approve")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this._approveRequest("Approve");
                        }
                      }}
                    >
                      <i className="fe-check-circle me-1"></i> Approve
                    </div>
                  </a>
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-warning waves-effect waves-light m-1"
                      onClick={(e) => this._reworkRequest("Rework")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this._reworkRequest("Rework");
                        }
                      }}
                    >
                      <i className="fe-corner-up-left me-1"></i> Rework
                    </div>
                  </a>
                  <a>
                    <div
                      role="button"
                      tabIndex={0}
                      className="btn btn-danger waves-effect waves-light m-1"
                      onClick={(e) => this._rejectRequest("Reject")}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          this._rejectRequest("Reject");
                        }
                      }}
                    >
                      <i className="fe-x-circle me-1"></i> Reject
                    </div>
                  </a>
                </>
              }
              {/* <DefaultButton text="Cancel" onClick={() => this.cancelRequest("myapproval")} /> */}

              <div
                role="button"
                tabIndex={0}
                className="btn cancel-btn waves-effect waves-light m-1"
                onClick={() => this.cancelRequest("myapproval")}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.cancelRequest("myapproval");
                  }
                }}
              >
                <i className="fe-x me-1"></i> Cancel
              </div>


            </div>
          }
          {this.state.edType === "view" && (!(this.state.Loading && setloading))
            &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

              {/* <DefaultButton text="Cancel" onClick={() => this.cancelRequest("Edcmain")} /> */}

              <div
                role="button"
                tabIndex={0}
                className="btn cancel-btn waves-effect waves-light m-1"
                onClick={() => this.cancelRequest("Edcmain")}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.cancelRequest("Edcmain");
                  }
                }}
              >
                <i className="fe-x me-1"></i> Cancel
              </div>



            </div>
          }

          {this.state.editSubmitStatus == "Yes" && (!this.state.Loading || !setloading) ?
            <section className='card card-body mt-2'>
              <form>
                <div
                  style={{ justifyContent: "left", textAlign: "left" }}
                  className="row"
                >
                  <div className="form-group col-md-12">
                    <h3
                      style={{ textAlign: "left" }}
                      className="text-dark text-left font-16 fw-bold mb-3"
                    >
                      Audit History {/* Divyansh Changes */}
                    </h3>
                    <p className="font-14 mb-3 text-muted">Below table describes the status of approval at various level </p>
                  </div>
                </div>
                <div style={{ display: 'grid' }} className='row'>
                  <table className='mtbalenew'>
                    <thead>
                      <tr>
                        <th style={{ minWidth: "70px", maxWidth: "70px" }}>
                          S No.
                        </th>
                        <th style={{ minWidth: "70px", maxWidth: "70px" }}>
                          Level
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Assigned To
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Assigned To Role {/** Divyansh Changes */}
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Requestor Name
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Requested Date
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Action Taken By
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Action Taken On
                        </th>
                        <th style={{ minWidth: "90px", maxWidth: "90px" }}>
                          Remarks
                        </th>
                        <th style={{ minWidth: "70px", maxWidth: "70px" }}>
                          Status
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {auditHistory}
                    </tbody>
                  </table>
                </div>
              </form>
            </section> : null}
          {
            <Modal show={this.state.ShowModalTemplateDoc} onHide={() => this.setState({ ShowModalTemplateDoc: false })} size={Showfile ? "xl" : "lg"} className='newmobmodal'>

              <Modal.Body className="" id="style-5">
                <>
                  {Showfile &&

                    <FileViewer showfile={Showfile} docurl={this.state.redirecturl} cancelAction={this.cancelModalAction} />

                  }
                </>
              </Modal.Body>
            </Modal>
          }
          {/* {
            <Modal show={this.state.ShowModalAtt} onHide={() => this.setState({ ShowModalAtt: false })} size={Showfile ? "xl" : "lg"} className='newmobmodal'>
              <Modal.Header closeButton>
                <Modal.Title> Attachment Details <br></br>
                  <p className='text-muted font-14 fw-400 mb-0'>Below are the attachment details for Change Request
                  </p>

                </Modal.Title>
                {/* {ImagepostArr1.length > 0 && showBannerModal && <Modal.Title>Media Images</Modal.Title>} */}
          {/* </Modal.Header>
              <Modal.Body className="" id="style-5">
                <>
                  {Showfile &&

                    <FileViewer showfile={Showfile} docurl={this.state.redirecturl} cancelAction={this.cancelModalAction} />

                  }
                </>
              </Modal.Body>
            </Modal>
          } */}
        </div>
      </section >
    )
  }
}
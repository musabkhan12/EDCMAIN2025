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
import CustomBreadcrumb from '../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
import { icon } from '@fortawesome/fontawesome-svg-core';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faDownload, faEye, faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Item } from '@pnp/sp/items';
import { redirect } from 'react-router-dom';
import { getMemoNumberAuditReport, getNCNumbers } from '../AnnualAuditReportComponent/AuditReportService';
let Approvallistitemid = 0;
let ApproverEmail = "";
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
let optionsmemoNumbernewobs: any[] = [];
let forwardisdisabled: boolean = false;
let Approveclicked: boolean = false;
let Rejectclicked: boolean = false;
let Reworkclicked: boolean = false;
let editforwardrecord: boolean = false;
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
  editNCNumberID: string | number;
  editdepartmentCode: string;
  editserialNo: number;
  notUpdateDepartmentCode: string;
  notUpdateSerialNo: number;
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
}
const optionsApp: IDropdownOption[] = [

  { key: 'All', text: 'Everyone' },
  { key: 'One', text: 'Anyone' }

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
      editdepartmentCode: "",
      editserialNo: 0,
      notUpdateDepartmentCode: "",
      notUpdateSerialNo: 0,
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
      approvers: [{ Role: "", Level: "", Name: "", Type: "One", Index: 0, appEx: "", itemId: "" }],
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
  private handleFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
    this.setState({ copyFil: [], exFiles: [] })
    if (e.target.files) {
      _self.setState({ fileCount: e.target.files.length });
      _self.setState({ files: e.target.files });
      var allfiles: any[] = [];
      [].forEach.call(e.target.files, function (file: File) {
        allfiles.push(file);
      })
      _self.setState({ copyFil: allfiles });
    }
  };
  private handleAuditeeFileChange(e: React.ChangeEvent<HTMLInputElement>, _self: any) {
    this.setState({ copyFilauditee: [], exFilesauditee: [] });
    if (e.target.files) {
      _self.setState({ fileCountauditee: e.target.files.length });
      _self.setState({ filesauditee: e.target.files });
      var allfiles: any[] = [];
      [].forEach.call(e.target.files, function (file: File) {
        allfiles.push(file);
      })
      _self.setState({ copyFilauditee: allfiles });
    }
  };

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
    debugger
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
  public changeMemoNumber = async (item: any): Promise<void> => {
    debugger
    const sp = spfi().using(SPFx(this.props.context));
    let nctypenew: string = this.state.editncType == "NC" ? "NC Number" : "Observation NUmber";
    //let NCNumberoptionnew = await getNCNumbers(sp, item.text, nctypenew);
    let optionsNCNumber: any = [];
    const NCNumberoptionnew = await getNCNumbers(sp, item.label, nctypenew);
    let existingrecords = await this.getNCdata(item.label);
    if (Array.isArray(NCNumberoptionnew) && NCNumberoptionnew.length > 0) {
      // Safely extract existing NCNumbers, even if the array is empty
      const existingNCNumbersSet = new Set(
        (Array.isArray(existingrecords) ? existingrecords[0] : []).map((rec: any) => rec.NCNumber)
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
    let optionsNCNumbernew: any[] = [];
    optionsNCNumbernew = await this.getUniqueBy(optionsNCNumber, "ncNo");
    optionsNCNumbernew.sort((a, b) => a.label.localeCompare(b.label));
    let approvedauditreportselected = this.state.editmemonumberOptions.filter((x: any) => x.value == item.value);
    this.setState({ editNCNumberOptions: optionsNCNumbernew })
    this.setState({ editApprovedAuditReport: item.value, editMemoNumber: item.memoNumber, ApprovedAuditSelected: approvedauditreportselected });
  };
  public async getNCdata(reportcode: string) {
    const sp = spfi().using(SPFx(this.props.context));
    let arr: any[] = []
    let arrs = []
    let bannerimg = []
    const currentUser = await sp.web.currentUser();
    await sp.web.lists.getByTitle("NonConformityList").items
      .select("*")
      .expand("")
      .filter(`ApprovedAuditReportMemoNumber eq '${reportcode}'`)
      .orderBy("Modified", false)
      ()
      .then((res: any) => {
        console.log(res, 'Memonumbers from audit report');

        arr.push(res)
        // arr = res;
      })
      .catch((error: any) => {
        console.log("Error fetching data: ", error);
      });
    console.log(arr, 'arr');
    return arr;
  }
  public changeNCNumber = (item: any): void => {
    let ncnumberselected = this.state.editNCNumberOptions.filter((x: any) => x.value == item.value);
    this.setState({ editNCNumber: item.label, editNCNumberID: item.value, NCNumberselected: ncnumberselected });
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
    debugger
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

  private _handleCheckboxChange = (stateKey: keyof IEditState, itemKey: number) =>
    (_ev: React.FormEvent<HTMLElement>, isChecked?: boolean) => {
      this.setState((prevState) => {
        const updatedValues = isChecked
          ? [...(prevState[stateKey] as number[]), itemKey]
          : (prevState[stateKey] as number[]).filter((key) => key !== itemKey);
        return { [stateKey]: updatedValues } as unknown as Pick<IEditState, keyof IEditState>;
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
    debugger
    const _sp = spfi().using(SPFx(this.props.context));
    const currentUser = await _sp.web.currentUser();
    CurrentuserEmail = currentUser.Email;
    // Extracting the part after `#/`
    const url = window.location.href;
    // alert(
    //   url + "url"
    // )
    const parts = url.split("#/")[1].split("/");

    const programName = decodeURIComponent(parts[0]); // "Non Confirmity"
    const editType = parts[1]; // "edit"
    const id = parts[2]; // "165"

    console.log("Program Name:", programName);
    console.log("Edit Type:", editType);
    console.log("ID:", id);
    // alert("Program Name:"+ programName)
    // alert("Edit Type:"+ editType)
    // alert("ID:"+ id)
    debugger
    if (id) {
      this.setState({ mainItemId: id }, async () => {
        // this.getAllapprovalitems(Number(id));
        // This will run AFTER the state update is completed
        // alert("Updated mainItemId: " + this.state.mainItemId);
        await this.getListData(); // Fetch list data after updating state
        await this.getGeneratedTemplateDocNC(Number(id))
      });
    }
    if (editType) {
      this.setState({ edType: editType })
    }
    // await this.getListData();
    await this.getnctypeoptions();
    await this.getDepartment();
    await this.getAuditreport();
    await this.getDataRoles();
    await this.getMainListName();
    await this.getRequestorRole();
    await this.getFormName();
    debugger
    if (editType === "edit") {
      this.setState({ showApprove: false });
      this.setState({ showSubmit: true });
      this.setState({ showForward: false });
      this.setState({ showReject: false });
      this.setState({ forwarDisable: false });
      if (this.state.editSubmitStatus === 'No' || (this.state.editStatus == "Rework" && this.state.editCurrentUserRole == "FirstInitiator")) {
        this.setState({ showDraft: true });
        this.setState({ isDisabled: false });
      }
      else {
        this.setState({ deptSectionDisable: false });
        this.setState({ showDelegate: false });
        this.setState({ showDraft: false });
        this.setState({ isDisabled: true });
      }
      if (this.state.editDelegateToSubmitStatus == "No" && this.state.editCurrentUserRole == "DelegateTo") {
        this.setState({ showDelegate: true });
      }
    }
    else if (editType === "view") {
      this.setState({ isDisabled: true });
      this.setState({ deptSectionDisable: true });
      this.setState({ showDelegate: true });
      this.setState({ forwarDisable: true });
      this.setState({ showApprove: false });
      this.setState({ showSubmit: false });
      this.setState({ showDraft: false });
      this.setState({ showForward: false });
      this.setState({ showReject: false });
    }
    else if (editType === "approve") {
      const approvalItemId = parts[3];
      // console.log("approvalItemId",approvalItemId)
      // alert("approvalItemId"+ approvalItemId)
      // alert("approvalItemId"+ typeof(approvalItemId))
      if (approvalItemId) {
        Approvallistitemid = Number(approvalItemId);
        this.setState({ approvalItemId: approvalItemId })
        // alert('here is my state ' + this.state.approvalItemId)
      }
      this.setState({ isDisabled: true });
      this.setState({ deptSectionDisable: true });
      this.setState({ showDelegate: true });
      this.setState({ showApprove: true });
      this.setState({ showSubmit: false });
      this.setState({ showDraft: false });
      this.setState({ showReject: false });
      if (this.state.editLastInitiatorSubmitStatus == "No" && this.state.editCurrentUserRole == "LastInitiator") {
        this.setState({ showApprove: false });
        this.setState({ showForward: true });
      }
      if (this.state.editLastInitiatorSubmitStatus == "Yes" && this.state.editCurrentUserRole == "LastInitiator") {
        this.setState({ showApprove: false });
        this.setState({ forwarDisable: true });
        this.setState({ showForward: false });
        // this.setState({ showReject: true });
      }
      if (this.state.editLastInitiatorSubmitStatus == "Yes") {
        if (this.state.editCurrentUserRole == "Approverrole") {
          this.setState({ showApprove: false });
          this.setState({ forwarDisable: true });
          this.setState({ showForward: false });
          this.setState({ showReject: true });
        }

        forwardisdisabled = true;
      }
      if (this.state.editStatus == "Approved" || this.state.editStatus == "Rejected") {
        this.setState({ showReject: false });
      }
    }
  }
  public async getGeneratedTemplateDocNC(itemId: number) {
    debugger
    const _sp = spfi().using(SPFx(this.props.context));
    let results: any = [];
    // for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("NonConformityGeneratedTemplateDoc").items
      .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
      .then((res) => {
        console.log(res, 'tem let arrs=[]');
        results = res;
        this.setState({ TemplateDoc: res })
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    // }
    console.log(results, 'results');
    return results;
  }
  public async getapprovalbyID(id: number, processName: string) {
    debugger
    const _sp = spfi().using(SPFx(this.props.context));
    let arr: any[] = []
    let arrs = []
    let bannerimg = []
    const currentUser = await _sp.web.currentUser();
    await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
      .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title,AssignedTo/EMail").expand("Author,RequesterName,AssignedTo")()
      .then((res) => {
        console.log(res, 'ghghghghgh let arrs=[]');
        if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName) {
          arr.push(res);
        }
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    console.log(arr, 'arr approval of current user');
    return arr;
  }
  public async getListData() {
    debugger
    let memoopt = await this.getAuditreport();
    let departopt = await this.getDepartment();
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const Items: any = await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId)
        .select("*, Category/Id, Category/Title, SubCategory/Id, Location/Id, Location/Title, SubCategory/Title, AssignedTo/Id, AssignedTo/Title,AssignedTo/EMail,DelegateTo/EMail, DelegateTo/Id, DelegateTo/Title, AnalyzedBy/Id, AnalyzedBy/Title, ReviewedBy/Id, ReviewedBy/Title, PersonAssigned/Id, PersonAssigned/Title,Author/Id,Author/Title,Author/EMail")
        .expand("Category, SubCategory, Location, AssignedTo, DelegateTo, AnalyzedBy, ReviewedBy, PersonAssigned,Author")();
      console.log("Itemsedit", Items);
      this.setState({
        ncItemId: Items.Id,
        editDepartment: Items.DepartmentId,
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
        remarks: Items.FinalRemarks,
        reworkremarks: Items.ReworkRemarks
      });
      const deptItems = await sp.web.lists.getByTitle("DepartmentMasterList").items();
      const optionsdept = deptItems.map((item: {
        DepartmentCode: any; Title: string; Id: number
      }) => ({
        value: item.Id,
        label: item.Title,
        data: { departmentCode: item.DepartmentCode },
      }));
      if (Items?.NCType) {
        this.setState({
          editncType: Items?.NCType,
          editmemonumberOptions: Items?.NCType == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs
        });
      }
      let nctypenew: string = Items.NCType == "NC" ? "NC Number" : "Observation NUmber";
      let NCNumberoptionnew = await getNCNumbers(sp, Items.ApprovedAuditReportMemoNumber, nctypenew);
      let optionsNCNumber: any = [];
      if (NCNumberoptionnew.length > 0) {
        optionsNCNumber = NCNumberoptionnew[0].map((item: any) => ({
          value: item.ID,
          label: item.NCNumber,
          ncNo: item.NCNumber,
          reportcode: item.ReportCode,
          nctype: item.NCType
        }));
      }
      let optionsNCNumbernew: any[] = [];
      optionsNCNumbernew = await this.getUniqueBy(optionsNCNumber, "ncNo");
      let editmemonumberOptionsselect = Items?.NCType == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs
      let approvedauditreportselected = editmemonumberOptionsselect.filter((x: any) => x.value == Items.ApprovedAuditReportId);
      let ncnumberselected = optionsNCNumbernew.filter((x: any) => x.value == Items.NCNumberID);
      let selecteddepartment = optionsdept.filter((x: any) => x.value == Items.DepartmentId);
      this.setState({ editNCNumberOptions: optionsNCNumbernew, NCNumberselected: ncnumberselected, ApprovedAuditSelected: approvedauditreportselected, departmentselected: selecteddepartment })
      //this.setState({ editApprovedAuditReport: item.key, editMemoNumber: item.memoNumber });
      RequesterEmail = Items.Author.EMail;
      const apprItems = await sp.web.lists
        .getByTitle("ProcessApprovalList")
        .items.select(
          "*",
          "AssignedTo/Title,AssignedTo/Id,AssignedTo/EMail,ActionTakenRole,ActionTakenRole/Role,RequesterName/Title,ActionTakenBy/Title"
        )
        .expand("AssignedTo,ActionTakenRole,RequesterName,ActionTakenBy")
        .filter(
          "ListItemId eq '" +
          this.state.mainItemId +
          "' and ProcessName eq 'Non Conformity'"
        )
        .orderBy("Id", false)();
      debugger
      let url = window.location.href;
      let parts = url.split("#/")[1].split("/");
      let editType = parts[1]; // "edit"
      let id = parts[2];
      let currentlevel: any;
      let finallevel: any;
      if (editType == "edit") {
        // if (currentApprover.CurrentUserRole !== "LastInitiator") {
        showreworkremarks = true;
        // }
      }
      if (editType === "approve") {

        let approvalItemIdnew = parts[3];
        let Approverdata = await this.getapprovalbyID(Number(approvalItemIdnew), "Non Conformity");
        console.log("Approverdata", Approverdata, "Approverdata0", Approverdata && Approverdata[0], Approvallistitemid, approvalItemIdnew);

        if (Approverdata.length > 0) {
          const currentApprover = Approverdata[0];
          const isAnalyzedRole = currentApprover.CurrentUserRole === "AnalyzedBy";
          const isFirstAssigned = currentApprover.CurrentUserRole === "FirstAssignedTo" || currentApprover.CurrentUserRole === "DelegateTo";
          const isCurrentUser = (currentApprover.CurrentUserRole === "FirstAssignedTo" && currentApprover.AssignedTo?.EMail === CurrentuserEmail) ||
            (currentApprover.CurrentUserRole === "DelegateTo" && currentApprover.DelegateTo?.EMail === CurrentuserEmail);

          const currentstatus = currentApprover.Status == "Approved";
          const finalstatus = Items.Status == "Approved";
          currentlevel = currentApprover.Level;
          finallevel = currentApprover.Maxlevel;
          ApproverEmail = currentApprover.AssignedTo?.EMail;

          if (this.state.editLastInitiatorSubmitStatus == "No" && this.state.editCurrentUserRole == "LastInitiator" && currentApprover.AssignedTo?.EMail === CurrentuserEmail) {
            forwardisdisabled = false
          } else {
            forwardisdisabled = true;
          }
          if (isAnalyzedRole && isCurrentUser) {
            IsAnalyzedBy = true;
            showimsupdated = true;
            isdisableims = false;
          }
          if (currentApprover.CurrentUserRole !== "LastInitiator") {
            showreworkremarks = true;
          }
          // This will override previous value only if role is "FirstAssignedTo"
          if (!isFirstAssigned && Items.IMSUpdated != "" && Items.IMSUpdated != null && Items.RiskOpportunitiesUpdated != "" && Items.RiskOpportunitiesUpdated != null && !(isAnalyzedRole && isCurrentUser)) {
            showimsupdated = true;
            isdisableims = true;
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

        if (Items.IMSUpdated != "" && Items.IMSUpdated != null && Items.RiskOpportunitiesUpdated != "" && Items.RiskOpportunitiesUpdated != null) {
          showimsupdated = true;
          isdisableims = true;
        }

        if (Items.Status == "Approved") {
          showcorrectionappicable = true;
          isdisablefinal = true;
        }
      }
      var cnt: any = 0;
      var appItems: any[] = [];
      console.log("apprItems111", apprItems);
      if (apprItems.length > 0) {
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
        if (Items.SubmitStatus == "Yes" && Items.Status == "Rework" && Items.CurrentUserRole == "FirstInitiator" && apprItems.length > 1) {
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
          objToAdd["AssignedTo"] = itm.AssignedTo.Title;
          objToAdd["AssignedToEmail"] = itm.AssignedTo.EMail;
          objToAdd["RequesterName"] = itm.RequesterName.Title;
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
        .filter("NonConformityId eq '" + this.state.mainItemId + "'")();
      if (upFiles.length > 0) {
        var obJFiles: any[] = [];
        let fCount: number = 0;
        upFiles.forEach(function (item: any) {
          obJFiles.push({
            "Name": item.File.Name,
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
        .filter("NonConformityId eq '" + this.state.mainItemId + "'")();
      if (upFilesauditee.length > 0) {
        var obJFiles: any[] = [];
        let fCount: number = 0;
        upFilesauditee.forEach(function (item: any) {
          obJFiles.push({
            "Name": item.File.Name,
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
      if (Items.substatus == "No") {
        this.setState({ ShowDeleteicon: true })
      }
      //AllProcessApproval Table data
      debugger
      const approvalItems = await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.select("*", "Approvers/Name", "Approvers/EMail", "Approvers/ID").expand("Approvers")
        .filter("MainListID eq '" + this.state.mainItemId + "'and ProcessName eq 'Non Conformity'")
        .orderBy("Level", true)();
      var allApp: any[] = [];
      var cnt: any = 0;
      const sorted = [...approvalItems].sort((a, b) => b.Level - a.Level);
     

      debugger
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
          objToAdd["itemId"] = itm.Id;
          objToAdd["Index"] = itm.Level;
          objToAdd["appEx"] = approve;


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
        const isFinalApprover =  finalApprover.some(
          (approver: any) => approver.email?.toLowerCase() === currentUserEmail?.toLowerCase()
        ) && Items.CurrentUserRole == "Approverrole";
        console.log("finallevel", finallevel, currentlevel);
        debugger
        if (editType === "approve") {
          if (Items.Status != "Approved" && isFinalApprover && finallevel == currentlevel && Items.CurrentUserRole == "Approverrole") {
            showcorrectionappicable = true;
            isdisablefinal = false;
            this.setState({ IsFinalapprover: isFinalApprover });
          }
        }

      }

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
    Swal.fire({
      title: 'Do you want to cancel this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(function (val) {
      if (val.isConfirmed) {
        Swal.fire({
          title: "Cancelled Successfully.",
          icon: "success"
        }).then(() => {
          if (redirectto == "myapproval") {
            window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
          } else {
            window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          }

          // window.location.reload();
        });
      }
    });
  }
  public async getAuditreport() {
    const sp = spfi().using(SPFx(this.props.context));
    debugger
    try {
      const memoItems = await getMemoNumberAuditReport(sp);
      let optionsNCNumber: any = [];
      let optionsObservationNumber: any = [];
      if (memoItems.length > 0) {
        const filteredItemsNC = memoItems[0].filter((item: any) => item.FailureofIntentNonconformity === "Yes");
        const filteredItemsObs = memoItems[0].filter((item: any) => item.Observations === "Yes");
        if (filteredItemsNC.length > 0) {
          optionsNCNumber = filteredItemsNC.map((item: any) => ({
            value: item.ID,
            label: item.ReportCode,
            itemId: item.ID,
            reportCode: item.ReportCode,
            ncNo: item.NCNumber
          }));
        }
        if (filteredItemsObs.length > 0) {
          optionsObservationNumber = filteredItemsObs.map((item: any) => ({
            value: item.ID,
            label: item.ReportCode,
            itemId: item.ID,
            reportCode: item.ReportCode,
            ncNo: item.NCNumber
          }));
        }
      }
      //this.state.ncType
      //let optionsmemoNumbernewnc: any[] = [];
      optionsmemoNumbernewnc = await this.getUniqueBy(optionsNCNumber, "reportCode");
      optionsmemoNumbernewnc = [...optionsmemoNumbernewnc].sort((a, b) =>
        a.label.localeCompare(b.label)
      );

      //let optionsmemoNumbernewobs: any[] = [];
      optionsmemoNumbernewobs = await this.getUniqueBy(optionsObservationNumber, "reportCode");
      optionsmemoNumbernewobs = [...optionsmemoNumbernewobs].sort((a, b) =>
        a.label.localeCompare(b.label)
      );
      this.setState({
        editmemonumberOptions: this.state.editncType == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs,
        editmemonumberOptionsall: this.state.editncType == "NC" ? optionsNCNumber : optionsObservationNumber
      });
    } catch (e) {
      console.error(e);
    }
  }
  // public async getAuditreport() {
  //   const sp = spfi().using(SPFx(this.props.context));
  //   debugger
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
      const deptItems = await sp.web.lists.getByTitle("DepartmentMasterList").items();
      const options = deptItems.map((item: {
        DepartmentCode: any; Title: string; Id: number
      }) => ({
        value: item.Id,
        label: item.Title,
        data: { departmentCode: item.DepartmentCode },
      }));
      this.setState({ editDepartmentOption: options });
      // Find the selected department's departmentCode and set it
      const selectedDeptArray = options.filter(opt => opt.value === this.state.editDepartment);
      if (selectedDeptArray.length > 0) {
        this.setState({ editdepartmentCode: selectedDeptArray[0].data.departmentCode });
      }

      await this.getCategory();
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
    if (this.validateFormSubmit()) {
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
      if (this.state.departmentselected.length == 0) editErrors.editDepartment = "Department is required";
      if (this.state.ApprovedAuditSelected.length == 0) editErrors.editApprovedAuditReport = "Report code is required";
      if (this.state.NCNumberselected.length == 0) editErrors.editNCNumber = "NCR number is required";
      if (!this.state.editCriteria) editErrors.editCriteria = "Criteria is required";
      if (!this.state.editCloseOutStatus) editErrors.editCloseOutStatus = "Close Out Status is required";
      if (this.state.editCategoryValueIsCheck.length == 0) {
        document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
          el.classList.add(styles.errCh);
        });
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
      if (!this.state.editReviewedBy) {
        editErrors.editReviewedBy = "Reviewed By is required";
        document.querySelectorAll("#reviewedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.add(styles.errCh);
        });
      } else {
        document.querySelectorAll("#reviewedBypeoplepicker .ms-BasePicker-text").forEach((el) => {
          el.classList.remove(styles.errCh);
        });
      }
      if (!this.state.editCorrectiveActionImplementedOn) editErrors.editCorrectiveActionImplementedOn = "Corrective Action Implemented on is required";
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
    if (this.validateFormForward()) {
      this.forwardRequest(formsubmode);
      console.log("Form Data:", this.state);
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

      if (isNameEmpty) {
        editErrors[`approvers[${index}].Name`] = 'Approver is required';
        hasError = true;

        document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`).forEach((el) => {
          el.classList.add(styles.errCh);
        });

      }
      else {
        document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`).forEach((el) => {
          el.classList.remove(styles.errCh);
        });
        document.querySelectorAll(`#approverpeoplepicker-${index} .ms-BasePicker-text`).forEach((el) => {
          el.classList.add(styles.peoplepickerstyleAuditteerrch);
        });
        
      }

      if (isRoleEmpty) {
        editErrors[`approvers[${index}].Role`] = 'Role is required';
        hasError = true;
      }
    });

    this.setState({ editErrors });

    return !hasError;
  };


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
    if (this.validateFormDraft()) {
      this._updateSubmitData(formsubmode);
      console.log("Form Data:", this.state);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public validateFormDraft = (): boolean => {
    let editErrors: { [key: string]: string } = {};

    if (this.state.departmentselected.length == 0) {
      editErrors.editDepartment = "Department is required";
      this.setState({ editErrors });
      Swal.fire('Please select a department.');
      return false;
    }
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
    if (this.validateFormRemark()) {
      this.approveRequest(formsubmode);
      console.log("Form Data:", this.state);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public _rejectRequest = (formsubmode: string) => {
    Rejectclicked = true;
    Reworkclicked = false;
    if (this.validateFormRemark()) {
      this.rejectRequest(formsubmode);
      console.log("Form Data:", this.state);
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
      console.log("Form Data:", this.state);
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
    if (!this.state.reworkremarks) editErrors.reworkremarks = "Remarks is required";
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

  private _getPeoplePickerItemsApp(items: any[], i: number) {
    debugger
    var arr: any[];
    arr = [];
    items.forEach(function (it) {
      arr.push(it.id);
    })
    this.state.approvers.filter(function (it) {
      if (it.Index == (editforwardrecord ? i + 1 : i)) {
        it.Name = arr
      }
    });
    this.setState({ approvers: this.state.approvers });
  }

  private addApprover() {
    debugger
    var itm = this.state.indApp + 1;
    this.setState({ indApp: itm });
    this.state.approvers.push({ Role: "", Level: "", Name: "", Index: itm, itemId: "" });
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

  public async updateData(_editsubmitStatus: string, firstInitiatorSubmitStatus: string, firstAssignedToSubmitStatus: string, delegateToSubmitStatus: string, analyzedBySubmitStatus: string, reviewedBySubmitStatus: string, lastAssignedToSubmitStatus: string, lastInitiatorSubmitStatus: string, currentUserRole: string, reworkById: any, serialNumber: number, documentCode: string, ncrnumber: String) {
    debugger
    const sp = spfi().using(SPFx(this.props.context));
    let observationStatus: string = "";
    let test1 = this.state.correctionApplicable ? "Yes" : "No";
    let test2 = this.state.notEffective ? "Yes" : "No";
    let test3 = this.state.effectiveClosed ? "Yes" : "No";
    let ncStatus = _editsubmitStatus == "Rework" ? "Rework" : "Pending";
    if (this.state.editncType == "Observation" && this.state.editDelegateToId != null && this.state.editCurrentUserRole == "LastAssignedTo") {
      observationStatus = "Approved"
    } else
      if (this.state.editncType == "Observation" && this.state.editDelegateToId == null && this.state.editCurrentUserRole == "ReviewedBy") {
        observationStatus = "Approved"
      } else {
        observationStatus = "Pending"
      };
    // let submitstatus: string = "";
    // submitstatus = ((_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo")
    //   || (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "DelegateTo")
    //   || (_editsubmitStatus == "draft")) ?
    //   "No" : "Yes";
    _editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "FirstAssignedTo"
    await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId).update({
      NCNumber: this.state.editNCNumber,
      NCNumberID: this.state.editNCNumberID,
      ApprovedAuditReportId: this.state.editApprovedAuditReport || null,
      ApprovedAuditReportMemoNumber: this.state.editMemoNumber,
      DepartmentId: this.state.editDepartment || null,
      Criteria: this.state.editCriteria,
      CloseOutStatus: this.state.editCloseOutStatus,
      //NCRNo: this.state.editNCRNo,
      ReferenceNumber: this.state.editReferenceNumber,
      RevisionNumber: this.state.editRevisionNo !== "" ? Number(this.state.editRevisionNo) : null,
      IssueNumber: this.state.editIssueNo !== "" ? Number(this.state.editIssueNo) : null,
      IssueDate: this.state.editIssueDate,
      RevisionDate: this.state.editRevisionDate,
      CategoryId: this.state.editCategoryValueIsCheck,
      SubCategoryId: this.state.editSubCategoryValueIsCheck,
      LocationId: this.state.editLocationValueIsCheck,
      AssignedToId: _editsubmitStatus == "submit" && this.state.editCurrentUserRole == null ? this.props.currentUserID : this.state.editAssignToId || null,
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
      SubmitStatus: _editsubmitStatus == "draft" ? "No" : "Yes",
      SubmiitedById: this.props.currentUserID || null,
      CurrentUserRole: currentUserRole,
      FirstInitiatorSubmitStatus: firstInitiatorSubmitStatus,
      FirstAssignedToSubmitStatus: firstAssignedToSubmitStatus,
      DelegateToSubmitStatus: delegateToSubmitStatus,
      AnalyzedBySubmitStatus: analyzedBySubmitStatus,
      ReviewedBySubmitStatus: reviewedBySubmitStatus,
      LastAssignedToSubmitStatus: lastAssignedToSubmitStatus,
      LastInitiatorSubmitStatus: lastInitiatorSubmitStatus,
      Status: this.state.editncType == "Observation" ? observationStatus : ncStatus,
      IsRework: _editsubmitStatus == "Rework" || _editsubmitStatus == "Reject" ? "Yes" : "No",
      ReworkById: reworkById,
      SerialNumber: serialNumber,
      NCRNo: ncrnumber,
      DocumentCode: this.state.editDocumentCode,
      IMSUpdated: this.state.isIMSUpdated,
      RiskOpportunitiesUpdated: this.state.riskandopportunitiesUpdated,

    })
  }

  //Update Function
  private _updateSubmitData = async (_editsubmitStatus: string) => {
    debugger
    let mText = "";
    let cText = "";
    //Start Flow condition
    let currentUserRole = "";
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

    if (_editsubmitStatus == "draft" && this.state.editCurrentUserRole == null) {
      firstInitiatorSubmitStatus = "No"
      firstAssignedToSubmitStatus = "No";
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
    else if (this.state.editCurrentUserRole == "FirstInitiator") {
      firstInitiatorSubmitStatus = "Yes";
      firstAssignedToSubmitStatus = "No";
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

    Swal.fire({
      title: "Do you want to " + mText + " this request?",
      showCancelButton: true,
      confirmButtonText: "Yes",
      cancelButtonText: "No"
    }).then(async (result) => {
      if (result.isConfirmed) {
        debugger
        setloading = true;
        console.log("ApprovallistitemidApprovallistitemid", Approvallistitemid, approvalItemId, _editsubmitStatus)
        this.setState({ Loading: true });
        await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber);
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
              const uploadPromises = this.state.copyFilauditee.map((file) => {
                const fileNamePath = encodeURI(file.name);
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
                Remark: editProblemDescription,
              });
            } else {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: editProblemDescription,
              });
            }

          }
          else {
            if (approvalItemId == null || approvalItemId == undefined || approvalItemId == "") {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: editProblemDescription,
              });
            } else {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: editProblemDescription,
              });
            }
            if (this.state.fileDeleteId.length > 0) {
              this.state.fileDeleteId.forEach(function (ids) {
                sp.web.lists.getByTitle("NonConformityDocs").items.getById(ids).delete();
              });
            }
            if (this.state.copyFil.length > 0) {
              this.state.copyFil.forEach(function (file) {
                var fileNamePath = encodeURI(file.name);
                sp.web.getFolderByServerRelativePath("NonConformityDocs").files.addUsingPath(fileNamePath, file, { Overwrite: true }).then(function (response) {
                  response.file.getItem().then(function (fileItem) {
                    fileItem.update({
                      NonConformityId: itemid
                    });
                  });
                });
              })
            }

          }
        }
        else if (_editsubmitStatus == "draft") {
          if (this.state.fileDeleteId.length > 0) {
            this.state.fileDeleteId.forEach(function (ids) {
              sp.web.lists.getByTitle("NonConformityDocs").items.getById(ids).delete();
            });
          }
          if (this.state.copyFil.length > 0) {
            this.state.copyFil.forEach(function (file) {
              var fileNamePath = encodeURI(file.name);
              sp.web.getFolderByServerRelativePath("NonConformityDocs").files.addUsingPath(fileNamePath, file, { Overwrite: true }).then(function (response) {
                response.file.getItem().then(function (fileItem) {
                  fileItem.update({
                    NonConformityId: itemid
                  });
                });
              });
            })
          }

        }
        setloading = false;
        this.setState({ Loading: false }, () => {
          Swal.fire({
            title: cText + " Successfully.",
            icon: "success"
          }).then(() => {
            (this.state.editCurrentUserRole == "DelegateTo" || this.state.editCurrentUserRole == "FirstAssignedTo") ?
              window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx" :
              window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          });
        });

      }
    })
      .catch(error => {
        console.error("Error while saving:", error);
      });
  }
  //Forward Call
  public forwardRequest = async (_editsubmitStatus: string) => {
    debugger
    var _self = this;
    //Start Flow condition
    let currentUserRole = "";
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
          debugger
          setloading = true;
          this.setState({ Loading: true });
          await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber);
          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
            Status: "Approved",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
          });
          //Forward Button
          console.log("this.state.....", this.state.approvers)
          debugger
          if (approvers.length > 0) {
            var maxLength = approvers.length;
            approvers.forEach(function (it: any, val: any) {
              if (it.itemId == 0) {
                sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
                  Title: "Title",
                  MainListNameId: _self.state.mainListId,
                  ApproverRoleId: it.Role,
                  Level: it.Index + 1,
                  LevelType: it.Type || "One",
                  SubmitStatus: "Yes",
                  Maxlevel: maxLength,
                  ContentTitle: _self.state.editProblemDescription,
                  RequestId: _self.state.editNCNumber,
                  RequesterNameId: _self.props.currentUserID,
                  RequestedDate: new Date(),
                  ProcessName: "Non Conformity",
                  FormNameId: _self.state.formNameId,
                  MainListID: _self.state.ncItemId,
                  RequesterRoleId: _self.state.reqRolId,
                  ApproversId: it.Name,
                  ApprovalType: "Approval",
                }).catch(function (ex) {
                  console.log(ex.errorMessage);
                })
              }
              else {
                sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(it.itemId).update({
                  Title: "Title",
                  MainListNameId: _self.state.mainListId,
                  ApproverRoleId: it.Role,
                  Level: it.Index,
                  LevelType: it.Type || "One",
                  SubmitStatus: "Yes",
                  Maxlevel: maxLength,
                  ContentTitle: _self.state.editProblemDescription,
                  RequestId: _self.state.editNCNumber,
                  RequesterNameId: _self.props.currentUserID,
                  RequestedDate: new Date(),
                  ProcessName: "Non Conformity",
                  FormNameId: _self.state.formNameId,
                  MainListID: _self.state.ncItemId,
                  RequesterRoleId: _self.state.reqRolId,
                  ApproversId: it.Name,
                  ApprovalType: "Approval",
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
          this.setState({ Loading: false }, () => {
            Swal.fire({
              title: "Forwarded Successfully.",
              icon: "success"
            }).then(() => {
              window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
            });
          });
        }
      });

  }
  //Approver call
  public approveRequest = async (_editsubmitStatus: string) => {
    //Start Flow condition
    debugger
    let currentUserRole = "";
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
      if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "ReviewedBy";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "ReviewedBy") {
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
      if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "ReviewedBy";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "ReviewedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "Yes";
        analyzedBySubmitStatus = "Yes";
        reviewedBySubmitStatus = "Yes";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "LastAssignedTo";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Approve" && this.state.editCurrentUserRole == "LastAssignedTo") {
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
    }
    //<-------- End Case2 --------->
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
          debugger
          setloading = true;
          this.setState({ Loading: true });
          if (editLastInitiatorSubmitStatus != "Yes") {
            await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber);
          }
          const sp = spfi().using(SPFx(this.props.context));
          if (this.state.IsFinalapprover) {
            let test1 = this.state.correctionApplicable ? "Yes" : "No";
            let test2 = this.state.notEffective ? "Yes" : "No";
            let test3 = this.state.effectiveClosed ? "Yes" : "No";
            await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId).update({
              Correctionapplicable: this.state.IsFinalapprover
                ? test1
                : "",
              NotEffective: this.state.IsFinalapprover
                ? test2
                : "",
              EffectiveandProblemClosed: this.state.IsFinalapprover
                ? test3
                : "",
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
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Approved Successfully.",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });

          } else {
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({

              Status: "Approved",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Approved Successfully.",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });
          }

        }
      });

  }
  //Reject call
  private rejectRequest = async (_editsubmitStatus: string) => {
    const { remarks } = this.state
    const { currentUserID, approvalItemId, context } = this.props
    const sp = spfi().using(SPFx(this.props.context));
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
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
              Status: "Rejected",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Rejected Successfully.",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });
          } else {
            sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(Approvallistitemid)).update({
              Status: "Rejected",
              ActionTakenById: currentUserID,
              ActionTakenOn: new Date(),
              Remark: remarks,
            });
            setloading = false;
            this.setState({ Loading: false }, () => {
              Swal.fire({
                title: "Rejected Successfully",
                icon: "success"
              }).then(() => {
                window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/MyApprovals.aspx";
              });
            });

          }

        }
      });

  }
  private OpenFile = (obj: any, sts: string) => {
    debugger
    let url = this.props.context.pageContext.web.absoluteUrl;
    let tenanturl = url.match(/^https:\/\/[^\/]+/)[0];
    console.log("obbbj", obj)
    const fileUrl = `${tenanturl}${obj.FileRef}`;

    if (sts == "Open") {
      if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {

        window.open(`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=default`, "_blank");
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
  //Rework Call
  private reworkRequest = async (_editsubmitStatus: string) => {
    //Start Flow condition
    debugger
    let currentUserRole = "";
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
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "ReviewedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "No";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "FirstAssignedTo";
        reworkById = this.state.editReviewedById || null;
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
    //<-------- End Case1 --------->
    //<-------- Start Case2 --------->
    else {
      if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "AnalyzedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "DelegateTo";
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
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "ReviewedBy") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "DelegateTo";
        reworkById = this.state.editReviewedById || null;
        serialNumber = this.state.notUpdateSerialNo;
        ncrnumber = this.state.notUpdateDepartmentCode;
      }
      else if (_editsubmitStatus == "Rework" && this.state.editCurrentUserRole == "LastAssignedTo") {
        firstInitiatorSubmitStatus = "Yes"
        firstAssignedToSubmitStatus = "Yes";
        delegateToSubmitStatus = "No";
        analyzedBySubmitStatus = "No";
        reviewedBySubmitStatus = "No";
        lastAssignedToSubmitStatus = "No";
        lastInitiatorSubmitStatus = "No";
        currentUserRole = "DelegateTo";
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
          await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode, ncrnumber);
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
        editmemonumberOptions: option?.text == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs
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
  public render(): React.ReactElement<IAuditPlanProps> {
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
    if(this.state.editCurrentUserRole != "FirstInitiator"){
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
    }else{
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
        console.log(this.state.approvers, "this.state.approvers")
        return (
          <tr key={i} className='tblCls'>
            {/* <td>
            <TextField value={(i + 1).toString()} disabled={true} className={styles.width} ></TextField>
          </td> */}
            <td style={{ minWidth: "30px", maxWidth: "30px" }}>
              <div
                style={{ marginLeft: "0px", overflow: 'inherit' }}
                className="indexdesign"
              >
                {i + 1}</div>
            </td>
            <td title={this.state.optionsRole.find(opt => opt.key === this.state.approvers[i].Role)?.text || "Select role"} style={{ overflow: 'inherit' }} className="ng-binding">
              <Dropdown disabled={this.state.forwarDisable || forwardisdisabled} placeholder="Select options"

                selectedKey={this.state.approvers[i].Role}
                options={this.state.optionsRole
                  .filter((opt) =>
                    !this.state.approvers.some((app, index) => index !== i && app.Role === opt.key)
                  )}
                onChange={(e, itm: IDropdownOption) => this.onRoleChange(e, itm, i)}
                className={this.state.editErrors?.[`approvers[${i}].Role`] ? 'dropdown-error' : ''}
              // className={this.state.editErrors?.approvers ? 'dropdown-error' : ''}
              />
            </td>
            <td title={`Level ${(i + 1).toString()}` || "Level "} style={{ minWidth: '70px', maxWidth: '70px', overflow: 'inherit' }}>
              <TextField value={`Level ${(i + 1).toString()}`} disabled={true}></TextField>
            </td>
            <td title={this.state.approvers[i].appEx && this.state.approvers[i].appEx.length > 0
              ? this.state.approvers[i].appEx.map((user: any) => user).join(', ')
              : "Select approver"}
              style={{ overflow: 'inherit' }} id={`approverpeoplepicker-${i}`}>

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
            <td title={optionsApp.find(opt => opt.key === this.state.approvers[i].Type)?.text || "Select Type"} style={{ overflow: 'inherit' }} className="ng-binding">
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
          <td style={{ minWidth: '50px', maxWidth: '50px' }}>{i + 1}</td>
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
      return (
        <tr >
          <td style={{ minWidth: '50px', maxWidth: '50px' }}>
            {i + 1}
          </td>
          <td title={decodeURIComponent(item.Name)}>
            {decodeURIComponent(item.Name)}
          </td>

          <td style={{ textAlign: 'center' }}>
            <span onClick={() => this.OpenFile(item && item, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon icon={faEye} /></span>
            <span onClick={() => this.OpenFile(item && item, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon icon={faDownload} /></span>

            {/* {<a href={item.Path} target="_blank">Link</a>} */}
          </td>
          <td title={moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")} style={{ minWidth: '100px' }} className="text-center">
            {moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")}
          </td>
          {this.state.ShowDeleteicon &&
            <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
              <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeleted(i)}></img>
            </td>
          }
        </tr>
      )
    });
    { console.log("this.state.fileDataAuditee", this.state.copyFilauditee) }
    var fileDataAuditee = this.state.copyFilauditee.map((item: any, i: number) => {
      return (
        <tr >
          <td>{i + 1}</td>
          <td>
            {item.name}
          </td>
          {/* <td>NA</td> */}
          <td className="text-center">{new Date().toLocaleDateString("en-GB", {
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
          <td style={{ minWidth: '50px', maxWidth: '50px' }}>
            {i + 1}
          </td>
          <td title={decodeURIComponent(item.Name)}>
            {decodeURIComponent(item.Name)}
          </td>

          <td style={{ textAlign: 'center' }}>
            <span onClick={() => this.OpenFile(item && item, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon icon={faEye} /></span>
            <span onClick={() => this.OpenFile(item && item, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
              <FontAwesomeIcon icon={faDownload} /></span>

            {/* {<a href={item.Path} target="_blank">Link</a>} */}
          </td>
          <td title={moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")} style={{ minWidth: '100px' }} className="text-center">
            {moment(new Date(item.Uploaded)).format("DD/MMM/YYYY")}
          </td>
          {this.state.ShowDeleteicon &&
            <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
              <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeletedauditee(i)}></img>
            </td>
          }
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
          <td title={`Level ${total - i}`} style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
            {/* {item.Level} */}
            Level {total - i}  {/* Reverse Level */}
          </td>
          <td title={item.AssignedTo} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.AssignedTo}
          </td>
          <td title={item.ActionTakenRole == "LastInitiator" ? "Initiator" : item.ActionTakenRole} style={{ minWidth: "90px", maxWidth: "90px" }}>
            {item.ActionTakenRole == "LastInitiator" ? "Initiator" : item.ActionTakenRole} {/* Divyansh Changes */}
          </td>
          <td title={item.RequesterName} style={{ minWidth: "90px", maxWidth: "90px" }}>
            {item.RequesterName}
          </td>
          <td title={moment(new Date(item.RequestedDate)).format("DD/MMM/YYYY HH:mm:ss")} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.RequestedDate ? moment(new Date(item.RequestedDate)).format("DD/MMM/YYYY HH:mm:ss") : ""}
          </td>
          <td title={item.ActionTakenBy} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.ActionTakenBy}
          </td>
          <td title={moment(new Date(item.ActionTakenOn)).format("DD/MMM/YYYY HH:mm:ss")} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.ActionTakenOn ? moment(new Date(item.ActionTakenOn)).format("DD/MMM/YYYY HH:mm:ss") : ""}
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
          {this.state.Loading || setloading ?

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
                    {this.state.TemplateDoc && this.state.TemplateDoc.length > 0 && (
                      <span
                        onClick={() => this.OpenFile(this.state.TemplateDoc[0], "Open")}
                        style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                      >
                        <div className="btn btn-primary">
                          <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../assets/noun-download-5006210.png")} ></img></div>
                      </span>
                    )}
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
                    <div className="form-group col-md-4 mb-3">
                      <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }}>Approved Report Code:<span className="text-danger1">*</span>
                      </label>
                      <TooltipHost
                        content={this.state.editmemonumberOptions.filter((item: any) => item.value == this.state.editApprovedAuditReport)[0]?.label || ""}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        
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
                        content={this.state.editNCNumberOptions.filter((item: any) => item.value == this.state.editNCNumberID)[0]?.label || ""}
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
                    <div className="form-group col-md-4 mb-3">
                      <TooltipHost
                        content={this.state.editRevisionNo}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <TextField label="Revision Number:" name='editRevisionNo' required value={this.state.editRevisionNo + ""} disabled={true} onChange={this.handleChange}

                        /></TooltipHost>
                    </div>

                  </div>

                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 mb-3">
                      <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }}>Department:<span className="text-danger1">*</span>
                      </label>
                      <TooltipHost
                        content={this.state.editDepartmentOption.filter((item: any) => item.value == this.state.editDepartment)[0]?.label || ""}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <Select
                          options={this.state.editDepartmentOption}
                          value={this.state.departmentselected}
                          isDisabled={this.state.isDisabled}
                          name="Department"
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
                        <TextField label="Close Out Status:" disabled={this.state.isDisabled} name='editCloseOutStatus' required value={this.state.editCloseOutStatus} onChange={this.handleChange}
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
                            <Checkbox label={item.text} disabled={this.state.isDisabled} checked={this.state.editCategoryValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("editCategoryValueIsCheck", item.key as number)} />
                          </div>
                        )
                      }
                      )}
                    </div>
                    <div className="form-group col-md-4" id="SubCategoryCheckbox">
                      <label>Sub Category: <span className={styles.textdanger}>*</span></label>
                      {this.state.editSubCategoryCheckOption.map((item: any) => {
                        return (
                          <div style={{ margin: "2px", padding: "3px" }}>
                            <Checkbox label={item.text} disabled={this.state.isDisabled} checked={this.state.editSubCategoryValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("editSubCategoryValueIsCheck", item.key as number)} />
                          </div>
                        )
                      }
                      )}
                    </div>
                    <div className="form-group col-md-4" id="locationCheckbox">
                      <label>Location: <span className={styles.textdanger}>*</span></label>
                      {this.state.editLocationCheckOption.map((item: any) => {
                        return (
                          <div style={{ margin: "2px", padding: "3px" }}>
                            <Checkbox label={item.text} disabled={this.state.isDisabled} checked={this.state.editLocationValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("editLocationValueIsCheck", item.key as number)}
                            />
                          </div>
                        )
                      }
                      )}
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 " id='AssigntoPeoplepicker'>
                      <TooltipHost
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
                        /></TooltipHost>
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
                      {this.state.fileCount > 0 ?
                        (<span style={{ fontSize: '0.875rem' }} onClick={this._OpenModal} className='newpo'>
                          <FontAwesomeIcon icon={faPaperclip} /> {this.state.fileCount} {this.state.fileCount > 0 ? "files" : "file"} Attached
                        </span>) : ""
                      }
                      {this.state.showDialog && <div id="myModal" className={styles.modal}>
                        <div className={styles.modalcontent}>
                          <span className={styles.close} onClick={() => this._CloseModal()}>&times;</span>
                          <h4 className="font-16 text-dark fw-bold mb-1">Attachment Details</h4>
                          <p className="text-muted font-14 mb-3 fw-400">Below are the attachment details for Non Conformity / Observation</p>

                          <table className='mtbalenew'>
                            <thead>
                              <tr>
                                <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                <th>File Name</th>
                                {this.state.exFiles.length > 0 && <th className="text-center">File Link</th>}
                                <th style={{ minWidth: '100px' }} className="text-center">Upload Date</th>
                                {(this.state.ShowDeleteicon || this.state.copyFil.length > 0) &&
                                  <th className="text-center" style={{ minWidth: "60px", maxWidth: "60px" }}>Action</th>
                                }
                              </tr>
                            </thead>
                            {upFiles}{fileData}
                          </table>
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
                      <DatePicker
                        disabled={this.state.deptSectionDisable}
                        formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                        placeholder="Select a Date"
                        value={this.state.editDate}
                        onSelectDate={(date: Date) => this.setState({ editDate: date })}
                        className={this.state.editErrors?.editDate ? 'textfield-error' : ''}
                      //styles={this.state.editErrors.editDate ? datePickerErrorStyles : {}}
                      />
                    </div>
                    <div className="form-group col-md-4">
                      <Label>
                        Deadline for completion <span className={styles.textdanger}>*</span>
                      </Label>
                      <DatePicker
                        formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                        disabled={this.state.deptSectionDisable}
                        placeholder="Select a Deadline"
                        value={this.state.editDeadlineCompletion}
                        onSelectDate={(date: Date) => this.setState({ editDeadlineCompletion: date })}
                        className={this.state.editErrors?.editDeadlineCompletion ? 'textfield-error' : ''}
                      //styles={this.state.editErrors.editDeadlineCompletion ? datePickerErrorStyles : {}}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-6">
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
                    <div className="form-group col-md-6">
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
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4 mb-3" id='delegatetopeoplepicker'>
                      <TooltipHost
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
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4" id='analyzedBypeoplepicker'>
                      <TooltipHost
                        content={this.state.editAnalyzedBy}
                        calloutProps={{ gapSpace: 0 }}
                        styles={{ root: { display: 'inline-block', width: '100%' } }}
                      >
                        <PeoplePicker
                          disabled={this.state.deptSectionDisable}
                          context={peoplePickerContext}
                          titleText="Analyzed By:"
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
                      </TooltipHost>
                    </div>
                    <div className="form-group col-md-4" id='reviewedBypeoplepicker'>
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
                    </div>
                    <div style={{ position: 'relative' }} className="col-lg-4 mt-1">
                      <label htmlFor="Attchments" style={{ marginRight: "10px" }}>Auditee Attachments </label>

                      <input
                        disabled={this.state.deptSectionDisable}
                        className="form-control"
                        type="file" name="myFile" onChange={(e) => this.handleAuditeeFileChange(e, this)} id="newfile" multiple
                      //className={`form-control ${this.state.errors?.Attachments} ? 'textfield-error' : ''`}
                      />
                      {this.state.fileCountauditee > 0 ?
                        (<span style={{ fontSize: '0.875rem' }} onClick={this._OpenModalauditee} className='newpo'>
                          <FontAwesomeIcon icon={faPaperclip} /> {this.state.fileCountauditee} {this.state.fileCountauditee > 0 ? "files" : "file"} Attached
                        </span>) : ""
                      }
                      {this.state.showDialogauditee && (
                        <div id="myModal" className={styles.modal}>
                          <div className={styles.modalcontent}>
                            {/* Close button */}
                            <span className={styles.close} onClick={() => this._CloseModalauditee()}>&times;</span>

                            {/* Modal title and subtitle */}
                            <h4 className="font-16 text-dark fw-bold mb-1">Attachment Details</h4>
                            <p className="text-muted font-14 mb-3 fw-400">Below are the attachment details for Non Conformity / Observation</p>

                            {/* Table */}
                            <table className={styles.mtbalenew}>
                              <thead>
                                <tr>
                                  <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                  <th>File Name</th>
                                  {this.state.exFilesauditee.length > 0 && <th className="text-center">File Link</th>}
                                  <th style={{ minWidth: '100px' }} className="text-center">Upload Date</th>
                                  {(this.state.ShowDeleteicon || this.state.copyFilauditee.length > 0) && <th className="text-center">Action</th>}
                                </tr>
                              </thead>
                              <tbody >
                                {upFilesauditee}
                                {fileDataAuditee}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-0'>Non Conformity / Observation Close Out Details</h3></div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
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
                          disabled={this.state.deptSectionDisable}
                          multiline rows={3} onChange={this.handleChange}
                          className={this.state.editErrors?.editCorrectiveActionImplementedOn ? 'textfield-error' : ''}
                        // styles={{
                        //   fieldGroup: {
                        //     backgroundColor: this.state.editErrors.editCorrectiveActionImplementedOn ? "#ffcccb" : "white", // Red tint for errors
                        //   }
                        // }}
                        />
                      </TooltipHost>
                    </div>
                  </div>
                  {(showimsupdated && isdisableims) &&
                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: '2rem', marginTop: '1rem' }}>

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
                  {console.log("isFinalApprover", this.state.IsFinalapprover, IsAnalyzedBy, isdisablefinal)}
                  {showcorrectionappicable && isdisablefinal &&
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
          {(this.state.editReviewedBySubmitStatus == "Yes" && this.state.editDelegateToId == null && (!this.state.Loading || !setloading)) ||
            (this.state.editLastAssignedToSubmitStatus == "Yes" && this.state.editDelegateToId != null && (!this.state.Loading || !setloading)) ?
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
                              <th style={{ borderBottomLeftRadius: "0px" }}>Role</th>
                              <th style={{ minWidth: '70px', maxWidth: '70px' }} >Level</th>
                              <th >Approver Name</th>
                              <th >Approval Criteria</th>
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

          {(((this.state.showApprove === true || this.state.showReject === true || showreworkremarks || (showimsupdated && !isdisableims) || (showcorrectionappicable && !isdisablefinal))
            && (!this.state.Loading || !setloading) && this.state.editCurrentUserRole != "FirstInitiator")
            ||
            ((this.state.editReviewedBySubmitStatus == "Yes" && this.state.editDelegateToId == null && (!this.state.Loading || !setloading)) ||
              (this.state.editLastAssignedToSubmitStatus == "Yes" && this.state.editDelegateToId != null && (!this.state.Loading || !setloading)))) && !this.state.showDraft
            ?
            <section style={{ justifyContent: 'left', textAlign: 'left' }} id="approvalSection" className='card card-body'>

              <TextField label="Remarks" required={(Approveclicked || Rejectclicked || Reworkclicked || this.state.showApprove === true || this.state.showReject === true) ? true : false} name="remarks" value={this.state.remarks} multiline rows={3} onChange={this.handleChange}

                className={this.state.editErrors?.remarks ? 'textfield-error' : ''}// styles={{
              />

              {(showimsupdated && !isdisableims) &&
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: '2rem', marginTop: '1rem' }}>

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
              {console.log("isFinalApprover", this.state.IsFinalapprover, IsAnalyzedBy, isdisablefinal)}
              {showcorrectionappicable && !isdisablefinal &&
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

            </section> : null
          }
          {/* Vishnu Changes  */}

          {this.state.showSubmit && (!this.state.Loading || !setloading) && ((this.state.editSubmitStatus == "No" && RequesterEmail == CurrentuserEmail) ||
            (this.state.editSubmitStatus == "Yes" && this.state.edType == "edit" && this.state.editStatus !== "Rework" && this.state.editAssignToEmail == CurrentuserEmail && this.state.editCurrentUserRole != "DelegateTo") ||
            (this.state.editSubmitStatus == "Yes" && this.state.edType == "edit" && this.state.editStatus !== "Rework" && this.state.editDelegateToEmail == CurrentuserEmail && this.state.editCurrentUserRole == "DelegateTo") ||
            (this.state.editStatus == "Rework" && RequesterEmail == CurrentuserEmail && this.state.editCurrentUserRole == "FirstInitiator") ||
            (this.state.edType == "edit" && this.state.editStatus == "Rework" && this.state.editAssignToEmail == CurrentuserEmail && this.state.editCurrentUserRole == "FirstAssignedTo") ||
            (this.state.edType == "edit" && this.state.editStatus == "Rework" && this.state.editDelegateToEmail == CurrentuserEmail && this.state.editCurrentUserRole == "DelegateTo")

          ) &&

            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }} className='newnbu'>

              {this.state.showDraft && RequesterEmail == CurrentuserEmail &&
                //this.state.editStatus != "Rework"&&

                <PrimaryButton text="Save as Draft" onClick={() => this.handleDraft("draft")} />

              }

              <PrimaryButton text="Submit" onClick={() => this.handleSubmit("submit")} />
              {(this.state.editCurrentUserRole === "FirstAssignedTo" || this.state.editCurrentUserRole === "DelegateTo") &&
                <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} />
              }
              <DefaultButton text="Cancel" onClick={() => this.cancelRequest("Edcmain")} />

            </div>
          }
          {console.log("ApproverEmailApproverEmail", ApproverEmail, CurrentuserEmail)}
          {this.state.showApprove && (!this.state.Loading || !setloading) &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>
              {ApproverEmail == CurrentuserEmail &&
                <>
                  <PrimaryButton text="Approve" onClick={() => this._approveRequest("Approve")} />
                  <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} />
                </>
              }

              <DefaultButton text="Cancel" onClick={() => this.cancelRequest("myapproval")} />

            </div>
          }
          {console.log("ApproverEmailforward", ApproverEmail, CurrentuserEmail)}
          {this.state.showForward && (!this.state.Loading || !setloading) &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>
              {RequesterEmail == CurrentuserEmail &&
                <>
                  <PrimaryButton text="Forward" onClick={() => this.handleForward("Forward")} />
                  <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} />
                </>
              }
              <a href='#/form'> <DefaultButton onClick={() => this.cancelRequest("myapproval")}>Cancel</DefaultButton></a>

            </div>
          }
          {console.log("Approverreject", this.state.showReject, ApproverEmail, CurrentuserEmail)}
          {this.state.showReject && (!this.state.Loading || !setloading) &&

            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>
              {ApproverEmail == CurrentuserEmail &&
                <>
                  <PrimaryButton text="Approve" onClick={() => this._approveRequest("Approve")} />
                  <PrimaryButton text="Reject" onClick={() => this._rejectRequest("Reject")} />
                  <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} />
                </>
              }
              <DefaultButton text="Cancel" onClick={() => this.cancelRequest("myapproval")} />

            </div>
          }
          {this.state.edType === "view" && (!this.state.Loading || !setloading)
            &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

              <DefaultButton text="Cancel" onClick={() => this.cancelRequest("Edcmain")} />

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
        </div>
      </section >
    )
  }
}
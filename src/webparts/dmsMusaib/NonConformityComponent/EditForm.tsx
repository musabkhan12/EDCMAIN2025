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
  Label
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

import moment from 'moment';
import CustomBreadcrumb from '../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
let Approvallistitemid = 0;
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
  mainItemId?: any;
  edType?: string;
  approvalItemId?: string;
  editDepartmentOption: IDropdownOption[];
  editDepartment: string | number;
  editdepartmentCode: string;
  editserialNo: number;
  notUpdateDepartmentCode: string;
  notUpdateSerialNo: number;
  editCriteria: string;
  editCloseOutStatus: string;
  editCategoryCheckOption: IDropdownOption[];
  editCategoryValueIsCheck: number[];
  editSubCategoryCheckOption: IDropdownOption[];
  editSubCategoryValueIsCheck: number[];
  editLocationCheckOption: IDropdownOption[];
  editLocationValueIsCheck: number[];
  editAssignTo: string;
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
  showDialog: boolean;
  copyFil: any[];
  fileCount: number;
  exFiles: any[];
  fileDeleteId: any[];
  files: FileList;
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

  { key: 'All', text: 'All' },
  { key: 'One', text: 'One' }

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
      mainItemId: '',
      edType: this.props.edType,
      approvalItemId: this.props.approvalItemId,
      editDepartmentOption: [],
      editDepartment: "",
      editdepartmentCode: "",
      editserialNo: 0,
      notUpdateDepartmentCode: "",
      notUpdateSerialNo: 0,
      editCriteria: "",
      editCloseOutStatus: "",
      editCategoryCheckOption: [],
      editCategoryValueIsCheck: [],
      editSubCategoryCheckOption: [],
      editSubCategoryValueIsCheck: [],
      editLocationCheckOption: [],
      editLocationValueIsCheck: [],
      editAssignTo: "",
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
      approvers: [{ Role: "", Level: "", Name: "", Type: "", Index: 0, appEx: "", itemId: "" }],
      optionsRole: [],
      apprDelId: [],
      indApp: 0,
      mainListId: 0,
      formNameId: 0,
      reqRolId: 0,
      ncItemId: null,
      remarks: "",
      showDialog: false,
      copyFil: [],
      fileCount: 0,
      exFiles: [],
      fileDeleteId: [],
      files: {} as FileList,
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
    this.toBeDeleted = this.toBeDeleted.bind(this);
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
  private Breadcrumb = [
    {
      MainComponent: "My Request",
      MainComponentURl: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EDCMAIN.aspx`,
    },
    {
      ChildComponent: "Non Conformity",
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
  public changeDepartment = (_event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ editDepartment: item.key, editdepartmentCode: item.data.departmentCode });
  };

  public handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      [name]: value,
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

  public async componentDidMount() {
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

    if (id) {
      this.setState({ mainItemId: id }, () => {
        // This will run AFTER the state update is completed
        // alert("Updated mainItemId: " + this.state.mainItemId);
        this.getListData(); // Fetch list data after updating state
      });
    }
    if (editType) {
      this.setState({ edType: editType })
    }
    // await this.getListData();
    await this.getDepartment();
    await this.getDataRoles();
    await this.getMainListName();
    await this.getRequestorRole();
    await this.getFormName();
    if (this.state.edType === "edit") {
      this.setState({ showApprove: false });
      this.setState({ showSubmit: true });
      this.setState({ showForward: false });
      this.setState({ showReject: false });
      this.setState({ forwarDisable: false });
      if (this.state.editSubmitStatus === 'No') {
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
    else if (this.state.edType === "view") {
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
    else if (this.state.edType === "approve") {
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
        this.setState({ showReject: true });
      }
      if (this.state.editStatus == "Approved" || this.state.editStatus == "Rejected") {
        this.setState({ showReject: false });
      }
    }
  }

  public async getListData() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const Items: any = await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId)
        .select("*, Category/Id, Category/Title, SubCategory/Id, Location/Id, Location/Title, SubCategory/Title, AssignedTo/Id, AssignedTo/Title, DelegateTo/Id, DelegateTo/Title, AnalyzedBy/Id, AnalyzedBy/Title, ReviewedBy/Id, ReviewedBy/Title, PersonAssigned/Id, PersonAssigned/Title")
        .expand("Category, SubCategory, Location, AssignedTo, DelegateTo, AnalyzedBy, ReviewedBy, PersonAssigned")();
      console.log(Items);
      this.setState({
        ncItemId: Items.Id,
        editDepartment: Items.DepartmentId,
        editCriteria: Items.Criteria,
        editCloseOutStatus: Items.CloseOutStatus,
        editCategoryValueIsCheck: Items.Category ? Items.Category.map((cat: any) => cat.Id) : [],
        editSubCategoryValueIsCheck: Items.SubCategory ? Items.SubCategory.map((sub: any) => sub.Id) : [],
        editLocationValueIsCheck: Items.Location ? Items.Location.map((loc: any) => loc.Id) : [],
        editAssignToId: Items.AssignedTo ? Items.AssignedTo.Id : null,
        editAssignTo: Items.AssignedTo ? Items.AssignedTo.Title : null,
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
        notUpdateDepartmentCode: Items.DocumentCode,
        notUpdateSerialNo: Items.SerialNumber,
      });
      //Process Approval List
      const apprItems = await sp.web.lists
        .getByTitle("ProcessApprovalList")
        .items.select(
          "*",
          "AssignedTo/Title,ActionTakenRole/Title,RequesterName/Title,ActionTakenBy/Title"
        )
        .expand("AssignedTo,ActionTakenRole,RequesterName,ActionTakenBy")
        .filter(
          "ListItemId eq '" +
            this.state.mainItemId +
            "' and ProcessName eq 'Non Conformity'"
        )
        .orderBy("Id", false)();
      var cnt: any = 0;
      var appItems: any[] = [];
      if (apprItems.length > 0) {
        apprItems.forEach(async function (itm: any) {
          //Audit Report
          var objToAdd: any = {};
          objToAdd["Level"] = itm.Level;
          objToAdd["AssignedTo"] = itm.AssignedTo.Title;
          objToAdd["RequesterName"] = itm.RequesterName.Title;
          objToAdd["ActionTakenRole"] = itm.ActionTakenRole?.Title||"";
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
          obJFiles.push({ "Name": item.File.Name, "type": "old", "Id": item.Id, "Uploaded": new Date(item.Modified).getDate() + "/" + new Date(item.Modified).getMonth() + "/" + new Date(item.Modified).getFullYear(), "Path": item.EncodedAbsUrl })
        })
        fCount = upFiles.length;
        this.setState({ exFiles: obJFiles, fileCount: fCount });
      }
      //AllProcessApproval Table data
      const approvalItems = await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.select("*", "Approvers/Name").expand("Approvers")
        .filter("MainListID eq '" + this.state.mainItemId + "'and ProcessName eq 'Non Conformity'")
        .orderBy("Level")();
      var allApp: any[] = [];
      var cnt: any = 0;
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
        this.setState({ approvers: allApp });
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
        var serialNo: any = "000";
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
    const listItems = await sp.web.lists.getByTitle("FormNameMaster").items.filter("FormName eq 'NonConformityList'")();
    this.setState({ formNameId: listItems[0].Id })

  }
  private async getRequestorRole() {
    const sp = spfi().using(SPFx(this.props.context));
    const listItems = await sp.web.lists.getByTitle("RequesterRoleMaster").items.filter("Role eq 'Initiator'")();
    this.setState({ reqRolId: listItems[0].Id })
  }
  public cancelRequest() {
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
          window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          // window.location.reload();
        });
      }
    });
  }
  public async getDepartment() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const deptItems = await sp.web.lists.getByTitle("DepartmentMasterList").items();
      const options = deptItems.map((item: {
        DepartmentCode: any; Title: string; Id: number
      }) => ({
        key: item.Id,
        text: item.Title,
        data: { departmentCode: item.DepartmentCode },
      }));
      this.setState({ editDepartmentOption: options });
      // Find the selected department's departmentCode and set it
      const selectedDeptArray = options.filter(opt => opt.key === this.state.editDepartment);
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
      if (!this.state.editDepartment) editErrors.editDepartment = "Department is required";
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
      if (!this.state.editAssignTo) editErrors.editAssignTo = "AssignTo is required";
      if (!this.state.editDueDate) editErrors.editDueDate = "dueDate is required";
      if (!this.state.fileCount && this.state.exFiles.length == 0) {
        editErrors.Attchments = "Attachments are required";
      }
      if (!this.state.editProblemDescription) editErrors.editProblemDescription = "Problem Description is required";
    }
    else if ((this.state.editCurrentUserRole == "FirstAssignedTo" && this.state.editDelegateToId == null) || (this.state.editCurrentUserRole == "DelegateTo" && this.state.editDelegateToId != null)) {
      if (!this.state.editPersonAssigned) editErrors.editPersonAssigned = "PersonAssigned is required";
      if (!this.state.editDate) editErrors.editDate = "Date is required";
      if (!this.state.editDeadlineCompletion) editErrors.editDeadlineCompletion = "Deadline for completion is required";
      if (!this.state.editCorrection) editErrors.editCorrection = "Correction is required";
      if (!this.state.editRootCause) editErrors.editRootCause = "Root Cause is required";
      if (!this.state.editCorrectiveAction) editErrors.editCorrectiveAction = "Corrective Action is required";
      if (!this.state.editAnalyzedBy) editErrors.editAnalyzedBy = "Analyzed By is required";
      if (!this.state.editReviewedBy) editErrors.editReviewedBy = "Reviewed By is required";
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
    let editErrors: { [key: string]: string } = {};
    if (this.state.approvers.length === 1) {
      if (!this.state.approvers[0].Role) {
        editErrors.approvers = "Approver role is required";
      }
    }
    this.setState({ editErrors });
    return Object.keys(editErrors).length === 0;
  };
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
  
    if (!this.state.editDepartment) {
      editErrors.editDepartment = "Department is required";
      this.setState({ editErrors });
      Swal.fire('Please select a department.');
      return false;
    }
  
    this.setState({ editErrors });
    return true;
  };
  public _approveRequest = (formsubmode: string) => {
    if (this.validateFormRemark()) {
      this.approveRequest(formsubmode);
      console.log("Form Data:", this.state);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public _rejectRequest = (formsubmode: string) => {
    if (this.validateFormRemark()) {
      this.rejectRequest(formsubmode);
      console.log("Form Data:", this.state);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
  public _reworkRequest = (formsubmode: string) => {
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
  //< ------- End Draft Validation -------->

  private onRoleChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Role = item.key;
    this.setState({ approvers: this.state.approvers });
  }

  private onTypeChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Type = item.text;
    this.setState({ approvers: this.state.approvers });
  }

  private _getPeoplePickerItemsApp(items: any[], i: number) {
    var arr: any[];
    arr = [];
    items.forEach(function (it) {
      arr.push(it.id);
    })
    this.state.approvers.filter(function (it) {
      if (it.Index == i) {
        it.Name = arr
      }
    });
    this.setState({ approvers: this.state.approvers });
  }

  private addApprover() {
    var itm = this.state.indApp + 1;
    this.setState({ indApp: itm });
    this.state.approvers.push({ Role: "", Level: "", Name: "", Index: itm, itemId: "" });
    this.setState({ approvers: this.state.approvers });
  }

  private deleteItemApp(i: number) {
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

  public async updateData(_editsubmitStatus: string, firstInitiatorSubmitStatus: string, firstAssignedToSubmitStatus: string, delegateToSubmitStatus: string, analyzedBySubmitStatus: string, reviewedBySubmitStatus: string, lastAssignedToSubmitStatus: string, lastInitiatorSubmitStatus: string, currentUserRole: string, reworkById: any, serialNumber: number, documentCode: string) {
    const sp = spfi().using(SPFx(this.props.context));
    await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId).update({
      DepartmentId: this.state.editDepartment || null,
      Criteria: this.state.editCriteria,
      CloseOutStatus: this.state.editCloseOutStatus,
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
      Status: _editsubmitStatus == "Rework" ? "Rework" : "Pending",
      IsRework: _editsubmitStatus == "Rework" || _editsubmitStatus == "Reject" ? "Yes" : "No",
      ReworkById: reworkById,
      SerialNumber: serialNumber,
      DocumentCode: documentCode,
    })
  }

  //Update Function
  private async _updateSubmitData(_editsubmitStatus: string) {
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
      documentCode = 'NC/' + this.state.editdepartmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.editserialNo;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode);
        if (_editsubmitStatus == "submit") {
          if (currentUserRole == "AnalyzedBy" || currentUserRole == "DelegateTo") {
            // alert("Approval Item ID" + approvalItemId)
            if (approvalItemId == null || approvalItemId == undefined || approvalItemId == "") {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: editProblemDescription,
              });
            } else {
              sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(approvalItemId)).update({
                Status: "Approved",
                ActionTakenById: currentUserID,
                ActionTakenOn: new Date(),
                Remark: editProblemDescription,
              });
            }

          }
          else {
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
        Swal.fire({
          title: cText + " Successfully.",
          icon: "success"
        }).then(() => {
           window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          // window.location.reload();
        });
      }
    })
      .catch(error => {
        console.error("Error while saving:", error);
      });
  }
  //Forward Call
  public async forwardRequest(_editsubmitStatus: string) {
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
        currentUserRole = "LastInitiator";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        documentCode = this.state.notUpdateDepartmentCode;
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
        currentUserRole = "LastInitiator";
        reworkById = null;
        serialNumber = this.state.notUpdateSerialNo;
        documentCode = this.state.notUpdateDepartmentCode;
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
    }).then(async function (val) {
      if (val.isConfirmed) {
        await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode);
        sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(approvalItemId)).update({
          Status: "Approved",
          ActionTakenById: currentUserID,
          ActionTakenOn: new Date(),
          Remark: remarks,
        });
        //Forward Button
        if (approvers.length > 0) {
          var maxLength = approvers.length;
          approvers.forEach(function (it: any, val: any) {
            if (it.itemId == 0) {
              sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
                Title: "Title",
                MainListNameId: _self.state.mainListId,
                ApproverRoleId: it.Role,
                Level: val + 1,
                LevelType: it.Type,
                SubmitStatus: "Yes",
                Maxlevel: maxLength,
                ContentTitle: _self.state.editProblemDescription,
                RequestId: documentCode,
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
                Level: val + 1,
                LevelType: it.Type,
                SubmitStatus: "Yes",
                Maxlevel: maxLength,
                ContentTitle: _self.state.editProblemDescription,
                RequestId: this.state.notUpdateDepartmentCode,
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
        Swal.fire({
          title: "Forwarded Successfully.",
          icon: "success"
        }).then(() => {
           window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          // window.location.reload();
        });
      }
    });

  }
  //Approver call
  public async approveRequest(_editsubmitStatus: string) {
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
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
    }).then(async function (val) {
      if (val.isConfirmed) {
        if (editLastInitiatorSubmitStatus != "Yes") {
          await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode);
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
          Swal.fire({
            title: "Approved Successfully.",
            icon: "success"
          }).then(() => {
             window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
            //window.location.reload();
          });
        } else {
          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(approvalItemId)).update({

            Status: "Approved",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
          });
          Swal.fire({
            title: "Approved Successfully.",
            icon: "success"
          }).then(() => {
             window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
            // window.location.reload();
          });
        }

      }
    });

  }
  //Reject call
  private async rejectRequest(_editsubmitStatus: string) {
    const { remarks } = this.state
    const { currentUserID, approvalItemId, context } = this.props
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: 'Do you want to reject this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(function (val) {
      if (val.isConfirmed) {
        if (approvalItemId == undefined || approvalItemId == null || approvalItemId == "") {
          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
            Status: "Rejected",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
          });
          Swal.fire({
            title: "Rejected Successfully.",
            icon: "success"
          }).then(() => {
             window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
            // window.location.reload();
          });
        } else {
          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(approvalItemId)).update({
            Status: "Rejected",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
          });
          Swal.fire({
            title: "Rejected Successfully.",
            icon: "success"
          }).then(() => {
             window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
            // window.location.reload();
          });


        }

      }
    });

  }
  //Rework Call
  private async reworkRequest(_editsubmitStatus: string) {
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
    let documentCode = "";

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
        reworkById = this.state.editAssignToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        documentCode = this.state.notUpdateDepartmentCode;
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
        reworkById = this.state.editAssignToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        documentCode = this.state.notUpdateDepartmentCode;
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
        reworkById = this.state.editDelegateToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        documentCode = this.state.notUpdateDepartmentCode;
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
        reworkById = this.state.editDelegateToId || null;
        serialNumber = this.state.notUpdateSerialNo;
        documentCode = this.state.notUpdateDepartmentCode;
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
        documentCode = this.state.notUpdateDepartmentCode;
      }
    }
    //<-------- End Case2 --------->


    const { remarks } = this.state;
    const { currentUserID, approvalItemId, context } = this.props;
    const sp = spfi().using(SPFx(this.props.context));
    const { updateData } = this;
    Swal.fire({
      title: 'Do you want to rework this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(async function (val) {
      if (val.isConfirmed) {
        await updateData(_editsubmitStatus, firstInitiatorSubmitStatus, firstAssignedToSubmitStatus, delegateToSubmitStatus, analyzedBySubmitStatus, reviewedBySubmitStatus, lastAssignedToSubmitStatus, lastInitiatorSubmitStatus, currentUserRole, reworkById, serialNumber, documentCode);
        if (approvalItemId == undefined || approvalItemId == null || approvalItemId == "") {

          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Approvallistitemid).update({
            Status: "Rework",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
            IsRework: "Yes"
          });

          Swal.fire({
            title: "Sent for Rework.",
            icon: "success"
          }).then(() => {
             window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
            // window.location.reload();
          });
        } else {

          sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(approvalItemId)).update({
            Status: "Rework",
            ActionTakenById: currentUserID,
            ActionTakenOn: new Date(),
            Remark: remarks,
            IsRework: "Yes"
          });

          Swal.fire({
            title: "Sent for Rework.",
            icon: "success"
          }).then(() => {
             window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
            // window.location.reload();
          });
        }


      }
    });
  }


  public render(): React.ReactElement<IAuditPlanProps> {
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
    var approval = this.state.approvers.map((item: any, i: number) => {
      return (
        <tr className='tblCls'>
          <td>
            <TextField value={(i + 1).toString()} disabled={true} className={styles.width} ></TextField>
          </td>
          <td>
            <Dropdown disabled={this.state.forwarDisable} placeholder="Select options" selectedKey={this.state.approvers[i].Role} options={this.state.optionsRole
              .filter((opt) =>
                !this.state.approvers.some((app, index) => index !== i && app.Role === opt.key)
              )}
              onChange={(e, itm: IDropdownOption) => this.onRoleChange(e, itm, i)}
              styles={{
                title: {
                  backgroundColor: this.state.editErrors.approvers ? "#ffcccb" : "white", // Light red when error
                }
              }} />
          </td>
          <td>
            <TextField value={(i + 1).toString()} disabled={true}></TextField>
          </td>
          <td>
            <PeoplePicker
              context={peoplePickerContext}
              personSelectionLimit={5}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              disabled={this.state.forwarDisable}
              ensureUser={true}
              defaultSelectedUsers={this.state.approvers[i].appEx ? this.state.approvers[i].appEx : []}
              onChange={(e) => this._getPeoplePickerItemsApp(e, i)}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              styles={{
                root: {
                  backgroundColor: this.state.editErrors.approvers ? "#ffcccb" : "white",
                }
              }}
            />
          </td>
          <td>
            <Dropdown disabled={this.state.forwarDisable} placeholder="Select options" selectedKey={this.state.approvers[i].Type} options={optionsApp} onChange={(e, itm: IDropdownOption) => this.onTypeChange(e, itm, i)}
              styles={{
                title: {
                  backgroundColor: this.state.editErrors.approvers ? "#ffcccb" : "white", // Light red when error
                }
              }} />
          </td>
          <td><button type="button" onClick={(e) => this.deleteItemApp(i)}>Delete</button></td>
        </tr>
      )
    })
    var fileData = this.state.copyFil.map((item: any, i: number) => {
      return (
        <tr style={{ display: 'table', width: '100%' }}>
          <td>
            {item.name}
          </td>
          <td></td>
          <td>{new Date().getDate() + "/" + new Date().getMonth() + "/" + new Date().getFullYear()}</td>
          <td>
            <button type="button" onClick={(e) => this.removeFiles(i)}>Delete</button>
          </td>
        </tr>
      )
    });
    var upFiles = this.state.exFiles.map((item: any, i: number) => {
      return (
        <tr style={{ display: 'table', width: '100%' }}>
          <td>
            {item.Name}
          </td>
          <td>
            {<a href={item.Path} target="_blank">Link</a>}
          </td>
          <td>
            {item.Uploaded}
          </td>
          <td>
            <button type="button" onClick={(e) => this.toBeDeleted(i)} disabled={this.state.isDisabled}>Delete</button>
          </td>
        </tr>
      )
    });
    var auditHistory = this.state.apprItems.map((item: any, i: number) => {
      return (
        <tr>
          <td style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
            {i + 1}
          </td>
          <td style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
            {item.Level}
          </td>
          <td style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.AssignedTo}
          </td>
          <td style={{ minWidth: "90px", maxWidth: "90px" }}>
            {item.ActionTakenRole} {/* Divyansh Changes */}
          </td>
          <td style={{ minWidth: "90px", maxWidth: "90px" }}>
            {item.RequesterName}
          </td>
          <td style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.RequestedDate ? moment(new Date(item.RequestedDate)).format("DD/MMM/YYYY HH:mm:ss") : ""}
          </td>
          <td style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.ActionTakenBy}
          </td>
          <td style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.ActionTakenOn ? moment(item.ActionTakenOn).format("DD/MMM/YYYY HH:mm:ss") : ""}
          </td>
          <td style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.Remarks}
          </td>
          <td style={{ minWidth: '70px', maxWidth: '70px' }}>
            {item.Status}
          </td>
        </tr>
      )

    });
    return (
      <section style={{ padding: '10px 0px' }}>
        <div className={styles.welcome}>
          <div className="row">
            <div className="col-lg-4 newbread">
              <CustomBreadcrumb Breadcrumb={this.Breadcrumb} />
            </div>

          </div>
          <section className='card card-body' >
            <fieldset>
              <form>
                {/* Start save as draft */}
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                  <div className="form-group col-md-12"><h3 className='text-dark text-left font-16 fw-bold mb-3'>Problem Details</h3></div>
                </div>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                  <div className="form-group col-md-4">
                    <Dropdown
                      required
                      disabled={this.state.isDisabled}
                      label="Department:"
                      options={this.state.editDepartmentOption}
                      defaultSelectedKey={this.state.editDepartment}
                      selectedKey={this.state.editDepartment}
                      onChange={this.changeDepartment}
                      styles={{
                        title: {
                          backgroundColor: this.state.editErrors.editDepartment ? "#ffcccb" : "white", // Light red when error
                        }
                      }}
                    />
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="Criteria:" disabled={this.state.isDisabled} name='editCriteria' required value={this.state.editCriteria} onChange={this.handleChange}
                      styles={{
                        fieldGroup: {
                          backgroundColor: this.state.editErrors.editCriteria ? "#ffcccb" : "white",
                        }
                      }}
                    />
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="Close Out Status:" disabled={this.state.isDisabled} name='editCloseOutStatus' required value={this.state.editCloseOutStatus} onChange={this.handleChange}
                      styles={{
                        fieldGroup: {
                          backgroundColor: this.state.editErrors.editCloseOutStatus ? "#ffcccb" : "white",
                        }
                      }}
                    />
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
                  <div className="form-group col-md-4">
                    <PeoplePicker
                      context={peoplePickerContext}
                      titleText="Assigned To:"
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
                  </div>
                  <div className="form-group col-md-4">
                    <Label>
                      Due Date <span className={styles.textdanger}>*</span>
                    </Label>
                    <DatePicker
                      disabled={this.state.isDisabled}
                      formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                      placeholder="Select a Due Date"
                      value={this.state.editDueDate}
                      onSelectDate={(date: Date) => this.setState({ editDueDate: date })}
                      styles={this.state.editErrors.editDueDate ? datePickerErrorStyles : {}}
                    />
                  </div>
                  <div style={{ position: 'relative' }} className="col-lg-4 mt-1">
                    <label htmlFor="Attchments" style={{ marginRight: "10px" }}>Attachments <span className={styles.textdanger}>*</span></label>

                    <input disabled={this.state.isDisabled} className="form-control" type="file" name="myFile" onChange={(e) => this.handleFileChange(e, this)} id="newfile" multiple
                      style={{
                        backgroundColor: this.state.editErrors.Attchments ? "#ffcccb" : "white",
                      }} />
                    <span onClick={this._OpenModal} className='newpo'>    {this.state.fileCount}</span>
                    {this.state.showDialog && <div id="myModal" className={styles.modal}>
                      <div className={styles.modalcontent}>
                        <span><b>Attachment Details</b></span>
                        <br />
                        <span>Below are the attachment details for the Initiative</span>
                        <span className={styles.close} onClick={e => this._CloseModal()}>&times;</span>
                        <table className='mtbalenew'>
                          <thead>
                            <tr style={{ display: 'table', width: '100%' }}>
                              <th>File Name</th>
                              <th>File Link</th>
                              <th>Upload Date</th>
                              <th>Delete</th>
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
                    <TextField label="Problem Description:"
                      required
                      name='editProblemDescription'
                      value={this.state.editProblemDescription}
                      multiline rows={5}
                      onChange={this.handleChange}
                      disabled={this.state.isDisabled}
                      styles={{
                        fieldGroup: {
                          backgroundColor: this.state.editErrors.editProblemDescription ? "#ffcccb" : "white", // Red tint for errors
                        }
                      }}
                    />
                  </div>
                </div>
                {/* End Save as draft */}
              </form>
            </fieldset>
          </section>
          {this.state.editSubmitStatus == "Yes" ? //this.state.editFirstInitiatorSubmitStatus == "Yes" || this.state.editFirstAssignedToSubmitStatus == "Yes" ?
            <section className='card card-body mt-2' >
              <fieldset>
                <form>
                  {/* Section 2 */}
                  <div className="row">
                    <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-3'>To be filled by Department Head</h3></div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4">
                      <PeoplePicker
                        context={peoplePickerContext}
                        disabled={this.state.showDelegate}
                        titleText="Person Assigned:"
                        personSelectionLimit={1}
                        required={true}
                        onChange={this._handlePeoplePickerChange("editPersonAssigned", "editPersonAssignedId")}
                        defaultSelectedUsers={[this.state.editPersonAssigned]}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        ensureUser={true}
                        styles={{
                          root: {
                            backgroundColor: this.state.editErrors.editPersonAssigned ? "#ffcccb" : "white",
                          }
                        }}
                      />
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
                        styles={this.state.editErrors.editDate ? datePickerErrorStyles : {}}
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
                        styles={this.state.editErrors.editDeadlineCompletion ? datePickerErrorStyles : {}}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-6">
                      <TextField label="Correction(Immediate Steps to stop the problem):" required name='editCorrection'
                        value={this.state.editCorrection}
                        multiline rows={3}
                        disabled={this.state.deptSectionDisable}
                        onChange={this.handleChange}
                        styles={{
                          fieldGroup: {
                            backgroundColor: this.state.editErrors.editCorrection ? "#ffcccb" : "white", // Red tint for errors
                          }
                        }}
                      />
                    </div>
                    <div className="form-group col-md-6">
                      <TextField label="Root Cause:" required name='editRootCause'
                        value={this.state.editRootCause} multiline rows={3}
                        onChange={this.handleChange}
                        disabled={this.state.deptSectionDisable}
                        styles={{
                          fieldGroup: {
                            backgroundColor: this.state.editErrors.editRootCause ? "#ffcccb" : "white", // Red tint for errors
                          }
                        }}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-12">
                      <TextField label="Corrective Action (Action to eliminate the root cause):"
                        required
                        name='editCorrectiveAction'
                        value={this.state.editCorrectiveAction} multiline rows={3}
                        onChange={this.handleChange}
                        disabled={this.state.deptSectionDisable}
                        styles={{
                          fieldGroup: {
                            backgroundColor: this.state.editErrors.editCorrectiveAction ? "#ffcccb" : "white", // Red tint for errors
                          }
                        }}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-4">
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
                    </div>
                    <div className="form-group col-md-4">
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
                            backgroundColor: this.state.editErrors.editAnalyzedBy ? "#ffcccb" : "white",
                          }
                        }}
                      />
                    </div>
                    <div className="form-group col-md-4">
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
                            backgroundColor: this.state.editErrors.editReviewedBy ? "#ffcccb" : "white",
                          }
                        }}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-0'>Problem Close Out Details</h3></div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                    <div className="form-group col-md-12">
                      <TextField label="Corrective Action Implemented On:"
                        required
                        name='editCorrectiveActionImplementedOn'
                        value={this.state.editCorrectiveActionImplementedOn}
                        disabled={this.state.deptSectionDisable}
                        multiline rows={3} onChange={this.handleChange}
                        styles={{
                          fieldGroup: {
                            backgroundColor: this.state.editErrors.editCorrectiveActionImplementedOn ? "#ffcccb" : "white", // Red tint for errors
                          }
                        }}
                      />
                    </div>
                  </div>
                  {/* Approve/Rework */}
                </form>
              </fieldset>
            </section> : null}
          {/* Approval Table */}
          {(this.state.editReviewedBySubmitStatus == "Yes" && this.state.editDelegateToId == null) || (this.state.editLastAssignedToSubmitStatus == "Yes" && this.state.editDelegateToId != null) ? (<section className={styles.sec}>
            <fieldset disabled={this.state.forwarDisable}>
              <form>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row'>
                  <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark text-left font-16 fw-bold mb-3'>Forward Detail</h3></div>
                </div>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                  <div className="container mt-4">
                    <button type="button" onClick={this.addApprover}>Add</button>
                    <table id="tblAppr" className='mtbalenew'>
                      <tbody>
                        <tr><td>SI No.</td>
                          <td>Role</td>
                          <td>Approver Level</td>
                          <td>Approver Name</td>
                          <td>Approval Type</td>
                          <td>Delete</td>
                        </tr>
                        {approval}
                      </tbody>
                    </table>
                  </div>
                </div>
              </form>
            </fieldset>
          </section>) : null}
          {this.state.editSubmitStatus == "Yes" ?
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
          {this.state.showApprove === true || this.state.showReject === true ?
            <section style={{ justifyContent: 'left', textAlign: 'left', display: 'grid' }} id="approvalSection" className='card card-body'>
              <TextField label="Remarks" required name="remarks" value={this.state.remarks} multiline rows={3} onChange={this.handleChange}
                styles={{
                  fieldGroup: {
                    backgroundColor: this.state.editErrors.remarks ? "#ffcccb" : "white", // Red tint for errors
                  }
                }} />
            </section> : null
          }
          {/* Vishnu Changes  */}
          {this.state.showSubmit &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

            {this.state.showDraft &&

                <PrimaryButton text="Save as Draft" onClick={() => this.handleDraft("draft")} />

              }

            <PrimaryButton text="Submit" onClick={() => this.handleSubmit("submit")} />

              <DefaultButton text="Cancel" onClick={() => this.cancelRequest()} />

            </div>
          }
          {this.state.showApprove &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

              <PrimaryButton text="Approve" onClick={() => this._approveRequest("Approve")} />


              <PrimaryButton text="Rework" onClick={() => this._reworkRequest("Rework")} />


              <DefaultButton text="Cancel" onClick={() => this.cancelRequest()} />

            </div>
          }
          {this.state.showForward &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

              <PrimaryButton text="Forward" onClick={() => this.handleForward("Forward")} />


              <a href='#/form'> <PrimaryButton onClick={this.cancelRequest}>Cancel</PrimaryButton></a>

            </div>
          }
          {this.state.showReject &&

            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

              <PrimaryButton text="Approve" onClick={() => this._approveRequest("Approve")} />


              <PrimaryButton text="Reject" onClick={() => this._rejectRequest("Reject")} />


              <DefaultButton text="Cancel" onClick={() => this.cancelRequest()} />

            </div>
          }
          {this.state.edType === "view" &&
            <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

              <DefaultButton text="Cancel" onClick={() => this.cancelRequest()} />

            </div>
          }
        </div>
      </section >
    )
  }
}
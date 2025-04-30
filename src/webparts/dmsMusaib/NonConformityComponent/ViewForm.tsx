import * as React from 'react';
import styles from './AuditPlan.module.scss';
import type { IAuditPlanProps } from './IAuditPlanProps';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PrimaryButton,
  TooltipHost,
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
import CustomBreadcrumb from '../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
import moment from 'moment';
import { getMemoNumberAuditReport } from '../AnnualAuditReportComponent/AuditReportService';

export class IViewState {
  mainItemId?: any | null;
  viewDepartmentOption: IDropdownOption[];
  memonumberOptions: any[];
  NCNumberOptions: any[];
  NCNumber: string;
  ApprovedAuditReport: string | number;
  NCNumberID: string | number;
  MemoNumber: string;
  viewDepartment: string | number;
  viewCriteria: string;
  viewncrNo: string;
  viewreferencenumber: String;
  viewissueNo: any;
  TemplateDoc: any[];
  viewrevisionNo: any;
  viewissueDate: any;
  viewRevisionDate: any;
  viewCloseOutStatus: string;
  viewCategoryCheckOption: IDropdownOption[];
  viewCategoryValueIsCheck: number[];
  viewSubCategoryCheckOption: IDropdownOption[];
  viewSubCategoryValueIsCheck: number[];
  viewLocationCheckOption: IDropdownOption[];
  viewLocationValueIsCheck: number[];
  viewAssignTo: string;
  viewAssignToId: number | null;
  viewProblemDescription: string;
  viewDueDate: any;
  viewPersonAssigned: string;
  viewPersonAssignedId: number | null;
  viewDate: any;
  viewDeadlineCompletion: any;
  viewCorrection: string;
  viewRootCause: string;
  viewCorrectiveAction: string;
  viewDelegateTo: string;
  viewDelegateToId: number | null;
  viewAnalyzedBy: string;
  viewAnalyzedById: number | null;
  viewReviewedBy: string;
  viewReviewedById: number | null;
  viewCorrectiveActionImplementedOn: string;
  viewSubmitedDate: any;
  viewSubmitStatus: string;
  viewStatus: string;
  viewIsRework: string;
  viewSerialNumber: string;
  viewDocumentCode: string;
  viewAttachmentPreArray: any[];
  viewAttachmentJson: any[];
  viewErrors: { [key: string]: string };
  viewCurrentUserRole: string;
  viewFirstInitiatorSubmitStatus: string;
  viewFirstAssignedToSubmitStatus: string;
  viewDelegateToSubmitStatus: string;
  viewAnalyzedBySubmitStatus: string;
  viewReviewedBySubmitStatus: string;
  viewLastAssignedToSubmitStatus: string;
  viewLastInitiatorSubmitStatus: string;
  disable: boolean;
  viewProcessListItemId: number | null;
  viewProcessCurrentUserRole: string;
  viewProcessStatus: string;
  viewProcessListNameId: number | null;
  viewProcessListName: string;
  viewProcessActionTakenById: number | null;
  viewProcessActionTakenBy: string;
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
  forwardBtn: string;
  showDialog: boolean;
  copyFil: any[];
  fileCount: number;
  exFiles: any[];
  fileDeleteId: any[];
  files: FileList;
  apprItems: any[];
}
const optionsApp: IDropdownOption[] = [

  { key: 'All', text: 'All' },
  { key: 'One', text: 'One' }

]
export default class EditForm extends React.Component<IAuditPlanProps, IViewState> {

  constructor(props: IAuditPlanProps) {
    super(props);
    const selectedTextDiv = document.getElementById('selectedText');
    selectedTextDiv.style.display = 'none';
    this.state = {
      mainItemId: props.edItm || null,
      viewDepartmentOption: [],
      memonumberOptions: [],
      NCNumberOptions: [],
      NCNumber: "",
      NCNumberID: "",
      ApprovedAuditReport: "",
      MemoNumber: "",
      viewDepartment: "",
      viewCriteria: "",
      viewncrNo: "",
      viewreferencenumber: "",
      viewissueNo: "",
      viewrevisionNo: "",
      viewissueDate: null,
      viewRevisionDate: null,
      TemplateDoc: [],
      viewCloseOutStatus: "",
      viewCategoryCheckOption: [],
      viewCategoryValueIsCheck: [],
      viewSubCategoryCheckOption: [],
      viewSubCategoryValueIsCheck: [],
      viewLocationCheckOption: [],
      viewLocationValueIsCheck: [],
      viewAssignTo: "",
      viewAssignToId: null,
      viewProblemDescription: "",
      viewDueDate: null,
      viewPersonAssigned: "",
      viewPersonAssignedId: null,
      viewDate: null,
      viewDeadlineCompletion: null,
      viewCorrection: "",
      viewRootCause: "",
      viewCorrectiveAction: "",
      viewDelegateTo: "",
      viewDelegateToId: null,
      viewAnalyzedBy: "",
      viewAnalyzedById: null,
      viewReviewedBy: "",
      viewReviewedById: null,
      viewCorrectiveActionImplementedOn: "",
      viewSubmitedDate: null,
      viewSubmitStatus: "",
      viewStatus: "",
      viewIsRework: "No",
      viewSerialNumber: "",
      viewDocumentCode: "",
      viewAttachmentPreArray: [],
      viewAttachmentJson: [],
      viewErrors: {},
      viewCurrentUserRole: "",
      viewFirstInitiatorSubmitStatus: "",
      viewFirstAssignedToSubmitStatus: "",
      viewDelegateToSubmitStatus: "",
      viewAnalyzedBySubmitStatus: "",
      viewReviewedBySubmitStatus: "",
      viewLastAssignedToSubmitStatus: "",
      viewLastInitiatorSubmitStatus: "",
      disable: true,
      viewProcessListItemId: null,
      viewProcessCurrentUserRole: "",
      viewProcessStatus: "",
      viewProcessListNameId: null,
      viewProcessListName: "",
      viewProcessActionTakenById: null,
      viewProcessActionTakenBy: "",
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
      forwardBtn: "",
      showDialog: false,
      copyFil: [],
      fileCount: 0,
      exFiles: [],
      fileDeleteId: [],
      files: {} as FileList,
      apprItems: [],
    };
    this.addApprover = this.addApprover.bind(this);
    this.deleteItemApp = this.deleteItemApp.bind(this);
    this._getPeoplePickerItemsApp = this._getPeoplePickerItemsApp.bind(this);
    this.onRoleChange = this.onRoleChange.bind(this);
    this.onTypeChange = this.onTypeChange.bind(this);
    this.approveRequest = this.approveRequest.bind(this);
    this.rejectRequest = this.rejectRequest.bind(this);
    this.reworkRequest = this.reworkRequest.bind(this);
    this._OpenModal = this._OpenModal.bind(this);
    this._CloseModal = this._CloseModal.bind(this);
    this.removeFiles = this.removeFiles.bind(this);
    this.toBeDeleted = this.toBeDeleted.bind(this);
    this.handleFileChange = this.handleFileChange.bind(this);
  }
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
    this.setState({ copyFil: items });
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
    this.setState({ exFiles: remFiles });
  }
  public changeDepartment = (_event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ viewDepartment: item.key });
  };

  public handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      [name]: value,
    }));
  };

  private _handlePeoplePickerChange = (field: keyof IViewState, idField: keyof IViewState) => (items: any[]) => {
    if (items.length > 0) {
      this.setState({
        [field]: items[0].text,
        [idField]: items[0].id,
      } as Pick<IViewState, keyof IViewState>);
    } else {
      this.setState({
        [field]: "",
        [idField]: null,
      } as unknown as Pick<IViewState, keyof IViewState>);
    }
  };

  private _handleCheckboxChange = (stateKey: keyof IViewState, itemKey: number) =>
    (_ev: React.FormEvent<HTMLElement>, isChecked?: boolean) => {
      this.setState((prevState) => {
        const updatedValues = isChecked
          ? [...(prevState[stateKey] as number[]), itemKey]
          : (prevState[stateKey] as number[]).filter((key) => key !== itemKey);
        return { [stateKey]: updatedValues } as unknown as Pick<IViewState, keyof IViewState>;
      });
    };

  public async componentDidMount() {
    const selectedTextDiv = document.getElementById('selectedText');
    if (selectedTextDiv) {
      selectedTextDiv.style.display = 'none';
    }
    await this.getAuditreport();
    await this.getGeneratedTemplateDocNC()
    await this.getDepartment();
    await this.getListData();
    await this.getProcessApprovalList();
    await this.getDataRoles();
    await this.getMainListName();
    await this.getRequestorRole();
    await this.getFormName();
    await this.getFiles();
  }
  public async getGeneratedTemplateDocNC() {
    debugger
    const _sp = spfi().using(SPFx(this.props.context));
    let results: any = [];
    // for (let itemId of AttachmentIds) {

    await _sp.web.lists.getByTitle("NonConformityGeneratedTemplateDoc").items
      .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${this.state.mainItemId}`)()
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
  public async getListData() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const Items: any = await sp.web.lists.getByTitle("NonConformityList").items.getById(this.state.mainItemId)
        .select("*, Category/Id, Category/Title, SubCategory/Id, Location/Id, Location/Title, SubCategory/Title, AssignedTo/Id, AssignedTo/Title, DelegateTo/Id, DelegateTo/Title, AnalyzedBy/Id, AnalyzedBy/Title, ReviewedBy/Id, ReviewedBy/Title, PersonAssigned/Id, PersonAssigned/Title")
        .expand("Category, SubCategory, Location, AssignedTo, DelegateTo, AnalyzedBy, ReviewedBy, PersonAssigned")();
      console.log("Items1", Items);
      this.setState({
        ncItemId: Items.Id,
        
        viewDepartment: Items.DepartmentId,
        viewCriteria: Items.Criteria,
        viewncrNo: Items.NCRNo,
        viewreferencenumber: Items.ReferenceNumber,
        viewissueNo: Items.IssueNumber,
        viewrevisionNo: Items.RevisionNumber,
        viewissueDate: Items.IssueDate,
        viewRevisionDate: Items.RevisionDate,
        viewDocumentCode: Items.DocumentCode,
        viewCloseOutStatus: Items.CloseOutStatus,
        viewCategoryValueIsCheck: Items.Category ? Items.Category.map((cat: any) => cat.Id) : [],
        viewSubCategoryValueIsCheck: Items.SubCategory ? Items.SubCategory.map((sub: any) => sub.Id) : [],
        viewLocationValueIsCheck: Items.Location ? Items.Location.map((loc: any) => loc.Id) : [],
        viewAssignToId: Items.AssignedTo ? Items.AssignedTo.Id : null,
        viewAssignTo: Items.AssignedTo ? Items.AssignedTo.Title : null,
        viewProblemDescription: Items.ProblemDescription,
        viewDueDate: Items.DueDate ? new Date(Items.DueDate) : null,
        viewPersonAssignedId: Items.AssignedTo ? Items.AssignedTo.Id : null,
        viewPersonAssigned: Items.AssignedTo ? Items.AssignedTo.Title : null,
        viewDate: Items.Date ? new Date(Items.Date) : null,
        viewDeadlineCompletion: Items.DueDate ? new Date(Items.DueDate) : null,
        viewCorrection: Items.Correctionproblem,
        viewRootCause: Items.RootCause,
        viewCorrectiveAction: Items.CorrectiveAction,
        viewDelegateToId: Items.DelegateTo ? Items.DelegateTo.Id : null,
        viewDelegateTo: Items.DelegateTo ? Items.DelegateTo.Title : null,
        viewAnalyzedById: Items.AnalyzedBy ? Items.AnalyzedBy.Id : null,
        viewAnalyzedBy: Items.AnalyzedBy ? Items.AnalyzedBy.Title : null,
        viewReviewedById: Items.ReviewedBy ? Items.ReviewedBy.Id : null,
        viewReviewedBy: Items.ReviewedBy ? Items.ReviewedBy.Title : null,
        viewCorrectiveActionImplementedOn: Items.CorrectiveActionImplementedOn,
        viewSubmitStatus: Items.SubmitStatus,
        viewCurrentUserRole: Items.CurrentUserRole,
        viewFirstInitiatorSubmitStatus: Items.FirstInitiatorSubmitStatus,
        viewFirstAssignedToSubmitStatus: Items.FirstAssignedToSubmitStatus,
        viewDelegateToSubmitStatus: Items.DelegateToSubmitStatus,
        viewAnalyzedBySubmitStatus: Items.AnalyzedBySubmitStatus,
        viewReviewedBySubmitStatus: Items.ReviewedBySubmitStatus,
        viewLastAssignedToSubmitStatus: Items.LastAssignedToSubmitStatus,
        viewLastInitiatorSubmitStatus: Items.LastInitiatorSubmitStatus,
        viewStatus: Items.Status,
        viewAttachmentPreArray: [],
        viewAttachmentJson: [],
        forwardBtn: Items.ForwardBtn,
      });
    } catch (e) {
      console.error(e);
    }
  }
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
  public async getAuditreport() {
    const sp = spfi().using(SPFx(this.props.context));
    debugger
    try {
      const memoItems = await getMemoNumberAuditReport(sp);
      let optionsmemoNumber: any = [];
      if (memoItems.length > 0) {
        optionsmemoNumber = memoItems[0].map((item: any) => ({
          key: item.ID,
          text: item.MemoNumber,
          itemId: item.ID,
          memoNumber: item.MemoNumber,
          ncNo: item.NCNumber
        }));
      }
      let optionsmemoNumbernew: any[] = [];
      optionsmemoNumbernew = await this.getUniqueBy(optionsmemoNumber, "memoNumber");
      this.setState({ memonumberOptions: optionsmemoNumbernew });
    } catch (e) {
      console.error(e);
    }
  }
  public changeMemoNumber = async (_event: React.FormEvent<HTMLDivElement>, item: any): Promise<void> => {
    debugger
    const optionsNCNumber = this.state.memonumberOptions.filter((x) => x.memoNumber == item.text).map((item: any) => ({
      key: item.key,
      text: item.ncNo,
      ncNo: item.ncNo
    }));
    let optionsNCNumbernew: any[] = [];
    optionsNCNumbernew = await this.getUniqueBy(optionsNCNumber, "ncNo")
    this.setState({ NCNumberOptions: optionsNCNumber })
    this.setState({ ApprovedAuditReport: item.key, MemoNumber: item.memoNumber });
  };
  public async getProcessApprovalList() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const Items: any = await sp.web.lists.getByTitle("ProcessApprovalList").items
        .select("*", "AssignedTo/Title", "AssignedTo/Id", "RequesterName/Title", "RequesterName/Id", "ActionTakenBy/Title", "ActionTakenBy/Id")
        .expand("AssignedTo,RequesterName,ActionTakenBy").filter("ListItemId eq '" + this.state.mainItemId + "'").orderBy("Level")();
      console.log(Items);
      Items.map((item: any) => {
        if (item.ListItemId == this.state.mainItemId) {  //&& item.ListName.Title == "NonConformityList"
          this.setState({
            processListItemID: item.ID,
            viewProcessListItemId: item.ListItemId,
            viewProcessCurrentUserRole: item.CurrentUserRole,
            viewProcessStatus: item.Status,
            viewProcessListNameId: item.ListName ? item.ListName.Id : null,
            viewProcessListName: item.ListName ? item.ListName.Title : null,
            viewProcessActionTakenById: item.ActionTakenBy ? item.ActionTakenBy.Id : null,
            viewProcessActionTakenBy: item.ActionTakenBy ? item.ActionTakenBy.Title : null,
          });
        }
      })

      const apprItems = await sp.web.lists.getByTitle("ProcessApprovalList").items.select("*", "AssignedTo/Title,RequesterName/Title,ActionTakenBy/Title").expand("AssignedTo,RequesterName,ActionTakenBy").filter("ListItemId eq '" + this.state.mainItemId + "'").orderBy("Level")();
      var cnt: any = 0;
      var appItems: any[] = [];
      if (apprItems.length > 0) {
        apprItems.forEach(async function (itm: any) {
          var objToAdd: any = {};
          objToAdd["Level"] = itm.Level;
          objToAdd["AssignedTo"] = itm.AssignedTo.Title;
          objToAdd["RequesterName"] = itm.RequesterName.Title;
          objToAdd["RequestedDate"] = itm.RequestedDate;
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

    } catch (e) {
      console.error(e);
    }
  }
  public async getFiles() {
    const sp = spfi().using(SPFx(this.props.context));
    const upFiles = await sp.web.lists.getByTitle("NonConformityDocs").items.select("*", "File/Name").expand("File").filter("NonConformityId eq '" + this.state.mainItemId + "'")();
    if (upFiles.length > 0) {
      var obJFiles: any[] = [];
      upFiles.forEach(function (item: any) {
        obJFiles.push({ "Name": item.File.Name, "type": "old", "Id": item.Id })
      })
      this.setState({ exFiles: obJFiles });
    }
    //Approval Table data
    const approvalItems = await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.select("*", "Approvers/Name").expand("Approvers").filter("MainListID eq '" + this.state.mainItemId + "'")();
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
  public cancelRequest() {
    alert("Cancel");
  }
  public async getDepartment() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const deptItems = await sp.web.lists.getByTitle("DepartmentMasterList").items();
      const options = deptItems.map((item: { Title: string; Id: number }) => ({
        key: item.Id,
        text: item.Title,
      }));
      this.setState({ viewDepartmentOption: options });
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
      this.setState({ viewCategoryCheckOption: options });
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
      this.setState({ viewSubCategoryCheckOption: options });
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
      this.setState({ viewLocationCheckOption: options });
    } catch (e) {
      console.error(e);
    }
  }
  private onRoleChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Role = item.key;
    this.setState({ approvers: this.state.approvers });
  }

  private onTypeChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i: number) {
    this.state.approvers[i].Type = item.text;
    this.setState({ approvers: this.state.approvers });
  }

  private _getPeoplePickerItemsApp(items: any[], i: number) {
    debugger;
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

  public async approveRequest() {
    const { remarks, processListItemID } = this.state
    const { currentUserID } = this.props
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: 'Do you want to approve this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(function (val) {
      if (val.isConfirmed) {
        sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(processListItemID)).update({
          Status: "Approved",
          ActionTakenById: currentUserID,
          ActionTakenOn: new Date(),
          Remark: remarks,
        });
        Swal.fire({ title: "Approved succesfully", icon: "success" });
      }
    });

  }
  private async rejectRequest() {
    const { remarks, processListItemID } = this.state
    const { currentUserID } = this.props
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: 'Do you want to reject this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(function (val) {
      if (val.isConfirmed) {
        sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(processListItemID)).update({
          Status: "Rejected",
          ActionTakenById: currentUserID,
          ActionTakenOn: new Date(),
          Remark: remarks,
        });
        Swal.fire({ title: "Rejected successfully", icon: "success" });
      }
    });

  }
  private async reworkRequest() {
    const { remarks, processListItemID } = this.state
    const { currentUserID } = this.props
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: 'Do you want to rework this request?',
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(function (val) {
      if (val.isConfirmed) {
        sp.web.lists.getByTitle("ProcessApprovalList").items.getById(Number(processListItemID)).update({
          Status: "Rework",
          ActionTakenById: currentUserID,
          ActionTakenOn: new Date(),
          Remark: remarks,
          IsRework: "Yes"
        });
        Swal.fire({ title: "Rework successfull", icon: "success" });
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
            <Dropdown disabled={this.state.disable} placeholder="Select options" selectedKey={this.state.approvers[i].Role} options={this.state.optionsRole} onChange={(e, itm: IDropdownOption) => this.onRoleChange(e, itm, i)} />
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
              disabled={this.state.disable}
              ensureUser={true}
              defaultSelectedUsers={this.state.approvers[i].appEx ? this.state.approvers[i].appEx : []}
              onChange={(e) => this._getPeoplePickerItemsApp(e, i)}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            />
          </td>
          <td>
            <Dropdown disabled={this.state.disable} placeholder="Select options" selectedKey={this.state.approvers[i].Type} options={optionsApp} onChange={(e, itm: IDropdownOption) => this.onTypeChange(e, itm, i)} />
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
            <button type="button" onClick={(e) => this.toBeDeleted(i)}>Delete</button>
          </td>
        </tr>
      )
    });
    var auditHistory = this.state.apprItems.map((item: any, i: number) => {
      return (
        <tr>
          <td style={{ minWidth: '70px', maxWidth: '70px' }}>
            {i + 1}
          </td>
          <td style={{ minWidth: '70px', maxWidth: '70px' }}>
            {item.Level}
          </td>
          <td title={item.AssignedTo} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.AssignedTo}
          </td>
          <td title={item.RequesterName} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.RequesterName}
          </td>
          <td title={item.Status !== 'test'
            ? `${new Intl.DateTimeFormat('en-GB', {
              day: '2-digit',
              month: 'short',
              year: 'numeric'
            }).format(new Date(item.RequestedDate)).replace(/ /g, "-")} ${new Date(item.RequestedDate).toLocaleTimeString('en-GB', {
              hour: '2-digit',
              minute: '2-digit',
              hour12: false
            })}`
            : ""}>
            {item.Status !== 'test'
              ? `${new Intl.DateTimeFormat('en-GB', {
                day: '2-digit',
                month: 'short',
                year: 'numeric'
              }).format(new Date(item.RequestedDate)).replace(/ /g, "-")} ${new Date(item.RequestedDate).toLocaleTimeString('en-GB', {
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
              })}`
              : ""}
          </td>
          <td title={item.ActionTakenBy} style={{ minWidth: '90px', maxWidth: '90px' }}>
            {item.ActionTakenBy}
          </td>
          <td title={item.Status !== 'Pending'
            ? `${new Intl.DateTimeFormat('en-GB', {
              day: '2-digit',
              month: 'short',
              year: 'numeric'
            }).format(new Date(item.ActionTakenOn)).replace(/ /g, "-")} ${new Date(item.ActionTakenOn).toLocaleTimeString('en-GB', {
              hour: '2-digit',
              minute: '2-digit',
              hour12: false
            })}`
            : ""}>
            {item.Status !== 'Pending'
              ? `${new Intl.DateTimeFormat('en-GB', {
                day: '2-digit',
                month: 'short',
                year: 'numeric'
              }).format(new Date(item.ActionTakenOn)).replace(/ /g, "-")} ${new Date(item.ActionTakenOn).toLocaleTimeString('en-GB', {
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
              })}`
              : ""}
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
            <div className="col-lg-4 newbread">
              <CustomBreadcrumb Breadcrumb={this.Breadcrumb} />
            </div>

          </div>
          <section className="card card-body">
            <fieldset>
              <form>
                {/* Start save as draft */}
                <div className="previewIcon">
                  <h4 style={{ textAlign: 'left' }} className="text-dark font-16 fw-bold mb-3">Problem Details</h4>
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
                  <div className="form-group col-md-4">
                    <TooltipHost
                      content={this.state.memonumberOptions.filter((item: any) => item.key == this.state.ApprovedAuditReport)[0]?.text || ""}
                      calloutProps={{ gapSpace: 0 }}
                      styles={{ root: { display: 'inline-block', width: '100%' } }}
                    >
                      <Dropdown
                        required
                        disabled={true}
                        placeholder="Approved Audit Report/Memo Number"
                        label="Approved Audit Report/Memo Number:"
                        options={this.state.memonumberOptions}
                        defaultSelectedKey={this.state.ApprovedAuditReport}
                        selectedKey={this.state.ApprovedAuditReport}
                        
                      />
                    </TooltipHost>
                  </div>
                  <div className="form-group col-md-4">
                    <TooltipHost
                      content={this.state.NCNumber || ""}
                      calloutProps={{ gapSpace: 0 }}
                      styles={{ root: { display: 'inline-block', width: '100%' } }}
                    >
                      <Dropdown
                        disabled={true}
                        required
                        placeholder="NCNumber"
                        label="NC Number:"
                        options={this.state.NCNumberOptions}
                        defaultSelectedKey={this.state.NCNumberID}
                        selectedKey={this.state.NCNumberID}
                        
                      />
                    </TooltipHost>
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="NCR No:" name='viewncrNo' required value={this.state.viewncrNo} disabled={true} onChange={this.handleChange}

                    />
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="Document Code:" name='viewDocumentCode' required value={this.state.viewDocumentCode} disabled={true} onChange={this.handleChange}

                    />
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="Issue Number:" name='viewIssueNo' required value={this.state.viewissueNo + ""} disabled={true} onChange={this.handleChange}

                    />
                  </div>
                </div>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                  <div className="form-group col-md-4">
                    <TextField label="Revision Number:" name='viewRevisionNo' required value={this.state.viewrevisionNo + ""} disabled={true} onChange={this.handleChange}

                    />
                  </div>
                  <div className="form-group col-md-4">
                    <Dropdown
                      required
                      disabled={this.state.disable}
                      label="Department:"
                      options={this.state.viewDepartmentOption}
                      defaultSelectedKey={this.state.viewDepartment}
                      selectedKey={this.state.viewDepartment}
                      onChange={this.changeDepartment}
                    />
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="Criteria:" disabled={this.state.disable} name='viewCriteria' required value={this.state.viewCriteria} onChange={this.handleChange}
                    />
                  </div>
                  <div className="form-group col-md-4">
                    <TextField label="Close Out Status:" disabled={this.state.disable} name='editCloseOutStatus' required value={this.state.viewCloseOutStatus} onChange={this.handleChange}
                    />
                  </div>
                </div>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                  <div className="form-group col-md-4">
                    <label>Category:</label>
                    {this.state.viewCategoryCheckOption.map((item: any) => {
                      return (
                        <div style={{ margin: "2px", padding: "3px" }}>
                          <Checkbox label={item.text} disabled={this.state.disable} checked={this.state.viewCategoryValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("viewCategoryValueIsCheck", item.key as number)} />
                        </div>
                      )
                    }
                    )}
                  </div>
                  <div className="form-group col-md-4">
                    <label>SubCategory:</label>
                    {this.state.viewSubCategoryCheckOption.map((item: any) => {
                      return (
                        <div style={{ margin: "2px", padding: "3px" }}>
                          <Checkbox label={item.text} disabled={this.state.disable} checked={this.state.viewSubCategoryValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("viewSubCategoryValueIsCheck", item.key as number)} />
                        </div>
                      )
                    }
                    )}
                  </div>
                  <div className="form-group col-md-4">
                    <label>Location:</label>
                    {this.state.viewLocationCheckOption.map((item: any) => {
                      return (
                        <div style={{ margin: "2px", padding: "3px" }}>
                          <Checkbox label={item.text} disabled={this.state.disable} checked={this.state.viewLocationValueIsCheck.indexOf(item.key) !== -1} onChange={this._handleCheckboxChange("viewLocationValueIsCheck", item.key as number)}
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
                      disabled={this.state.disable}
                      ensureUser={true}
                      defaultSelectedUsers={this.state.viewAssignTo ? [this.state.viewAssignTo] : []}
                      onChange={this._handlePeoplePickerChange("viewAssignTo", "viewAssignToId")}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                    />
                  </div>
                  <div className="form-group col-md-4">
                    <DatePicker
                      label="Due Date"
                      isRequired
                      disabled={this.state.disable}
                      placeholder="Select a Duedate"
                      value={this.state.viewDueDate}
                      onSelectDate={(date: Date) => this.setState({ viewDueDate: date })}
                    />
                  </div>
                  <div style={{ position: 'relative' }} className="col-lg-4">
                    <label htmlFor="Attachments">Attachments <span className="text-danger">*</span></label>
                    <input disabled={this.state.disable} className="form-control" type="file" name="myFile" onChange={(e) => this.handleFileChange(e, this)} id="newfile" multiple />
                    <span onClick={this._OpenModal} className='newpo'>{this.state.fileCount}</span>
                    <table>
                      <tbody>
                        {fileData}
                      </tbody>
                    </table>
                    {this.state.showDialog && <div id="myModal" className={styles.modal}>
                      <div className={styles.modalcontent}>
                        <span className={styles.close} onClick={e => this._CloseModal()}>&times;</span>
                        <table className='mtbalenew'>
                          <tbody>
                            <tr style={{ display: 'table', width: '100%' }}>
                              <td>File Name</td></tr>
                            {fileData}
                          </tbody>
                        </table>
                        <table>
                          <tr style={{ display: 'table', width: '100%' }}>
                            <td>Uploaded Files</td></tr>
                          {upFiles}
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
                      value={this.state.viewProblemDescription}
                      multiline rows={5}
                      onChange={this.handleChange}
                      disabled={this.state.disable}
                      styles={{
                        fieldGroup: {
                          backgroundColor: this.state.viewErrors.editProblemDescription ? "#ffcccb" : "white", // Red tint for errors
                        }
                      }}
                    />
                  </div>
                </div>
                {/* End svae as draft */}
              </form>
            </fieldset>
          </section>
          {this.state.viewFirstInitiatorSubmitStatus == "Yes" || this.state.viewFirstAssignedToSubmitStatus == "Yes" ?
            <section className='card card-body mt-2' >
              <fieldset>
                <form>
                  {/* Section 2 */}
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark font-16 text-left fw-bold mb-3'>To be filled by Department Head</h3></div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-4">
                      <PeoplePicker
                        context={peoplePickerContext}
                        disabled={this.state.disable}
                        titleText="Person Assigned:"
                        personSelectionLimit={1}
                        required={true}
                        onChange={this._handlePeoplePickerChange("viewPersonAssigned", "viewPersonAssignedId")}
                        defaultSelectedUsers={[this.state.viewPersonAssigned]}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        ensureUser={true}
                      />
                    </div>
                    <div className="form-group col-md-4">
                      <DatePicker
                        isRequired
                        disabled={this.state.disable}
                        label="Date"
                        placeholder="Select a Date"
                        isMonthPickerVisible={false}
                        value={this.state.viewDate}
                        onSelectDate={(date: Date) => this.setState({ viewDate: date })}
                      />
                    </div>
                    <div className="form-group col-md-4">
                      <DatePicker
                        isRequired
                        disabled={this.state.disable}
                        label="Deadline for completion"
                        placeholder="Select a Deadline"
                        isMonthPickerVisible={false}
                        value={this.state.viewDeadlineCompletion}
                        onSelectDate={(date: Date) => this.setState({ viewDeadlineCompletion: date })}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-6">
                      <TextField label="Correction(Immediate Steps to stop the problem):" required name='editCorrection'
                        value={this.state.viewCorrection}
                        multiline rows={3}
                        disabled={this.state.disable}
                        onChange={this.handleChange}
                      />
                    </div>
                    <div className="form-group col-md-6">
                      <TextField label="Root Cause:" required name='editRootCause'
                        value={this.state.viewRootCause} multiline rows={3}
                        onChange={this.handleChange}
                        disabled={this.state.disable}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-12">
                      <TextField label="Corrective Action (Action to eliminate the root cause):"
                        required
                        name='editCorrectiveAction'
                        value={this.state.viewCorrectiveAction} multiline rows={3}
                        onChange={this.handleChange}
                        disabled={this.state.disable}
                      />
                    </div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-4">
                      <PeoplePicker
                        disabled={true}
                        context={peoplePickerContext}
                        titleText="Delegate To:"
                        personSelectionLimit={1}
                        required={false}
                        onChange={this._handlePeoplePickerChange("viewDelegateTo", "viewDelegateToId")}
                        defaultSelectedUsers={[this.state.viewDelegateTo]}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        ensureUser={true}
                      />
                    </div>
                    <div className="form-group col-md-4">
                      <PeoplePicker
                        disabled={this.state.disable}
                        context={peoplePickerContext}
                        titleText="Analyzed By:"
                        personSelectionLimit={1}
                        required={true}
                        onChange={this._handlePeoplePickerChange("viewAnalyzedBy", "viewAnalyzedById")}
                        defaultSelectedUsers={[this.state.viewAnalyzedBy]}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        ensureUser={true}
                      />
                    </div>
                    <div className="form-group col-md-4">
                      <PeoplePicker
                        context={peoplePickerContext}
                        disabled={this.state.disable}
                        titleText="Reviewed By:"
                        personSelectionLimit={1}
                        required={true}
                        onChange={this._handlePeoplePickerChange("viewReviewedBy", "viewReviewedById")}
                        defaultSelectedUsers={[this.state.viewReviewedBy]}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        ensureUser={true}
                      />
                    </div>
                  </div>
                  {/* Section 3 */}
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark font-16 text-left fw-bold mb-0'>Problem Close Out Details</h3></div>
                  </div>
                  <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                    <div className="form-group col-md-12">
                      <TextField label="Corrective Action Implemented On:"
                        required
                        name='editCorrectiveActionImplementedOn'
                        value={this.state.viewCorrectiveActionImplementedOn}
                        disabled={this.state.disable}
                        multiline rows={3} onChange={this.handleChange}
                      />
                    </div>
                  </div>
                  {/* Approve/Rework */}
                </form>
              </fieldset>
            </section> : null}
          {/* Approval Table */}
          {(this.state.viewReviewedBySubmitStatus == "Yes" && this.state.viewDelegateToId == null) || (this.state.viewLastAssignedToSubmitStatus == "Yes" && this.state.viewDelegateToId != null) ? (<section className={styles.sec}>
            <fieldset disabled={this.state.disable}>
              <form>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row'>
                  <div className="form-group col-md-12"><h3 style={{ textAlign: 'left' }} className='text-dark font-16 text-left fw-bold mb-0'>Forward Detail</h3></div>
                </div>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                  <div className="container mt-4">
                    <button type="button" onClick={this.addApprover}>Add</button>
                    <table id="tblAppr" className='mtbalenew'>
                      <tbody>
                        <tr><td>Sl.</td>
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
          <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row'>
            {this.state.viewLastInitiatorSubmitStatus == "Yes" ? (<section id="approvalSection">
              <TextField label="Remarks" name="remarks" value={this.state.remarks} multiline rows={3} onChange={this.handleChange} />
            </section>) : null}
          </div>
          {this.state.viewSubmitStatus == "Yes" ?
            <section className='card card-body mt-2' >
              <form>
                <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row'>
                  <div className="form-group col-md-12"><h3>Audit Report</h3></div>
                </div>
                <div style={{ justifyContent: 'left', textAlign: 'left', display: 'grid' }} className='row'>
                  <table className='mtbalenew'>
                    <thead>
                      <tr>
                        <th style={{ minWidth: '70px', maxWidth: '70px' }}>S No.</th>
                        <th style={{ minWidth: '70px', maxWidth: '70px' }}>Level</th>
                        <th style={{ minWidth: '90px', maxWidth: '90px' }}>Assigned To</th>
                        <th style={{ minWidth: '90px', maxWidth: '90px' }}>Requestor Name</th>
                        <th style={{ minWidth: '90px', maxWidth: '90px' }}>Requested Date</th>
                        <th style={{ minWidth: '90px', maxWidth: '90px' }}>Action Taken By</th>
                        <th style={{ minWidth: '90px', maxWidth: '90px' }}>Action Taken On</th>
                        <th style={{ minWidth: '90px', maxWidth: '90px' }}>Remarks</th>
                        <th style={{ minWidth: '70px', maxWidth: '70px' }}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {auditHistory}
                    </tbody>
                  </table>
                </div>
              </form>
            </section> : null}

          <a href='#/form'> <PrimaryButton onClick={() => this.cancelRequest()}>Cancel</PrimaryButton></a>
        </div>
      </section >
    )
  }
}
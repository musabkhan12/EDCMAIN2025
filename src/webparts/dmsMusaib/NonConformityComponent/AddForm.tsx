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
import moment from 'moment';
import CustomBreadcrumb from '../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
//import { CONTENTTYPE_NonComformity } from '../../../Shared/Constants';
import { getLatestChangeRequestTemplateType, getMemoNumberAuditReport, getNCNumbers } from '../AnnualAuditReportComponent/AuditReportService';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import "./nonconformity.scss";
import { CONTENTTYPE_NonComformity } from '../ChangerequestComponent/Constants';
import Select from "react-select";

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


export class IState {
  Loading: boolean;
  departmentOption: any[];
  departmentselected: any;
  memonumberOptions: any[];
  typeoptions: any[];
  ncType: any;
  memonumberOptionsall: any[];
  NCNumberOptions: any[];
  NCNumberselected: any[];
  ApprovedAuditSelected: any;
  NCNumber: string;
  ApprovedAuditReport: string | number;
  NCNumberID: string | number;
  MemoNumber: string;
  department: string | number;
  departmentCode: string;
  serialNo: number;
  criteria: string;
  revisionNo: any;
  issueNo: any;
  ncrNo: string;
  documentCode: string;
  revisionDate: any;
  issueDate: any;
  referenceNo: string;
  closeOutStatus: string;
  categoryCheckOption: IDropdownOption[];
  categoryValueIsCheck: number[];
  subCategoryCheckOption: IDropdownOption[];
  subCategoryIsCheck: number[];
  locationCheckOption: IDropdownOption[];
  locationValueIsCheck: number[];
  assignTo: string;
  assignToId: number | null;
  problemDescription: string;
  dueDate: any;
  submitStatus: string;
  currentUserRole: string;
  firstInitiatorSubmitStatus: string;
  attachmentPreArray: any[];
  attachmentJson: any[];
  errors: { [key: string]: string };
  showDialog: boolean;
  copyFil: any[];
  fileCount: number;
  exFiles: any[];
  fileDeleteId: any[];
  files: FileList;
  siteurl: any;
}
let optionsmemoNumbernewnc: any[] = [];
let optionsmemoNumbernewobs: any[] = [];
export default class AuditPlan extends React.Component<IAuditPlanProps, IState> {
  constructor(props: IAuditPlanProps) {
    super(props);
    const selectedTextDiv = document.getElementById('selectedText');
    selectedTextDiv.style.display = 'none';
    this.state = {
      Loading: false,
      departmentOption: [],
      departmentselected: [],
      memonumberOptions: [],
      typeoptions: [],
      ncType: "",
      memonumberOptionsall: [],
      NCNumberOptions: [],
      NCNumberselected: [],
      ApprovedAuditSelected: [],
      NCNumber: "",
      NCNumberID: "",
      ApprovedAuditReport: "",
      MemoNumber: "",
      department: "",
      departmentCode: "",
      serialNo: 0,
      criteria: "",
      revisionNo: "",
      issueNo: "",
      ncrNo: "",
      documentCode: "",
      revisionDate: null,
      referenceNo: "",
      issueDate: null,
      closeOutStatus: "",
      categoryCheckOption: [],
      categoryValueIsCheck: [],
      subCategoryCheckOption: [],
      subCategoryIsCheck: [],
      locationCheckOption: [],
      locationValueIsCheck: [],
      assignTo: "",
      assignToId: null,
      problemDescription: "",
      dueDate: null,
      submitStatus: "",
      currentUserRole: "",
      firstInitiatorSubmitStatus: "",
      attachmentPreArray: [],
      attachmentJson: [],
      errors: {},
      showDialog: false,
      copyFil: [],
      fileCount: 0,
      exFiles: [],
      fileDeleteId: [],
      files: {} as FileList,
      siteurl: this.props.context.pageContext.web.absoluteUrl,
    };
    this.handleFileChange = this.handleFileChange.bind(this);
    this._OpenModal = this._OpenModal.bind(this);
    this._CloseModal = this._CloseModal.bind(this);
    this.removeFiles = this.removeFiles.bind(this);
    this.toBeDeleted = this.toBeDeleted.bind(this);

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
  private _OpenModal() {
    this.setState({
      showDialog: true
    });
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
    });
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
  public changeDepartment = (item: any): void => {
    let selecteddepartment = this.state.departmentOption.filter((x: any) => x.value == item.value);
    this.setState({ department: item.value, departmentCode: item.data.departmentCode, departmentselected: selecteddepartment });
  };
  // private onChangenctype = (name: string, value: string) => {
  //   debugger
  //   this.setState({ ncType: value })
  // }
  private onChangenctype = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({
        ncType: option.key as string,
        memonumberOptions: option?.text == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs
      });
    }

  };

  public changeMemoNumber = async (item: any): Promise<void> => {
    const sp = spfi().using(SPFx(this.props.context));
    const nctypenew: string = this.state.ncType === "NC" ? "NC Number" : "Observation Number";
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

    const uniqueOptions = await this.getUniqueBy(optionsNCNumber, "ncNo");
    uniqueOptions.sort((a, b) => a.label.localeCompare(b.label));
    let approvedauditreportselected = this.state.memonumberOptions.filter((x: any) => x.value == item.value);
    this.setState({
      NCNumberOptions: uniqueOptions,
      ApprovedAuditReport: Number(item.value),
      MemoNumber: item.reportCode,
      ApprovedAuditSelected: approvedauditreportselected
    });
    console.log("ApprovedAuditSelected", approvedauditreportselected);
  };

  // public changeMemoNumber = (_event: React.FormEvent<HTMLDivElement>, item: any): void => {
  //   const seen = new Set<string>(); // to track unique NC numbers
  //   debugger
  //   const optionsNCNumber = this.state.memonumberOptions
  //     .filter((x) => x.memoNumber === item.text)
  //     .filter((x) => {
  //       if (!x.ncNo || seen.has(x.ncNo)) {
  //         return false;
  //       }
  //       seen.add(x.ncNo);
  //       return true;
  //     })
  //     .map((item: any) => ({
  //       key: item.ID,
  //       text: item.ncNo,
  //       ncNo: item.ncNo,
  //     }));

  //   this.setState({
  //     NCNumberOptions: optionsNCNumber,
  //     ApprovedAuditReport: item.key,
  //     MemoNumber: item.memoNumber,
  //   });
  // };

  public changeNCNumber = (item: any): void => {
    let ncnumberselected = this.state.NCNumberOptions.filter((x: any) => x.value == item.value);
    this.setState({ NCNumber: item.label, NCNumberID: item.value, NCNumberselected: ncnumberselected });
  };
  public handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    debugger
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      [name]: value,
    }));
  };
  public handleChangeobsdescriptiob = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    debugger
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      problemDescription: value,
    }));
  };
  private _handlePeoplePickerChange = (field: keyof IState, idField: keyof IState) => (items: any[]) => {
    if (items.length > 0) {
      this.setState({
        [field]: items[0].text,
        [idField]: items[0].id,
      } as Pick<IState, keyof IState>);
    } else {
      this.setState({
        [field]: "",
        [idField]: null,
      } as unknown as Pick<IState, keyof IState>);
    }
  };

  private _handleCheckboxChange = (stateKey: keyof IState, itemKey: number) =>
    (_ev: React.FormEvent<HTMLElement>, isChecked?: boolean) => {
      this.setState((prevState) => {
        const updatedValues = isChecked
          ? [...(prevState[stateKey] as number[]), itemKey]
          : (prevState[stateKey] as number[]).filter((key) => key !== itemKey);
        return { [stateKey]: updatedValues } as unknown as Pick<IState, keyof IState>;
      });
    };

  public async componentDidMount() {
    await this.getDepartment();
    await this.getFiles();
    await this.getchangerequestdetails();
    //await this.getDepartment();
    await this.getAuditreport();
    await this.getnctypeoptions();

  }
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
  public async getchangerequestdetails() {
    const sp = spfi().using(SPFx(this.props.context));
    let ChangeRequestTemplateType = await getLatestChangeRequestTemplateType(sp, CONTENTTYPE_NonComformity);
    debugger
    if (ChangeRequestTemplateType.length > 0) {
      const template = ChangeRequestTemplateType[0];
      console.log("template", template);
      this.setState(prevData => ({
        ...prevData,
        documentCode: template.DocumentCode || "",
        revisionNo: template.RevisionNumber,
        issueNo: template.IssueNumber,
        referenceNo: template.ReferenceNumber || "",
        revisionDate: template.RevisionDate == null ? null : new Date(template.RevisionDate).toLocaleDateString("en-CA"),
        issueDate: template.IssueDate == null ? null : new Date(template.IssueDate).toLocaleDateString("en-CA"),
      }));
    } else {
      this.setState(prevData => ({
        ...prevData,
        documentCode: "",
        revisionNo: "",
        issueNo: "",
        referenceNo: "",
        revisionDate: "",
        issueDate: "",
      }));
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
  // private getnctypeoptions = async () => {
  //   let arr = []
  //   const _sp = spfi().using(SPFx(this.props.context));
  //   const field2 = await _sp.web.lists.getByTitle("NonConformityList").fields.getByInternalNameOrTitle("NCType")()
  //   console.log(field2, 'field2');
  //   arr = field2["Choices"];
  //   this.setState({ typeoptions: arr })
  //   return arr;
  // }
  private getnctypeoptions = async () => {
    const _sp = spfi().using(SPFx(this.props.context));
    const field2 = await _sp.web.lists.getByTitle("NonConformityList").fields.getByInternalNameOrTitle("NCType")();
    const choices: string[] = field2["Choices"];

    // Convert string array to dropdown option objects
    const dropdownOptions = choices.map(choice => ({
      key: choice,
      text: choice
    }));

    this.setState({ typeoptions: dropdownOptions });
  };

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
      optionsmemoNumbernewnc.sort((a, b) => a.label.localeCompare(b.label));
      //let optionsmemoNumbernewobs: any[] = [];
      optionsmemoNumbernewobs = await this.getUniqueBy(optionsObservationNumber, "reportCode");
      optionsmemoNumbernewobs.sort((a, b) => a.label.localeCompare(b.label));
      this.setState({
        memonumberOptions: this.state.ncType == "NC" ? optionsmemoNumbernewnc : optionsmemoNumbernewobs,
        memonumberOptionsall: memoItems.length > 0 ? memoItems : []
      });
    } catch (e) {
      console.error(e);
    }
  }

  // public async getAuditreport() {
  //   debugger
  //   const sp = spfi().using(SPFx(this.props.context));
  //   try {
  //     const memoItems = await getMemoNumberAuditReport(sp);
  //     let optionsmemoNumber:any[];
  //     // Create a Set to track unique memo numbers
  //     const seen = new Set<string>();
  //     if (memoItems.length >0){
  //        optionsmemoNumber = memoItems[0]
  //     }

  //     optionsmemoNumber.filter((item: any) => {
  //         if (!item.MemoNumber || seen.has(item.MemoNumber)) {
  //           return false;
  //         }
  //         seen.add(item.MemoNumber);
  //         return true;
  //       })
  //       .map((item: any) => ({
  //         key: item.ID,
  //         text: item.MemoNumber,
  //         itemId: item.ID,
  //         memoNumber: item.MemoNumber,
  //         ncNo: item.NCNumber
  //       }));

  //     this.setState({ memonumberOptions: optionsmemoNumber });
  //   } catch (e) {
  //     console.error(e);
  //   }
  // }

  public async getFiles() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const upFiles = await sp.web.lists.getByTitle("NonConformityDocs").items.select("*", "File/Name,FileLeafRef,FileRef,EncodedAbsUrl").expand("File").filter("NonConformityId eq '" + this.props.edItm + "'")();
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
      //Get the NC list   
      const latestItem = await sp.web.lists
        .getByTitle("NonConformityList").items.select("Id", "SerialNumber", "Created", "SubmitStatus")
        .orderBy("SerialNumber", false).top(1)();
      if (latestItem.length > 0) {
        var sNo = latestItem[0].SerialNumber + 1;
        if (sNo < 999)
          sNo = ("0000" + sNo).slice(-3);
        this.setState({ serialNo: sNo })
      }
      else {
        var sNo: any = "000";
        this.setState({ serialNo: sNo })
      }
    } catch (e) {
      console.error(e);
    }
  }

  // private getCurrentUser = async (sp, siteUrl) => {
  //   let arr = []
  //   await _sp.web.currentUser()
  //     .then(async (res) => {
  //       console.log(res);
  //       arr = res;
  //       const ProfilePic = `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${res.Email}`
  //       //await getUserProfilePicture(res.Id,_sp)
  //     })
  //     .catch((error) => {
  //       console.log("Error fetching data: ", error);
  //     });
  //   return arr;
  // }

  public async getDepartment() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      // const deptItems = await sp.web.lists.getByTitle("DepartmentMasterList").items();
      const deptItems = await sp.web.lists.getByTitle("DepartmentMasterList").items.orderBy("Title", true)();
      const options = deptItems.map((item: {
        DepartmentCode: any; Title: string; Id: number
      }) => ({
        value: item.Id,
        label: item.Title,
        data: { departmentCode: item.DepartmentCode },
      }));
      this.setState({ departmentOption: options });
      //const Currusers: any = await this.getCurrentUser(sp, this.state.siteurl);
      const userProfile = await sp.profiles.myProperties();
      const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
      let currentuserdepartment = UserDept == "IT" ? "Information Technology" : UserDept;
      const selectedOption = options.find(user => user?.label === currentuserdepartment);
      this.setState({ department: selectedOption?.value, departmentCode: selectedOption?.data.departmentCode, departmentselected: selectedOption });
      //this.setState({ department: selectedOption?.key });
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
      debugger
      console.log("options cate", options);
      this.setState({ categoryCheckOption: options });
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
      debugger
      console.log("options sub", options);
      this.setState({ subCategoryCheckOption: options });
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
      console.log("options loc", options);
      this.setState({ locationCheckOption: options });
    } catch (e) {
      console.error(e);
    }
  }
  // public validateFormSubmit = (): boolean => {
  //   let errors: { [key: string]: string } = {};
  //   if (!this.state.department) errors.department = "Department is required";
  //   if (!this.state.criteria) errors.criteria = "Criteria is required";
  //   if (!this.state.closeOutStatus) errors.closeOutStatus = "Close Out Status is required";
  //   if (this.state.categoryValueIsCheck.length == 0) {
  //     Swal.fire({ title: "Please select at least one category!" });
  //     document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   }
  //   else {
  //     document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }
  //   if (this.state.subCategoryIsCheck.length == 0) {
  //     Swal.fire({ title: "Please select at least one Sub category!" });
  //     document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   }
  //   else {
  //     document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }
  //   if (this.state.locationValueIsCheck.length == 0) {
  //     Swal.fire({ title: "Please select at least one location!" });
  //     document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   }
  //   else {
  //     document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }
  //   if (!this.state.assignTo) errors.assignTo = "AssignTo is required";
  //   if (!this.state.dueDate) errors.dueDate = "dueDate is required";
  //   if (!this.state.fileCount && this.state.exFiles.length == 0) {
  //     errors.Attchments = "Attachments are required";
  //   }
  //   if (!this.state.problemDescription) errors.problemDescription = "Problem Description is required";
  //   this.setState({ errors });
  //   return Object.keys(errors).length === 0;
  // };
  public validateFormSubmit = (): boolean => {
    let errors: { [key: string]: string } = {};
    let isValid = true;
    if (!this.state.ncType) {
      errors.nctype = "NC Type is required";
      isValid = false;
    }
    if (this.state.ApprovedAuditSelected.length == 0) {
      errors.approvedauditreport = "Approved Report code is required";
      isValid = false;
    }
    if (this.state.NCNumberselected.length == 0) {
      errors.ncnumber = "NC/Observation number is required";
      isValid = false;
    }
    // if (this.state.departmentselected.length == 0) {
    //   errors.department = "Department is required";
    //   isValid = false;
    // }
    if (!this.state.criteria) {
      errors.criteria = "Criteria is required";
      isValid = false;
    }
    if (!this.state.closeOutStatus) {
      errors.closeOutStatus = "Close Out Status is required";
      isValid = false;
    }

    // Validate Category
    if (this.state.categoryValueIsCheck.length == 0) {
      errors.category = "At least one Category is required";
      isValid = false;
      document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.add(styles.errCh);
      });
    } else {
      document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.remove(styles.errCh);
      });
    }

    // Validate SubCategory
    if (this.state.subCategoryIsCheck.length == 0) {
      errors.subCategory = "At least one Sub Category is required";
      isValid = false;
      document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.add(styles.errCh);
      });
    } else {
      document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.remove(styles.errCh);
      });
    }

    // Validate Location
    if (this.state.locationValueIsCheck.length == 0) {
      errors.location = "At least one Location is required";
      isValid = false;
      document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.add(styles.errCh);
      });
    } else {
      document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.remove(styles.errCh);
      });
    }

    if (!this.state.assignTo) {
      errors.assignTo = "AssignTo is required";
      isValid = false;
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
    if (!this.state.dueDate) {
      errors.dueDate = "Due Date is required";
      isValid = false;
    }
    // if (!this.state.fileCount && this.state.exFiles.length == 0) {
    //   errors.Attchments = "Attachments are required";
    //   isValid = false;
    //   document.querySelectorAll("#newfile").forEach((el) => {
    //     el.classList.remove(styles.errCh);
    //   });

    // } else {
    //   document.querySelectorAll("#newfile").forEach((el) => {
    //     el.classList.remove(styles.errCh);
    //   });
    // }
    if (!this.state.problemDescription) {
      errors.problemDescription = "NC/Observation Description is required";
      isValid = false;
    }

    this.setState({ errors });

    if (!isValid) {
      Swal.fire('Please fill all the mandatory fields.');
    }

    return isValid;
  };
  public validateFormDraft = (): boolean => {
    let errors: { [key: string]: string } = {};

    // Allow Save as Draft if at least department is selected
    if (!this.state.ncType) {
      errors.nctype = "NC Type is required";

    }
    if (this.state.ApprovedAuditSelected.length == 0) {
      errors.approvedauditreport = "Approved Report code is required";

    }
    if (this.state.NCNumberselected.length == 0) {
      errors.ncnumber = "NC/Observation number is required";

    }
    // if (this.state.departmentselected.length == 0) {
    //   errors.department = "Department is required";
    //   Swal.fire('Please select a Department.');
    // }

    // You can skip checking other fields if department is selected
    this.setState({ errors });
    return Object.keys(errors).length === 0 || (this.state.department && Object.keys(errors).length === 2 && errors.hasOwnProperty("department") === false);
  };
  // public validateFormDraft = (): boolean => {
  //   let errors: { [key: string]: string } = {};

  //   // Department validation
  //   // if (!this.state.department) errors.department = "Department is required";
  //   if (!this.state.department) {
  //     errors.department = "Department is required";
  //     Swal.fire('Please select a Department.');
  //   }

  //   // Criteria validation
  //   if (!this.state.criteria) errors.criteria = "Criteria is required";

  //   // Close Out Status validation
  //   if (!this.state.closeOutStatus) errors.closeOutStatus = "Close Out Status is required";

  //   // Category validation (at least one checkbox must be selected)
  //   if (this.state.categoryValueIsCheck.length == 0) {
  //     errors.category = "Please select at least one category!";
  //     document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   } else {
  //     document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }

  //   // Sub-Category validation (at least one checkbox must be selected)
  //   if (this.state.subCategoryIsCheck.length == 0) {
  //     errors.subCategory = "Please select at least one Sub-Category!";
  //     document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   } else {
  //     document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }

  //   // Location validation (at least one checkbox must be selected)
  //   if (this.state.locationValueIsCheck.length == 0) {
  //     errors.location = "Please select at least one location!";
  //     document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.add(styles.errCh);
  //     });
  //   } else {
  //     document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
  //       el.classList.remove(styles.errCh);
  //     });
  //   }

  //   // Assigned To validation
  //   if (!this.state.assignTo) errors.assignTo = "Assigned To is required";

  //   // Due Date validation
  //   if (!this.state.dueDate) errors.dueDate = "Due Date is required";

  //   // Attachments validation (either new files or existing files must be present)
  //   if (!this.state.fileCount && this.state.exFiles.length == 0) {
  //     errors.Attchments = "Attachments are required";
  //   }

  //   // Problem Description validation
  //   if (!this.state.problemDescription) errors.problemDescription = "Problem Description is required";

  //   this.setState({ errors });
  //   return Object.keys(errors).length === 0 || (this.state.department && Object.keys(errors).length === 1 && errors.hasOwnProperty("department") === false);
  // };

  public handleSubmit = (formsubmode: string) => {
    if (this.validateFormSubmit()) {
      this._saveData(formsubmode);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };

  public handleSubmitDraft = (formsubmode: string) => {
    if (this.validateFormDraft()) {
      this._saveData(formsubmode);
    }
    else {
      Swal.fire('Please fill the mandatory fields.');
    }
  };
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
          // window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/NC.aspx#/listing";
          window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          // window.location.reload();
        });
      }
    });
  }
  //Save Function
  private _saveData = async (_submitStatus: string) => {
    debugger
    let mText = "";
    let cText = "";
    let substatus = "";
    let currentUserRole = "";
    let firstInitiatorSubmitStatus = ""
    let firstAssignedToSubmitStatus = "";
    let serialNumber: any;
    let documentCode = "";
    let ncrnumber = "";

    if (_submitStatus == "submit") {
      substatus = "Yes";
      currentUserRole = "FirstAssignedTo";
      firstInitiatorSubmitStatus = "Yes";
      firstAssignedToSubmitStatus = "No";
      mText = "submit";
      cText = "Submitted";
      serialNumber = this.state.serialNo;
      ncrnumber = 'NC/' + this.state.departmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.serialNo;
      //documentCode = 'NC/' + this.state.departmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.serialNo;
    }
    else {
      substatus = "No";
      currentUserRole = "";
      firstInitiatorSubmitStatus = "No";
      firstAssignedToSubmitStatus = "No";
      mText = "save";
      cText = "Saved";
      serialNumber = 0;
      ncrnumber = "";
      documentCode = "";
    }
    const { context } = this.props;
    const sp = spfi().using(SPFx(this.props.context));
    let IssueNumber = this.state.issueNo !== "" ? Number(this.state.issueNo) : null;
    let RevisionNumber = this.state.revisionNo !== "" ? Number(this.state.revisionNo) : null;
    Swal.fire({
      title: "Do you want to " + mText + " this request?",
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(async (result) => {
      if (result.isConfirmed) {
        debugger

        this.setState({ Loading: true });
        await sp.web.lists.getByTitle("NonConformityList").items.add({
          NCNumber: this.state.NCNumber,
          NCType: this.state.ncType,
          NCNumberID: Number(this.state.NCNumberID),
          ApprovedAuditReportId: this.state.ApprovedAuditReport || null,
          ApprovedAuditReportMemoNumber: this.state.MemoNumber,
          NCRNo: ncrnumber,
          DocumentCode: this.state.documentCode,
          IssueNumber: this.state.issueNo !== "" ? Number(this.state.issueNo) : null,
          RevisionNumber: this.state.revisionNo !== "" ? Number(this.state.revisionNo) : null,
          IssueDate: this.state.issueDate || null,
          RevisionDate: this.state.revisionDate || null,
          ReferenceNumber: this.state.referenceNo,
          DepartmentId: this.state.department || null,
          Criteria: this.state.criteria,
          CloseOutStatus: this.state.closeOutStatus,
          CategoryId: this.state.categoryValueIsCheck,
          SubCategoryId: this.state.subCategoryIsCheck,
          LocationId: this.state.locationValueIsCheck,
          AssignedToId: this.state.assignToId || null,
          DueDate: this.state.dueDate,
          ProblemDescription: this.state.problemDescription,
          SubmiitedDate: new Date(),
          SubmitStatus: substatus,
          SubmiitedById: this.props.currentUserID,
          CurrentUserRole: currentUserRole,
          FirstInitiatorSubmitStatus: firstInitiatorSubmitStatus,
          FirstAssignedToSubmitStatus: firstAssignedToSubmitStatus,
          Status: "Pending",
          SerialNumber: serialNumber,
          //DocumentCode: documentCode,
        }).then((i: any) => {

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
                    NonConformityId: i.data.Id
                  });
                });
              });
            })
          }
          this.setState({ Loading: false });
        });

        Swal.fire({
          title: cText + " Successfully.",
          icon: "success"
        }).then(() => {
          // window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/NC.aspx#/listing";
          window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/EDCMAIN.aspx";
          // window.location.reload();
        });
      }
    })
      .catch(error => {
        console.error("Error while saving:", error);
      });
  }

  public render(): React.ReactElement<IAuditPlanProps> {
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
    document.querySelectorAll("#AssigntoPeoplepicker .ms-BasePicker-text").forEach((el) => {
      el.classList.add(styles.peoplepickerstyleAuditte);
    });
    var fileData = this.state.copyFil.map((item: any, i: number) => {
      return (
        <tr>
          <td>{i + 1}</td>
          <td>
            {item.name}
          </td>
          {/* <td></td> */}
          <td title={moment(item.Uploaded).format("DD/MMM/YYYY")}>{moment(item.Uploaded).format("DD/MMM/YYYY")}</td>
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
          <td style={{ minWidth: "60px", maxWidth: "60px", textAlign: 'center' }}>
            <img src={require("../assets/del.png")} className='' onClick={() => this.toBeDeleted(i)}></img>
          </td>

        </tr>
      )
    });
    return (
      <section className='card card-body'>
        <div className={styles.welcome}>
          <div className="row">
            <div className="col-lg-6 newbread">
              <CustomBreadcrumb Breadcrumb={this.Breadcrumb} />
            </div>

          </div>
          {this.state.Loading ?

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
            <form>
              {/* Section 1 */}
              <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row">
                <div className="form-group col-md-12"><h3 className='text-dark font-16 text-left fw-bold mb-3'>Non Conformity / Observation Details</h3></div>
              </div>
              <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                {console.log("this.state.typeoptions", this.state.typeoptions)}
                <div className="form-group col-md-4 mb-3">
                  <TooltipHost
                    content={this.state.ncType || ""}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block', width: '100%' } }}
                  >
                    <Dropdown
                      required
                      placeholder="Category"
                      label="Category:"
                      options={this.state.typeoptions}
                      selectedKey={this.state.ncType}
                      onChange={this.onChangenctype}
                      className={this.state.errors?.nctype ? 'dropdown-error' : ''}
                    />

                    {/* <Dropdown
                      required
                      placeholder="Type"
                      label="Type:"
                      options={this.state.typeoptions}
                      defaultSelectedKey={this.state.ncType}
                      selectedKey={this.state.ncType}
                      onChange={(event, option) =>
                        this.onChangenctype("ncType", option?.key)
                      }
                      onChange={(e: any) =>
                        this.onChangenctype(e.target.name, e.target.value)
                      }
                      className={this.state.errors?.nctype ? 'dropdown-error' : ''}
                    /> */}
                  </TooltipHost>
                </div>
                <div className="form-group col-md-4 mb-3">
                  <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }} >Approved Report Code:<span className="text-danger1">*</span>
                  </label>
                  <TooltipHost
                    content={this.state.memonumberOptions.filter((item: any) => item.key == this.state.ApprovedAuditReport)[0]?.text || ""}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block', width: '100%' } }}
                  >
                    <Select
                      options={this.state.memonumberOptions}
                      value={this.state.ApprovedAuditSelected}
                      name="AuditPlan"
                      isClearable={true}
                      isSearchable={true}
                      className={` ${this.state.errors?.approvedauditreport ? "border-on-error" : ""}`}
                      //className={? 'dropdown-error' : ''}
                      onChange={(selectedOption: any) => this.changeMemoNumber(selectedOption)}
                      placeholder={"Approved Report Code"}

                    />
                    {/* <Dropdown
                      required
                      placeholder="Approved Report Code"
                      label="Approved Report Code:"
                      options={this.state.memonumberOptions}
                      defaultSelectedKey={this.state.ApprovedAuditReport}
                      selectedKey={this.state.ApprovedAuditReport}
                      onChange={this.changeMemoNumber}
                      className={this.state.errors?.approvedauditreport ? 'dropdown-error' : ''}
                    /> */}
                  </TooltipHost>
                </div>
                <div className="form-group col-md-4 mb-3">
                  <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }} >NC/Observation Number:<span className="text-danger1">*</span>
                  </label>
                  <TooltipHost
                    content={this.state.NCNumber || ""}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block', width: '100%' } }}
                  >
                    <Select
                      options={this.state.NCNumberOptions}
                      value={this.state.NCNumberselected}
                      name="ncnumber"
                      isClearable={true}
                      isSearchable={true}
                      className={this.state.errors?.ncnumber ? 'border-on-error' : ''}
                      onChange={(selectedOption: any) => this.changeNCNumber(selectedOption)}
                      placeholder={"NC/Observation Number"}

                    />
                    {/* <Dropdown
                      required
                      placeholder="NC/Observation Number"
                      label="NC/Observation Number:"
                      options={this.state.NCNumberOptions}
                      defaultSelectedKey={this.state.NCNumberID}
                      selectedKey={this.state.NCNumberID}
                      onChange={this.changeNCNumber}
                      className={this.state.errors?.ncnumber ? 'dropdown-error' : ''}
                    /> */}
                  </TooltipHost>
                </div>
                {/* <div className="form-group col-md-4">
                <TextField label="NCR No:" name='NCRNo' required value={this.state.ncrNo} disabled={true} onChange={this.handleChange}

                // styles={{
                //   fieldGroup: {
                //     backgroundColor: this.state.errors.ncrno ? "#ffcccb" : "white",
                //   },
                // }}
                />
              </div> */}
                <div className="form-group col-md-4 mb-3">
                  <TextField label="Document Code:" name='DocumentCode' required value={this.state.documentCode} disabled={true} onChange={this.handleChange}
                    styles={{
                      fieldGroup: {
                        backgroundColor: this.state.errors.documentcode ? "#ffcccb" : "white",
                      },
                    }}
                  />
                </div>
                <div className="form-group col-md-4 mb-3">
                  <TextField label="Issue Number:" name='IssueNumber' required value={this.state.issueNo + ""} disabled={true} onChange={this.handleChange}
                    styles={{
                      fieldGroup: {
                        backgroundColor: this.state.errors.issuenumber ? "#ffcccb" : "white",
                      },
                    }}
                  />
                </div>
                <div className="form-group col-md-4 mb-3">
                  <TextField label="Revision Number:" name='RevisionNumber' required value={this.state.revisionNo + ""} disabled={true} onChange={this.handleChange}
                    styles={{
                      fieldGroup: {
                        backgroundColor: this.state.errors.revisionnumber ? "#ffcccb" : "white",
                      },
                    }}
                  />
                </div>
                <div className="form-group col-md-4 mb-3">
                  <label htmlFor="DocumentCode" style={{ marginBottom: '10px' }} >Department:<span className="text-danger1">*</span>
                  </label>
                  <Select
                    options={this.state.departmentOption}
                    value={this.state.departmentselected}
                    name="Department"
                    isClearable={true}
                    isSearchable={true}
                    className={this.state.errors?.department ? 'border-on-error' : ''}
                    onChange={(selectedOption: any) => this.changeDepartment(selectedOption)}
                    placeholder={"Department"}

                  />
                  {/* <Dropdown
                    required
                    placeholder="Department"
                    label="Department:"
                    options={this.state.departmentOption}
                    defaultSelectedKey={this.state.department}
                    selectedKey={this.state.department}
                    onChange={this.changeDepartment}
                    className={this.state.errors?.department ? 'dropdown-error' : ''}
                  // styles={{
                  //   title: {
                  //     backgroundColor: this.state.errors.department ? "#ffcccb" : "white", // Light red when error
                  //   },
                  // }}
                  /> */}
                </div>
                <div className="col-md-4 mb-3">
                  <TextField label="Criteria:" name='criteria' required value={this.state.criteria} onChange={this.handleChange}
                    className={this.state.errors?.criteria ? 'textfield-error' : ''}
                  // styles={{
                  //   fieldGroup: {
                  //     border: this.state.errors?.criteria ? '1px solid red !important' : undefined,
                  //     backgroundColor: this.state.errors?.criteria ? '#ffcccb !important' : 'white',
                  //   },
                  //   field: {
                  //     backgroundColor: this.state.errors?.criteria ? '#ffcccb' : 'white',
                  //   },
                  // }}
                  />

                </div>
                <div className="form-group col-md-4 mb-3">
                  <TextField label="Close Out Status:" name='closeOutStatus' required value={this.state.closeOutStatus} onChange={this.handleChange}
                    className={this.state.errors?.closeOutStatus ? 'textfield-error' : ''}
                  // styles={{
                  //   fieldGroup: {
                  //     backgroundColor: this.state.errors.closeOutStatus ? "#ffcccb" : "white",
                  //   },
                  // }}
                  />
                </div>
              </div>
              <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                <div className="form-group col-md-4" id="categoryCheckbox">
                  <label>Category: <span className={styles.textdanger}>*</span></label>
                  {this.state.categoryCheckOption.map((item: any) => {
                    return (
                      <div style={{ margin: "2px", padding: "3px" }}>
                        <Checkbox label={item.text} onChange={this._handleCheckboxChange("categoryValueIsCheck", item.key as number)} />
                      </div>
                    );
                  }
                  )}
                </div>
                <div className="form-group col-md-4" id="SubCategoryCheckbox">
                  <label>SubCategory: <span className={styles.textdanger}>*</span></label>
                  {this.state.subCategoryCheckOption.map((item: any) => {
                    return (
                      <div style={{ margin: "2px", padding: "3px" }}>
                        <Checkbox label={item.text} onChange={this._handleCheckboxChange("subCategoryIsCheck", item.key as number)} />
                      </div>
                    );
                  }
                  )}
                </div>
                <div className="form-group col-md-4" id="locationCheckbox">
                  <label>Location: <span className={styles.textdanger}>*</span></label>
                  {this.state.locationCheckOption.map((item: any) => {
                    return (
                      <div style={{ margin: "2px", padding: "3px" }}>
                        <Checkbox label={item.text}
                          onChange={this._handleCheckboxChange("locationValueIsCheck", item.key as number)}
                          className={this.state.errors?.criteria ? 'textfield-error' : ''}
                        />
                      </div>
                    );
                  }
                  )}
                </div>
              </div>
              <div style={{ justifyContent: 'left', textAlign: 'left' }} className="row mb-3">
                <div className="form-group col-md-4" id="AssigntoPeoplepicker">
                  <PeoplePicker
                    context={peoplePickerContext}
                    titleText="Auditee:"
                    personSelectionLimit={1}
                    required={true}
                    groupName={""} // Leave this blank in case you want to filter from all users
                    showtooltip={true}
                    disabled={false}
                    ensureUser={true}
                    onChange={this._handlePeoplePickerChange("assignTo", "assignToId")}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    styles={{
                      root: {
                        backgroundColor: this.state.errors.assignTo ? "#ffcccb" : "white",
                      },
                    }}
                  />
                </div>
                <div className="form-group col-md-4">
                  <Label>
                    Due Date <span className={styles.textdanger}>*</span>
                  </Label>
                  <DatePicker
                    formatDate={(date: Date) => moment(date).format("DD/MMM/YYYY")}
                    placeholder="Select a Due Date"
                    value={this.state.dueDate}
                    onSelectDate={(date: Date) => this.setState({ dueDate: date })}
                    //styles={this.state.errors.dueDate ? datePickerErrorStyles : {}}
                    className={this.state.errors?.dueDate ? 'textfield-error' : ''}
                  />
                </div>
                <div style={{ position: 'relative' }} className="col-lg-4 mt-1">
                  <label htmlFor="Attchments" style={{ marginRight: "10px" }}>Attachments </label>

                  <input
                    className="form-control"
                    type="file" name="myFile" onChange={(e) => this.handleFileChange(e, this)} id="newfile" multiple
                  // style={{
                  //   backgroundColor: this.state.errors.Attchments ? "#ffcccb" : "white",
                  //   borderColor: this.state.errors.Attchments ? 'red' : '#dee2e6'
                  // }}
                  //className={`form-control ${this.state.errors?.Attachments} ? 'textfield-error' : ''`}
                  />
                  {this.state.fileCount > 0 ?
                    (<span style={{ fontSize: '0.875rem' }} onClick={this._OpenModal} className='newpo'>
                      <FontAwesomeIcon icon={faPaperclip} /> {this.state.fileCount} {this.state.fileCount > 0 ? "files" : "file"} Attached
                    </span>) : ""
                  }

                  {/* {this.state.showDialog && <div id="myModal" className={styles.modal}>
                  <div className={styles.modalcontent}>
                    <span><b>Attachment Details</b></span>
                    <br />
                    <span>Below are the attachment details for Non Comfirmity</span>
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
                </div>} */}
                  {this.state.showDialog && (
                    <div id="myModal" className={styles.modal}>
                      <div className={styles.modalcontent}>
                        {/* Close button */}
                        <span className={styles.close} onClick={() => this._CloseModal()}>&times;</span>

                        {/* Modal title and subtitle */}
                        <h4 className="font-16 text-dark fw-bold mb-1">Attachment Details</h4>
                        <p className="text-muted font-14 mb-3 fw-400">Below are the attachment details for Non Conformity / Observation</p>

                        {/* Table */}
                        <table className={styles.mtbalenew}>
                          <thead>
                            <tr>
                              <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                              <th>File Name</th>
                              {/* <th>File Link</th> */}
                              <th style={{ minWidth: '100px' }} className="text-center">Upload Date</th>
                              <th className="text-center">Action</th>
                            </tr>
                          </thead>
                          <tbody >
                            {upFiles}
                            {fileData}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              </div>
              <div style={{ justifyContent: 'left', textAlign: 'left' }} className='row mb-3'>
                <div className="form-group col-md-12 newdes">
                  <TextField label="NC/Observation Description:"
                    required
                    name='problemDescription'
                    value={this.state.problemDescription}
                    multiline rows={5}
                    //onChange={(e) => this.setState({ ...formData, issueNo: e.target.value })}
                    onChange={this.handleChangeobsdescriptiob}
                    //errorMessage={this.state.errors.problemDescription}
                    className={this.state.errors?.problemDescription ? 'textfield-error' : ''}
                  // styles={{
                  //   fieldGroup: {
                  //     backgroundColor: this.state.errors.problemDescription ? "#ffcccb" : "white", // Red tint for errors
                  //   },
                  // }}
                  />
                </div>
              </div>
              {/* Button Section 4 */}
              <div style={{ margin: '10px', justifyContent: 'center', display: 'flex', gap: '5px' }}>

                <PrimaryButton text="Save as Draft" onClick={() => this.handleSubmitDraft("draft")} />

                <PrimaryButton text="Submit" onClick={() => this.handleSubmit("submit")} />


                <DefaultButton text="Cancel" onClick={() => this.cancelRequest()} />

              </div>
            </form>
          }
        </div>

      </section>
    );
  }
}
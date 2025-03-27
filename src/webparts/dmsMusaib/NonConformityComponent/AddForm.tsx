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
  departmentOption: IDropdownOption[];
  department: string | number;
  departmentCode: string;
  serialNo: number;
  criteria: string;
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
}

export default class AuditPlan extends React.Component<IAuditPlanProps, IState> {
  constructor(props: IAuditPlanProps) {
    super(props);
    this.state = {
      departmentOption: [],
      department: "",
      departmentCode: "",
      serialNo: 0,
      criteria: "",
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
  public changeDepartment = (_event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ department: item.key, departmentCode: item.data.departmentCode });
  };

  public handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    this.setState((prevState) => ({
      ...prevState,
      [name]: value,
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
  }

  public async getFiles() {
    const sp = spfi().using(SPFx(this.props.context));
    try {
      const upFiles = await sp.web.lists.getByTitle("NonConformityDocs").items.select("*", "File/Name,FileLeafRef,FileRef,EncodedAbsUrl").expand("File").filter("NonConformityId eq '" + this.props.edItm + "'")();
      if (upFiles.length > 0) {
        var obJFiles: any[] = [];
        let fCount: number = 0;
        upFiles.forEach(function (item: any) {
          obJFiles.push({ "Name": item.File.Name, "type": "old", "Id": item.Id, "Uploaded": new Date(item.Modified).getDate() + "/" + new Date(item.Modified).getMonth() + "/" + new Date(item.Modified).getFullYear(), "Path": item.EncodedAbsUrl })
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
      this.setState({ departmentOption: options });
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
      this.setState({ locationCheckOption: options });
    } catch (e) {
      console.error(e);
    }
  }
  public validateFormSubmit = (): boolean => {
    let errors: { [key: string]: string } = {};
    if (!this.state.department) errors.department = "Department is required";
    if (!this.state.criteria) errors.criteria = "Criteria is required";
    if (!this.state.closeOutStatus) errors.closeOutStatus = "Close Out Status is required";
    if (this.state.categoryValueIsCheck.length == 0) {
      Swal.fire({ title: "Please select at least one category!" });
      document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.add(styles.errCh);       
      });      
    }
    else {
      document.querySelectorAll("#categoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.remove(styles.errCh);
      });
    }
    if (this.state.subCategoryIsCheck.length == 0) {
      Swal.fire({ title: "Please select at least one Sub category!" });
      document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.add(styles.errCh);        
      });     
    }
    else {
      document.querySelectorAll("#SubCategoryCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.remove(styles.errCh);
      });
    }
    if (this.state.locationValueIsCheck.length == 0) {
      Swal.fire({ title: "Please select at least one location!" });
      document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.add(styles.errCh);
      });      
    }
    else {
      document.querySelectorAll("#locationCheckbox .ms-Checkbox-checkbox").forEach((el) => {
        el.classList.remove(styles.errCh);
      });
    }
    if (!this.state.assignTo) errors.assignTo = "AssignTo is required";
    if (!this.state.dueDate) errors.dueDate = "dueDate is required";
    if (!this.state.fileCount && this.state.exFiles.length == 0) {
      errors.Attchments = "Attachments are required";
    }
    if (!this.state.problemDescription) errors.problemDescription = "Problem Description is required";
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };
  public validateFormDraft = (): boolean => {
    let errors: { [key: string]: string } = {};
    if (!this.state.problemDescription) errors.problemDescription = "Problem Description is required";
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

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
          window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/NC.aspx#/listing";
        });
      }
    });
  }
  //Save Function
  private async _saveData(_submitStatus: string) {
    let mText = "";
    let cText = "";
    let substatus = "";
    let currentUserRole = "";
    let firstInitiatorSubmitStatus = ""
    let firstAssignedToSubmitStatus = "";
    let serialNumber: any;
    let documentCode = "";

    if (_submitStatus == "submit") {
      substatus = "Yes";
      currentUserRole = "FirstAssignedTo";
      firstInitiatorSubmitStatus = "Yes";
      firstAssignedToSubmitStatus = "No";
      mText = "submit";
      cText = "Submitted";
      serialNumber = this.state.serialNo;
      documentCode = 'NC/' + this.state.departmentCode + "/" + moment(new Date()).format("MM") + "/" + this.state.serialNo;
    }
    else {
      substatus = "No";
      currentUserRole = "";
      firstInitiatorSubmitStatus = "No";
      firstAssignedToSubmitStatus = "No";
      mText = "save";
      cText = "Saved";
      serialNumber = 0;
      documentCode = "";
    }
    const { context } = this.props;
    const sp = spfi().using(SPFx(this.props.context));
    Swal.fire({
      title: "Do you want to " + mText + " this request?",
      showCancelButton: true,
      confirmButtonText: 'Yes',
      cancelButtonText: 'No'
    }).then(async (result) => {
      if (result.isConfirmed) {
        await sp.web.lists.getByTitle("NonConformityList").items.add({
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
          DocumentCode: documentCode,
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
        });
        Swal.fire({
          title: cText + " Successfully.",
          icon: "success"
        }).then(() => {
          window.location.href = context.pageContext.web.absoluteUrl + "/SitePages/NC.aspx#/listing";
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
    var fileData = this.state.copyFil.map((item: any, i: number) => {
      return (
        <tr>
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
        <tr>
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
            <button type="button" onClick={(e) => this.toBeDeleted(i)} >Delete</button>
          </td>
        </tr>
      )
    });
    return (
      <section className={`${styles.auditPlan} `}>
        <div className={styles.welcome}>
          <form>
            {/* Section 1 */}
            <div className="row">
              <div className="form-group col-md-12"><h3>Problem Details</h3></div>
            </div>
            <div className="row">
              <div className="form-group col-md-4">
                <Dropdown
                  required
                  placeholder="Department"
                  label="Department:"
                  options={this.state.departmentOption}
                  defaultSelectedKey={this.state.department}
                  selectedKey={this.state.department}
                  onChange={this.changeDepartment}
                  styles={{
                    title: {
                      backgroundColor: this.state.errors.department ? "#ffcccb" : "white", // Light red when error
                    },
                  }}
                />
              </div>
              <div className="form-group col-md-4">
                <TextField label="Criteria:" name='criteria' required value={this.state.criteria} onChange={this.handleChange}
                  styles={{
                    fieldGroup: {
                      backgroundColor: this.state.errors.criteria ? "#ffcccb" : "white",
                    },
                  }}
                />
              </div>
              <div className="form-group col-md-4">
                <TextField label="Close Out Status:" name='closeOutStatus' required value={this.state.closeOutStatus} onChange={this.handleChange}
                  styles={{
                    fieldGroup: {
                      backgroundColor: this.state.errors.closeOutStatus ? "#ffcccb" : "white",
                    },
                  }}
                />
              </div>
            </div>
            <div className="row">
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
                      <Checkbox label={item.text} onChange={this._handleCheckboxChange("locationValueIsCheck", item.key as number)}
                      />
                    </div>
                  );
                }
                )}
              </div>
            </div>
            <div className="row">
              <div className="form-group col-md-4">
                <PeoplePicker
                  context={peoplePickerContext}
                  titleText="Assigned To:"
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
                  styles={this.state.errors.dueDate ? datePickerErrorStyles : {}}
                />
              </div>
              <div className="col-lg-4">
                <label htmlFor="Attchments" style={{ marginRight: "10px" }}>Attachments <span className={styles.textdanger}>*</span></label>
                <button type="button" onClick={this._OpenModal} >{this.state.fileCount}</button>
                <input className="form-control" type="file" name="myFile" onChange={(e) => this.handleFileChange(e, this)} id="newfile" multiple
                  style={{
                    backgroundColor: this.state.errors.Attchments ? "#ffcccb" : "white",
                  }} />                
                {this.state.showDialog && <div id="myModal" className={styles.modal}>
                  <div className={styles.modalcontent}>
                    <span><b>Attachment Details</b></span>
                    <br />
                    <span>Below are the attachment details for the Initiative</span>
                    <span className={styles.close} onClick={e => this._CloseModal()}>&times;</span>
                    <table>
                      <thead>
                        <th>File Name</th>
                        <th>File Link</th>
                        <th>Upload Date</th>
                        <th>Delete</th>
                      </thead>
                      {upFiles}{fileData}
                    </table>
                  </div>
                </div>}
              </div>
            </div>
            <div className='row'>
              <div className="form-group col-md-12">
                <TextField label="Problem Description:"
                  required
                  name='problemDescription'
                  value={this.state.problemDescription}
                  multiline rows={3}
                  onChange={this.handleChange}
                  //errorMessage={this.state.errors.problemDescription}
                  styles={{
                    fieldGroup: {
                      backgroundColor: this.state.errors.problemDescription ? "#ffcccb" : "white", // Red tint for errors
                    },
                  }}
                />
              </div>
            </div>
            {/* Button Section 4 */}
            <div className="row" style={{ margin: '10px' }}>
              <div className="form-group col-md-2">
                <PrimaryButton text="Submit" onClick={() => this.handleSubmit("submit")} />
              </div>
              <div className="form-group col-md-2">
                <PrimaryButton text="Save as Draft" onClick={() => this.handleSubmitDraft("draft")} />
              </div>
              <div className="form-group col-md-2">
                <DefaultButton text="Cancel" onClick={() => this.cancelRequest()} />
              </div>
            </div>
          </form>
        </div>
      </section>
    );
  }
}
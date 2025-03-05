import * as React from 'react';
import styles from '../AuditApp.module.scss';
import type { IFormProps } from './IFormProps';
import {Checkbox,  PrimaryButton ,TextField} from '@fluentui/react';
import { Field, Textarea } from "@fluentui/react-components";

import CustomBreadcrumb from '../../ChangerequestComponent/CustomBreadcrumb/CustomBreadcrumb';
import { IState } from '../IState';
import  Swal from 'sweetalert2';
import "@pnp/sp/site-users/web";
import { Caching } from "@pnp/queryable";
// import { getSP } from "../PNPJsConfig";
// import { SPFI, spfi } from "@pnp/sp";
import { Dropdown,  IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { DateTimePicker, DateConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker'; 
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { getSP } from "../../loc/pnpjsConfig";
import { SPFI , spfi} from "@pnp/sp";
//import {  useHistory } from 'react-router-dom'
//const _dropdown = 
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};
//const navigate:any= useHistory();
const dropdownStylesNew: Partial<IDropdownStyles> = {
  dropdown: { width: 80 },
};
const optionsYear: IDropdownOption[] = [
  
  { key: 2024, text: '2024' },
  { key: 2025, text: '2025' },
  { key: 2026, text: '2026' }

 
];
const optionsApp: IDropdownOption[] = [
  
  { key: 'All', text: 'All' },
  { key: 'One', text: 'One' }

]
const optionsHours: IDropdownOption[] = [
  
  { key: '1', text: '1' },
  { key: '2', text: '2' },
  { key: '3', text: '3' },
  { key: '4', text: '4' },
  { key: '5', text: '5' },
  { key: '6', text: '6' },
  { key: '7', text: '7' },
  { key: '8', text: '8' },
  { key: '9', text: '9' },
  { key: '10', text: '10' },
  { key: '11', text: '11' },
  { key: '12', text: '12' },
  { key: '13', text: '13' },
  { key: '14', text: '14' },
  { key: '15', text: '15' },
  { key: '16', text: '16' },
  { key: '17', text: '17' },
  { key: '18', text: '18' },
  { key: '19', text: '19' },
  { key: '20', text: '20' },
  { key: '21', text: '21' },
  { key: '22', text: '22' },
  { key: '23', text: '23' },
  { key: '24', text: '24' }
];
const optionsMinutes: IDropdownOption[] = [
  { key: '1', text: '1' },
  { key: '2', text: '2' },
  { key: '3', text: '3' },
  { key: '4', text: '4' },
  { key: '5', text: '5' },
  { key: '6', text: '6' },
  { key: '7', text: '7' },
  { key: '8', text: '8' },
  { key: '9', text: '9' },
  { key: '10', text: '10' },
  { key: '11', text: '11' },
  { key: '12', text: '12' },
  { key: '13', text: '13' },
  { key: '14', text: '14' },
  { key: '15', text: '15' },
  { key: '16', text: '16' },
  { key: '17', text: '17' },
  { key: '18', text: '18' },
  { key: '19', text: '19' },
  { key: '20', text: '20' },
  { key: '21', text: '21' },
  { key: '22', text: '22' },
  { key: '23', text: '23' },
  { key: '24', text: '24' },
  { key: '25', text: '25' },
  { key: '26', text: '26' },
  { key: '27', text: '27' },
  { key: '28', text: '28' },
  { key: '29', text: '29' },
  { key: '30', text: '30' },
  { key: '31', text: '31' },
  { key: '32', text: '32' },
  { key: '33', text: '33' },
  { key: '34', text: '34' },
  { key: '35', text: '35' },
  { key: '36', text: '36' },
  { key: '37', text: '37' },
  { key: '38', text: '38' },
  { key: '39', text: '39' },
  { key: '40', text: '40' },
  { key: '41', text: '41' },
  { key: '42', text: '42' },
  { key: '43', text: '43' },
  { key: '44', text: '44' },
  { key: '45', text: '45' },
  { key: '46', text: '46' },
  { key: '47', text: '47' },
  { key: '48', text: '48' },
  { key: '49', text: '49' },
  { key: '50', text: '50' },
  { key: '51', text: '51' },
  { key: '52', text: '52' },
  { key: '53', text: '53' },
  { key: '54', text: '54' },
  { key: '55', text: '55' },
  { key: '56', text: '56' },
  { key: '57', text: '57' },
  { key: '58', text: '58' },
  { key: '59', text: '59' },
  { key: '60', text: '60' }
];
const optionMonths: IDropdownOption[] = [
  
  { key: 'January', text: 'January' },
  { key: 'February', text: 'February' },
  { key: 'March', text: 'March' },
  { key: 'April', text: 'April' },
  { key: 'May', text: 'May' },
  { key: 'June', text: 'June' },
  { key: 'July', text: 'July' },
  { key: 'August', text: 'August' },
  { key: 'September', text: 'September' },
  { key: 'October', text: 'October' },
  { key: 'November', text: 'November' },
  { key: 'December', text: 'December' },
];
//const [rows, setRows] = React.useState([{ name: '', age: '' }]);
//const [file, setFile] = React.useState<File>();
export class FormComponent extends React.Component<IFormProps, IState> {
  private _sp: SPFI;

 // parseInt(((((it[0].OPE* it[0].GrossConveyorSpeed)/100)*parseFloat(it[0].Shiftwiseshopwiseworkinghours))-(currItmTot[0].Volume)));
					
 
  //private dept: IDropdownOption;
 
  constructor (props : any, state:  IState){
    super(props);
    this._sp = getSP();
   
   //this.dept= this.getDataDepartment() ;
    this.createDraft = this.createDraft.bind(this);
    this.submitDraft = this.submitDraft.bind(this);
    this.state={
      Breadcrumb: [],
      status:"Ready",
      items:[],
      rows:[{Section:"",Date:"",Auditor:""}],
      approvers:[{Role:"",Level:"",Name:"",Type:"One",Index:0, appEx:"", itemId:"",errLevel:"",errType:""}],
      isChecked:false,
      requestSign:false,
      forApproval:false,
      memoNumber: "",
       issueNumber: 200,
       revisionNumber:100,
       departmentVal:"",
  referenceNumber: "",
  subject:"",  
  background:"",
  issues: "",
  recommendations: [{section:"",date:null,approver:"",approverDef:"",hours:"",minutes:"" ,itemId:"",index:0,errSec:'',errHr:'',errMin:'',errTo:''}],
  recommendedApproval: "",
  optionsDepartment:[],
  optionsRole:[],
  options:[],
  errSub:"",
  errDep:true,
  errTo:"",
  errCC:"",
  errRecoApp:"",
  errback:"",
  errIss:"",
  department:0,
  currUserId:0,
  toArr:[],
  ccArr:[],
  dueDate:new Date(),
  filePickerResult:[],
  ind:0,
  showMonth:false,
  indApp:0,
  isInfo:true,
  type:0,
  year:0,
  month:"",
  itemId:0,
  showDialog:false,
//files:[],
  files: {} as FileList,
  copyFil:[],
      fileCount:0,
      toUsers:[],
      ccUsers:[],
      exFiles:[],
      recomDeleteId:[],
      fileDeleteId:[],
      apprDelId:[],
      mainListId:0,
      formNameId:0,
      apprItems:[],
      remarks:"",
      edItm:'',
      edType:'',
      showApprove:false,
      allDepartments:[],
      showSubmit:true,
      approvalItemId:'',
      isDisabled:false,
      memoSerialNo:0,
      reqRolId:0,
      isSubmit:false,
      isSave:false,
      isRework:false
    };
    this.subOnChange=this.subOnChange.bind(this);
    this.onChangeBack=this.onChangeBack.bind(this);
    this.onChangeIss=this.onChangeIss.bind(this);
    this.onChangeRecomApproval=this.onChangeRecomApproval.bind(this);
    this.onInfoChange=this.onInfoChange.bind(this);
    this.onSignChange=this.onSignChange.bind(this);
    this.onAppChange=this.onAppChange.bind(this);
    this.getDataDepartment=this.getDataDepartment.bind(this);
    this.onDepartmentChange = this.onDepartmentChange.bind(this);
    this.addRow= this.addRow.bind(this);
    this.addApprover = this.addApprover.bind(this);
    this._getPeoplePickerItemsTo = this._getPeoplePickerItemsTo.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._getPeoplePickerItemsAud=this._getPeoplePickerItemsAud.bind(this);
    this.onSectionChange= this.onSectionChange.bind(this);
    this.onHoursChange = this.onHoursChange.bind(this);
    this.onMinutesChange = this.onMinutesChange.bind(this);
    this.onDateChange= this.onDateChange.bind(this);
    this.deleteItem = this.deleteItem.bind(this);
    this.deleteItemApp= this.deleteItemApp.bind(this);
    this.onYearSelect = this.onYearSelect.bind(this);
    this.onRoleChange = this.onRoleChange.bind(this);
    this.onTypeChange = this.onTypeChange.bind(this);
    this.handleFileChange= this.handleFileChange.bind(this);
    this.getAuditTypes= this.getAuditTypes.bind(this);
    this.onYearChange= this.onYearChange.bind(this);
    this.onMonthChange = this.onMonthChange.bind(this);
    this.saveAsDraft = this.saveAsDraft.bind(this);
   this.getData = this.getData.bind(this);
 //  this.onIdChange = this.onIdChange.bind(this);
   this._OpenModal = this._OpenModal.bind(this);
   this._CloseModal = this._CloseModal.bind(this);
   this.updateDraft= this.updateDraft.bind(this);
   this.setDueDate = this.setDueDate.bind(this);
   this.removeFiles =this.removeFiles.bind(this);
   this.toBeDeleted = this.toBeDeleted.bind(this);
   this.updateRecommendations=  this.updateRecommendations.bind(this);
   this.getMainListName= this.getMainListName.bind(this);
   this.getFormName = this.getFormName.bind(this);
   this.getMemoNumber = this.getMemoNumber.bind(this);
   this.getRequestorRole = this.getRequestorRole.bind(this);
   this._getPeoplePickerItemsApp = this._getPeoplePickerItemsApp.bind(this);
   this.approveRequest = this.approveRequest.bind(this);
   this.rejectRequest= this.rejectRequest.bind(this);
   this.reworkRequest= this.reworkRequest.bind(this);
   this.cancelRequest = this.cancelRequest.bind(this);
   this.onRemarksChange = this.onRemarksChange.bind(this);
   this.cancelDraft= this.cancelDraft.bind(this);
   this.subOnClick= this.subOnClick.bind(this);
   this.onDepClick= this.onDepClick.bind(this);
   this.bacOnClick = this.bacOnClick.bind(this);
   this.issOnClick= this.issOnClick.bind(this);
   this.onRecClick=  this.onRecClick.bind(this);
  }
 


  async componentDidMount(){
    await this.getDataDepartment();
    await this.getDataRoles();
    await this.getAuditTypes();
    await this.getMainListName();
    await this.getFormName();
    await this.getMemoNumber();
    await this.getRequestorRole();

    // Extracting the part after `#/`
    const url=window.location.href;
    
    this.setState({
      Breadcrumb: [
        {
          MainComponent: "Home",
          MainComponentURl: `${url}`,
        },
        {
          ChildComponent: "Annual Audit Program",
          ChildComponentURl: `${url}`,
        }
      ],
    });
    
    const parts = url.split("#/")[1].split("/");

    const programName = decodeURIComponent(parts[0]); // "Annual Audit Program"
    const editType = parts[1]; // "edit"
    const id = parts[2]; // "165"

    console.log("Program Name:", programName);
    console.log("Edit Type:", editType);
    console.log("ID:", id);
    // alert(`${programName},${editType},${id}`);

    // Set state
    // this.setState((prevState) => ({
    //   edItm: id ? id : prevState.edItm, // Update only if `id` exists
    //   edType: editType ? editType : prevState.edType // Update only if `editType` exists
    // }));
    if(id){
      this.setState({edItm:id})
    }
    if(editType){
      this.setState({edType:editType})
    }
    
    // alert(`after setting edtype and edItem ${this.state.edItm},${this.state.edType}`);
    if(this.state.edType=="approve"){
    //  alert("approve");
    const approvalItemId= parts[3];
    if(approvalItemId){
      this.setState({approvalItemId:approvalItemId})
    }
    
      this.getData();
      this.setState({isDisabled:true})
      this.setState({showApprove:true});
      this.setState({showSubmit: false});
     
    }
    else if(this.state.edType =="edit"){
//alert("edit");
this.getData();
this.setState({isDisabled:false})
this.setState({showApprove:false});
this.setState({showSubmit: true});

    }
    else if(this.state.edType=="view"){
      this.getData();
      this.setState({isDisabled:true})
this.setState({showApprove:false});
this.setState({showSubmit: false});

    }
    else{
//alert("new");

    }
  }

  componentWillReceiveProps() {  

   }

   
  public render(): React.ReactElement<IFormProps> {   
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
  }; 

 
 
  var items=  this.state.recommendations.map((item: any,i:number) => {  
    var idtxt="txtRec_"+i;
    var idHr="txtHr_"+i;
    var idMin="txtMin_"+i
   // var idAss="ppAss_"+i;
 return(
    <tr className='tblCls'>
     <td>
     <TextField id={idtxt} className={this.state.recommendations[i].errSec} value={this.state.recommendations[i].section} onChange={(e) => this.onSectionChange(e,i)} disabled={this.state.isDisabled} ></TextField>
     </td>
     <td>
     <DateTimePicker  dateConvention={DateConvention.Date}  value={this.state.recommendations[i].date} disabled={this.state.isDisabled}
        showLabels={false}    formatDate={(date: Date) => date.toLocaleDateString()}
        onChange={(date: Date) =>this.setDueDate(date,i)}/>
          </td>
     <td>
     <Dropdown id={idHr} className={this.state.recommendations[i].errHr}  placeholder="Select" selectedKey={this.state.recommendations[i].hours}  options={optionsHours} styles={dropdownStylesNew} onChange={(e,itm:IDropdownOption) => this.onHoursChange(e,itm,i)} disabled={this.state.isDisabled}/>
    </td><td>
     <Dropdown id={idMin} className={this.state.recommendations[i].errMin}  placeholder="Select" selectedKey={this.state.recommendations[i].minutes}  options={optionsMinutes} styles={dropdownStylesNew} onChange={(e,itm:IDropdownOption) => this.onMinutesChange(e,itm,i)} disabled={this.state.isDisabled}/>
       
  
     </td>
     <td>
     <PeoplePicker peoplePickerWPclassName={this.state.recommendations[i].errTo} ensureUser={true}  disabled={this.state.isDisabled}  context={peoplePickerContext} personSelectionLimit={1} groupName={""} showtooltip={true}  searchTextLimit={5} onChange={(e) =>this._getPeoplePickerItemsAud(e,i)} 
         defaultSelectedUsers={this.state.recommendations[i].approverDef ? this.state.recommendations[i].approverDef : []}   principalTypes={[PrincipalType.User]} resolveDelay={1000} />
     </td>
    
     <td><button   onClick={(e) =>this.deleteItem(i)}>Delete</button></td>
     </tr>
    
  ) 
   }) ;
   var fileData = this.state.copyFil.map((item: any,i:number) => {
    return(
      <tr>
        <td>
        {item.name }  
        </td>
        <td>
          <button onClick={(e) =>this.removeFiles(i)}>Delete</button>
        </td>
      </tr> 
       )    
     
  });

  var upFiles = this.state.exFiles.map((item: any,i:number) => {
    return(
      <tr>
        <td>
        {item.Name }  
        </td>
        <td>
      {<a href={item.Path} target="_blank">Link</a>}
        </td>
        <td>
      {item.Uploaded}
        </td>
        <td>
          <button onClick={(e) =>this.toBeDeleted(i)} disabled={this.state.isDisabled}>Delete</button>
        </td>
      </tr> 
       )    
     
  });
   var approval = this.state.approvers.map((item: any,i:number) => {  
    return(
       <tr className='tblCls'>
        <td>
        <TextField value={(i+1).toString()} disabled={true} className={styles.width} ></TextField>
        </td>
        <td>
        <Dropdown className={this.state.approvers[i].errLevel} placeholder="Select" disabled={this.state.isDisabled} selectedKey={this.state.approvers[i].Role} options={this.state.optionsRole} styles={dropdownStyles} onChange={(e,itm:IDropdownOption) => this.onRoleChange(e,itm,i)}/>
        </td>
        <td>
        <TextField value={(i+1).toString()} disabled={true}></TextField>
        </td>
        <td>
        <PeoplePicker peoplePickerWPclassName={this.state.approvers[i].errType} ensureUser={true}  context={peoplePickerContext} personSelectionLimit={5} groupName={""} showtooltip={true}  disabled={this.state.isDisabled} searchTextLimit={5} onChange={(e) =>this._getPeoplePickerItemsApp(e,i)}
           defaultSelectedUsers={this.state.approvers[i].appEx ? this.state.approvers[i].appEx : []}  principalTypes={[PrincipalType.User]} resolveDelay={1000} />
        </td>
        <td>
        <Dropdown disabled={this.state.isDisabled} placeholder="Select" selectedKey={this.state.approvers[i].Type} options={optionsApp} styles={dropdownStylesNew} onChange={(e,itm:IDropdownOption) => this.onTypeChange(e,itm,i)}/>
     </td>
        <td><button   onClick={(e) =>this.deleteItemApp(i)}>Delete</button></td>
        </tr>
       
     ) 
      }) 
  
      var auditHistory = this.state.apprItems.map((item: any,i:number) => {
        return(
          <tr>
            <td>
            {i+1}
            </td>
            <td>
                {item.Level}                  
            </td>
            <td>
                {item.AssignedTo}
            </td>
            <td>
                {item.RequesterName}
            </td>
            <td>
            {new Date(item.RequestedDate).getDate()+"/" +new Date(item.RequestedDate).getMonth()+"/"+ new Date(item.RequestedDate).getFullYear()}
            </td>
            <td>
            {item.ActionTakenBy}
            </td>
            <td>
            {item.ActionTakenOn==''?'': new Date(item.ActionTakenOn).getDate()+"/" +new Date(item.ActionTakenOn).getMonth()+"/"+ new Date(item.ActionTakenOn).getFullYear()}                    
            </td>
            <td>
              {item.Remarks}
            </td>
            <td>
            {item.Status}
            </td>
          </tr> 
           )    
         
      }); 

    return (
    
      <section>
        <CustomBreadcrumb Breadcrumb={this.state.Breadcrumb} />
          <section className={styles.sec} >
            <fieldset disabled={this.state.isDisabled}>
            <h3>Memo Details</h3>
          <div className={styles['cl-12']}>
           <label>For Information *</label> 
           <Checkbox className={this.state.isInfo ? "":styles.errCls} checked={this.state.isChecked} label="For Information"  onChange={this.onInfoChange}  disabled={this.state.isDisabled}/>
             <Checkbox id="reqSign" checked={this.state.requestSign} className={styles['cl-3']}  label="Request for Signing" onChange={this.onSignChange} disabled={this.state.isDisabled}/>
           <Checkbox  className={styles['cl-3']}  checked={this.state.forApproval} label="For Approval" onChange={this.onAppChange} disabled={this.state.isDisabled}/>
          </div>
          <Dropdown  id="dept" placeholder="Select" label="Department" required={true} options={this.state.optionsDepartment} styles={dropdownStyles}  selectedKey={(this.state.department)}  onFocus={this.onDepClick} onChange={this.onDepartmentChange} disabled={this.state.isDisabled}/>
          <TextField label="Memo Number" className={this.state.errDep ? "":styles.errCls}  required={ true } value={(this.state.memoNumber)} name='memoNumber' disabled={true}/>
          <TextField label="From" required={ true } defaultValue={(this.props.userDisplayName)} name='lastname' disabled={true}/>
          <PeoplePicker  errorMessage={this.state.errTo} ensureUser={true}  context={peoplePickerContext}  titleText="To"  personSelectionLimit={3}   groupName={""}   showtooltip={true}  required={true} disabled={this.state.isDisabled}  searchTextLimit={5}
             onChange={this._getPeoplePickerItemsTo} defaultSelectedUsers={this.state.toUsers ? this.state.toUsers : []}  principalTypes={[PrincipalType.User]} resolveDelay={1000} />
          <PeoplePicker  errorMessage={this.state.errCC} ensureUser={true}  context={peoplePickerContext} titleText="CC"  personSelectionLimit={3} validateOnFocusOut={true} groupName={""} showtooltip={true} required={true} disabled={this.state.isDisabled} searchTextLimit={5} onChange={this._getPeoplePickerItems}
              principalTypes={[PrincipalType.User]} defaultSelectedUsers={this.state.ccUsers ? this.state.ccUsers : []} resolveDelay={1000} />
           <TextField label="Subject" id="sub" errorMessage={this.state.errSub} required={ true }  onNotifyValidationResult={this.subOnClick}  validateOnFocusOut={true} value={this.state.subject} onClick={this.subOnClick}   onChange ={this.subOnChange} name='lastname' disabled={this.state.isDisabled}/>
            <DateTimePicker label="Date" showLabels={false} dateConvention={DateConvention.Date}  value={this.state.dueDate}
               formatDate={(date: Date) => date.toLocaleDateString()}
        onChange={(date: Date) => this.setState({ dueDate: date })} disabled={this.state.isDisabled}/>
           <TextField id="bac" errorMessage={this.state.errback} multiline autoAdjustHeight value={this.state.background}  onNotifyValidationResult={this.bacOnClick}  validateOnFocusOut={true} onChange={this.onChangeBack} required={ true } label="Background"  disabled={this.state.isDisabled}/>
          <TextField id="iss" errorMessage={this.state.errIss} multiline autoAdjustHeight value={this.state.issues}  onNotifyValidationResult={this.issOnClick}  validateOnFocusOut={true} onChange={this.onChangeIss} required={ true } label="Issues" disabled={this.state.isDisabled} />
          <label>Recommendation</label>
          <button  onClick={this.addRow}>Add
</button>     
    <table id="tabRec">
         
          <tbody>  
            <tr><td>Section</td>
            <td>Date</td>
            <td colSpan={2}>Time</td>
            <td>Auditor</td>
            <td>Delete</td>
            </tr>     
        { items }
        </tbody>
         
        </table>

         
         
           <TextField  id="rec" errorMessage={this.state.errRecoApp} multiline autoAdjustHeight value={this.state.recommendedApproval} onNotifyValidationResult={this.onRecClick}  validateOnFocusOut={true}   onChange={this.onChangeRecomApproval} required={ true }  label="Recommendation for Approval" disabled={this.state.isDisabled}/>   
          
       
            </fieldset>
            </section>
         
         <section className={styles.sec}>
          <fieldset >
          <h3> Audit Programme Detail</h3>

<Dropdown id="drpType" placeholder="Select" label="Type *" onChange={this.onYearSelect} options={this.state.options}  selectedKey={(this.state.type)} styles={dropdownStyles} disabled={this.state.isDisabled}/>
<Dropdown id="drpYear"
placeholder="Select"
label="Year *" selectedKey={(this.state.year)} 
options={optionsYear}
styles={dropdownStyles} onChange={this.onYearChange} disabled={this.state.isDisabled}
/>
<Dropdown id="drpMonths"
placeholder="Select" selectedKey={(this.state.month)} 
label="Months *" disabled={this.state.showMonth || this.state.isDisabled}
options={optionMonths}
styles={dropdownStyles} onChange={this.onMonthChange} 
/>
<input className="form-control" type="file" disabled={this.state.isDisabled} name="myFile" onChange={(e)=>this.handleFileChange(e,this)}  id="newfile" multiple/>
      <label>FIles Selected:</label>  <button  onClick={this._OpenModal}>{this.state.fileCount}</button>
         <table>
          <tbody>
            {fileData}
          </tbody>
         </table>
          </fieldset>
         
         </section>
         <section className={styles.sec}>
         <fieldset disabled={this.state.isDisabled}>
         <h3> Approval Detail</h3>
         <button  onClick={this.addApprover}>Add
         </button>     
         <table id="tblAppr">
         
         <tbody>       
          <tr><td>Sl.</td>
          <td>Role</td>
          <td>Approver Level</td>
          <td>Approver Name</td>
          <td>Approval Type</td>
          <td>Delete</td>
          </tr>
       { approval }
       </tbody>
        
       </table>
        </fieldset>
         </section>
         

   {this.state.showSubmit &&  <section id="editDetails">
      <PrimaryButton className={(styles as any)['mar-10']} onClick={this.createDraft}>Save As Draft</PrimaryButton>
      <PrimaryButton className={(styles as any)['mar-10']} onClick={this.submitDraft}>Submit</PrimaryButton> 
      <PrimaryButton className={(styles as any)['mar-10']} onClick={this.cancelDraft}>Cancel</PrimaryButton> 
      </section>}
      {this.state.showApprove &&
      <section id="approvalSection">
           <Field label="Remarks">
            <Textarea id="comm" value={this.state.remarks} onChange={this.onRemarksChange} />
          </Field>
                     <PrimaryButton className={(styles as any)['mar-10']} onClick={this.approveRequest}>Approve</PrimaryButton>
                      <PrimaryButton  className={(styles as any)['mar-10']} onClick={this.rejectRequest}>Reject</PrimaryButton>
                       <PrimaryButton className={(styles as any)['mar-10']} onClick={this.reworkRequest}>Rework</PrimaryButton>
                       <a href='#/listing'> <PrimaryButton onClick={this.cancelRequest}>Cancel</PrimaryButton></a>
        
      </section>}
      <section id='audit'>
        
        <label>Audit Trial</label>
        <table>
          <thead>
            <tr>
            <th>Sl No</th>
            <th>Level</th>
              <th>Assigned To</th>
              <th>Requestor Name</th>
              <th>Requested Date</th>
              <th>Action Taken By</th>
              <th>Action Taken On</th>
              <th>Remarks</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
                {auditHistory}
          </tbody>
        </table>
      </section>
          {this.state.showDialog && <div id="myModal" className={styles.modal}>
           
      <div className={styles.modalcontent}>
      <span><b>Attachment Details</b></span>
            <br/>
            <span>Below are the attachment details for the Initiative</span>
          <span className={styles.close} onClick={e => this._CloseModal()}>&times;</span>
          <table>
            <thead>
            <th>File Name</th>
              <th>File Link</th>
              <th>Upload Date</th>
              <th>Delete</th>
            </thead>
         
            {upFiles}
          </table>
         
      </div>
   </div>}
      </section>
    );

    
  }
 /* private async getDepartment(): IDropdownOption{
    const spCache = spfi(this._sp).using(Caching({store:"session"}));

    const items = await spCache.web.lists.getByTitle("DepartmentMasterList").items;
    return items;
  }*/

    //save as draft
    private async createDraft(){ 
      this.clearAllValidations();
      this.setState({errDep:true});
      this.setState({isSubmit:false});
      this.setState({isSave:true}); 
      var _self= this;
      var isValid= true;
          if(_self.state.department ==0){
            document.getElementById('dept-option')?.classList.add(styles.errCls);
            isValid = false;
          }
          if(_self.state.isChecked ==false && _self.state.requestSign ==false && _self.state.forApproval){
            document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.add(styles.errCh);
            document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.add(styles.errCh);
            document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.add(styles.errCh);
            isValid = false;    
          }
          //if(_self.state.requestSign ==false){
         //   document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.add(styles.errCh);
         //   isValid = false;

       //   }
           
          if(isValid){
            Swal.fire({ title: 'Do you want to save this request?',  
              showCancelButton: true,  
              confirmButtonText: 'Yes',
              cancelButtonText:'No'
              }).then(function(val){
               if(val.isConfirmed){
                _self.saveAsDraft= _self.saveAsDraft.bind(_self);
            if(_self.state.itemId ==0){
          _self.saveAsDraft();
            }
            else{
              _self.updateDraft();
            }
                }          
              });           
          }
          else{
            Swal.fire({title:"Please fill the mandatory fields."});
          }
     
     
      } 

      private updateRecommendations(id:number, it:any){
        const spCache = spfi(this._sp).using(Caching({store:"session"}));
       console.log(spCache);
      }

  //update draft item
private async updateDraft(){
  var itemIdAudit = this.state.itemId;
  var _self= this;
  //const spCache = spfi(_self._sp).using(Caching({store:"session"}));
    const item = await spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.getById(this.state.itemId).update({
      Title: "Title",
      Info: this.state.isChecked,
      Sign:this.state.requestSign,
      Approval: this.state.forApproval,
      MemoNumber:this.state.memoNumber,
     DepartmentId:this.state.department,
     FromId:this.state.currUserId,
     ToId: this.state.toArr,
     CcId: this.state.ccArr,
      Date:this.state.dueDate,
      Subject:this.state.subject,
      Background: this.state.background,
      Issues:this.state.issues,
      RecommendedforApproval: this.state.recommendedApproval,
     AuditProgramTypeId: this.state.type,
      Year: this.state.year,
      MonthName: this.state.month,
        SubmitStatus:"No",
          Status:"Pending"
    });
    if(_self.state.recommendations.length>0){
      _self.state.recommendations.forEach(function(it){  
        var app;
        if(it.approver =='' || it.approver == null){
          app= null;
        } 
        else{
app= it.approver;
        }
        if(it.itemId ==""){
          spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.add({
            Title: "Title",
            Section:it.section,
            AnnualAuditProgramId:itemIdAudit,
            Date:it.date,
            Time:it.hours+":"+it.minutes,
            AuditorId:app
          });
        } 
        else{
          spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.getById(it.itemId).update({
            Title: "Title",
            Section:it.section,
            AnnualAuditProgramId:itemIdAudit,
            Date:it.date,
            Time:it.hours+":"+it.minutes,
            AuditorId:app
          }).catch(function(ex){
            console.log(ex.errorMessage);
          });
        
        }      
      
        });

    }
   
    if(_self.state.approvers.length>0){
      var maxLength = _self.state.approvers.length;
      _self.state.approvers.forEach(function(it,val){
        if(it.itemId ==0)    {
          spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
            Title: "Title",
            MainListNameId:_self.state.mainListId,
            ApproverRoleId:it.Role,
            Level:val+1,
            LevelType:it.Type,
            SubmitStatus:"No",
            Maxlevel:maxLength,
            ContentTitle:_self.state.subject,
            RequestId: _self.state.memoNumber,
            RequesterNameId:_self.state.currUserId,
            RequestedDate:new Date(),
            ProcessName:"Annual Audit Program",
           FormNameId: _self.state.formNameId,
           MainListID:_self.state.itemId,
           RequesterRoleId: _self.state.reqRolId,
           ApproversId: it.Name
          }).catch(function(ex){
            console.log(ex.errorMessage);
          })
        } 
        else{
          spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(it.itemId).update({
            Title: "Title",
            MainListNameId:_self.state.mainListId,
            ApproverRoleId:it.Role,
            Level:val+1,
            LevelType:it.Type,
            SubmitStatus:"No",
            Maxlevel:maxLength,
            ContentTitle:_self.state.subject,
            RequestId: _self.state.memoNumber,
            RequesterNameId:_self.state.currUserId,
            RequestedDate:new Date(),
            ProcessName:"Annual Audit Program",
           FormNameId: _self.state.formNameId,
           MainListID:_self.state.itemId,
           RequesterRoleId: _self.state.reqRolId,
           ApproversId: it.Name
          }).catch(function(ex){
            console.log(ex.errorMessage);
          })
        }     
        
        });
    }

    if(_self.state.recomDeleteId.length>0){
      _self.state.recomDeleteId.forEach(function(ids){
        spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.getById(ids).delete();
      });
     
    }
    if(_self.state.apprDelId.length>0){
      _self.state.apprDelId.forEach(function(ids){
        spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(ids).delete();
      });
     
    }
    
    if(_self.state.fileDeleteId.length>0){
      _self.state.fileDeleteId.forEach(function(ids){
        spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramDocs").items.getById(ids).delete();
      });
     
    }

    if(_self.state.copyFil.length>0){
      _self.state.copyFil.forEach(function(file){
        var fileNamePath = encodeURI(file.name);
        spfi(_self._sp).web.getFolderByServerRelativePath("AnnualAuditProgramDocs").files.addUsingPath(fileNamePath, file, { Overwrite: true }).then(function(response){
         response.file.getItem().then(function(fileItem){
          fileItem.update({
            AnnualAuditId: itemIdAudit
          });
         });
        });
        
      })
     
    }
   Swal.fire({title:"Saved Successfully.", icon:"success"});
   window.location.href="#"+'/listing';
   console.log(item);
}

//create draft Item
private async saveAsDraft(){
var _self= this;
    //const spCache = spfi(_self._sp).using(Caching({store:"session"}));
      const item = await spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.add({
        Title: "Title",
        Info: this.state.isChecked,
        Sign:this.state.requestSign,
        Approval: this.state.forApproval,
        MemoNumber:this.state.memoNumber,
       DepartmentId:this.state.department,
       FromId:this.state.currUserId,
       ToId: this.state.toArr,
       CcId: this.state.ccArr,
         Date:this.state.dueDate,
        Subject:this.state.subject,
        Background: this.state.background,
        Issues:this.state.issues,
        RecommendedforApproval: this.state.recommendedApproval,
       AuditProgramTypeId: this.state.type,
        Year: this.state.year,
        MonthName: this.state.month,
        SubmitStatus:"No",
          Status:"Pending"
      });
      if(_self.state.recommendations.length>0){
        _self.state.recommendations.forEach(function(it){ 
          var app;  
          if(it.approver == "" || it.approver == null){
            app=null;
          }   
          else app= it.approver;    
          spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.add({
              Title: "Title",
              Section:it.section,
              AnnualAuditProgramId:item.data.Id,
              Date:it.date,
              Time:it.hours+":"+it.minutes,
              AuditorId:app
            });
          });

      }
      if(_self.state.approvers.length>0){
        var maxLength = _self.state.approvers.length;
        _self.state.approvers.forEach(function(it,val){          
          spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
              Title: "Title",
              MainListNameId:_self.state.mainListId,
              ApproverRoleId:it.Role,
              Level:val+1,
              LevelType:it.Type,
              SubmitStatus:"No",
              Maxlevel:maxLength,
              ContentTitle:_self.state.subject,
             // RequestId: _self.state.memoNumber,
              RequesterNameId:_self.state.currUserId,
              MainListID:item.data.Id,
              ApprovalType:"Approval",
              RequestedDate:new Date(),
              ProcessName:"Annual Audit Program",
             FormNameId: _self.state.formNameId,
             ApproversId: it.Name,
             RequesterRoleId: _self.state.reqRolId
            }).catch(function(ex){
              console.log(ex.errorMessage);
            })
          });
      }
      if(_self.state.copyFil.length>0){
        _self.state.copyFil.forEach(function(file){
          var fileNamePath = encodeURI(file.name);
          spfi(_self._sp).web.getFolderByServerRelativePath("AnnualAuditProgramDocs").files.addUsingPath(fileNamePath, file, { Overwrite: true }).then(function(response){
           response.file.getItem().then(function(fileItem){
            fileItem.update({
              AnnualAuditId: item.data.Id
            });
           });
          });
          
        })
       
      }
     
     Swal.fire({title:"Saved Successfully.", icon:"success"});
     window.location.href="#"+'/listing';
     console.log(item);
 
}
private clearAllValidations(){
  document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.remove(styles.errCh);
      document.getElementById('dept-option')?.classList.remove(styles.errCls);
      document.getElementsByClassName('ms-BasePicker-text')[0]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-BasePicker-text')[1]?.classList.remove(styles.errCh);
      document.getElementById('sub')?.classList.remove(styles.errCls);
      document.getElementById('bac')?.classList.remove(styles.errCls);
      document.getElementById('iss')?.classList.remove(styles.errCls);
      document.getElementById('rec')?.classList.remove(styles.errCls);
      document.getElementById('drpType-option')?.classList.remove(styles.errCls);
      document.getElementById('drpYear-option')?.classList.remove(styles.errCls);
      document.getElementById('drpMonths-option')?.classList.remove(styles.errCls);
      this.state.approvers.forEach(function(itm){
        itm.errName ='';
        itm.errType='';
        itm.errLevel ='';
      });

      this.state.recommendations.forEach(function(itm){
        itm.errSec ='';
        itm.errHr ='';
        itm.errMin ='';
        itm.errTo ='';
      });
}

  private async submitAllDraft(){
         // var _self= this; 
         
     // const spCache = spfi(this._sp).using(Caching({store:"session"}));
  var _self= this;
  var depCode:any = this.state.allDepartments.filter(function(it){
    return it['Id'] == _self.state.department;
  })
//   const getUserId = async (loginName:any) => {
//     try {
//         const user = await spfi(this._sp).web.ensureUser(loginName);
//         alert(user.data.Id)
//         return user.data.Id;
//     } catch (error) {
//         console.error("Error getting user ID for:", loginName, error);
//         return null;
//     }
// };

// // Convert all `ToId` and `CcId` users to their respective IDs
// const convertUsersToIds = async (userArray:any) => {
//     const userIds = await Promise.all(userArray.map(getUserId));
//     return userIds.filter(id => id !== null); // Remove any null values in case of errors
// };

      if(this.state.itemId !=0){

        const item = await spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.getById(this.state.itemId).update({
          Title: "Title",
          Info: this.state.isChecked,
          Sign:this.state.requestSign,
          Approval: this.state.forApproval,
          MemoNumber: depCode[0].DepartmentCode+"/"+(new Date().getMonth()+1)+"/" +this.state.memoSerialNo,
         DepartmentId:this.state.department,
         FromId:this.state.currUserId,
         ToId: this.state.toArr,
         CcId: this.state.ccArr,
           Date:this.state.dueDate,
          Subject:this.state.subject,
          Background: this.state.background,
          Issues:this.state.issues,
          RecommendedforApproval: this.state.recommendedApproval,
         AuditProgramTypeId: this.state.type,
          Year: this.state.year,
          MonthName: this.state.month,
          MemoSerialNumber: this.state.memoSerialNo,
          SubmitStatus:"Yes",
          Status:"Pending",
          SubmiitedDate: new Date()

        });
        // const user = await spfi(_self._sp).web.ensureUser(_self.props.userid);
        if(this.state.recommendations.length>0){
          this.state.recommendations.forEach(function(it){   
            if(it.itemId ==""){
              spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.add({
                Title: "Title",
                Section:it.section,
                AnnualAuditProgramId:_self.state.itemId,
                Date:it.date,
                Time:it.hours+":"+it.minutes,
                AuditorId:it.approver
              });
            } 
            else{
              spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.getById(it.itemId).update({
                Title: "Title",
                Section:it.section,
                AnnualAuditProgramId:_self.state.itemId,
                Date:it.date,
                Time:it.hours+":"+it.minutes,
                AuditorId:it.approver
              });
            }      
          
            });
    
        }
        if(_self.state.approvers.length>0){
          var maxLength = _self.state.approvers.length;
          _self.state.approvers.forEach(function(it,val){
            if(it.itemId ==0)    {
              spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
                Title: "Title",
                MainListNameId:_self.state.mainListId,
                ApproverRoleId:it.Role,
                Level:val+1,
                LevelType:it.Type,
                SubmitStatus:"Yes",
                MainListID:_self.state.itemId,
                ApprovalType:"Approval",
                Maxlevel:maxLength,
                ContentTitle:_self.state.subject,
                RequestId: depCode[0].DepartmentCode+"/"+(new Date().getMonth()+1)+"/" +_self.state.memoSerialNo,
                RequesterNameId:_self.state.currUserId,
                RequestedDate:new Date(),
                ProcessName:"Annual Audit Program",
               FormNameId: _self.state.formNameId,
               RequesterRoleId: _self.state.reqRolId,
               
               ApproversId: it.Name
              }).catch(function(ex){
                console.log(ex.errorMessage);
              })
            } 
            else{
              spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(it.itemId).update({
                Title: "Title",
                MainListNameId:_self.state.mainListId,
                ApproverRoleId:it.Role,
                Level:val+1,
                LevelType:it.Type,
                SubmitStatus:"Yes",
                MainListID:_self.state.itemId,
                ApprovalType:"Approval",
                Maxlevel:maxLength,
                ContentTitle:_self.state.subject,
                RequestId: depCode[0].DepartmentCode+"/"+(new Date().getMonth()+1)+"/" +_self.state.memoSerialNo,
                RequesterNameId:_self.state.currUserId,
                RequestedDate:new Date(),
                ProcessName:"Annual Audit Program",
               FormNameId: _self.state.formNameId,
               RequesterRoleId: _self.state.reqRolId,
               ApproversId: it.Name
              }).catch(function(ex){
                console.log(ex.errorMessage);
              })
            } 
          })
        }
        if(_self.state.isRework){
          const userdata = await spfi(_self._sp).web.currentUser();
          spfi(_self._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(parseInt(_self.state.approvalItemId)).update({
            Status: "Approved",
            ActionTakenById:userdata.Id,
            ActionTakenOn:new Date()
          });

        }
        if(_self.state.recomDeleteId.length>0){
          _self.state.recomDeleteId.forEach(function(ids){
            spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.getById(ids).delete();
          });
         
        }

        if(_self.state.apprDelId.length>0){
          _self.state.apprDelId.forEach(function(ids){
            spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(ids).delete();
          });
         
        }
        if(_self.state.fileDeleteId.length>0){
          _self.state.fileDeleteId.forEach(function(ids){
            spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramDocs").items.getById(ids).delete();
          });
         
        }
        if(_self.state.copyFil.length>0){
          _self.state.copyFil.forEach(function(file){
            var fileNamePath = encodeURI(file.name);
            spfi(_self._sp).web.getFolderByServerRelativePath("AnnualAuditProgramDocs").files.addUsingPath(fileNamePath, file, { Overwrite: true }).then(function(response){
             response.file.getItem().then(function(fileItem){
              fileItem.update({
                AnnualAuditId: _self.state.itemId
              });
             });
            });
            
          })
         
        }
       Swal.fire({title:"Submitted Successfully.", icon:"success"});
       window.location.href="#"+'/listing';
       var tab= document.getElementsByClassName("tblCls");
       console.log(tab);
        //swal({title:"Request created successfully.",icon:"success"});
       console.log(item);
      }
      else{
        // const toUserIds = await convertUsersToIds(_self.state.toArr);
        // const ccUserIds = await convertUsersToIds(_self.state.ccArr);
        // const approvaluser=await convertUsersToIds()
        const item = await spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.add({
          Title: "Title",
        Info: _self.state.isChecked,
        Sign:_self.state.requestSign,
        Approval: _self.state.forApproval,
      // MemoNumber:_self.state.memoNumber,
       DepartmentId:_self.state.department,
       FromId:this.props.userid,
       ToId: _self.state.toArr,
       CcId: _self.state.ccArr,
         Date:_self.state.dueDate,
        Subject:_self.state.subject,
        Background: _self.state.background,
        Issues:_self.state.issues,
        RecommendedforApproval: _self.state.recommendedApproval,
       AuditProgramTypeId: _self.state.type,
        Year: _self.state.year,
        MonthName: _self.state.month,
        MemoNumber: depCode[0].DepartmentCode+"/"+(new Date().getMonth()+1)+"/" +this.state.memoSerialNo,
        MemoSerialNumber:this.state.memoSerialNo,
        SubmitStatus:"Yes",
          Status:"Pending",
          SubmiitedDate: new Date()
        });
      
      var itemId=item.data.Id;
        if(_self.state.recommendations.length>0){
          _self.state.recommendations.forEach(async function(it){   
            if(it.itemId ==""){
              spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.add({
                Title: "Title",
                Section:it.section,
                AnnualAuditProgramId:itemId,
                Date:it.date,
                Time:it.hours+":"+it.minutes,
                AuditorId:it.approver
              });
            } 
            else{
              spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.getById(it.itemId).update({
                Title: "Title",
                Section:it.section,
                AnnualAuditProgramId:itemId,
                Date:it.date,
                Time:it.hours+":"+it.minutes,
                AuditorId:it.approver
              });
            }      
          
            });
    
        }

        if(_self.state.approvers.length>0){
          var maxLength = _self.state.approvers.length;
          _self.state.approvers.forEach(async function(it,val){          
            spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.add({
                Title: "Title",
                MainListNameId:_self.state.mainListId,
                ApproverRoleId:it.Role,
                Level:val+1,
                LevelType:it.Type,
                SubmitStatus:"Yes",
                MainListID:itemId,
                ApprovalType:"Approval",
                Maxlevel:maxLength,
                ContentTitle:_self.state.subject,
                RequestId: _self.state.departmentVal+"/"+(new Date().getMonth()+1)+"/" +_self.state.memoSerialNo,
                RequesterNameId:_self.state.currUserId,
                RequestedDate:new Date(),
                ProcessName:"Annual Audit Program",
               FormNameId: _self.state.formNameId,
               RequesterRoleId: _self.state.reqRolId,
               ApproversId: it.Name
              }).catch(function(ex){
                console.log(ex.errorMessage);
              })
            });
        }
        if(this.state.copyFil.length>0){
          this.state.copyFil.forEach(function(file){
            var fileNamePath = encodeURI(file.name);
            spfi(_self._sp).web.getFolderByServerRelativePath("AnnualAuditProgramDocs").files.addUsingPath(fileNamePath, file, { Overwrite: true }).then(function(response){
             response.file.getItem().then(function(fileItem){
              fileItem.update({
                AnnualAuditId: itemId
              });
             });
            });
            
          })
         
        }
  
       Swal.fire({title:"Submitted Successfully.", icon:"success"});
       var tab= document.getElementsByClassName("tblCls");
       window.location.href="#"+'/listing';
       console.log(tab);
        //swal({title:"Request created successfully.",icon:"success"});
       console.log(item);
      }
    
      
     
  }

private cancelDraft(){
  window.location.href="#";
}
  
  private async submitDraft() {
    var _self= this;
    let  isValid = true;
    _self.clearAllValidations();
    _self.setState({isSubmit:true});
    _self.setState({isSave:false});
    if(_self.state.isChecked ==false && _self.state.requestSign ==false && _self.state.forApproval == false){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.add(styles.errCh);
      isValid = false;    
    }
    if(_self.state.department ==0){
      document.getElementById('dept-option')?.classList.add(styles.errCls);
      isValid = false;
    }
    if(_self.state.toArr.length ==0){
      document.getElementsByClassName('ms-BasePicker-text')[0]?.classList.add(styles.errCh);
      isValid = false;
    }
    if(_self.state.ccArr.length ==0){
      document.getElementsByClassName('ms-BasePicker-text')[1]?.classList.add(styles.errCh);
      isValid = false;
    }
    if(_self.state.subject=='' || _self.state.subject== null){
      document.getElementById('sub')?.classList.add(styles.errCls);
      isValid = false;
    }
    if(_self.state.background =='' || _self.state.background== null)
      {
        document.getElementById('bac')?.classList.add(styles.errCls);
        isValid = false;            
      }
    if(_self.state.issues =='' || _self.state.issues== null)
      {
        document.getElementById('iss')?.classList.add(styles.errCls);
        isValid = false; 
      }
      if(_self.state.type ==0 || _self.state.type== null){
        document.getElementById('drpType-option')?.classList.add(styles.errCls);
        isValid = false;
      }
      if(_self.state.year ==0 || _self.state.year== null){
        document.getElementById('drpYear-option')?.classList.add(styles.errCls);
        isValid = false;
      }
    if(_self.state.recommendedApproval =='' || _self.state.recommendedApproval== null)
      {
        document.getElementById('rec')?.classList.add(styles.errCls);
        isValid = false;
      }
      if(_self.state.showMonth == false && ( _self.state.month =='' || _self.state.month ==null)){
        document.getElementById('drpMonths-option')?.classList.add(styles.errCls);
        isValid = false;
      }
      _self.state.recommendations.forEach(function(itm){
        if(itm.section =="" || itm.date =="" || itm.hours =="" || itm.minutes =="" || itm.approver =='')
        {
          if(itm.section ==""){
            itm.errSec =styles.errCls;
          }
          if(itm.hours ==""){
            itm.errHr =styles.errCls;
          }
          if(itm.minutes ==""){
            itm.errMin =styles.errCls;
          }
          if(itm.approver ==""){
            itm.errTo =styles.errCls;
          }
        isValid = false;
        }   
        else{
          itm.errSec ='';
          itm.errHr ='';
          itm.errMin ='';
          itm.errTo ='';
        }     
      })
      _self.state.approvers.forEach(function(itm){
        
        if(itm.Role =="" || itm.Name =="" || itm.Type=="")
        {
        if(itm.Role ==""){
          itm.errLevel =styles.errCls;
        }
        if(itm.Name =='')
        {
          itm.errType= styles.errCls;
        }
        isValid = false;
        }
        else{
          itm.errLevel ='';
          itm.errName='';
          itm.errType ='';
        }        
      });
      if(_self.state.copyFil.length==0 && this.state.exFiles.length==0){
        isValid = false;
      }
      _self.setState({approvers:_self.state.approvers})
      if(isValid){
            var isRecMiss=false;
            var isAppMiss= false;
            if(_self.state.recommendations.length==1 )
              {
                if(_self.state.recommendations[0].section =="" ||_self.state.recommendations[0].approver =="" ){
                isRecMiss = true ;
                Swal.fire({title:"Please fill the mandatory fields."});
              }
             
            }
            if(_self.state.approvers.length ==1 && isRecMiss == false){
              if(_self.state.approvers[0].Role =="" ||_self.state.recommendations[0].Name =="" ){
                Swal.fire({title:"Please fill the mandatory fields."});
                isAppMiss = true;
              }
             
             
            }
            if(isRecMiss ==false && isAppMiss == false ){
              Swal.fire({ title: 'Do you want to submit this request?',  
                showCancelButton: true,  
                  confirmButtonText: 'Yes',
                  cancelButtonText:'No'  }).then(function(val){
                 if(val.isConfirmed){
                  _self.submitAllDraft=_self.submitAllDraft.bind(_self);
           _self.submitAllDraft();
                     
                 }
                });
          
            }
            }
            else{
              Swal.fire({title:"Please fill the mandatory fields."});
            }

 


   


    
   /* if(this.state.errTo.length ==0){
      this.setState({errTo:"This field is required"})
      isValid = false;
    }
    else{
      this.setState({errTo:""})
    }
    if(this.state.errCC.length ==0){
      this.setState({errCC:"This field is required"})
      
      isValid = false;
    }
    else{
      this.setState({errCC:""})
    }*/
   
  }  
  private _getPeoplePickerItemsTo(items: any[]) {
     var arr:any[];
     arr=[];
    items.forEach(function(it){
      arr.push(it.id);
    })
      this.setState({toArr:arr});
  }
  private _getPeoplePickerItemsAud(items:any[], i: number){
    this.state.recommendations.filter(function(it){
      if(it.index ==i){
        it.approver =items[0].id;
        it.errTo='';
      }
    });
    this.setState({recommendations: this.state.recommendations});
  }
  private _getPeoplePickerItemsApp(items:any[], i: number){
    var arr:any[];
     arr=[];
    items.forEach(function(it){
      arr.push(it.id);
    })
    this.state.approvers.filter(function(it){
      if(it.Index ==i){
        it.Name =arr;
        it.errType='';
      }
    });
    this.setState({approvers: this.state.approvers});
  }

  
  private onDateChange(e:any, i:number){
    this.state.recommendations.filter(function(it){
      if(it.index ==i){
        it.date =e.target.value
      }
    });
    this.setState({recommendations: this.state.recommendations});
  }
  private _getPeoplePickerItems(items: any[]) {
    var arr:any[];
    arr=[];
   items.forEach(function(it){
     arr.push(it.id);
   })
    this.setState({ccArr:arr});
  }
  private subOnChange(event:any)
  {
    this.setState({subject:event.target.value});
    
  }
  private subOnClick(){
    if(this.state.isSubmit == true && this.state.subject ==''){
      document.getElementById('sub')?.classList.add(styles.errCls);
    }
  }
  private bacOnClick(){
    if(this.state.isSubmit == true && this.state.background ==''){
      document.getElementById('bac')?.classList.add(styles.errCls);
    }
  }
  private issOnClick(){
    if(this.state.isSubmit == true && this.state.issues ==''){
      document.getElementById('iss')?.classList.add(styles.errCls);
    }
  }
  private onRecClick(){
    if(this.state.isSubmit == true && this.state.recommendedApproval ==''){
      document.getElementById('rec')?.classList.add(styles.errCls);
    }
  }
  private onDepClick(){
   // alert('Hi');
  }
  private onSectionChange(event:any, i:number){
    this.state.recommendations.filter(function(it){
      if(it.index ==i){
        it.section = event.target.value;
        it.errSec='';
      }
    });
    this.setState({recommendations: this.state.recommendations});
  }
  private onHoursChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i:number){
    this.state.recommendations.filter(function(it){
      if(it.index ==i){
        it.hours = item.text;
        it.errHr='';
      }
    });
    this.setState({recommendations: this.state.recommendations});
  }
  private onMinutesChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i:number){
    this.state.recommendations.filter(function(it){
      if(it.index ==i){
        it.minutes =item.text;
        it.errMin='';
      }
    });
    this.setState({recommendations: this.state.recommendations});
  }
  

  
  private onRoleChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i:number){
   /* this.state.approvers.filter(function(it){
      if(it.Index ==i){
        it.Role = item.text
      }
    });*/
    var isExists:Boolean=false;
    this.state.approvers.forEach(function(it){
      if(it.Role == item.key)
      {
        Swal.fire("","Role already selected!");
        isExists= true;
      }
    })
    if(!isExists){
      this.state.approvers[i].Role = item.key;
      this.state.approvers[i].errLevel = '';
      this.setState({approvers: this.state.approvers});
    }
   
  }
  private onTypeChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, i:number){
    /* this.state.approvers.filter(function(it){
       if(it.Index ==i){
         it.Role = item.text
       }
     });*/
     this.state.approvers[i].Type = item.text;
    
     this.setState({approvers: this.state.approvers});
   }
  

  private handleFileChange(e: React.ChangeEvent<HTMLInputElement>,_self:any){
    if (e.target.files) {
     // const files =e.target.files;
     _self.setState({fileCount:e.target.files.length});
     _self.setState({files:e.target.files});
      //console.log(_self.state.files.length);
      var allfiles:any[]= [];
    //  var onlyName:any[]=_self.state.copyFil;
      [].forEach.call(e.target.files,function(file:File){
       //onlyName.push({"Name":file.name,"type":"new","Id":""});
        
        allfiles.push(file);
      })
      _self.setState({copyFil: allfiles});
    //  _self.setState({exFiles:allfiles});
    }
  };
  private onChangeBack(event:any)
  {
    this.setState({background:event.target.value});    
  }
  private onChangeIss(event:any)
  {
    this.setState({issues:event.target.value});    
  }  
  private onChangeRecomApproval(event:any)
  {
    this.setState({recommendedApproval:event.target.value});    
  }
  private onInfoChange(event:any)
  {
    this.setState({isChecked:event.target.checked});    
    if(event.target.checked== true){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.remove(styles.errCh);
    }
 /*   else if(this.state.isSubmit == true && this.state.requestSign == false && this.state.forApproval == false){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.add(styles.errCh);
    }*/
  }
  private onSignChange(event:any)
  {
    this.setState({requestSign:event.target.checked}); 
    if(event.target.checked== true){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.remove(styles.errCh);
    }
   /* else if(this.state.isSubmit == true && this.state.isChecked == false && this.state.forApproval == false){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.add(styles.errCh);
    }  */ 
  }
  private onAppChange(event:any)
  {
    this.setState({forApproval:event.target.checked});
    if(event.target.checked== true){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.remove(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.remove(styles.errCh);
    }  
  /*  else if(this.state.isSubmit == true && this.state.isChecked == false && this.state.requestSign == false){
      document.getElementsByClassName('ms-Checkbox-checkbox')[0]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[1]?.classList.add(styles.errCh);
      document.getElementsByClassName('ms-Checkbox-checkbox')[2]?.classList.add(styles.errCh);
    }    */
  }
  private onDepartmentChange =(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) =>{
    var _self= this;
    this.setState({department:item.key as number});
    //this.setState({departmentVal:item.text});
   
   this.state.allDepartments.filter(function(itm){
    if(itm["ID"] == item.key)
    {
      _self.setState({departmentVal: itm["DepartmentCode"]})
    }
   })

  }
private  deleteItem(i:number){
  if(this.state.recommendations[i].itemId !=""){
    var delIDs=[];
    delIDs.push(this.state.recommendations[i].itemId )
    this.setState({recomDeleteId:delIDs});    
  }
  var items= this.state.recommendations.filter(function(it,val){
    return val != i
  });
  this.setState({rows: items});
  this.setState({recommendations: items});
}
private removeFiles(i:number){
var items= this.state.copyFil.filter(function(it,val){
  if(val != i){
    return it;
  }
})
this.setState({copyFil:items});
}


/*private async getVersionHistory(){
  const spCache = spfi(this._sp).using(Caching({store:"session"}));
//   const user = await spCache.web.ensureUser(this.props.userid);
const listItems = await spCache.web.lists.getByTitle("ProcessApprovalList").items.select('Id,RequestId,Title,ProcessName,ActionTakenBy/Title,ActionTakenOn,Status,RequestedDate,ListItemId').expand('ActionTakenBy').filter("ListItemId eq '"+this.state.listItemId+"' and Status ne 'Pending'")();
//const user = await spCache.web.ensureUser(listItems.RequesterNameId);
this.setState({items:listItems});
if(listItems.length==0){
  document.getElementById('audit')?.classList.add(styles.none);
}
}*/
private async approveRequest(){
  // const spCache = spfi(this._sp).using(Caching({store:"session"}));
  // const user = await spCache.web.ensureUser(this.props.userid);
  var _self= this;
  const userdata = await spfi(_self._sp).web.currentUser();

  Swal.fire({ title: 'Do you want to approve this request?',  
    showCancelButton: true,  
      confirmButtonText: 'Yes',
      cancelButtonText:'No'  }).then(function(val){
     if(val.isConfirmed){  
      // alert(`_self.state.approvalItemId ${_self.state.approvalItemId}`)  
      spfi(_self._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(parseInt(_self.state.approvalItemId)).update({
          Status: "Approved",
          ActionTakenById:userdata.Id,
          ActionTakenOn:new Date(),
          Remark:_self.state.remarks,
        });
        Swal.fire({title:"Approved Successfully.",icon:"success"});
       window.location.href="#"+'/listing';
     }
});
  
}
private async rejectRequest(){
  var _self= this;
    // const spCache = spfi(this._sp).using(Caching({store:"session"}));
    // const user = await spCache.web.ensureUser(this.props.userid);
    const userdata = await spfi(_self._sp).web.currentUser();
    Swal.fire({ title: 'Do you want to reject this request?',  
      showCancelButton: true,  
        confirmButtonText: 'Yes',
        cancelButtonText:'No'  }).then(function(val){
       if(val.isConfirmed){   
        spfi(_self._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(parseInt(_self.state.approvalItemId)).update({
          Status: "Rejected",
            ActionTakenById:userdata.Id,
            ActionTakenOn:new Date(),
            Remark:_self.state.remarks,
          });
           Swal.fire({title:"Rejected Successfully.",icon:"success"});
           window.location.href="#"+'/listing';
           //navigate('/listing');
       }
        });
  
}
private async reworkRequest(){
  var _self= this;
  // const spCache = spfi(_self._sp).using(Caching({store:"session"}));
  // const user = await spCache.web.ensureUser(_self.props.userid);
  const userdata = await spfi(_self._sp).web.currentUser();
  Swal.fire({ title: 'Do you want to rework this request?',  
    showCancelButton: true,  
      confirmButtonText: 'Yes',
      cancelButtonText:'No'  }).then(function(val){
     if(val.isConfirmed){   
      spfi(_self._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(parseInt(_self.state.approvalItemId)).update({
        Status: "Rework",
          ActionTakenById:userdata.Id,
          ActionTakenOn:new Date(),
          Remark:_self.state.remarks,
          IsRework:"Yes"
        });
        spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.getById(parseInt(_self.state.edItm)).update({
          Status: "Rework",
            IsRework:"Yes"
          });
         Swal.fire({title:"Sent for rework.",icon:"success"});
         window.location.href="#"+'/listing';
        // navigate('/listing');
     }
    });
 
}
private onRemarksChange(event:any){
  this.setState({remarks:event.target.value});
}
private cancelRequest(){
//  history.go(1);

}
private toBeDeleted( i:number){
  var itemId =this.state.exFiles.filter(function(it,val){
    return val == i
  })

var remFiles= this.state.exFiles.filter(function(it,val){
  return val != i
})
  //var exFiles:any[]= 
  this.state.fileDeleteId.push(itemId[0].Id)
  this.setState({fileDeleteId:this.state.fileDeleteId});
  this.setState({exFiles:remFiles});
}
private deleteItemApp(i:number){
  var items= this.state.approvers.filter(function(it,val){
    return val != i
  });
  
  this.setState({approvers: items});

  if(this.state.approvers[i].itemId !=""){
    var delIDs=[];
    delIDs.push(this.state.approvers[i].itemId )
    this.setState({apprDelId:delIDs});    
  }
  var items= this.state.approvers.filter(function(it,val){
    return val != i
  });
  this.setState({approvers: items});

}

  private addRow(){
var num = this.state.recommendations.length;
this.state.recommendations.push({section:"",date:null,approver:"",approverDef:"",hours:"",minutes:"" ,itemId:"",index:num,errSec:'',errHr:'',errMin:'',errTo:''});
this.setState({recommendations: this.state.recommendations});       
  }
  
private addApprover(){
  var itm=this.state.indApp+1;
  this.setState({indApp:itm});
  this.state.approvers.push({Role:"",Level:"",Name:"", Index:itm,itemId:"",Type:"One",errLevel:'',errType:''});
  this.setState({approvers:this.state.approvers});
 

  }
private onYearSelect=(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) =>{
  this.setState({type:item.key as number});
  if(item.text =="Annual"){
    this.setState({showMonth:false});
    
  }
  else{
    this.setState({showMonth:true});
  }
}

private async getData()
{
  var _self= this;
  this.setState({itemId:parseInt(this.state.edItm)});
 // const spCache = spfi(this._sp).using(Caching({store:"session"}));
 console.log(this.state.itemId , typeof(this.state.itemId) , "items id");
  const listItems = await  spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.getById(this.state.itemId).select("*", "To/Name","Cc/Name,Department/DepartmentCode").expand("To","Department","Cc")();
    this.setState({department:listItems.DepartmentId});
    if(listItems.DepartmentId !="")
    this.setState({departmentVal:listItems.Department.DepartmentCode});
    this.setState({memoNumber:listItems.MemoNumber});
    this.setState({subject:listItems.Subject});
    //this.setState({dueDate:listItems.Date});
    this.setState({background:listItems.Background});
    this.setState({issues:listItems.Issues});
    this.setState({recommendedApproval:listItems.RecommendedforApproval});
    this.setState({type:listItems.AuditProgramTypeId}); 
    this.setState({year:listItems.Year});
    this.setState({month:listItems.MonthName});
    this.setState({isChecked: listItems.Info});
    this.setState({requestSign:listItems.Sign});
    this.setState({forApproval:listItems.Approval});
    if(listItems.Status=="Rework"){
      this.setState({isRework:true});
    }
    if(listItems.SubmitStatus =="Yes" && listItems.Status=="Pending"){
      this.setState({isDisabled:true})
    
//this.setState({showApprove:true});
this.setState({showSubmit: false});
    }
else if(listItems.SubmitStatus =="Yes" && listItems.Status=="Rework"){
  this.setState({isDisabled:false})
this.setState({showApprove:false});
this.setState({showSubmit: true});

}
else if(listItems.SubmitStatus =="Yes" && (listItems.Status=="Approved" || listItems.Status=="Rejected")){
  this.setState({isDisabled:true})
this.setState({showApprove:false});
this.setState({showSubmit: false});

}
    let usernamearr: string[] = [];
   // let usernameToIds : any[]=[];
    if(listItems.ToId !=null){
    listItems.To.forEach(function(user:any){
     // usernameToIds.push(user.Id);
        usernamearr.push(user.Name.split('|membership|')[1].toString());
    });
    this.setState({toUsers:usernamearr});
    this.setState({toArr:listItems.ToId});
  }
    let userNameCc: string[] = [];
    //let userNameCcIds: any[] =[];
    if(listItems.CcId != null){
    listItems.Cc.forEach(function(user:any){
    //  userNameCcIds.push(user.Id);
      userNameCc.push(user.Name.split('|membership|')[1].toString());
    });
    this.setState({ccArr:listItems.CcId});
    this.setState({ccUsers:userNameCc});
  }
    this.setState({dueDate:new Date(listItems.Date)});
    const recItems = await  spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.select("*","Auditor/Name").expand("Auditor").filter("AnnualAuditProgramId eq '"+this.state.itemId+"'")();
     var cnt:any=0;
     var recoItms:any[]=[];
     if(recItems.length>0){
    recItems.forEach(function(itm:any){
      var objToAdd:any={};
      let approve: string[] = [];
      if(itm.AuditorId !=null)
      approve.push(itm.Auditor.Name.split('|membership|')[1].toString());
      objToAdd["section"] =itm.Section;
      objToAdd["date"] =new Date(itm.Date);
      objToAdd["approverDef"] =approve;
      objToAdd["approver"]= itm.AuditorId;
      objToAdd["hours"] =itm.Time.split(":")[0];
      objToAdd["minutes"] =itm.Time.split(":")[1];
      objToAdd["itemId"] =itm.Id;
      objToAdd["index"] =cnt;
      objToAdd["errSec"] ='';
      objToAdd["errHr"] ='';
      objToAdd["errMin"] ='';
      objToAdd["errTo"] ='';
  
      recoItms.push(objToAdd);
        cnt= cnt+1;
      });
      this.setState({recommendations:recoItms});
    }
    const apprItems = await  spfi(_self._sp).web.lists.getByTitle("ProcessApprovalList").items.select("*","AssignedTo/Title,RequesterName/Title,ActionTakenBy/Title,ListName/Title").expand("AssignedTo,RequesterName,ActionTakenBy,ListName").filter("ListItemId eq '"+this.state.itemId+"' and ProcessName eq 'Annual Audit Program'").orderBy("Id", false)();
    var cnt:any=0;
    var appItems:any[]=[];
   // var allApprovalItems:any[]=[];
    if(apprItems.length>0){
      apprItems.forEach(async function(itm:any){
     var objToAdd:any={};    
     objToAdd["Level"] =itm.Level;
     objToAdd["AssignedTo"] =itm.AssignedTo.Title;
     objToAdd["RequesterName"] =itm.RequesterName.Title;
     if(itm.RequestedDate  =='' || itm.RequestedDate ==null ){
      objToAdd["RequestedDate"] ='' ;
     }
     else{
      objToAdd["RequestedDate"] =itm.RequestedDate ;
     }
    
     if(itm.ActionTakenById != null){
      objToAdd["ActionTakenBy"]= itm.ActionTakenBy.Title;
     }
     else{
      objToAdd["ActionTakenBy"]= "";
     }
    if(itm.ActionTakenOn =='' || itm.ActionTakenOn == null){
      objToAdd["ActionTakenOn"]= '';
    }
    else{
      objToAdd["ActionTakenOn"]= itm.ActionTakenOn;
    }
    
     objToAdd["Remarks"]= itm.Remark;
     objToAdd["Status"]= itm.Status;
     objToAdd["Index"] =itm.Level;
    
    
     appItems.push(objToAdd);
       cnt= cnt+1;
       //const user = await spCache.web.ensureUser(listItems.RequesterNameId);
       _self.setState({apprItems:appItems});
      
     });
  
   }
   const approvalItems = await  spfi(_self._sp).web.lists.getByTitle("AllProcessApprovalLevelList").items.select("*","Approvers/Name").expand("Approvers").filter("MainListID eq '"+this.state.itemId+"' and ProcessName eq 'Annual Audit Program'")();
   var allApp:any[]=[];
   var cnt:any=0;
   if(approvalItems.length>0){
    approvalItems.forEach(async function(itm:any){
var objToAdd:any={};
let approve: string[] = [];
itm.Approvers.forEach(function(it:any){
  approve.push(it.Name.split('|membership|')[1].toString());
 // ids.push(it.Id);
 }) 
objToAdd["Role"] =itm.ApproverRoleId;
objToAdd["Type"] =itm.LevelType;
objToAdd["Name"] =itm.ApproversId ;
objToAdd["itemId"]= itm.Id;
objToAdd["Index"] =itm.Level;
objToAdd["appEx"] =approve;
objToAdd["errLevel"]='';
objToAdd["errType"]='';


allApp.push(objToAdd);
  cnt= cnt+1;
    });
    _self.setState({approvers:allApp});
   } 
  
 
      const upFiles= await  spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgramDocs").items.select("*","File/Name,FileLeafRef,FileRef,EncodedAbsUrl").expand("File").filter("AnnualAuditId eq '"+_self.state.itemId+"'")();
      if(upFiles.length>0){
        var obJFiles:any[]=[];
        upFiles.forEach(function(item:any){
            obJFiles.push({"Name":item.File.Name,"type":"old","Id":item.Id,"Uploaded":new Date(item.Modified).getDate()+"/"+new Date(item.Modified).getMonth()+"/"+new Date(item.Modified).getFullYear(),"Path":item.EncodedAbsUrl})
        })
       // this.setState({fileCount:obJFiles.length});
        this.setState({exFiles:obJFiles});
       // console.log(this.state.exFiles);
      }

    } 
private onYearChange=(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) =>{
  this.setState({year:item.key as number});
}

private onMonthChange=(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) =>{
  this.setState({month:item.text});
}
private async getMainListName(){
  var _self=this;
  const spCache = spfi(_self._sp).using(Caching({store:"session"}));
  const listItems = await spCache.web.lists.getByTitle("ListNameMaster").items.filter("ListName eq 'AnnualAuditProgram'")();
  this.setState({mainListId: listItems[0].Id})

}



private async getFormName(){
  var  _self=this;
  const spCache = spfi(_self._sp).using(Caching({store:"session"}));
  const listItems = await spCache.web.lists.getByTitle("FormNameMaster").items.filter("FormName eq 'Annual Audit Programme'")();
  this.setState({formNameId: listItems[0].Id})

}

private async getMemoNumber(){
  var  _self=this;
  const listItems = await  spfi(_self._sp).web.lists.getByTitle("AnnualAuditProgram").items.orderBy("MemoSerialNumber",false).top(1)();
 if(listItems.length>0){
  var memo=listItems[0].MemoSerialNumber+1;
  if(memo<999)
  memo=("0000"+memo).slice(-3); 
  this.setState({memoSerialNo: memo})
 }
 else{
  var mem:any= "000";
  this.setState({memoSerialNo: mem})
 }
 
}
private async getRequestorRole(){
  var _self=this;
  const listItems = await  spfi(_self._sp).web.lists.getByTitle("RequesterRoleMaster").items.filter("Role eq 'Initiator'")();

  this.setState({reqRolId: listItems[0].Id})
}

 private async getDataDepartment(){
    // const spCache = spfi(this._sp).using(Caching({store:"session"}));
    // const user = await spCache.web.ensureUser(this.props.userid);
    //console.log(userId);
  //  this.setState({currUserId:this.props.userid});
  try {
    var _self=this;
    const listItems = await spfi(_self._sp).web.lists.getByTitle("DepartmentMasterList").items.filter("Active eq 'Yes'")();
    
    let dropdownItems: IDropdownOption[] =[];
    let allItems:any=[];
    listItems.map(item =>{
             dropdownItems.push({
                key: item.Id,
                text: item.Title
             })
             allItems.push(item);
     });
     this.setState({allDepartments:allItems});
     
    this.setState({optionsDepartment: dropdownItems});
  } catch (error) {
    console.log("Error getting deoartment data",error)
  }
    
    }
   
    private setDueDate(date:Date,i:number){
      this.state.recommendations.filter(function(it){
        if(it.index ==i){
          it.date =date
        }
      });
      this.setState({recommendations: this.state.recommendations});
    }

    private async getDataRoles(){
      var _self=this;
      const spCache = spfi(_self._sp).using(Caching({store:"session"}));
      const listItems = await spCache.web.lists.getByTitle("ApproverRoleMaster").items();
      
      let dropdownItems: IDropdownOption[] =[];
      listItems.map(item =>{
               dropdownItems.push({
                  key: item.Id,
                  text: item.Role
               })
       });
      this.setState({optionsRole: dropdownItems});
      }
      private async getAuditTypes(){
        var  _self=this;
        const spCache = spfi(_self._sp).using(Caching({store:"session"}));
        const listItems = await spCache.web.lists.getByTitle("AuditProgramTypeMaster").items();
        
        let dropdownItems: IDropdownOption[] =[];
        listItems.map(item =>{
                 dropdownItems.push({
                    key: item.Id,
                    text: item.Title
                 })
         });
        this.setState({options: dropdownItems});
        }
      
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

}

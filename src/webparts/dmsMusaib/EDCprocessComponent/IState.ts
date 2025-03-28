import { IListItem } from './IListItem';  
import {  IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
export interface IState { 
  From:any; 
  Breadcrumb:any[];
  status: string;  
  items: IListItem[];  
  rows:any[];
  approvers: any[];
  isChecked: boolean;
  requestSign: boolean,
  forApproval:boolean,
  memoNumber: string,
  issueNumber: number,
  revisionNumber:number,
  referenceNumber: string,
  departmentVal:string,
  subject:string,
  dueDate:Date,
  background:string,
  issues: string,
  recommendations: any[],
  recommendedApproval: string,
  optionsDepartment:IDropdownOption[],
  options:IDropdownOption[],
  optionsRole:IDropdownOption[],
  errSub:any,
  errDep:boolean,
  errTo:string,
  errCC:string,
  errback:string,
  errIss:string,
  errRecoApp:string,
  department:number,
  toArr:any[],
  ccArr:any [],
  currUserId:number,
  filePickerResult:IFilePickerResult[],
  ind:number,
  showMonth:boolean,
  indApp:number,
  isInfo:boolean,
  type:number,
  year:number,
  month:string,
  itemId:number,
  fileCount:number,
  showDialog:boolean,
  files:FileList,
  copyFil:any[],
  toUsers:any[],
  ccUsers:any[],
  exFiles:any[],
  recomDeleteId: any[],
  fileDeleteId:any[],
  mainListId: number,
  formNameId: number,
  apprDelId:any[],
  apprItems:any[],
  remarks:string,
   edItm:any,
   edType:any,
  showApprove:boolean,
  showSubmit:boolean,
   approvalItemId: any,
  isDisabled: boolean,
  memoSerialNo: number,
  reqRolId: number,
  isSubmit: boolean,
  allDepartments: any[],
 
  isSave: boolean,
  isRework: boolean,

}  




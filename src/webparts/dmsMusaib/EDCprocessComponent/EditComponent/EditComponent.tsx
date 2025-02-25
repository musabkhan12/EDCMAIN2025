import * as React from 'react';
import { IEditProps } from './IEditProps';
import { IEditState } from './IEditState';
import { getSP } from "../../loc/pnpjsConfig";
import { SPFI , spfi} from "@pnp/sp";
// import { getSP } from "../PNPJsConfig";
import { Caching } from "@pnp/queryable";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
// import { SPFI, spfi } from "@pnp/sp";
import { PrimaryButton ,TextField} from '@fluentui/react';
import { Field, Textarea } from "@fluentui/react-components";
import Swal from 'sweetalert2';
import styles from '../AuditApp.module.scss';
//import { useHistory } from 'react-router-dom';
export class EditComponent extends React.Component<IEditProps,IEditState> {
        _sp: SPFI;
    constructor (props : any, state:  IEditState){
        super(props);
         this._sp = getSP();
       this.state={
        reqId:'',
        title: '',
  processName: '',
  requestedBy:'',
  requestedDate: '',
  status: '',
  remarks:'',
        // itemId:props["match"]["params"]["postId"],
        items:[],
        listItemId:0
       }
       //var itemId=this.props["match"]["params"]["id"];
       this.getAllItems=this.getAllItems.bind(this);
       this.approveRequest= this.approveRequest.bind(this);
       this.rejectRequest = this.rejectRequest.bind(this);
       this.reworkRequest= this.reworkRequest.bind(this);
       this.cancelRequest=this.cancelRequest.bind(this);
       this.getVersionHistory = this.getVersionHistory.bind(this);
       this.onRemarksChange= this.onRemarksChange.bind(this);
      }
 
      async componentDidMount(){
        await this.getAllItems();
        await this.getVersionHistory();
      }
    public render(): React.ReactElement<IEditProps> {   
    //  const history = useHistory();
     //   history.push('/listing');
    
      var allItems = this.state.items.map((item: any,i:number) => {
        return(
          <tr>
            <td style={{minWidth:'60px',maxWidth:'60px'}}>
            {i+1}
            </td>
            <td style={{minWidth:'80px',maxWidth:'80px'}}>
                {item.RequestId}                  
            </td>
            <td>
                {item.Title}
            </td>
            <td style={{minWidth:'80px',maxWidth:'80px'}}>
                {item.ProcessName}
            </td>
            <td style={{minWidth:'100px',maxWidth:'100px'}}>
            {item.ActionTakenBy.Title}
            </td>
            <td style={{minWidth:'100px',maxWidth:'100px'}}>
            { new Date(item.ActionTakenOn).getDate()+"/" +new Date(item.ActionTakenOn).getMonth()+"/"+ new Date(item.ActionTakenOn).getFullYear()}                    
            </td>
            <td style={{minWidth:'60px',maxWidth:'60px'}}>
            {item.Status}
            </td>
          </tr> 
           )    
         
      });

    return (
        <section >
          <div className='card card-body'>
          <div className='row'>
            <div className='col-sm-4 mb-3'>
            <TextField label="Request Id"   id="sub"  value={this.state.reqId} disabled={true} />

            </div>
            <div className='col-sm-4 mb-3'>
            <TextField label="Title"   id="title"  value={this.state.title}  disabled={true}/>

              </div>

              <div className='col-sm-4 mb-3'>
              <TextField label="Process Name"   id="process"  value={this.state.processName} disabled={true}/>

</div>
<div className='col-sm-4 mb-3'>
<TextField label="Requested By"   id="reqBy"  value={this.state.requestedBy} disabled={true}/>

</div>
<div className='col-sm-4 mb-3'>
<TextField label="Requested Date"   id="reqDate"  value={this.state.requestedDate}  disabled={true}/>
</div>
<div className='col-sm-4 mb-3'>
<TextField label="Status"   id="Status"  value={this.state.status} disabled={true}/>
</div>

<div className='col-sm-12 mb-3'>
<Field label="Remarks">
    <Textarea id="comm" value={this.state.remarks} onChange={this.onRemarksChange} />
  </Field>

</div>


          </div>
             
             
             
            
            
              
              <div style={{display:'flex',justifyContent:'center', gap:'10px'}}>
             <PrimaryButton className='btn btn-success' onClick={this.approveRequest}>Approve</PrimaryButton>
              <PrimaryButton className='btn btn-danger' onClick={this.rejectRequest}>Reject</PrimaryButton>
               <PrimaryButton className='btn btn-warning' onClick={this.reworkRequest}>Rework</PrimaryButton>
               <a href='#/listing'> <PrimaryButton className='btn cancel-btn' onClick={this.cancelRequest}>Cancel</PrimaryButton></a>
               </div>
               </div>

                <section style={{display:'grid'}} id='audit' className='mt-2'>
                  <div className='card card-body'>
                  <h3 style={{margin:'inherit'}} className='text-dark font-16 mb-3'>Audit Trial</h3>
                  <table className='mtbalenew'>
                    <thead>
                      <tr>
                      <th style={{minWidth:'60px',maxWidth:'60px'}}>S.No</th>
                        <th style={{minWidth:'60px',maxWidth:'60px'}}>Request Id</th>
                        <th>Title</th>
                        <th style={{minWidth:'80px',maxWidth:'80px'}}>Process Name</th>
                        <th style={{minWidth:'100px',maxWidth:'100px'}}>Action Taken By</th>
                        <th style={{minWidth:'100px',maxWidth:'100px'}}>Action Taken On</th>
                        <th style={{minWidth:'60px',maxWidth:'60px'}}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
{allItems}
                    </tbody>
                  </table>
                  </div>
                </section>
        </section>
    )};
 private async getAllItems(){

   // alert(this.state.itemId);
  //  const spCache = spfi(this._sp).using(Caching({store:"session"}));
 //   const user = await spCache.web.ensureUser(this.props.userid);            select('Id,RequesterNameId,RequestId,Title,ProcessName,RequesterName/Title,Status,AssignedToId,RequestedDate').expand('RequesterName')
    const listItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('Id,ListItemId,RequesterNameId,RequestId,Title,ProcessName,RequesterName/Title,Status,AssignedToId,RequestedDate').expand('RequesterName').filter("Id eq '"+277+"'")();
    //const user = await spCache.web.ensureUser(listItems.RequesterNameId);
    this.setState({reqId: listItems[0].RequestId});
   this.setState({title: listItems[0].Title});
   this.setState({processName: listItems[0].ProcessName}); 
   var req='';
   if(listItems[0].RequestedDate !=''){
req = new Date(listItems[0].RequestedDate).getDate()+"/"+new Date(listItems[0].RequestedDate).getMonth()+"/"+new Date(listItems[0].RequestedDate).getFullYear();
   }  
   this.setState({requestedDate: req});
   this.setState({status: listItems[0].Status});
   if(listItems[0].RequesterNameId != ''){
    this.setState({requestedBy: listItems[0].RequesterName.Title});
   }
 
   this.setState({listItemId: listItems[0].ListItemId})
    }

    private async getVersionHistory(){
      // const spCache = spfi(this._sp).using(Caching({store:"session"}));
 //   const user = await spCache.web.ensureUser(this.props.userid);
    const listItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('Id,RequestId,Title,ProcessName,ActionTakenBy/Title,ActionTakenOn,Status,RequestedDate,ListItemId').expand('ActionTakenBy').filter("ListItemId eq '"+277+"' and Status ne 'Pending'")();
    //const user = await spCache.web.ensureUser(listItems.RequesterNameId);
    this.setState({items:listItems});
    if(listItems.length==0){
      document.getElementById('audit')?.classList.add(styles.none);
    }
    }
    private async approveRequest(){
        // const spCache = spfi(this._sp).using(Caching({store:"session"}));
        const user = await spfi(this._sp).web.ensureUser(this.props.userid);
        spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(277).update({
            Status: "Approved",
            ActionTakenById:user.data.Id,
            ActionTakenOn:new Date(),
            Remark:this.state.remarks,
          });
          Swal.fire({title:"Approved succesfully",icon:"success"});
    }
    private async rejectRequest(){
        // const spCache = spfi(this._sp).using(Caching({store:"session"}));
        const user = await spfi(this._sp).web.ensureUser(this.props.userid);
        spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(277).update({
          Status: "Rejected",
            ActionTakenById:user.data.Id,
            ActionTakenOn:new Date(),
            Remark:this.state.remarks,
          });
           Swal.fire({title:"Rejected successfully",icon:"success"});
    }
    private async reworkRequest(){
      // const spCache = spfi(this._sp).using(Caching({store:"session"}));
      const user = await spfi(this._sp).web.ensureUser(this.props.userid);
      spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.getById(277).update({
        Status: "Rework",
          ActionTakenById:user.data.Id,
          ActionTakenOn:new Date(),
          Remark:this.state.remarks,
        });
         Swal.fire({title:"Rework successfull",icon:"success"});
    }
    private onRemarksChange(event:any){
      this.setState({remarks:event.target.value});
    }
    private cancelRequest(){
    //  history.go(1);
    
    }
    
}
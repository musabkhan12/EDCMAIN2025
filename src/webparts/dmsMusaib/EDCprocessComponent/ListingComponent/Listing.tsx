import * as React from 'react';
import { IListingProps } from './IListingProps';
import { IListingState } from './IListingState';
import { Profiles } from "@pnp/sp/profiles";
// import { getSP } from "../PNPJsConfig";
import { getSP } from "../../loc/pnpjsConfig";
//import { Caching } from "@pnp/queryable";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { SPFI, spfi } from "@pnp/sp";
import { EditComponent } from '../EditComponent/EditComponent';
export class Listing extends React.Component<IListingProps,IListingState> {
      private _sp: SPFI;
    constructor (props : IListingProps, state:  IListingState){
        super(props);
         this._sp = getSP();
       this.state={
        items:[],
        showform:false
       }
       this.getAllItems=this.getAllItems.bind(this);
       this.editItem = this.editItem.bind(this);
      }
 
      async componentDidMount(){
        await this.getAllItems();
      }

      private editItem(item: any) {
        console.log("Editing item:", item);
        this.setState({ showform: true });
        // alert(this.state.showform + "showform");
      }
    public render(): React.ReactElement<IListingProps> {   
        var showform = this.state.showform;
        // alert(this.state.showform + "showform");
        var allItems = this.state.items.map((item: any,i:number) => {
            var path='#/approve/'+item.MainListId+'/'+item.Id;
            // alert(item.RequestId + "item.RequestId");
            return(
              <tr>
                <td style={{minWidth:'60px',maxWidth:'60px'}}>
               <div className='indexdesign'> {i+1}</div>
                </td>
                <td style={{minWidth:'85px',maxWidth:'85px'}}>
                    {item.RequestId}                  
                </td>
                <td>
                    {item.Title}
                </td>
                <td style={{minWidth:'85px',maxWidth:'85px'}}>
                    {item.ProcessName}
                </td>
                <td style={{minWidth:'85px',maxWidth:'85px'}}>
                {item.ReqName}
                </td>
                <td style={{minWidth:'85px',maxWidth:'85px'}}>
                {item.ReqDt}                    
                </td>
                <td style={{minWidth:'75px',maxWidth:'75px'}}>
                {item.Status}
                </td>
                <td style={{minWidth:'75px',maxWidth:'75px'}}>

                
                <a href={path} onClick={() => this.editItem(item)}>
                <img src={require("../../assets/edit.png")} className="fas fa-trash"   alt="delete" />

                </a>
               
                </td>
              </tr> 
               )    
             
          });
       
    return (
    
        <div>
        {showform ? 
        <div>
        <EditComponent userid={this.props.userid} context={this.props.context} />
        </div>
        : 
        <section style={{display:'grid'}}>
        <table id="tabAllItems" className='mtbalenew'>
            <thead>
                <tr>
                    <th style={{minWidth:'60px',maxWidth:'60px'}}>S.No</th>
                    <th style={{minWidth:'85px',maxWidth:'85px'}}>Request Id</th>
                    <th>Title</th>
                    <th style={{minWidth:'85px',maxWidth:'85px'}}>Process Name</th>
                    <th style={{minWidth:'85px',maxWidth:'85px'}}>Requested By</th>
                    <th style={{minWidth:'85px',maxWidth:'85px'}}>Requested Date</th>
                    <th style={{minWidth:'75px',maxWidth:'75px'}}>Status</th>
                    <th style={{minWidth:'75px',maxWidth:'75px'}}>Action</th>
                </tr>
            </thead>
            <tbody>
            {allItems}
            </tbody>

        </table>
    </section>
        }


        </div>
    
    )};



 private async getAllItems(){
    var _self= this;
   // const spCache = spfi(this._sp).using(Caching({store:"session"}));
   console.log(this.props.userid , "this.props.userid ");

    // const user = await spfi(this._sp).web.ensureUser(this.props.userid);
    // console.log(user.data.Id, "user.data.Id");
    const listItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('Id,RequesterNameId,RequestId,Title,ProcessName,ApprovalLevelListItemId,RequesterName/Title,Status,AssignedToId,RequestedDate,AssignedToId,ListItemId').expand('RequesterName').filter("AssignedToId eq '"+String(this.props.userid)+"' and Status eq 'Pending'")();
     console.log(listItems, "listItems in list");
  
    var allItems: any[]=[];
    listItems.forEach( function(itm){
    // var itemId = await spCache.web.lists.getByTitle("AllProcessApprovalLevelList").items.select('Id,MainListID').filter("Id eq '"+itm.ApprovalLevelListItemId+"'")();

    if(itm.RequesterNameId != "")
    {
        itm["ReqName"]= itm.RequesterName.Title;
    }
    else{
        itm["ReqName"]='';
    }
    if(itm.RequestedDate !=''){
        itm["ReqDt"] = new Date(itm.RequestedDate).getDate()+"/"+new Date(itm.RequestedDate).getMonth()+"/"+new Date(itm.RequestedDate).getFullYear();
 
    }
    else{
        itm["ReqDt"] ='';
    }
    itm["MainListId"]= itm.ListItemId;
   // var item= itemId[0].MainListID;
  //  itm["MainListId"]=item;
    allItems.push(itm);
    _self.setState({items: allItems});
   })
  

    }

} 
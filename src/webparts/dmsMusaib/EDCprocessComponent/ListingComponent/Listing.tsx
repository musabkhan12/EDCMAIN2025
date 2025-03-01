import * as React from 'react';
import { IListingProps } from './IListingProps';
import { IListingState } from './IListingState';
import { FormComponent } from '../FormComponent/Form';

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
import type { IFormProps } from '../FormComponent/IFormProps';
import ChangeDocumentRequest from '../../ChangerequestComponent/ChangeDocumentRequest';
import DocumentCancellationProcess from '../DocumentCancellation/DocumentCancellationProcess';
export class Listing extends React.Component<IListingProps, IListingState, IFormProps> {
    private _sp: SPFI;
    constructor(props: IListingProps, state: IListingState) {
        super(props);
         this._sp = getSP();
       this.state={
        edItm: null,
        items:[],
        showform:false,
        process: "",
        siteUrl: ""
       }
       this.getAllItems=this.getAllItems.bind(this);
       this.editItem = this.editItem.bind(this);
      }
 
      async componentDidMount(){
        await this.getAllItems();
    }

    private editItem(item: any) {
        console.log("Editing item:", item);
        this.setState({ showform: true, process: item.ProcessName });
        // alert(this.state.showform + "showform");
    }
    public render(): React.ReactElement<IListingProps> {
        var showform = this.state.showform;
        // alert(this.state.showform + "showform");
        console.log(this.state.items.length, "this.state.items.length");
        var allItems = this.state.items.map((item: any, i: number) => {
            console.log(item.ProcessName + "item.ProcessName");
            console.log(item.MainListId + "item.MainListId");
            var path = `#/${item.ProcessName}/${item.MainListId}`
            // var path=`#/${item.ProcessName}/${item.MainListId}/${item.Id}`;
            // alert(item.RequestId + "item.RequestId");

            //here iam  condionally creating path
     if(item.ProcessName =="Document Cancellation" && item.Status =="Rework"){
        let actionType ="approve";
        

        var path = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId}`;

     } 
     else{
        let actionType = (item.Status === "Save As Draft" || item.Status === "Rework" || item.Status === "Save as draft") ? "edit" : "view";

        var path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;

     }      
    
            return(
              <tr>
                <td style={{minWidth:'60px',maxWidth:'60px'}}>
               <div style={{marginLeft:'10px'}} className='indexdesign'> {i+1}</div>
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
                            <img src={require("../../assets/edit.png")} className="fas fa-trash" alt="delete" />

                        </a>

                    </td>
                </tr>
            )

        });

        return (

            <div>
                {showform ?
                    <div>
                        {/* <EditComponent userid={this.props.userid} context={this.props.context} /> */}
                        {this.state.process == "Change Request" &&
                            <ChangeDocumentRequest description={''} isDarkTheme={false} environmentMessage={''} hasTeamsContext={false} userDisplayName={''} context={undefined} siteUrl={''}></ChangeDocumentRequest>
                        }
                        {this.state.process == "Document Cancellation" &&
                            <DocumentCancellationProcess description={''} isDarkTheme={false} environmentMessage={''} hasTeamsContext={false} userDisplayName={''} context={undefined} siteUrl={''}></DocumentCancellationProcess>
                        }
                        {this.state.process == "Annual Audit Program" &&
                            <FormComponent userDisplayName={''} userid={this.props.userid} context={this.props.context} item={this.state.edItm} onClose={this.closeForm} />
                        }
                    </div>
                    :
                    <section style={{ display: 'grid' }}>
                        <table id="tabAllItems" className='mtbalenew'>
                            <thead>
                                <tr>
                                    <th style={{ minWidth: '60px', maxWidth: '60px' }}>S.No</th>
                                    <th style={{ minWidth: '85px', maxWidth: '85px' }}>Request Id</th>
                                    <th>Title</th>
                                    <th style={{ minWidth: '85px', maxWidth: '85px' }}>Process Name</th>
                                    <th style={{ minWidth: '85px', maxWidth: '85px' }}>Requested By</th>
                                    <th style={{ minWidth: '85px', maxWidth: '85px' }}>Requested Date</th>
                                    <th style={{ minWidth: '75px', maxWidth: '75px' }}>Status</th>
                                    <th style={{ minWidth: '75px', maxWidth: '75px' }}>Action</th>
                                </tr>
                            </thead>
                            <tbody>
                                {allItems}
                            </tbody>

                        </table>
                    </section>
                }


            </div>

        )
    };

    closeForm = () => {
        this.setState({ showform: false });
    };

    // private async getAllItems() {
    //     var _self = this;
    //     console.log(this.props.userid, "this.props.userid");


    //     const processApprovalItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList")
    //         .items.select('Id,RequesterNameId,RequestId,Title,ProcessName,ApprovalLevelListItemId,RequesterName/Title,Status,AssignedToId,RequestedDate,AssignedToId,ListItemId')
    //         .expand('RequesterName')
    //         .filter("AssignedToId eq '" + String(this.props.userid) + "' and Status eq 'Pending'")();

    //     let allItems: any[] = [];


    //     processApprovalItems.forEach(itm => {
    //         allItems.push({
    //             RequestId: itm.RequestId,
    //             Title: itm.Title,
    //             ProcessName: itm.ProcessName,
    //             ReqName: itm.RequesterName ? itm.RequesterName.Title : '',
    //             ReqDt: itm.RequestedDate ? new Date(itm.RequestedDate).toLocaleDateString() : '',
    //             Status: itm.Status,
    //             MainListId: itm.ListItemId,
    //             Id: itm.Id
    //         });
    //     });


    //     const auditItems = await spfi(this._sp).web.lists.getByTitle("AnnualAuditProgram")
    //         .items.select('Id,MemoNumber,Title,Author/Title,Created,Status')
    //         .expand('Author')();


    //     auditItems.forEach(itm => {
    //         allItems.push({
    //             RequestId: itm.MemoNumber,
    //             Title: itm.Title, 
    //             ProcessName: "Annual Audit Program", 
    //             ReqName: itm.Author ? itm.Author.Title : '', 
    //             ReqDt: itm.Created ? new Date(itm.Created).toLocaleDateString() : '', 
    //             Status: itm.Status, 
    //             MainListId: itm.Id, 
    //             Id: itm.Id
    //         });
    //     });

    //     _self.setState({ items: allItems });
    // }

    private async getAllItems() {
        var _self = this;
        console.log(this.props.userid, "this.props.userid");


        // const processApprovalItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList")
        //     .items.select('Id,RequesterNameId,RequestId,Title,ProcessName,ApprovalLevelListItemId,RequesterName/Title,Status,AssignedToId,RequestedDate,AssignedToId,ListItemId')
        //     .expand('RequesterName')
        //     .filter("AssignedToId eq '" + String(this.props.userid) + "' and Status eq 'Pending'")();

        let allItems: any[] = [];


        // processApprovalItems.forEach(itm => {
        //     allItems.push({
        //         RequestId: itm.RequestId,
        //         Title: itm.Title,
        //         ProcessName: itm.ProcessName,
        //         ReqName: itm.RequesterName ? itm.RequesterName.Title : '',
        //         ReqDt: itm.RequestedDate ? new Date(itm.RequestedDate).toLocaleDateString() : '',
        //         Status: itm.Status,
        //         MainListId: itm.ListItemId,
        //         Id: itm.Id
        //     });
        // });


        const auditItems = await spfi(this._sp).web.lists.getByTitle("AnnualAuditProgram")
            .items.select('Id,MemoNumber,Title,Author/Title,Created,Status')
            .expand('Author')();


        auditItems.forEach(itm => {
            allItems.push({
                RequestId: itm.MemoNumber,
                Title: itm.Title,
                ProcessName: "Annual Audit Program",
                ReqName: itm.Author ? itm.Author.Title : '',
                ReqDt: itm.Created ? new Date(itm.Created).toLocaleDateString() : '',
                Status: itm.Status,
                MainListId: itm.Id,
                Id: itm.Id
            });
        });

        const ChangeRequestDocumentCancellationListItems = await spfi(this._sp).web.lists.getByTitle("ChangeRequestDocumentCancellationList")
            .items.select('Id,RequesterName/Title,RequesterName/Id,Title,Author/Title,RequestDate,Status,DocumentCode')
            .expand('Author', 'RequesterName')();

        ChangeRequestDocumentCancellationListItems.forEach(async (item) => {

            if(item.Status =="Rework"){

                const processItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList")
                .items.select('*,Title,Author/Title,Created,Status')
                .expand('Author').filter("IsInitiator eq 'Yes' and Status eq 'Pending' and ProcessName eq 'Document Cancellation' and ListItemId eq '"+item.Id+"'")();
        
        
                processItems.forEach(itm => {
                    allItems.push({
                        RequestId: item.DocumentCode,
                        Title: item.Title,
                        ProcessName: "Document Cancellation",
                        ReqName: item.RequesterName?.Title ? item.RequesterName?.Title : '',
                        ReqDt: item.RequestDate ? new Date(item.RequestDate).toLocaleDateString() : '',
                        Status: item.Status,
                        MainListId: item.Id,
                        Id: item.Id,
                        ProcessItemId:itm.Id
                    })
                
                 });
               

            }
            else{

                allItems.push({
                    RequestId: item.DocumentCode,
                    Title: item.Title,
                    ProcessName: "Document Cancellation",
                    ReqName: item.RequesterName?.Title ? item.RequesterName?.Title : '',
                    ReqDt: item.RequestDate ? new Date(item.RequestDate).toLocaleDateString() : '',
                    Status: item.Status,
                    MainListId: item.Id,
                    Id: item.Id
                });

            }

           
           
        })

        const ChangeRequestListItems = await spfi(this._sp).web.lists.getByTitle("ChangeRequestList")
            .items.select('Id,RequesterName/Title,RequesterName/Id,Title,Author/Title,RequestDate,Status,DocumentCode')
            .expand('Author', 'RequesterName')();

        ChangeRequestListItems.forEach((item) => {
            allItems.push({
                RequestId: item.DocumentCode,
                Title: item.Title,
                ProcessName: "Change Request",
                ReqName: item.RequesterName?.Title ? item.RequesterName?.Title : '',
                ReqDt: item.RequestDate ? new Date(item.RequestDate).toLocaleDateString() : '',
                Status: item.Status,
                MainListId: item.Id,
                Id: item.Id
            });
        })
        _self.setState({ items: allItems });


       
    }
} 
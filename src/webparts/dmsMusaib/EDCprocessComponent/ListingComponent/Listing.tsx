import * as React from 'react';
import { IListingProps } from './IListingProps';
import { IListingState } from './IListingState';
import FormComponent from '../FormComponent/Form';
import { Profiles } from "@pnp/sp/profiles";
import { getSP } from "../../loc/pnpjsConfig";
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
import moment from 'moment';
import AnnualAuditPlan from '../../AnnualAuditPlanComponent/AnnualAuditPlan';
import AnnualAuditReport from '../../AnnualAuditReportComponent/AnnualAuditReport';
import NonConformity from '../../NonConformityComponent/EditForm';
import MemoComponent from '../MemorandumComponent/Memorandum';
let currentuserid: any;
let currentusertitle: any;
let setloading: boolean = false;
export class Listing extends React.Component<IListingProps, IListingState, IFormProps> {
    private _sp: SPFI;

    constructor(props: IListingProps, state: IListingState) {
        super(props);
        this._sp = getSP();
        this.state = {
            edItm: null,
            items: [],
            showform: !1,
            process: "",
            siteUrl: "",
            currentPage: 1,
            itemsPerPage: 10,
            totalItems: 0,
            loading:false,
            sortColumn: 'ReqDt', // Track the currently sorted column
            sortDirection: 'desc', // Track the sort direction
            searchValues: { // Track search input values for each column
                RequestId: '',
                Title: '',
                ProcessName: '',
                ReqName: '',
                ReqDt: '',
                Status: ''
            }
        };
        this.getAllItems = this.getAllItems.bind(this);
        this.editItem = this.editItem.bind(this);
        this.handlePageChange = this.handlePageChange.bind(this);
        this.handleItemsPerPageChange = this.handleItemsPerPageChange.bind(this);
        this.handleSort = this.handleSort.bind(this);
        this.handleSearchChange = this.handleSearchChange.bind(this);
    }

    async componentDidMount() {
       debugger
        this.setState({loading:true});
        const userdata = await this._sp.web.currentUser();

        console.log(userdata, "user data edc")
        console.log(userdata.Id, "user data edc")
        currentuserid = userdata.Id;
        currentusertitle = userdata.Title
        // alert(currentusertitle + "currentusertitle")
        await this.getAllItems();
        this.setState({ loading: false });
    }

    private editItem(item: any) {
        console.log("Editing item:", item);

        if (item.ProcessName === 'Non Conformity') {
            // 1
            //const actionType = item.SubmitStatus === "No" ? "edit" : "view";
            // 2
            // const actionType = item.SubmitStatus === "No" ? "edit" : "view";
            // const newPath = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId || ""}`;
            // window.location.hash = newPath;
            // this.setState({ process: item.ProcessName, showform: true });
            // 3
            const actionType = item.SubmitStatus === "No" ? "edit"
                : (item.SubmitStatus === "Yes" && item.Status === "Rework" ? "edit" : "view");

            const newPath = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId || ""}`;

            window.location.hash = newPath;
            this.setState({ process: item.ProcessName, showform: true });

        } else {
            // alert("else Non Conformity");
            this.setState({ showform: true, process: item.ProcessName });
        }

    }
    //  private editItem(item: any) {
    //         console.log("Editing item:", item);
    //         this.setState({ showform: !0, process: item.ProcessName });
    //     }

    private handlePageChange(pageNumber: number) {
        this.setState({ currentPage: pageNumber });
    }

    private handleItemsPerPageChange(event: React.ChangeEvent<HTMLSelectElement>) {
        this.setState({ itemsPerPage: parseInt(event.target.value, 10), currentPage: 1 });
    }

    private handleSort(column: string) {
        const { sortColumn, sortDirection } = this.state;
        let newSortDirection = 'asc';
        debugger
        if (column === sortColumn) {
            newSortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
        }

        this.setState({ sortColumn: column, sortDirection: newSortDirection });
    }

    private handleSearchChange(column: string, value: string) {
        const { searchValues } = this.state;
        this.setState({
            searchValues: { ...searchValues, [column]: value },
            currentPage: 1 // Reset to the first page when searching
        });
    }

    public render(): React.ReactElement<IListingProps> {
        const { showform, items, currentPage, itemsPerPage, totalItems, sortColumn, sortDirection, searchValues } = this.state;

        // Filter items based on search values
        const filteredItems = items.filter(item => {
            return Object.keys(searchValues).every(key => {
                const columnValue = item[key]?.toString().toLowerCase();
                const searchValue = searchValues[key].toLowerCase();
                return columnValue?.includes(searchValue);
            });
        });
        let filitems = filteredItems.sort((a, b) => b.ReqDt - a.ReqDt)
        // Sort items based on the selected column and direction
        const sortedItems = filitems.sort((a, b) => {
            // debugger
            if (sortColumn) {
                const aValue = a[sortColumn];
                const bValue = b[sortColumn];
                if (aValue < bValue) return sortDirection === 'asc' ? -1 : 1;
                if (aValue > bValue) return sortDirection === 'asc' ? 1 : -1;
            }
            return 0;
        });

        // Calculate the items to display on the current page
        const indexOfLastItem = currentPage * itemsPerPage;
        const indexOfFirstItem = indexOfLastItem - itemsPerPage;
        const currentItems = sortedItems.slice(indexOfFirstItem, indexOfLastItem);

        // Generate table rows
        const allItems = currentItems.map((item: any, i: number) => {
            let path = `#/${item.ProcessName}/${item.MainListId}`;

            if ((item.ProcessName == "Document Cancellation" || item.ProcessName == "Change Request") && item.Status == "Rework") {
                if (item.ProcessItemId) {
                    let actionType = "approve";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId}`;
                } else {
                    let actionType = "view";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                }
            }
            else if ((item.ProcessName == "Annual Audit Plan" || item.ProcessName == "IMS Audit Plan") && item.Status == "Pending") {
                if (item.ProcessItemId) {
                    let actionType = "approve";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId}`;
                } else {
                    let actionType = "view";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                }
            }
            else if ((item.ProcessName == "Annual Audit Report" || item.ProcessName == "IMS Audit Report and Checklist") && item.Status == "Pending") {
                if (item.ProcessItemId) {
                    let actionType = "approve";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId}`;
                } else {
                    let actionType = "view";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                }
            }
            else if ((item.ProcessName == "Non Conformity") && item.Status == "Pending") {

                // if (item.ProcessItemId) {
                //     let actionType = "approve";
                //     path = `#/${item.ProcessName}/${actionType}/${item.MainListId}/${item.ProcessItemId}`;
                // } else 
                if (item.SubmitStatus == "Yes") {
                    // let actionType = "view";
                    // path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                    // 1
                    // let actionType = "view";
                    // path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                    if (item.Status == "Rework") {
                        let actionType = "edit";
                        path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                    } else {
                        let actionType = "view";
                        path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                    }
                } else if (item.SubmitStatus == "No") {
                    let actionType = "edit";
                    path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
                }
            }
            else {
                let actionType = (item.Status === "Save As Draft" || item.Status === "Rework" || item.Status === "Save as draft") ? "edit" : "view";
                path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
            }

            return (
                <tr key={i}>
                    <td style={{ minWidth: '40px', maxWidth: '40px' }}>
                        <div style={{ marginLeft: '5px' }} className='indexdesign'>{i + 1}</div>
                    </td>
                    <td title={item.ProcessName == "Non Conformity" ? item?.NCNumber : item?.RequestId} style={{ minWidth: '105px', maxWidth: '105px' }}>{item.ProcessName == "Non Conformity" ? item?.NCNumber : item?.RequestId}</td>
                    <td title={item.ProcessName == "Non Conformity" ? item?.ProblemDescription : item.Title} style={{ minWidth: '105px', maxWidth: '105px' }}>{item.ProcessName == "Non Conformity" ? item?.ProblemDescription : item.Title}</td>
                    <td title={item?.ProcessName === "Annual Audit Plan"
                        ? "IMS Audit Plan"
                        : item?.ProcessName === "Annual Audit Program"
                            ? "IMS Annual Audit Program"
                            : item?.ProcessName === "Annual Audit Report"
                                ? "IMS Audit Report and Checklist"
                                : item?.ProcessName} style={{ minWidth: '105px', maxWidth: '105px' }}>
                        {item?.ProcessName === "Annual Audit Plan"
                            ? "IMS Audit Plan"
                            : item?.ProcessName === "Annual Audit Program"
                                ? "IMS Annual Audit Program"
                                : item?.ProcessName === "Annual Audit Report"
                                    ? "IMS Audit Report and Checklist"
                                    : item?.ProcessName}
                    </td>
                    <td title={item.ReqName} style={{ minWidth: '80px', maxWidth: '80px' }}>{item.ReqName}</td>
                    <td title={moment(item.ReqDt).format("DD-MMM-YYYY")} style={{ minWidth: '85px', maxWidth: '85px' }}>{moment(item.ReqDt).format("DD-MMM-YYYY")}</td>
                    <td title={item.SubmitStatus == "No" ? "Save as Draft" : item.Status} style={{ minWidth: '70px', maxWidth: '70px' }}>{item.SubmitStatus == "No" ? "Save as Draft" : item.Status}</td>
                    <td style={{ minWidth: '50px', maxWidth: '50px' }}>
                        <a href={path} onClick={() => this.editItem(item)}>
                            {/* <a  onClick={() => this.editItem(item)}> */}
                            <img src={require("../../assets/edit.png")} className="fas fa-trash" alt="delete" />
                        </a>
                    </td>
                </tr>
            );
        });

        // Calculate the total number of pages
        const totalPages = Math.ceil(filteredItems.length / itemsPerPage);

        // Generate page numbers for pagination
        const pageNumbers = [];
        for (let i = 1; i <= totalPages; i++) {
            pageNumbers.push(i);
        }
        console.log("let setloading: boolean = false;", setloading);
        return (
            <div>
                {showform ? (
                    <div>
                        {this.state.process == "Change Request" && <ChangeDocumentRequest description={''} isDarkTheme={!1} environmentMessage={''} hasTeamsContext={!1} userDisplayName={''} context={undefined} siteUrl={''}></ChangeDocumentRequest>}
                        {this.state.process == "Document Cancellation" && <DocumentCancellationProcess description={''} isDarkTheme={!1} environmentMessage={''} hasTeamsContext={!1} userDisplayName={''} context={undefined} siteUrl={''}></DocumentCancellationProcess>}
                        {(this.state.process == "Annual Audit Program" || this.state.process == "IMS Annual Audit Program") && <FormComponent userDisplayName={''} userid={this.props.userid} context={this.props.context} item={this.state.edItm} onClose={this.closeForm} />}
                        {(this.state.process == "Annual Audit Plan" || this.state.process == "IMS Audit Plan") && <AnnualAuditPlan description={''} isDarkTheme={!1} environmentMessage={''} hasTeamsContext={!1} userDisplayName={''} context={undefined} siteUrl={''} />}
                        {(this.state.process == "Annual Audit Report" || this.state.process == "IMS Audit Report and Checklist") && <AnnualAuditReport description={''} isDarkTheme={!1} environmentMessage={''} hasTeamsContext={!1} userDisplayName={''} context={undefined} siteUrl={''} />}
                        {this.state.process == "Non Conformity" && <NonConformity description={''} context={this.props.context} currentUserID={this.props.userid} userDisplayName={currentusertitle} />}
                        {this.state.process == "Memorandum" && <MemoComponent userDisplayName={''} userid={this.props.userid} context={this.props.context} item={this.state.edItm} onClose={this.closeForm} />}

                    </div>
                ) : (
                    <section style={{ display: 'grid' }}>
                        <table id="tabAllItems" className='mtbalenew'>
                            <thead>
                                <tr>
                                    <th style={{ minWidth: '40px', textAlign: 'center', maxWidth: '40px' }}>

                                        <div style={{
                                            width: '100%', height: '80px', clear: 'both', display: 'flex', justifyContent: 'start', textAlign
                                                : 'center'
                                        }} className='pb-3'> S.No</div>
                                    </th>
                                    {/* {['RequestId', 'Title', 'ProcessName', 'ReqName', 'ReqDt', 'Status'].map(column => (
                                        <th key={column} style={{ minWidth: '90px', textAlign: 'center', maxWidth: '90px' }}>
                                            <div>
                                                <div onClick={() => this.handleSort(column)} style={{ cursor: 'pointer', display: 'flex', height: '35px' }}>
                                                    {/* {column} */}
                                    {/* {column === 'ProcessName'
                                                        ? 'Process Name'
                                                        : column === 'RequestId'
                                                            ? 'Request ID'
                                                            : column === 'ReqName'
                                                                ? 'Request Name'
                                                                : column === 'ReqDt'
                                                                    ? 'Request Date'
                                                                    : column} {/* Dynamically update column names */}
                                    {/* {sortColumn === column && (
                                                        <span>
                                                            {sortDirection === 'asc' ? ' ▲' : ' ▼'}
                                                        </span>
                                                    )}
                                                </div> */}
                                    {/* <input
                                                    type="text"
                                                    placeholder={`Search ${column}`}
                                                    value={searchValues[column]}
                                                    onChange={(e) => this.handleSearchChange(column, e.target.value)}
                                                    style={{ width: '100%', marginTop: '5px' }}
                                                /> */}
                                    {/* </div> */}
                                    {/* </th> */}
                                    {/* ))} */}
                                    {['RequestId', 'Title', 'ProcessName', 'ReqName', 'ReqDt', 'Status'].map(column => {
                                        let columnWidth = { minWidth: '90px', maxWidth: '90px' }; // default

                                        if (['RequestId', 'Title', 'ProcessName'].includes(column)) {
                                            columnWidth = { minWidth: '105px', maxWidth: '105px' };
                                        } else if (['ReqName', 'ReqDt'].includes(column)) {
                                            columnWidth = { minWidth: '85px', maxWidth: '85px' };
                                        } else if (column === 'Status') {
                                            columnWidth = { minWidth: '70px', maxWidth: '70px' };
                                        }

                                        return (
                                            <th key={column} style={{ ...columnWidth, textAlign: 'center' }}>
                                                <div>
                                                    <div
                                                        onClick={() => this.handleSort(column)}
                                                        style={{ cursor: 'pointer', height: '35px' }}
                                                        // display: 'flex',
                                                    >
                                                        {column === 'ProcessName'
                                                            ? 'Process Name'
                                                            : column === 'RequestId'
                                                                ? 'Request ID'
                                                                : column === 'ReqName'
                                                                    ? 'Request Name'
                                                                    : column === 'ReqDt'
                                                                        ? 'Request Date'
                                                                        : column}
                                                        {sortColumn === column && (
                                                            <span>
                                                                {sortDirection === 'asc' ? ' ▲' : ' ▼'}
                                                            </span>
                                                        )}
                                                    </div>
                                                    <input
                                                        type="text"
                                                        placeholder={`Search ${column}`}
                                                        value={searchValues[column]}
                                                        onChange={(e) => this.handleSearchChange(column, e.target.value)}
                                                        style={{ width: '100%', marginTop: '5px' }}
                                                    />
                                                </div>
                                            </th>
                                        );
                                    })}

                                    <th style={{ minWidth: '50px', textAlign: 'center', maxWidth: '50px' }}>
                                        <div style={{
                                            width: '100%', height: '80px', display: 'flex', justifyContent: 'start', textAlign
                                                : 'center'
                                        }} className='pb-3'>Action</div>

                                    </th>
                                </tr>
                            </thead>
                            {this.state.loading ?
                                <div style={{ position: 'fixed', zIndex: '9', left: '0%', top: '0%' }} className="loadernewadd mt-10">
                                    <div>
                                        <img
                                            src={require("../../assets/edc-gif.gif")}
                                            className="alignrightl"
                                            alt="Loading..."
                                        />
                                    </div>
                                    <span>Loading </span>{" "}
                                    <span>
                                        <img
                                            src={require("../../assets/edcnew.gif")}
                                            className="alignrightl"
                                            alt="Loading..."
                                        />
                                    </span>
                                </div> :
                                <tbody>{allItems}</tbody>
                            }
                        </table>
                        <div className="pagination">
                            <button onClick={() => this.handlePageChange(currentPage - 1)} disabled={currentPage === 1}>Previous</button>
                            {pageNumbers.map(number => (
                                <button
                                    key={number}
                                    onClick={() => this.handlePageChange(number)}
                                    disabled={number === currentPage}
                                    style={{ fontWeight: number === currentPage ? 'bold' : 'normal' }}
                                >
                                    {number}
                                </button>
                            ))}

                            <select style={{ height: '38px', marginTop: '19px' }} value={itemsPerPage} onChange={this.handleItemsPerPageChange}>
                                <option value={5}>5</option>
                                <option value={10}>10</option>
                                <option value={20}>20</option>
                                <option value={50}>50</option>
                            </select>
                            <button onClick={() => this.handlePageChange(currentPage + 1)} disabled={currentPage === totalPages}>Next</button>
                        </div>
                    </section>
                )}
            </div>
        );
    }

    closeForm = () => {
        this.setState({ showform: !1 });
    };

    private async getAllItems() {
        const _self = this;
        console.log(this.props.userid, "this.props.userid");
        let allItems: any[] = [];
        const auditItems = await spfi(this._sp).web.lists.getByTitle("AnnualAuditProgram").items.select('Id,MemoNumber,Title,Author/Title,Created,Status,Subject').expand('Author').filter(`Author/ID eq '${this.props.userid}'`)();
        auditItems.forEach(itm => {
            allItems.push({
                RequestId: itm.MemoNumber || "",
                Title: itm.Subject || "",
                ProcessName: "IMS Annual Audit Program",
                ReqName: itm.Author ? itm.Author.Title : '',
                ReqDt: new Date(itm.Created),
                Status: itm.Status,
                MainListId: itm.Id,
                Id: itm.Id,
                SubmitStatus: ''
            });
        });

        const MemoItems = await spfi(this._sp).web.lists.getByTitle("Memorandum").items.select('Id,MemoNumber,Title,Author/Title,Created,Status,Subject').expand('Author').filter(`Author/ID eq '${this.props.userid}'`)();
        MemoItems.forEach(itm => {
            allItems.push({
                RequestId: itm.MemoNumber || "",
                Title: itm.Subject || "",
                ProcessName: "Memorandum",
                ReqName: itm.Author ? itm.Author.Title : '',
                ReqDt: new Date(itm.Created),
                Status: itm.Status,
                MainListId: itm.Id,
                Id: itm.Id,
                SubmitStatus: ''
            });
        });

        const AnnualAuditPlanList = await spfi(this._sp).web.lists.getByTitle("AnnualAuditPlanList").items.select('Id,MemoNumber,Title,Author/Title,Created,Status,ReferenceNumber,Subject').expand('Author').filter(`Author/ID eq '${this.props.userid}'`)();
        AnnualAuditPlanList.forEach(async itm => {

            if (itm.Status === "Pending") {
                const processItems1 = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('*,Title,Author/Title,Created,Status').expand('Author').filter(`IsInitiator eq 'Yes' and (Status eq 'Pending' or Status eq 'Save as draft') and ProcessName eq 'Annual Audit Plan' and ListItemId eq ${itm.Id}`)();
                if (processItems1.length > 0) {
                    for (const itom of processItems1) {

                        allItems.push({
                            RequestId: itm.MemoNumber ? itm.MemoNumber : "",
                            Title: itm.Subject ? itm.Subject : "",
                            ProcessName: "IMS Audit Plan",
                            ReqName: itm.Author ? itm.Author.Title : '',
                            ReqDt: new Date(itm.Created),
                            Status: itm.Status,
                            MainListId: itm.Id,
                            Id: itm.Id,
                            ProcessItemId: itom.Id,
                            SubmitStatus: ''
                        });

                    }

                }
                else {

                    allItems.push({
                        RequestId: itm.MemoNumber ? itm.MemoNumber : "",
                        Title: itm.Subject ? itm.Subject : "",
                        ProcessName: "IMS Audit Plan",
                        ReqName: itm.Author ? itm.Author.Title : '',
                        ReqDt: new Date(itm.Created),
                        // ? moment(itm.Created).format("DD-MMM-YYYY") : ''
                        Status: itm.Status,
                        MainListId: itm.Id,
                        Id: itm.Id,
                        SubmitStatus: ''
                    });

                }
            }
            else {

                allItems.push({
                    RequestId: itm.MemoNumber ? itm.MemoNumber : "",
                    Title: itm.Subject ? itm.Subject : "",
                    ProcessName: "IMS Audit Plan",
                    ReqName: itm.Author ? itm.Author.Title : '',
                    ReqDt: new Date(itm.Created),
                    Status: itm.Status,
                    MainListId: itm.Id,
                    Id: itm.Id,
                    SubmitStatus: ''
                });

            }

        });

        const ChangeRequestDocumentCancellationListItems = await spfi(this._sp).web.lists.getByTitle("ChangeRequestDocumentCancellationList").items.select('Id,RequesterName/Title,RequesterName/Id,Title,Author/Title,RequestDate,Status,DocumentCode,ReferenceNumber,Created').expand('Author', 'RequesterName').filter(`Author/ID eq '${this.props.userid}'`).orderBy("Modified", false)();
        for (const item of ChangeRequestDocumentCancellationListItems) {
            if (item.Status === "Rework") {
                const processItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('*,Title,Author/Title,Created,Status').expand('Author').filter(`IsInitiator eq 'Yes' and (Status eq 'Pending' or Status eq 'Save as draft') and ProcessName eq 'Document Cancellation' and ListItemId eq ${item.Id}`)();
                if (processItems.length > 0) {
                    for (const itm of processItems) {
                        allItems.push({
                            RequestId: item.DocumentCode,
                            Title: item.ReferenceNumber,
                            ProcessName: "Document Cancellation",
                            ReqName: item.RequesterName?.Title || '',
                            ReqDt: new Date(item.Created),
                            Status: item.Status,
                            MainListId: item.Id,
                            Id: item.Id,
                            ProcessItemId: itm.Id,
                            SubmitStatus: ''
                        });
                    }
                } else {
                    allItems.push({
                        RequestId: item.DocumentCode,
                        Title: item.ReferenceNumber,
                        ProcessName: "Document Cancellation",
                        ReqName: item.RequesterName?.Title || '',
                        ReqDt: new Date(item.Created),
                        Status: item.Status,
                        MainListId: item.Id,
                        Id: item.Id,
                        SubmitStatus: ''
                    });
                }
            } else {
                allItems.push({
                    RequestId: item.DocumentCode,
                    Title: item.ReferenceNumber,
                    ProcessName: "Document Cancellation",
                    ReqName: item.RequesterName?.Title || '',
                    ReqDt: new Date(item.Created),
                    Status: item.Status,
                    MainListId: item.Id,
                    Id: item.Id,
                    SubmitStatus: ''
                });
            }
        }

        const ChangeRequestListItems = await spfi(this._sp).web.lists.getByTitle("ChangeRequestList").items.select('Id,ReferenceNumber,RequesterName/Title,RequesterName/Id,Title,Author/Title,RequestDate,Status,DocumentCode,Created').expand('Author', 'RequesterName').filter(`Author/ID eq '${this.props.userid}'`).orderBy("Modified", false)();
        for (const item of ChangeRequestListItems) {
            if (item.Status === "Rework") {
                const processItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('*,Title,Author/Title,Created,Status').expand('Author').filter(`IsInitiator eq 'Yes' and Status eq 'Pending' and ProcessName eq 'Change Request' and ListItemId eq ${item.Id}`)();
                if (processItems.length > 0) {
                    for (const itm of processItems) {
                        allItems.push({
                            RequestId: item.DocumentCode == "" || item.DocumentCode == null ? " " : item.DocumentCode,
                            Title: item.ReferenceNumber == "" || item.ReferenceNumber == null ? " " : item.ReferenceNumber,
                            ProcessName: "Change Request",
                            ReqName: item.RequesterName?.Title || '',
                            ReqDt: new Date(item.Created),
                            Status: item.Status,
                            MainListId: item.Id,
                            Id: item.Id,
                            ProcessItemId: itm.Id,
                            SubmitStatus: ''
                        });
                    }
                } else {
                    allItems.push({
                        RequestId: item.DocumentCode == "" || item.DocumentCode == null ? " " : item.DocumentCode,
                        Title: item.ReferenceNumber == "" || item.ReferenceNumber == null ? " " : item.ReferenceNumber,
                        ProcessName: "Change Request",
                        ReqName: item.RequesterName?.Title || '',
                        ReqDt: new Date(item.Created),
                        Status: item.Status,
                        MainListId: item.Id,
                        Id: item.Id,
                        SubmitStatus: ''
                    });
                }
            } else {
                allItems.push({
                    RequestId: item.DocumentCode == "" || item.DocumentCode == null ? " " : item.DocumentCode,
                    Title: item.ReferenceNumber == "" || item.ReferenceNumber == null ? " " : item.ReferenceNumber,
                    ProcessName: "Change Request",
                    ReqName: item.RequesterName?.Title || '',
                    ReqDt: new Date(item.Created),
                    // ? moment(item.RequestDate).format("DD-MMM-YYYY") : ''
                    Status: item.Status,
                    MainListId: item.Id,
                    Id: item.Id,
                    SubmitStatus: ''
                });
            }
        }

        const AnnualAuditReportList = await spfi(this._sp).web.lists.getByTitle("AnnualAuditReportList").items.select('*,ReportCode,Id,Title,Author/Title,Created,Status,ReferenceNumber').expand('Author').filter(`Author/ID eq '${this.props.userid}'`)();
        AnnualAuditReportList.forEach(async itm => {

            if (itm.Status === "Pending") {
                const processItems2 = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('*,Title,Author/Title,Created,Status').expand('Author').filter(`IsInitiator eq 'Yes' and (Status eq 'Pending' or Status eq 'Save as draft') and ProcessName eq 'Annual Audit Report' and ListItemId eq ${itm.Id}`)();
                if (processItems2.length > 0) {
                    for (const itom of processItems2) {

                        allItems.push({
                            RequestId: itm.ReportCode ? itm.ReportCode : "",
                            ReportCode: itm.ReportCode ? itm.ReportCode : "",
                            Title: itm.MemoNumber ? itm.MemoNumber : "",
                            ProcessName: "IMS Audit Report and Checklist",
                            ReqName: itm.Author ? itm.Author.Title : '',
                            ReqDt: new Date(itm.Created),
                            Status: itm.Status,
                            MainListId: itm.Id,
                            Id: itm.Id,
                            ProcessItemId: itom.Id,
                            SubmitStatus: ''
                        });

                    }

                }
                else {

                    allItems.push({
                        RequestId: itm.ReportCode ? itm.ReportCode : "",
                        ReportCode: itm.ReportCode ? itm.ReportCode : "",
                        Title: itm.MemoNumber ? itm.MemoNumber : "",
                        ProcessName: "IMS Audit Report and Checklist",
                        ReqName: itm.Author ? itm.Author.Title : '',
                        ReqDt: new Date(itm.Created),
                        Status: itm.Status,
                        MainListId: itm.Id,
                        Id: itm.Id,
                        SubmitStatus: ''
                    });

                }
            }
            else {

                allItems.push({
                    RequestId: itm.ReportCode ? itm.ReportCode : "",
                    ReportCode: itm.ReportCode ? itm.ReportCode : "",
                    Title: itm.MemoNumber ? itm.MemoNumber : "",
                    ProcessName: "IMS Audit Report and Checklist",
                    ReqName: itm.Author ? itm.Author.Title : '',
                    ReqDt: new Date(itm.Created),
                    Status: itm.Status,
                    MainListId: itm.Id,
                    Id: itm.Id,
                    SubmitStatus: ''
                });

            }

        });

        const nonconfirmity = await spfi(this._sp).web.lists.getByTitle('NonConformityList').items.select('Id,NCNumber, DocumentCode ,Author/Title , Status , Created,ProblemDescription , SubmitStatus , NCRNo').expand('Author').filter(`Author/ID eq '${this.props.userid}'`).orderBy("Modified", false)();
        console.log(nonconfirmity, "nonconfirmity")
        for (const item of nonconfirmity) {
            // alert (item.DocumentCode + "item.DocumentCode" )
            console.log(item.DocumentCode, "item.DocumentCode")
            let Doccode = ''
            if (item.DocumentCode == "" || item.DocumentCode == null) {
                Doccode = ""
            } else {
                Doccode = item.DocumentCode
            }
            allItems.push({
                RequestId: item.DocumentCode == "" || item.DocumentCode == null ? " " : item.DocumentCode,
                //RequestId:  item.NCRNo,

                NCRNo: item.NCRNo,
                NCNumber: item.NCNumber,
                Title: item.DocumentCode == "" || item.DocumentCode == null ? " " : item.DocumentCode,
                //Title: item.ProblemDescription,
                ProblemDescription: item.ProblemDescription,
                ProcessName: "Non Conformity",
                ReqName: item.Author?.Title || '',
                // ReqDt: item.Created ? moment(item.Created).format("DD-MMM-YYYY") : '',
                ReqDt: new Date(item.Created),
                Status: item.Status,
                MainListId: item.Id,
                Id: item.Id,
                SubmitStatus: item.SubmitStatus
            });

        }
        const sortedAllItems = allItems.sort((a, b) => moment(b.ReqDt, "DD-MMM-YYYY").toDate().getTime() - moment(a.ReqDt, "DD-MMM-YYYY").toDate().getTime());
        _self.setState({ items: sortedAllItems, totalItems: sortedAllItems.length });
        // _self.setState({ items: allItems, totalItems: allItems.length });
        // console.log(allItems, "all items");
        // console.log(allItems.length, "all items length");
        // console.log(this.state.items, "this.state.items");
    }
}
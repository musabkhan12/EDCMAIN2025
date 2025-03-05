import * as React from 'react';
import { IListingProps } from './IListingProps';
import { IListingState } from './IListingState';
import { FormComponent } from '../FormComponent/Form';
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
            sortColumn: null, // Track the currently sorted column
            sortDirection: 'asc', // Track the sort direction
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
        await this.getAllItems();
    }

    private editItem(item: any) {
        console.log("Editing item:", item);
        this.setState({ showform: !0, process: item.ProcessName });
    }

    private handlePageChange(pageNumber: number) {
        this.setState({ currentPage: pageNumber });
    }

    private handleItemsPerPageChange(event: React.ChangeEvent<HTMLSelectElement>) {
        this.setState({ itemsPerPage: parseInt(event.target.value, 10), currentPage: 1 });
    }

    private handleSort(column: string) {
        const { sortColumn, sortDirection } = this.state;
        let newSortDirection = 'asc';

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

        // Sort items based on the selected column and direction
        const sortedItems = filteredItems.sort((a, b) => {
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
            } else {
                let actionType = (item.Status === "Save As Draft" || item.Status === "Rework" || item.Status === "Save as draft") ? "edit" : "view";
                path = `#/${item.ProcessName}/${actionType}/${item.MainListId}`;
            }

            return (
                <tr key={i}>
                    <td style={{ minWidth: '60px', maxWidth: '60px' }}>
                        <div style={{ marginLeft: '10px' }} className='indexdesign'>{i + 1}</div>
                    </td>
                    <td style={{ minWidth: '85px', maxWidth: '85px' }}>{item.RequestId}</td>
                    <td>{item.Title}</td>
                    <td title={item.ProcessName} style={{ minWidth: '85px', maxWidth: '85px' }}>{item.ProcessName}</td>
                    <td style={{ minWidth: '85px', maxWidth: '85px' }}>{item.ReqName}</td>
                    <td style={{ minWidth: '85px', maxWidth: '85px' }}>{moment(item.ReqDt).format("DD-MMM-YYYY")}</td>
                    <td style={{ minWidth: '75px', maxWidth: '75px' }}>{item.Status}</td>
                    <td style={{ minWidth: '75px', maxWidth: '75px' }}>
                        <a href={path} onClick={() => this.editItem(item)}>
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

        return (
            <div>
                {showform ? (
                    <div>
                        {this.state.process == "Change Request" && <ChangeDocumentRequest description={''} isDarkTheme={!1} environmentMessage={''} hasTeamsContext={!1} userDisplayName={''} context={undefined} siteUrl={''}></ChangeDocumentRequest>}
                        {this.state.process == "Document Cancellation" && <DocumentCancellationProcess description={''} isDarkTheme={!1} environmentMessage={''} hasTeamsContext={!1} userDisplayName={''} context={undefined} siteUrl={''}></DocumentCancellationProcess>}
                        {this.state.process == "Annual Audit Program" && <FormComponent userDisplayName={''} userid={this.props.userid} context={this.props.context} item={this.state.edItm} onClose={this.closeForm} />}
                    </div>
                ) : (
                    <section style={{ display: 'grid' }}>
                        <table id="tabAllItems" className='mtbalenew'>
                            <thead>
                                <tr>
                                    <th style={{ minWidth: '60px', maxWidth: '60px' }}>
                                        S.No
                                    </th>
                                    {['RequestId', 'Title', 'ProcessName', 'ReqName', 'ReqDt', 'Status'].map(column => (
                                        <th key={column} style={{ minWidth: '85px', maxWidth: '85px' }}>
                                            <div>
                                                <div onClick={() => this.handleSort(column)} style={{ cursor: 'pointer' }}>
                                                    {column}
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
                                    ))}
                                    <th style={{ minWidth: '75px', maxWidth: '75px' }}>Action</th>
                                </tr>
                            </thead>
                            <tbody>{allItems}</tbody>
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
                            <button onClick={() => this.handlePageChange(currentPage + 1)} disabled={currentPage === totalPages}>Next</button>
                            <select value={itemsPerPage} onChange={this.handleItemsPerPageChange}>
                                <option value={5}>5</option>
                                <option value={10}>10</option>
                                <option value={20}>20</option>
                                <option value={50}>50</option>
                            </select>
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
        const auditItems = await spfi(this._sp).web.lists.getByTitle("AnnualAuditProgram").items.select('Id,MemoNumber,Title,Author/Title,Created,Status').expand('Author')();
        auditItems.forEach(itm => {
            allItems.push({
                RequestId: itm.MemoNumber,
                Title: itm.Title,
                ProcessName: "Annual Audit Program",
                ReqName: itm.Author ? itm.Author.Title : '',
                ReqDt: itm.Created ? moment(itm.Created).format("DD-MMM-YYYY") : '',
                Status: itm.Status,
                MainListId: itm.Id,
                Id: itm.Id
            });
        });

        const ChangeRequestDocumentCancellationListItems = await spfi(this._sp).web.lists.getByTitle("ChangeRequestDocumentCancellationList").items.select('Id,RequesterName/Title,RequesterName/Id,Title,Author/Title,RequestDate,Status,DocumentCode,ReferenceNumber').expand('Author', 'RequesterName')();
        for (const item of ChangeRequestDocumentCancellationListItems) {
            if (item.Status === "Rework") {
                const processItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('*,Title,Author/Title,Created,Status').expand('Author').filter(`IsInitiator eq 'Yes' and Status eq 'Pending' and ProcessName eq 'Document Cancellation' and ListItemId eq ${item.Id}`)();
                if (processItems.length > 0) {
                    for (const itm of processItems) {
                        allItems.push({
                            RequestId: item.DocumentCode,
                            Title: item.ReferenceNumber,
                            ProcessName: "Document Cancellation",
                            ReqName: item.RequesterName?.Title || '',
                            ReqDt: item.RequestDate ? moment(item.RequestDate).format("DD-MMM-YYYY") : '',
                            Status: item.Status,
                            MainListId: item.Id,
                            Id: item.Id,
                            ProcessItemId: itm.Id
                        });
                    }
                } else {
                    allItems.push({
                        RequestId: item.DocumentCode,
                        Title: item.ReferenceNumber,
                        ProcessName: "Document Cancellation",
                        ReqName: item.RequesterName?.Title || '',
                        ReqDt: item.RequestDate ? moment(item.RequestDate).format("DD-MMM-YYYY") : '',
                        Status: item.Status,
                        MainListId: item.Id,
                        Id: item.Id
                    });
                }
            } else {
                allItems.push({
                    RequestId: item.DocumentCode,
                    Title: item.ReferenceNumber,
                    ProcessName: "Document Cancellation",
                    ReqName: item.RequesterName?.Title || '',
                    ReqDt: item.RequestDate ? moment(item.RequestDate).format("DD-MMM-YYYY") : '',
                    Status: item.Status,
                    MainListId: item.Id,
                    Id: item.Id
                });
            }
        }

        const ChangeRequestListItems = await spfi(this._sp).web.lists.getByTitle("ChangeRequestList").items.select('Id,RequesterName/Title,RequesterName/Id,Title,Author/Title,RequestDate,Status,DocumentCode').expand('Author', 'RequesterName')();
        for (const item of ChangeRequestListItems) {
            if (item.Status === "Rework") {
                const processItems = await spfi(this._sp).web.lists.getByTitle("ProcessApprovalList").items.select('*,Title,Author/Title,Created,Status').expand('Author').filter(`IsInitiator eq 'Yes' and Status eq 'Pending' and ProcessName eq 'Change Request' and ListItemId eq ${item.Id}`)();
                if (processItems.length > 0) {
                    for (const itm of processItems) {
                        allItems.push({
                            RequestId: item.DocumentCode,
                            Title: item.Title,
                            ProcessName: "Change Request",
                            ReqName: item.RequesterName?.Title || '',
                            ReqDt: item.RequestDate ? moment(item.RequestDate).format("DD-MMM-YYYY") : '',
                            Status: item.Status,
                            MainListId: item.Id,
                            Id: item.Id,
                            ProcessItemId: itm.Id
                        });
                    }
                } else {
                    allItems.push({
                        RequestId: item.DocumentCode,
                        Title: item.Title,
                        ProcessName: "Change Request",
                        ReqName: item.RequesterName?.Title || '',
                        ReqDt: item.RequestDate ? moment(item.RequestDate).format("DD-MMM-YYYY") : '',
                        Status: item.Status,
                        MainListId: item.Id,
                        Id: item.Id
                    });
                }
            } else {
                allItems.push({
                    RequestId: item.DocumentCode,
                    Title: item.Title,
                    ProcessName: "Change Request",
                    ReqName: item.RequesterName?.Title || '',
                    ReqDt: item.RequestDate ? moment(item.RequestDate).format("DD-MMM-YYYY") : '',
                    Status: item.Status,
                    MainListId: item.Id,
                    Id: item.Id
                });
            }
        }

        _self.setState({ items: allItems, totalItems: allItems.length });
    }
}
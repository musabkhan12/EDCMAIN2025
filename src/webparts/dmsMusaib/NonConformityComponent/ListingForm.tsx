import * as React from 'react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { spfi, SPFx } from '@pnp/sp';
import type { IAuditPlanProps } from './IAuditPlanProps';

export class IListingState {
    items: any;
}
export default class ListingForm extends React.Component<IAuditPlanProps, IListingState> {

    constructor(props: IAuditPlanProps) {
        super(props);
        this.state = {
            items: []
        }
        this.getAllItems = this.getAllItems.bind(this);
    }

    async componentDidMount() {
        await this.getAllItems();
    }
      
    public render(): React.ReactElement<IAuditPlanProps> {
        var allItems = this.state.items.map((item: any, i: number) => {
            var path = "";
            if (item.CurrentUserRole == "FirstAssignedTo" || item.CurrentUserRole == "DelegateTo") {
                path = '#/edit/' + item.MainListId + '/' + item.Id;
            }
            else {
                path = '#/approve/' + item.MainListId + '/' + item.Id;
            }

            return (
                <tr>
                    <td>
                        {i + 1}
                    </td>
                    <td>
                        {item.RequestId}
                    </td>
                    <td>
                        {item.RequestId}
                    </td>
                    <td>
                        {item.ProcessName}
                    </td>
                    <td>
                        {item.ReqName}
                    </td>
                    <td>
                        {item.ReqDt}
                    </td>
                    <td>
                        {item.Status}
                    </td>
                    <td>
                        <a href={path}>Edit</a>

                    </td>
                </tr>
            )

        });

        return (
            <section>
                <table id="tabAllItems">
                    <thead>
                        <tr>
                            <th>Sl No</th>
                            <th>Request Id</th>
                            <th>Title</th>
                            <th>Process Name</th>
                            <th>Requested By</th>
                            <th>Requested Date</th>
                            <th>Status</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {allItems}
                    </tbody>

                </table>
            </section>
        )
    };   

    private async getAllItems() {
        const sp = spfi().using(SPFx(this.props.context));
        var _self = this;
        const listItems = await sp.web.lists.getByTitle("ProcessApprovalList").items
        .select('Id,RequesterNameId,RequestId,Title,CurrentUserRole,ProcessName,ApprovalLevelListItemId,RequesterName/Title,Status,AssignedToId,RequestedDate,AssignedToId,ListItemId').expand('RequesterName')
        .filter("AssignedToId eq '" + this.props.currentUserID + "' and Status eq 'Pending'").orderBy("Id", false)();
        var allItems: any[] = [];
        listItems.forEach(function (itm) {

            if (itm.RequesterNameId != "") {
                itm["ReqName"] = itm.RequesterName?.Title;
            }
            else {
                itm["ReqName"] = '';
            }
            if (itm.RequestedDate != '') {
                itm["ReqDt"] = new Date(itm.RequestedDate).getDate() + "/" + new Date(itm.RequestedDate).getMonth() + "/" + new Date(itm.RequestedDate).getFullYear();

            }
            else {
                itm["ReqDt"] = '';
            }
            itm["MainListId"] = itm.ListItemId;
            allItems.push(itm);
            _self.setState({ items: allItems });
        })
    }
} 
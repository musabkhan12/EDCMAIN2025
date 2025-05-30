import { escape } from "@microsoft/sp-lodash-subset";

import React, { useState } from "react";
// import { updateItemApproval, updateItemApproval2 } from "../../APISearvice/ApprovalService";
import { updateItemApproval, updateItemApproval2 } from "../../APISearvice/ApprovalService";
// import { getSP } from "../../webparts/addDynamicBanner/loc/pnpjsConfig";
import { getSP } from "../../webparts/dmsMusaib/loc/pnpjsConfig";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Swal from "sweetalert2";
// import { updateItem, updateItemChangeRequestList } from "../../APISearvice/DocumentCancellation";
import { updateItem, updateItemChangeRequestList } from "../../APISearvice/DocumentCancellation";

export interface IWorkflowActionProps {
  currentItem: any;
  ContentType: string;
  ctx: WebPartContext;
  DisableApproval?: boolean;
  DisableRework?: boolean;
  DisableReject?: boolean;
  DisableCancel?: boolean;

}

export const WorkflowAction = (props: IWorkflowActionProps) => {


  const siteUrl = props.ctx.pageContext.site.absoluteUrl;

  const sp = getSP(props.ctx);

  const [formData, setFormData] = React.useState({
    Remark: '',
  });
    const [ValidRemark, setValidRemark] = React.useState(true);

  const onChange = (name: string, value: string) => {

    debugger

    setFormData((prevData) => ({

      ...prevData,

      [name]: value,

    }));

  };
  const handleCancel = () => {
    window.location.href = `${siteUrl}/SitePages/MyApprovals.aspx`;
  }
  const handleFromSubmit = async (e: React.MouseEvent<HTMLButtonElement, MouseEvent>, Status: string) => {

    e.preventDefault();
    setValidRemark(true);
    let postPayload = {}
    let postPayload2 = {}
    if ((Status === 'Rework' || Status === 'Rejected') && formData.Remark === "") {
      setValidRemark(false);
      Swal.fire('Please fill the mandatory fields', '', 'warning');
      return;
    }


    if (props.ContentType == "Document Cancellation" || props.ContentType == "Change Request" || props.ContentType == "Annual Audit Plan" || props.ContentType == "IMS Audit Plan" || props.ContentType == "Annual Audit Report"|| props.ContentType == "IMS Audit Report and Checklist"|| props.ContentType == "Annual Audit Program"|| props.ContentType == "IMS Annual Audit Program"|| props.ContentType == "Memorandum") {
      const currentUser = await sp.web.currentUser();

      postPayload = {

        Remark: formData.Remark,

        Status: Status,

        ActionTakenById: currentUser.Id,
        ActionTakenOn: new Date().toLocaleDateString("en-CA")

      };

      postPayload2 = {

        Status: Status,
        OESSubmitStatus: "No",
        InitiatorSubmitStatus: "No",
        CurrentUserRole: "Initiator",
        SubmitStatus: "No",

      };

    }
    else {
      postPayload = {

        Remark: formData.Remark,

        Status: Status

      };
    }


    console.log(postPayload);
    let confirmation, resultmessage = "";

    if (Status == 'Approved') {
      confirmation = "Do you want to approve this request?";
      resultmessage = 'Approved successfully.'

    }
    else if (Status == 'Rejected') {
      confirmation = "Do you want to reject this request?";
      resultmessage = 'Rejected successfully.'
    }

    else if (Status == 'Rework') {
      confirmation = "Do you want to rework this request?";
      resultmessage = 'Sent for rework.'
    }

    Swal.fire({
      // title: 'Do you want to save?',
      title: confirmation,
      showConfirmButton: true,
      showCancelButton: true,
      confirmButtonText: "Yes",
      cancelButtonText: "No"

    }).then(async (result) => {
      console.log(result)
      if (result.isConfirmed) {
        var postResult;
        if (props.ContentType == "Document Cancellation" || props.ContentType == "Change Request") {
          postResult = await updateItemApproval2(postPayload, sp, props.currentItem.Id);
          if (Status == 'Rework') {
            const postResult2 = props.ContentType == "Document Cancellation" ? await updateItem(postPayload2, sp, Number(props.currentItem.ListItemId)) :
              await updateItemChangeRequestList(postPayload2, sp, Number(props.currentItem.ListItemId))

          }
          // const postId = postResult?.data?.ID;

        }
        else if (props.ContentType == "Annual Audit Plan"|| props.ContentType == "Annual Audit Program" || props.ContentType == "IMS Audit Plan"|| props.ContentType == "IMS Annual Audit Program"|| props.ContentType == "Memorandum") {
          postResult = await updateItemApproval2(postPayload, sp, props.currentItem.Id);

        }
        else if (props.ContentType == "Annual Audit Report" || props.ContentType == "IMS Audit Report and Checklist") {
          postResult = await updateItemApproval2(postPayload, sp, props.currentItem.Id);

        }
        else {
          postResult = await updateItemApproval(postPayload, sp, props.currentItem.Id);

        }

        if (postResult) {
          Swal.fire(resultmessage, '', 'success');
          setTimeout(() => {

            // window.location.reload()

            location.href = `${siteUrl}/SitePages/MyApprovals.aspx`;

          }, 1000);


        }

      }
    })


  }

  return (

    <div className="card">

      <div className="card-body">

        <div className="row">
          {

            (props.currentItem.Status == "Pending" || props.currentItem.IsRework) && (<div className="col-lg-12">

              <div className="mb-0" >

                <label htmlFor="example-textarea" className="form-label text-dark font-14">Remarks <span className="text-danger1"> *</span></label>

                <textarea style={{ height: '80px' }} className={`form-control ${(!ValidRemark) ? "border-on-error" : ""}`} id="example-textarea" rows={5} name="Remark" value={formData.Remark}

                  onChange={(e) => onChange(e.target.name, e.target.value)}></textarea>

              </div>

            </div>)

          }


        </div>

        {


          (props.currentItem.Status == "Pending" || props.currentItem.IsRework) && (

            <div className="row mt-3">

              <div className="col-12 text-center">

                {!props.DisableApproval ? (<a href="my-approval.html">

                  <button type="button" className="btn btn-success waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Approved')}>

                    <i className="fe-check-circle me-1"></i> Approve

                  </button>

                </a>) : (<div></div>)}

                {!props.DisableApproval ? (<a href="my-approval.html">

                  <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Rework')}>

                    <i className="fe-corner-up-left me-1"></i> Rework

                  </button>

                </a>) : (<div></div>)}

                {!props.DisableApproval ? (<a href="my-approval.html">

                  <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Rejected')}>

                    <i className="fe-x-circle me-1"></i> Reject

                  </button>

                </a>) : (<div></div>)}

                {!props.DisableCancel ? (<button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={(e) => handleCancel()}>

                  <i className="fe-x me-1"></i> Cancel

                </button>) : (<div></div>)}

              </div>

            </div>

          )

        }


      </div>

    </div>
  )

}
import * as React from "react";
import { useParams } from "react-router-dom";
import ViewForm from "./ViewForm"
import { IAuditPlanProps } from "./IAuditPlanProps";

const ViewFormWrapper: React.FC<IAuditPlanProps> = (props) => {
  const { type, postId, itmId } = useParams();
  
    return (
      <ViewForm
        {...props}
        edType={type}
        edItm={postId}
        approvalItemId={itmId}
      />
    );
  };

export default ViewFormWrapper;

import * as React from "react";
import { useParams } from "react-router-dom";
import EditForm from "./EditForm";
import { IAuditPlanProps } from "./IAuditPlanProps";

const EditFormWrapper: React.FC<IAuditPlanProps> = (props) => {
  // Extract parameters using destructuring
  const { type, postId, itmId } = useParams();

  return (
    <EditForm
      {...props}
      edType={type}
      edItm={postId}
      approvalItemId={itmId}
    />
  );
};

export default EditFormWrapper;


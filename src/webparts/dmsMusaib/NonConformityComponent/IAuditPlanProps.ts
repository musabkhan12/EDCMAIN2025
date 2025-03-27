export interface IAuditPlanProps {
  description: string;  
  context: any|null;  
  currentUserID: string;
  userDisplayName: string;
  edType?: string;
  edItm?: string; 
  approvalItemId?: string;
}

import React, { useEffect } from 'react'
import "bootstrap/dist/css/bootstrap.min.css";
import Select from "react-select";
import { useRef, useState } from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import Swal from "sweetalert2";


interface CreateFolderProps {
  OthProps: { [key: string]: string };
  onReturnToMain: () => void;
}

let IsApprovalColumnId:number;
let IsUpdate:boolean;
const ManageWorkFlow :React.FC<CreateFolderProps> = ({
  OthProps,
  onReturnToMain,
}) =>{
    console.log("props",OthProps)
    const sp: SPFI = getSP();
    const currentUserEmailRef = useRef('');

    const getcurrentuseremail = async()=>{
        const userdata = await sp.web.currentUser();
        currentUserEmailRef.current = userdata.Email;
        setDefaultValues();
       }  


  const [users, setUsers] = React.useState<any[]>([]);
  const [toggleApprover, setToggleApprover]=React.useState<string>();


  const handleSetNewApprover=()=>{
    setToggleApprover("Yes");
  }
  console.log("Users Array", users);


  const setDefaultValues=async()=>{
    try {
            // Fetch the Data from the DMSPreviewFormMaster and check if approver is added to the document library or not
            const approverDetails=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("IsApproval","Id").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and IsDocumentLibrary eq 1`)();
            console.log("approverDetails",approverDetails);
            IsUpdate=approverDetails[0].IsApproval;
            IsApprovalColumnId=approverDetails[0].Id;

            
            if(!IsUpdate){
              setToggleApprover("No");
              setRows([{ id: 0, selectionType: "One", approvedUserList: [] }]);
            }else{
              setToggleApprover("Yes");
            }
            const LibraryApproverDdetails = await sp.web.lists
            .getByTitle("DMSFolderPermissionMaster")
            .items.select("CurrentUser" , "SiteName" , "DocumentLibraryName" , "Permissions","ApprovalType","Level","ApprovalUser/Title","ApprovalUser/Id","ID").expand("ApprovalUser")
            .filter(`CurrentUser eq '${currentUserEmailRef.current}' and SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName} '`)();

            const groupedByLevel: { [key: number]: { id: number; selectionType:"All" | "One"; approvedUserList: any[] } } = {};

            LibraryApproverDdetails.forEach(async(item)=>{
              const level = item.Level; 
              let approvalType: "All" | "One" = item.ApprovalType ? "All" : "One";

              // Check if the level already exists in the groupedByLevel object
              if (!groupedByLevel[level]) {
                // If not, initialize an object for this level
                
                groupedByLevel[level] = {
                  id: level-1,                
                  selectionType:approvalType,        
                  approvedUserList: []
                };
              }

              const approvalUserDetails={
                email:item.CurrentUser,
                label:item.ApprovalUser.Title,
                value:item.ApprovalUser.Title, 
                userId:item.ApprovalUser.Id
              }
              groupedByLevel[level].approvedUserList.push(approvalUserDetails);
            })

            // const levelArray = Object.values(groupedByLevel);
            const levelArray = Object.keys(groupedByLevel).map(key => groupedByLevel[parseInt(key)]);

            console.log("levelArray",levelArray);
            console.log("groupedByLevel",groupedByLevel);
            console.log("Library Details",LibraryApproverDdetails);
            setRows(levelArray);
    } catch (error) {
      console.log("Error from setting default value",error)
    }
    
  }


    const [rows, setRows] = React.useState<
                 { id: number; selectionType: "All" | "One"; approvedUserList: string[] }[]
          >([{ id: 0, selectionType: "One", approvedUserList: [] }]);

          console.log("Rows",rows);

     // erroe for user selection
  const [errorsForUserSelection,setErrorsForUserSelection]=useState<{ [key: number]: { userSelect?: string} }>({});
  

  const validateUsersSelect = () => {
    let isValid = true;
    const newErrors: { [key: number]: { userSelect?: string} } = {};

    rows.forEach((row) => {
      if (row.approvedUserList.length === 0) {
        newErrors[row.id] = {userSelect: 'Please select at least one user.' };
        isValid = false;
      }
    });

    setErrorsForUserSelection(newErrors);
    return isValid;
  };


    React.useEffect(() => {
        getcurrentuseremail();
        console.log(currentUserEmailRef.current ,"my current id")
        const fetchUsers = async () => {
          try {
            // start

            const siteContext = await sp.site.openWebById(OthProps.SiteID);
            const user0 = await siteContext.web.siteUsers();

            const combineUsersArray=user0.map((user)=>(
                  {
                    userId:user.Id,
                    value: user.Title,
                    label: user.Title,
                    email: user.Email,
                  }
            ))
            setUsers(combineUsersArray);
            console.log("Sub site users",combineUsersArray);
            // const user0 = await sp.web.siteUsers();
            // const [
            //   users,
            //   users1,
            //   users2,
            //   users3,
            //   users4,
            // ] = await Promise.all([
            //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Read`).users(),
            //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Initiator`).users(),
            //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Contribute`).users(),
            //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Admin`).users(),
            //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_View`).users(),
            // ]);
            // console.log(users, "users ", users1,users2,users3,users4);
            // const combineArray = [
            //   ...(users || []),
            //   ...(users1 || []),
            //   ...(users2 || []),
            //   ...(users3 || []),
            //   ...(users4 || []),
            // ];
            // setUsers(
            //   combineArray.map((user) => ( 
            //   {
            //     userId:user.Id,
            //     value: user.Title,
            //     label: user.Title,
            //     email: user.Email,
            //   }
            // ))
            // );
            // console.log("combineArray", combineArray);
            // end
          } catch (error) {
            console.error("Error fetching site users:", error);
          }
        };
    
        fetchUsers();
      }, []);



    const handleUserSelect = (selected: any, id: any) => {
        // console.log(selected, "selected ");
        const newRows = rows.map((row) =>
          row.id === id ? { ...row, approvedUserList: selected } : row
        );
        console.log("Selected items", selected, id);
        // console.log(rows.length);
        setRows(newRows);
      };



    const handleAddRow = (
        event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
      ) => {
        event.preventDefault();
        const newId = rows.length ? rows[rows.length - 1].id + 1 : 0;
        // setRows([...rows, { id: newId, approvedUser: "", searchTerm: "", filteredUsers: [] }]);
    
        // start
        setRows([
          ...rows,
          { id: newId, selectionType: "One", approvedUserList: [] },
        ]);
        //end
      };

    const handleRemoveRow = (
        id: number,
        event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
      ) => {
        event.preventDefault();
        setRows(rows.filter((row) => row.id !== id));
    };

  const handleSelectionModeChange = (id: number, type: "All" | "One") => {
    const newRows = rows.map((row) =>
      row.id === id ? { ...row, selectionType: type } : row
    );
    setRows(newRows);
  };

const handleCreate = async(e: any) => {
    e.preventDefault();
    if(!validateUsersSelect()){
          console.log("User errors checks called");
          return;
    }

    try {

          // update the column IsApproval if its false.
          if(!IsUpdate){
              // Update list item by ID
              await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.getById(IsApprovalColumnId).update({
                IsApproval:true
              });
              console.log("Item Updated in DMSPreviewFormMaster");
         
              
          }
          const LibraryApproverDdetails = await sp.web.lists
          .getByTitle("DMSFolderPermissionMaster")
          .items.select("CurrentUser" , "SiteName" , "DocumentLibraryName" , "Permissions","ApprovalType","Level","ApprovalUser/Title","ApprovalUser/Id","ID").expand("ApprovalUser")
          .filter(`CurrentUser eq '${currentUserEmailRef.current}' and SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName} '`)();

          console.log("LibraryApproverDdetails",LibraryApproverDdetails);

          LibraryApproverDdetails.forEach(async(item)=>{

                try {
                  const itemId = item.Id; 
                  await sp.web.lists.getByTitle("DMSFolderPermissionMaster").items.getById(itemId).delete();
                  console.log(`Deleted item with ID: ${itemId}`);
                } catch (error) {
                  console.log("Error in deleting the data in the DMSFolderPermissionMaster",error)
                }
                
          })


          console.log("Approved User list",rows);

          rows.forEach((row)=>{

            let payloadForFolderPermissionMaster={
              SiteName:OthProps.SiteTitle,
              DocumentLibraryName:OthProps.DocumentLibraryName,
              CurrentUser:currentUserEmailRef.current,
            }
    
            row.approvedUserList.forEach(async(user:any)=>{
                    console.log("user",user.value);
                    console.log("userID",user.userId);
                    console.log("id",row.id);
    
                    
                    if(row.selectionType === "All"){
                      (payloadForFolderPermissionMaster as any).ApprovalType=1;
                    }else if(row.selectionType === "One"){
                      (payloadForFolderPermissionMaster as any).ApprovalType=0;
                    };
    
    
                    (payloadForFolderPermissionMaster as any).ApprovalUserId=user.userId;
    
                    (payloadForFolderPermissionMaster as any).Level=row.id + 1;
                    console.log("payloadForFolderPermissionMaster",payloadForFolderPermissionMaster);
    
                    // Add the payload DMSFolderPermissionMaster
                    try {
                      const addedItem = await sp.web.lists.getByTitle("DMSFolderPermissionMaster").items.add(payloadForFolderPermissionMaster);
                      console.log("Item added successfully in the payloadForFolderPermissionMaster", addedItem);
                    } catch (error) {
                      console.log("Error adding items to DMSFolderPermissionMaster",error);
                    }
                   
            })

        })

        try {
          const { web } = await sp.site.openWebById(`${OthProps.SiteID}`);
        // Fetch all the groups in the subsite
        interface IMember {
          PrincipalType: number;
          Title:String;
          Id:number 
        }
        interface IRoleAssignmentInfo {
          Member?: IMember; 
        }
        const groups:IRoleAssignmentInfo[] = await web.roleAssignments.expand("Member")();
        console.log("groups",groups);
        const filteredMembers=groups.filter(roleAssignment => {
          return roleAssignment.Member.PrincipalType === 8;
        });
     
        const filteredObject = filteredMembers.filter(item => item.Member.Title === `${OthProps.SiteTitle}_Approval`);

        console.log("filteredObject",filteredObject);
        const roleDefinition = await web.roleDefinitions.getByName("Edit")();
        const roleDefinitionId = roleDefinition.Id;
        const principalId = filteredObject[0].Member.Id;
        console.log("Approval group added successfully")

        const libraryNestedData=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}'`)();
        console.log("documentNestedData",libraryNestedData);
        if(libraryNestedData.length > 0){
          for(let item of libraryNestedData){
            try {
              let securableObject: any;
              if(item.IsLibrary === true){
                securableObject =await web.lists.getByTitle(`${item.DocumentLibraryName}`);
                console.log("securableObject",securableObject);
                // Break inheritance if needed (optional)
                const hasUniquePermissions = await securableObject.hasUniqueRoleAssignments;
                if (!hasUniquePermissions) {
                    await securableObject.breakRoleInheritance(true); 
                }
                await securableObject.roleAssignments.add(principalId, roleDefinitionId);
              }else if(item.IsFolder === true){
                const folder =await web.getFolderByServerRelativePath(`${item.FolderPath}`).getItem();
                securableObject=folder;
                const itemData = await folder.select("HasUniqueRoleAssignments")();
                const breaKRole=itemData.HasUniqueRoleAssignments;
                if (!breaKRole) {
                  await folder.breakRoleInheritance(true);
                  console.log("Inheritance broken, retaining previous permissions.");
                }
                await securableObject.roleAssignments.add(principalId, roleDefinitionId);
              }
            } catch (error) {
              console.log(`Error in adding Approvals group `,error)
            }
          }
        }
        } catch (error) {
          console.log("Error in adding Approval group",error)
        }
        Swal.fire('Added','Users Added Successfully','success');
    } catch (error) {
        console.log("Erroe in LibraryApproverDdetails",error);
    }

                  
    // Clear form on successful submission
    clearForm();
}


  const clearForm = () => {
    console.log("Clear Form called");
    // setRows([{ id: 0, selectionType: "One", approvedUserList: [] }])

  };
  
  return (

    <div className="container mt-0 second">
        <div className="modal show d-block" tabIndex={-1}>
          <div className="modal-dialog">
            <div className="modal-content" style={{
                width:"160%",
                padding:"0px"
                
                }}>
                  {toggleApprover === "Yes"  ?
                  (<div>
                      <div className="" style={{ height: "auto", width: "100%" }}>
                      <div  className='row'>
                      <div className='col-sm-8 w90'>
                                        <h5 className="mb-1 Permissionsectionstyle">
                                          Approval Hierarchy
                                        </h5>
                                        <p className="subheadernew font-14">
                                          Define approval hierarchy for the documents submitted by Team
                                          members in this folder.
                                        </p>

                                        </div>
                                        <div className='col-sm-4 w20'>
                                      
                                            <div className="mb-0">
                                          <div style={{height:'20px'}} className="col-12 d-flex justify-content-end">
                                            <a onClick={handleAddRow}>
                                              <img className="bi bi-plus" src={require("../assets/plus.png")} alt="add" style={{ width: "50px", position:'relative', height: "50px", top:'0' }} />
                                            </a>
                                          </div>
                                        </div>
                                            </div>

                                            <div style={{borderBottom:'1px solid #ccc', marginBottom:'15px', height:'15px', float:'left', width:'100%',  paddingBottom:'10px'}}></div>
                                          </div>
                                       
                                        <div className="row mb-1 approvalheirarcystyle">
                                        <table className="mtbalenew mtbalenewn createc">
    <thead>
      <tr>
        <th  style={{minWidth:'60px',maxWidth:'60px'}}> Level</th>
        <th> Approver</th>
        <th style={{minWidth:'80px',maxWidth:'80px'}}> &nbsp;</th>
        <th style={{minWidth:'40px',maxWidth:'40px'}}> Action</th>
      </tr>
    </thead>
    <tbody>
                                          {/* <div className="col-12 col-md-4">
                                            <label htmlFor="level" className="form-label approvalhierarcyfont">
                                              Level
                                            </label>
                                          </div> */}
                                          {/* <div className="col-12 col-md-6">
                                            <label htmlFor="approver" className="form-label approvalhierarcyfont">
                                              Approver
                                            </label>
                                          </div> */}
                                       
                                        {rows.map((row) => (
                                          <tr className="approvalheirarchyfield" key={row.id}>
                                            <td  style={{minWidth:'60px',maxWidth:'60px'}}>
                                              <input type="text" style={{height:'36px'}} className="form-control" id={`level-${row.id}`} value={`Level ${row.id + 1}`} disabled />
                                            </td>
                                            <td>
                                              <Select
                                                value={row.approvedUserList}
                                                isMulti
                                                options={users}
                                                onChange={(selected: any) => handleUserSelect(selected, row.id)}
                                                placeholder="Enter names or email addresses..."
                                                noOptionsMessage={() => "No User Found..."}
                                              />
                                              {errorsForUserSelection[row.id]?.userSelect && (
                                                <span className="text-danger">{errorsForUserSelection[row.id].userSelect}</span>
                                              )}
                                            </td>
                                            <td style={{minWidth:'80px',maxWidth:'80px', textAlign:'center'}}>
                                            <div style={{gap:'10px', justifyContent:'center'}} className="d-flex">
                                              <div className="form-check">
                                                <input
                                                  className="form-check-input"
                                                  type="radio"
                                                  name={`selection-${row.id}`}
                                                  id={`all-${row.id}`}
                                                  value="all"
                                                  checked={row.selectionType === "All"}
                                                  onChange={() => handleSelectionModeChange(row.id, "All")}
                                                />
                                                <label className="form-check-label" htmlFor={`all-${row.id}`}>
                                                  All
                                                </label>
                                              </div>
                                              <div className="form-check">
                                                <input
                                                  className="form-check-input"
                                                  type="radio"
                                                  name={`selection-${row.id}`}
                                                  id={`one-${row.id}`}
                                                  value="one"
                                                  checked={row.selectionType === "One"}
                                                  onChange={() => handleSelectionModeChange(row.id, "One")}
                                                />
                                                <label className="form-check-label" htmlFor={`one-${row.id}`}>
                                                  One
                                                </label>
                                              </div>
                                            </div>
                                            </td>
                                            <td style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}>
                                            {row.id === 0 ? null : (
                                              <div style={{textAlign:'center', justifyContent:'center'}} className="d-flex align-items-end">
                                                <a onClick={(e) => handleRemoveRow(row.id, e)} style={{ width: "50px", height: "50px", cursor: "pointer" }}>
                                                  <img className="fas fa-trash" src={require("../assets/del.png")} alt="delete" />
                                                </a>
                                              </div>
                                            )}

                                            </td>
                                           
                                          </tr>
                                        ))}
                                        </tbody>
                                        </table>
                                      </div>  </div>
                                      <div className="modal-footer">
                                      <button type="button" className="btn btn-primary" 
                                      onClick={handleCreate}
                                      >
                                        Create
                                      </button>
                                      <button type="button" className="btn btn-secondary" 
                                      //   onClick={toggleModal}
                                      onClick={onReturnToMain}
                                      >
                                          Cancel{" "}
                                      </button>
                                    </div> 
                  </div>) : 
                  null
                  }
                  {toggleApprover === "No"  ? (<div>
                       <h6>"This document library does not have an approver. Would you like to set one?"</h6>
                        <div className='setApprover' style={{
                          display:"flex",
                          marginLeft:"280px",
                          gap:"10px"
                        }}>
                        <button type="button" className="btn btn-primary" onClick={handleSetNewApprover}>Set Approver</button>
                        <button type="button" className="btn btn-secondary" 
                          onClick={onReturnToMain}
                        >
                          Cancel{" "}
                        </button>
                      </div>
                  </div>): null}

            </div>
          </div>
        </div>
    </div>
  
  )
}

export default ManageWorkFlow
import * as React from "react";
// import Provider from '../../../GlobalContext/provider';
import "bootstrap/dist/css/bootstrap.min.css";
// import './SideBar';
import { useRef, useState } from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/site-groups"
import "@pnp/sp/folders"; 
import "@pnp/sp/webs"; 
import "./CreateFoldercss";
import Select from "react-select";
import Swal from "sweetalert2";
// import Form from "react-bootstrap/Form";

// let selectedArrayForUserPermission:{
//   userId:number,
//   value: String,
//   label: String,
//   email:String
// }[];

// let selectedPermissionValue:String;

interface CreateFolderProps {
  OthProps: { [key: string]: string };
  onReturnToMain: () => void;
  // myRequest: (event: React.MouseEvent<HTMLButtonElement>) => void;
}

let togglecolumneDetails=true;
let toggleaddFieldsButton=true;
let togglefolderPrivacy=true;
// let toggleApprovalForFolder=true;
let locationPath=window.location.pathname.match(/\/sites\/[^\/]+/)[0];

const CreateFolder: React.FC<CreateFolderProps> = ({
  OthProps,
  onReturnToMain,
}) => {
  console.log(OthProps, "oth props");
  const sp: SPFI = getSP();

  const siteUrl = window.location.origin;
  // alert(`siteUrl : ${siteUrl}${locationPath}/SitePages/DMSAdmin.aspx`)
  const [toggleApproval, setToggleApproval] = React.useState(false);
  const [approvalOption, setApprovalOption]=useState("");
  console.log("Approval option",approvalOption);

  console.log("Location URL",window.location.pathname.match(/\/sites\/[^\/]+/)[0]);


  
  // new code for permission.
  // const [permission, setPermission]=React.useState(false);

  // const handlePermissionToggle=(set:any)=>{
  //   console.log("Set permission called");
  //   console.log(set);
  //   setPermission(set);
  // }

  const permissionArray:{value:string,label:string}[]=[
    {value:"Full Control",label:"Full Control"},
    {value:"Contribute",label:"Contribute"},
    {value:"Edit",label:"Edit"},
    {value:"Read",label:"Read"},
    {value:"View",label:"View"}
  ];
  //new code end 
// New Code for Adding permission dynamically on click of + button.
  const [rowsForPermission, setRowsForPermission] = React.useState<
        { id: number; selectedUserForPermission: string[]; selectedPermission:String }[]
        >([{ id: 0, selectedUserForPermission: [],selectedPermission:"" }]);
  console.log("rowsForPermission",rowsForPermission);
  // Add new row for permission
  const handleAddRowForPermission=(event: React.MouseEvent<HTMLAnchorElement, MouseEvent>)=>{
    event.preventDefault();
    const newId = rowsForPermission.length ? rowsForPermission[rowsForPermission.length - 1].id + 1 : 0;
    setRowsForPermission([
      ...rowsForPermission,
      { id: newId, selectedUserForPermission: [], selectedPermission:"" },
    ]);
  }

  // Remove new row for permission
  const handleRemoveRowForPermission = (
    id: number,
    event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
  ) => {
    event.preventDefault();
    setRowsForPermission(rowsForPermission.filter((row) => row.id !== id));
  };
  
  const handleUserSelectForPermission=(selectedUser:any, id:any)=>{
    // console.log("selectedArrayForUserPermission",selectedArrayForUserPermission)
    console.log("Selected user for permission",selectedUser);
    // selectedArrayForUserPermission=selectedUser;
    const newRows = rowsForPermission.map((row) =>
      row.id === id ? { ...row, selectedUserForPermission: selectedUser } : row
    );
    setRowsForPermission(newRows);
  }

  const handlePermissionSelect=(selectedPermission:any,id:any)=>{
    // console.log("Before selectedPermissionValue",selectedPermissionValue)
    console.log("Selected Permission",selectedPermission)
    // selectedPermissionValue=selectedPermission.value;
    const newRows = rowsForPermission.map((row) =>
      row.id === id ? { ...row, selectedPermission: selectedPermission.value } : row
    );
    setRowsForPermission(newRows);
  }

  // new code end

  // Toggle the Folder Privacy and Column details
  if(OthProps.DocumentLibrary !==""){
    togglecolumneDetails=false;
    toggleaddFieldsButton=false;
    togglefolderPrivacy=false;
    // toggleApprovalForFolder=false;

  }else{
    togglecolumneDetails=true;
    toggleaddFieldsButton=true;
    togglefolderPrivacy=true;
    // toggleApprovalForFolder=true;
  }


  const currentUserEmailRef = useRef('');


  const getcurrentuseremail = async()=>{
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    const groups = await sp.web.siteGroups();

    console.log("All Groups:", groups);
    const allusers = await sp.web.siteUsers();

    console.log("All allusers:", allusers);

    // const groups = await sp.web.siteGroups();

    // console.log("All Groups:", groups);

    // Get all site users
    const users = await sp.web.siteUsers();

    console.log("All Users:", users);
   }

  const [users, setUsers] = React.useState<any[]>([]);
  console.log("Users Array", users);

  const handleToggleApproval = (event:any) => {
    // event.preventDefault();
    setApprovalOption(event.target.value);
    setToggleApproval(true);
  };

  const handleToggleRemove = (event:any) => {
    setApprovalOption(event.target.value);
    setToggleApproval(false);
    
  };
  const [rows, setRows] = React.useState<
    { id: number; selectionType: "All" | "One"; approvedUserList: string[] }[]
  >([{ id: 0, selectionType: "One", approvedUserList: [] }]);
  // end

//   Errors for field selection
  const [errors1, setErrors1] = useState<{ [key: number]: { fieldName?: string; selectField?: string } }>({});

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

  //start
//   store the form field and its type.
  const [formFields, setFormFields] = useState([
    { id:0, fieldName: '', selectField: ''}
  ]);

//   add field in the formField arry
  const handleInputChange = (id:number, event:any) => {
    const values = formFields.map(field =>
        field.id === id
          ? { ...field, fieldName: event.target.value } 
          : field
      );
      setFormFields(values);

        // Reset error when user enters a value
    if (event.target.value.trim() !== '') {
        setErrors1((prevErrors) => ({
          ...prevErrors,
          [id]: { ...prevErrors[id], [event.target.name]: '' }
        }));
      }
  };

//   add type in the formField array
  const handleSelectedType=(id:number,event:any)=>{

    const values = formFields.map(field =>
        field.id === id
          ? { ...field, selectField: event.target.value } 
          : field
      );
      setFormFields(values);

      // Reset error when user selects a value
    if (event.target.value !== '') {
        setErrors1((prevErrors) => ({
          ...prevErrors,
          [id]: { ...prevErrors[id], selectField: '' }
        }));
      }
  }
//   add new field row
  const handleAddFields = () => {
    const newId = formFields.length ? formFields[formFields.length - 1].id + 1 : 0;
    setFormFields([
      ...formFields,
      { id: newId, fieldName: "", selectField:"" },
    ]);
  };
  console.log("FormsField Array",formFields);

//   remove field row
  const handleRemoveField=(id:number,event:any)=>{
    event.preventDefault();
    // console.log("index",id);
    // console.log("Remove Field Called");

    setFormFields(formFields.filter((field) => field.id !== id));
    
  }

// Handle validation and error state update
const validateFields = () => {
   
  let isValid = true;
  const newErrors: { [key: number]: { fieldName?: string; selectField?: string } } = {};

  formFields.forEach((field) => {
    const nonAlphaNumericForEntity = field.fieldName.replace(/[^a-zA-Z0-9 -]/g, '');
    // if (!field.fieldName.trim()) {
    //   newErrors[field.id] = { ...newErrors[field.id], fieldName: 'Field Name is required' };
    //   isValid = false;
    // }
    // if (!field.selectField) {
    //   newErrors[field.id] = { ...newErrors[field.id], selectField: 'Field Type is required' };
    //   isValid = false;
    // }

     // Check if field name has been filled and if select field is empty
    if (field.fieldName.trim() && !field.selectField) {
      newErrors[field.id] = { ...newErrors[field.id], selectField: 'Field type is required.'};
      isValid = false;
    }

    if(field.fieldName !== nonAlphaNumericForEntity){
        newErrors[field.id] = { ...newErrors[field.id], fieldName: 'Special characters are not allowed.' };
        isValid = false;
    }

      // Check if field type is selected but field name is empty
      if (!field.fieldName.trim() && field.selectField) {
        newErrors[field.id] = { ...newErrors[field.id], fieldName: 'Field name is required.' };
        isValid = false;
      }
      
         if(field.fieldName.trim() === folderName.trim()){
        newErrors[field.id] = { ...newErrors[field.id], fieldName: 'Meta tags and folder names should not be the same.' };
        isValid = false;
    }
  });

  setErrors1(newErrors);
  return isValid;
};
  const [siteUsers,setSiteUsers]=React.useState<any[]>([]);
  console.log("siteUsers --> ",siteUsers)
  React.useEffect(()=>{

    const fetchUserFromSitLevel=async()=>{
        // start
        // const siteContext = await sp.site.openWebById(OthProps.siteID);
        // const user0 = await siteContext.web.siteUsers();

        // const combineUsersArray=user0.map((user)=>(
        //       {
        //       userId:user.Id,
        //       value: user.Title,
        //       label: user.Title,
        //       email: user.Email,
        //   }
        // ))
        // console.log("Site Users",combineUsersArray);
        // setSiteUsers(combineUsersArray);

        if(OthProps.IsExternal === 'true'){
          const user0 = await sp.web.siteUsers();
          const user1 = await sp.web.siteGroups();
          const groupsArray=user1.map((user)=>(
            {
            PrincipalType:user.PrincipalType,
            userId:user.Id,
            value: user.Title,
            label: user.Title,
            email: user.Title,
            }
          ))
          const combineUsersArray=user0.map((user)=>(
                {
                userId:user.Id,
                value: user.Title,
                label: user.Title,
                email: user.Email,
            }
          ))
          let resultArray =[...combineUsersArray, ...groupsArray];
          console.log("resultArray --->",resultArray)
          setSiteUsers(resultArray);
        }else{
          // fetch the data from site Gropus
          const [
            users,
            users1,
            users2,
            users3,
            users4,
            users5,
            users6,
            users7
          ] = await Promise.all([
            sp.web.siteGroups.getByName(`${OthProps.Entity}_Read`).users(),
            sp.web.siteGroups.getByName(`${OthProps.Entity}_Initiator`).users(),
            sp.web.siteGroups.getByName(`${OthProps.Entity}_Contribute`).users(),
            sp.web.siteGroups.getByName(`${OthProps.Entity}_Admin`).users(),
            sp.web.siteGroups.getByName(`${OthProps.Entity}_View`).users(),
            sp.web.siteGroups.getByName(`${OthProps.Entity}_AllUsers`).users(),
            sp.web.siteGroups.getByName(`${OthProps.Entity}_Approval`).users(),
            sp.web.siteGroups.getByName(`DMSSuper_Admin`).users(),
          ]);
  
          const combineArray = [
            ...(users || []),
            ...(users1 || []),
            ...(users2 || []),
            ...(users3 || []),
            ...(users4 || []),
            ...(users5 || []),
            ...(users6 || []),
            ...(users7 || []),
          ];
          setSiteUsers(
            combineArray.map((user) => ( 
            {
              userId:user.Id,
              value: user.Title,
              label: user.Title,
              email: user.Email,
            }
          ))
          );
          console.log("combineArray", combineArray);
        }
        // // fetch the data from site Gropus
        //   const [
        //   users,
        //   users1,
        //   users2,
        //   users3,
        //   users4,
        // ] = await Promise.all([
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Read`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Initiator`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Contribute`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Admin`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_View`).users(),
        // ]);

        // const combineArray = [
        //   ...(users || []),
        //   ...(users1 || []),
        //   ...(users2 || []),
        //   ...(users3 || []),
        //   ...(users4 || []),
        // ];
        // setSiteUsers(
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

    }
    fetchUserFromSitLevel();
  },[]);

  //end   

  // Fetch users from SharePoint
  React.useEffect(() => {
    getcurrentuseremail();
    createBreadCrumb()
    console.log(currentUserEmailRef.current ,"my current id")
    const fetchUsers = async () => {
      try {
        // start
        const siteContext = await sp.site.openWebById(OthProps.siteID);
        // const user0 = await siteContext.web.siteUsers();
        const approvalGroupUsers=await siteContext.web.siteGroups.getByName(`${OthProps.Entity}_Approval`).users();
        // console.log("approvalGroupUsers",approvalGroupUsers);
        const finalArray=approvalGroupUsers.map((user)=>(
                {
                    userId:user.Id,
                    value: user.Title,
                    label: user.Title,
                    email: user.Email,
                }
        ));
        console.log("finalArray",finalArray);
        setUsers(finalArray);
        // const combineUsersArray=user0.map((user)=>(
        //       {
        //       userId:user.Id,
        //       value: user.Title,
        //       label: user.Title,
        //       email: user.Email,
        //   }
        // ))
        // setUsers(combineUsersArray);
        // console.log("Sub site users",combineUsersArray);

        // const [
        //   users,
        //   users1,
        //   users2,
        //   users3,
        //   users4,
        // ] = await Promise.all([
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Read`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Initiator`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Contribute`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_Admin`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.Entity}_View`).users(),
        // ]);
          
        // const users=await sp.web.siteGroups.getByName(`${OthProps.Entity}_Read`).users();
        // const users1=await sp.web.siteGroups.getByName(`${OthProps.Entity}_Initiator`).users();
        // const users2 =await sp.web.siteGroups.getByName(`${OthProps.Entity}_Contribute`).users();
        // const users3=await sp.web.siteGroups.getByName(`${OthProps.Entity}_Admin`).users();
        // const users4=await sp.web.siteGroups.getByName(`${OthProps.Entity}_View`).users();
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

  const userOptions = users.map((user: any) => ({
    label: user.Title, // Display name
    value: user.Email, // Value for selection
  }));
  console.log(userOptions, "userOptions");

  console.log("component rendered", rows);

  const handleUserSelect = (selected: any, id: any) => {
    console.log(selectedUsers, "selectedUsers");
    console.log(selectedUsers, "selectedUsers");
    setSelectedUsers(selected || []);
    console.log(selected, "selected ");
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

//   remove new row
  const handleRemoveRow = (
    id: number,
    event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
  ) => {
    event.preventDefault();
    setRows(rows.filter((row) => row.id !== id));
  };

  // start
  const handleSelectionModeChange = (id: number, type: "All" | "One") => {
    const newRows = rows.map((row) =>
      row.id === id ? { ...row, selectionType: type } : row
    );
    setRows(newRows);
  };
  // end

  ///////////////////// form validation //////////////////////////////////////
  type FormErrors = {
    folderName?: string;
    folderPrivacy?: string;
    folderOverview?: string;
    selectedUsers?: any;
    fieldName? : any
    selectField? : any
    approvalOption?:any
  };
  // Define state variables to manage form input
  const [folderName, setFolderName] = useState("");
//   const [fieldName, setFieldName] = useState("");
  const [folderPrivacy, setFolderPrivacy] = useState("");
  console.log("Folder Privacy",folderPrivacy);
  const [folderOverview, setFolderOverview] = useState("");
  const [Approver, setApprover] = useState("");
  const [selectedUsers, setSelectedUsers] = useState<any[]>([]); // Assuming multiple users
//   const [selectField, setSelectField] = useState(""); // For dropdown selection
  const [errors, setErrors] = useState<FormErrors>({}); // For validation errors
  const [errorsForPermissionSelection, setErrorsForPermissionSelection] = useState<{ [key: number]: { userSelect?: string, permissionSelect?: string } }>({});
  const validatePermissionsSelect = () => {
    let isValid = true;
    const newErrors: { [key: number]: { userSelect?: string, permissionSelect?: string } } = {};
  
    rowsForPermission.forEach((row) => {
      if (!row.selectedUserForPermission || row.selectedUserForPermission.length === 0) {
        newErrors[row.id] = { ...newErrors[row.id], userSelect: 'Please select at least one user.' };
        isValid = false;
      }
      if (!row.selectedPermission) {
        newErrors[row.id] = { ...newErrors[row.id], permissionSelect: 'Please select a permission.' };
        isValid = false;
      }
    });
  
    setErrorsForPermissionSelection(newErrors);
    return isValid;
  };
  
  // select the delete option
  // const [deleteOption, setDeleteOption]=useState("");

  // const handleDeleteOption=(event:any)=>{
  //     event.preventDefault();
  //     console.log("Target value of delete option", event.target.value);
  //     setDeleteOption(event.target.value);
  // }

  // Handle form submission (Create button click)
  // const checkDuplicateFolderNameValidation=async()=>{
  //   let isValid = true;
  //   if(OthProps.DocumentLibrary !== ""){
  //     // alert('check for folder');
  //     const getDMSFolderMasterData=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${OthProps.Entity}' and DocumentLibraryName eq '${OthProps.DocumentLibrary}'`)();
  //     console.log("getDMSFolderMasterData",getDMSFolderMasterData);
  //     if(getDMSFolderMasterData.length > 0){
  //       for(const item of getDMSFolderMasterData){
  //         if(item?.FolderName === folderName.trim()){
  //           isValid=false;
  //         }
  //       }
  //     }
  //   }else if(OthProps.DocumentLibrary === ""){
  //     const getDMSFolderMasterData=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${OthProps.Entity}'`)();
  //     console.log("getDMSFolderMasterData",getDMSFolderMasterData);
  //     if(getDMSFolderMasterData.length > 0){
  //       for(const item of getDMSFolderMasterData){
  //         if(item?.DocumentLibraryName === folderName.trim()){
  //           isValid=false;
  //         }
  //       }
  //     }
  //   }
  //   return isValid;
  // }
  const checkFolderNameValidation=async ()=>{
    let isValid=true;
    if(OthProps.DocumentLibrary === ""){
        const getFolderData=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${OthProps.Entity}' and DocumentLibraryName eq '${folderName.trim()}' and FolderPath eq '${locationPath}/${OthProps.Entity}/${folderName.trim()}'`)();

        if(getFolderData.length > 0 ){
          isValid=false;
        }
    }else if(OthProps.DocumentLibrary !== ""){

      const getFolderData=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${OthProps.Entity}' and DocumentLibraryName eq '${OthProps.DocumentLibrary}' and FolderPath eq '${OthProps.folderpath}/${folderName.trim()}'`)();

      if(getFolderData.length > 0 ){
        isValid=false;
      }

    }

    return isValid;
  }
  const handleCreate = async(e: any) => {
    e.preventDefault();

    let validateColumns=false;
    let validateUser=false;
    let formFieldValidation=false;
    let validatePermissionAndUser=false;
    // console.log("Handcreate called");
    const nonAlphaNumericForEntity = folderName.replace(/[^a-zA-Z0-9 -]/g, '');
    // Validate the form
    let validationErrors: FormErrors = {};

    if(OthProps.DocumentLibrary !== ""){
      console.log("create Folder");
      if (!folderName.trim()) {
        validationErrors.folderName = "Folder Name is required.";
      }
      if(folderName !== nonAlphaNumericForEntity){
        validationErrors.folderName = "Special charaters are not allowed.";
      }
      if(nonAlphaNumericForEntity.length > 50){
        validationErrors.folderName = "Input cannot exceed 50 characters in the folder name field.";
      }
      if (!folderOverview.trim()) {
        validationErrors.folderOverview = "Folder Overview is required.";
      }
      if(!validatePermissionsSelect() && showDiv){
        validatePermissionAndUser=true;
      }
      // if(!await checkDuplicateFolderNameValidation()){
      //   validationErrors.folderName = "Folder name already exist.";
      // }
      if(!await checkFolderNameValidation()){
        validationErrors.folderName = "Folder already exists. Please change the folder name."
     }
    }else{
      console.log("create document library");
      if (!folderName.trim()) {
        validationErrors.folderName = "Folder Name is required.";
      }
      if(folderName !== nonAlphaNumericForEntity){
        validationErrors.folderName = "Special charaters are not allowed.";
      }
      if(nonAlphaNumericForEntity.length > 50){
        validationErrors.folderName = "Input cannot exceed 50 characters in the folder name field.";
      }
      if(!approvalOption.trim()){
        validationErrors.approvalOption = "Approval Option is required.";
      }
      if (!folderPrivacy) {
        validationErrors.folderPrivacy = "Please select folder privacy.";
      }
      if (!folderOverview.trim()) {
        validationErrors.folderOverview = "Folder Overview is required.";
      }
      if(!validateUsersSelect() && toggleApproval){
        console.log("User errors checks called");
        validateUser=true;
      }
      if(!validateFields()){
          // console.log("select the fiels or type");
          validateColumns=true
      }

      if(!validatePermissionsSelect() && showDiv){
        validatePermissionAndUser=true;
      }
      // if(!await checkDuplicateFolderNameValidation()){
      //   validationErrors.folderName = "Folder name already exist.";
      // }

      if(!await checkFolderNameValidation()){
         validationErrors.folderName = "Folder already exists. Please change the folder name."
      }
    }
    
    // Validation for forbidden column names
    const forbiddenNames = ["status", "isdeleted"];
    const invalidFields = formFields.filter((field) => forbiddenNames.includes(field.fieldName.trim().toLowerCase()));
    if (invalidFields.length > 0) {
      formFieldValidation=true;
          // return;
    }
    // If errors exist, set them to the state and prevent submission
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors);
    }
    else if(validateColumns){
        // alert("Add Columns Fields and Type");
    }
    else if(validateUser){
        // alert("Please select at least one user");
    }else if(formFieldValidation){
      Swal.fire(
        'Validation Error',
        `The column names "${invalidFields.map(f => f.fieldName).join(', ')}" are not allowed. Please choose different names.`,
        'error'
      );
    }else if(validatePermissionAndUser){

    }
    else {
      const createFolderButton=document.getElementById('CreateFolderInsideSharePoint') as HTMLButtonElement;
      createFolderButton.disabled=true;
      const payloadForFolderMaster={
        SiteTitle:OthProps.Entity,
        CurrentUser:currentUserEmailRef.current
      }

      if(OthProps.DocumentLibrary === ""){
        (payloadForFolderMaster as any).DocumentLibraryName=folderName.trim();
        (payloadForFolderMaster as any).FolderPath=`${locationPath}/${OthProps.Entity}/${folderName.trim()}`;
        //  (payloadForFolderMaster as any).FolderPath=`/sites/ED/${OthProps.Entity}/${folderName}`;
        (payloadForFolderMaster as any).IsLibrary=true;
        (payloadForFolderMaster as any).IsActive=false;
        if(folderPrivacy === "private"){
          (payloadForFolderMaster as any).IsPrivate=true;
        }else if(folderPrivacy === "public"){
          (payloadForFolderMaster as any).IsPrivate=false;
        }
        if(OthProps.IsFolderDeligationUser === "true"){
          (payloadForFolderMaster as any).IsFolderDeligation=true;
        }
        if(OthProps.IsExternal === 'true'){
          (payloadForFolderMaster as any).External=true;
        }
      }else{
        (payloadForFolderMaster as any).DocumentLibraryName=OthProps.DocumentLibrary;
        (payloadForFolderMaster as any).FolderPath=`${OthProps.folderpath}/${folderName.trim()}`;
        (payloadForFolderMaster as any).IsFolder=true;
        if(OthProps.IsFolderDeligationUser === "true"){
          (payloadForFolderMaster as any).IsActive=false;
        }else if(OthProps.IsFolderDeligationUser === "false"){
          (payloadForFolderMaster as any).IsActive=true;
        }

        if(OthProps.Folder ===  ""){
            (payloadForFolderMaster as any).FolderName=folderName.trim();
        }else{
            (payloadForFolderMaster as any).FolderName=folderName.trim();
            (payloadForFolderMaster as any).ParentFolderId=OthProps.Folder;

            const parentIdData=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`FolderPath eq '${OthProps.folderpath}'`)();
            console.log("parentIdData",parentIdData);

            (payloadForFolderMaster as any).ParentID=parentIdData[0].ID;
            
        }
        if(folderPrivacy === "private"){
          (payloadForFolderMaster as any).IsPrivate=true;
        }else if(folderPrivacy === "public"){
          (payloadForFolderMaster as any).IsPrivate=false;
        }

        if(OthProps.IsFolderDeligationUser === "true"){
          (payloadForFolderMaster as any).IsFolderDeligation=true;
        }
        if(OthProps.IsExternal === 'true'){
          (payloadForFolderMaster as any).External=true;
        }
      }

      if(OthProps.Department !== ""){
        (payloadForFolderMaster as any).Department=OthProps.Department
      }
      if(OthProps.Devision !== ""){
        (payloadForFolderMaster as any).Devision=OthProps.Devision
      }

      console.log("payloadForFolderMaster",payloadForFolderMaster);
      console.log("Approved User list",rows);

      

      const addedItem = await sp.web.lists.getByTitle("DMSFolderMaster").items.add(payloadForFolderMaster);
      console.log("Item added successfully in the DMSFolderMaster", addedItem);


      // new code for Creating Folder inside the document library
      if(OthProps.DocumentLibrary !== ""){

          try {
          
            console.log("Create Folder Inside this Document Library -",OthProps.DocumentLibraryName);
            const {web}=await sp.site.openWebById(OthProps.siteID);
            const folderAddResult = await web.folders.addUsingPath(`${OthProps.folderpath}/${folderName.trim()}`);
            console.log("Folder created successfully -",folderAddResult);

            if(folderPrivacy === "public"){
              const folder =await web.getFolderByServerRelativePath(`${OthProps.folderpath}/${folderName.trim()}`).getItem();
              const itemData = await folder.select("HasUniqueRoleAssignments")();
              const breaKRole=itemData.HasUniqueRoleAssignments;
              if (!breaKRole) {
                await folder.breakRoleInheritance(true);
                console.log("Inheritance broken, retaining previous permissions.");
              }
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
                console.log("groups3",groups);
                const filteredMembers=groups.filter(roleAssignment => {
                  return roleAssignment.Member.PrincipalType === 8;
                });

                const filteredGroups = filteredMembers.map((object) => ({
                    value: object.Member.Title,
                    label: object.Member.Title,
                    Id: object.Member.Id,
                }));
                console.log("filteredGroups",filteredGroups);
                console.log("filteredMembers",filteredMembers);

                const updatedData = filteredGroups.map(item => {
                  let permission = "";
              
                  if (item.value.includes("_Admin")) permission = "Full Control";
                  else if (item.value.includes("_View")) permission = "View";
                  else if (item.value.includes("_Read")) permission = "Read";
                  else if (item.value.includes("_Contribute")) permission = "Contribute";
                  else if (item.value.includes("_Initiator")) permission = "Edit";
                  else if (item.value.includes("_Approval")) permission = "Edit";
                  else if (item.value.includes("_AllUsers")) permission = "Edit";
                  else if (item.value.includes("_FolderDeligation")) permission = "Contribute";
              
                  return { ...item, permission };
              });

              console.log("updatedData",updatedData);
              updatedData.forEach(async(item)=>{
                try {
                  const roleDefinition = await web.roleDefinitions.getByName(item.permission)();
                  const roleDefinitionId = roleDefinition.Id;
                  const principalId =item.Id;
                  await folder.roleAssignments.add(principalId, roleDefinitionId);
                  console.log(`Adding ${item.value} (${principalId}) with ${item.permission} permissions`);
                } catch (error) {
                  console.log("Error Adding groups to the folders",error)
                }
 
              })
            }
          } catch (error) {
            console.log("Error In creating Folder Inside the Document Library",error);
          }
        
      
      }
      // END NEW CODE

      if(OthProps.DocumentLibrary === "" && toggleApproval){

            let payloadForFolderPermissionMaster={
              SiteName:OthProps.Entity,
              DocumentLibraryName:folderName.trim(),
              CurrentUser:currentUserEmailRef.current,
            }

            rows.forEach((row)=>{

              payloadForFolderPermissionMaster={
                SiteName:OthProps.Entity,
                DocumentLibraryName:folderName.trim(),
                CurrentUser:currentUserEmailRef.current,
  
              }

              row.approvedUserList.forEach(async(user:any)=>{
                // (payloadForFolderPermissionMaster as any).ApprovalUser=user.value
                console.log("user",user.value);
                console.log("userID",user.userId);
                console.log("id",row.id);

                
                if(row.selectionType === "All"){
                  (payloadForFolderPermissionMaster as any).ApprovalType=1;
                }else if(row.selectionType === "One"){
                  (payloadForFolderPermissionMaster as any).ApprovalType=0;
                };


                // (payloadForFolderPermissionMaster as any).ApprovalUser={
                //   "__metadata": {"type": "SP.FieldUserValue" },
                //   LookupId: user.userId
                // };

                // const ensureUser=await sp.web.ensureUser(user.email);  
                // console.log("user to update",ensureUser);

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
      }
      

    if(OthProps.DocumentLibrary === ""){

          console.log("Add the Columns when create document library");
          const payloadForPreviewFormMaster={
            SiteName:OthProps.Entity,
            DocumentLibraryName:folderName.trim(),
            IsRequired:true,
            AddorRemoveThisColumn:"Add To Library",
            IsInProgress:true
          }

          // console.log("payloadForPreviewFormMaster",payloadForPreviewFormMaster)
          
          let optionSelectedForPrivacy:boolean;
          if(folderPrivacy === "private"){
            optionSelectedForPrivacy=true;
          }else if(folderPrivacy === "public"){
            optionSelectedForPrivacy=false;
          }
          let optionSelectedForApprovals:boolean;
          if(approvalOption === "Yes"){
            optionSelectedForApprovals=true;
          }else if(approvalOption === "No"){
            optionSelectedForApprovals=false;
          }

          const payload={
            SiteName:OthProps.Entity,
            DocumentLibraryName:folderName.trim(),
            IsDocumentLibrary:true,
            IsPrivate:optionSelectedForPrivacy,
            IsHardDelete:false,
            IsApproval:optionSelectedForApprovals
          }
          console.log("payload for DMSPreviewFormField for IsDocumentLibrary",payload)
          const addedItem = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.add(payload);
          console.log("Item added successfully in the DMSPreviewFormField for IsDocumentLibrary", addedItem);
          
          if(formFields.length > 0){
            // if(formFields[0].fieldName !== '' && formFields[0].selectField !== ''){
              formFields.forEach(async(field)=>{
                // type.replace(/\s+/g, '').toLowerCase();
                if (field.fieldName.trim() !== '') {
                    (payloadForPreviewFormMaster as any).ColumnName=field.fieldName.replace(/\s+/g,'');
                    (payloadForPreviewFormMaster as any).ColumnType=field.selectField
                    console.log("Call the Api with this payload",payloadForPreviewFormMaster)
    
                    const addedItem = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.add(payloadForPreviewFormMaster);
                    console.log("Item added successfully in the DMSPreviewFormField", addedItem);
                }
                    
              })
            // }
          }
          // formFields.forEach(async(field)=>{
          //   // type.replace(/\s+/g, '').toLowerCase();
          //       (payloadForPreviewFormMaster as any).ColumnName=field.fieldName.replace(/\s+/g,'');
          //       (payloadForPreviewFormMaster as any).ColumnType=field.selectField
          //       console.log("Call the Api with this payload",payloadForPreviewFormMaster)

          //       const addedItem = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.add(payloadForPreviewFormMaster);
          //       console.log("Item added successfully in the DMSPreviewFormField", addedItem);
                
          // })
    }

    // new code  creating payload for DMSFolderPrivacy and add the data
      // if(OthProps.DocumentLibrary === "" && permission === true){
        // if(permission === true){

    // Add permission  to all whetehr its document library folder or subfolder
    let Id:any;
    if(OthProps.DocumentLibrary === ""){
      const getIDDetailsOfTheCurrentFolder=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`FolderPath eq '${locationPath}/${OthProps.Entity}/${folderName.trim()}'`)();

      Id=getIDDetailsOfTheCurrentFolder[0].ID;
    }else{
      const getIDDetailsOfTheCurrentFolder=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`FolderPath eq '${OthProps.folderpath}/${folderName.trim()}'`)();

      Id=getIDDetailsOfTheCurrentFolder[0].ID;
    }
          const payloadForDMSFolderPrivacy={
            SiteName:OthProps.Entity,
            CurrentUser:currentUserEmailRef.current,
            IsModified:false,
            FolderID:Id
            // DocumentLibraryName:folderName
          }
          if(OthProps.DocumentLibrary === ""){
            (payloadForDMSFolderPrivacy as any).DocumentLibraryName=folderName.trim();
          }else{
            (payloadForDMSFolderPrivacy as any).DocumentLibraryName=OthProps.DocumentLibrary;
            (payloadForDMSFolderPrivacy as any).FolderName=folderName.trim();
          }

          if(folderPrivacy === "private"){
            (payloadForDMSFolderPrivacy as any).PublicFolderPermission=false;
          }else if(folderPrivacy === "public"){
            (payloadForDMSFolderPrivacy as any).PublicFolderPermission=true;
          }

          console.log("Payload for DMSFolderPrivacy without selected field",payloadForDMSFolderPrivacy);
          const addedItem1 = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacy);
          console.log("Added data to DMSFolderPrivacy without selected field",addedItem1);

          rowsForPermission.forEach((row)=>{
              console.log("row",row.selectedPermission);
              row.selectedUserForPermission.forEach(async(user:any)=>{
                (payloadForDMSFolderPrivacy as any).User=user.value;
                (payloadForDMSFolderPrivacy as any).UserID=user.userId;
                (payloadForDMSFolderPrivacy as any).UserPermission=row.selectedPermission;
                payloadForDMSFolderPrivacy.IsModified=false;
                console.log("Payload for DMSFolderPrivacy after slecetd value",payloadForDMSFolderPrivacy);
                try {
                  const addedItem = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacy);
                  console.log("Item added to the list DMSFolderPrivacy after selected value ",addedItem);
                } catch (error) {
                  console.log("Erroe in adding items in the DMSFolderPrivacy after selected value",error);
                }
              })
          })
          // selectedArrayForUserPermission.forEach(async(user)=>{
          //   (payloadForDMSFolderPrivacy as any).User=user.value;
          //   (payloadForDMSFolderPrivacy as any).UserID=user.userId;
          //   (payloadForDMSFolderPrivacy as any).UserPermission=selectedPermissionValue;
          //   payloadForDMSFolderPrivacy.IsModified=false;

          //   console.log("Payload for DMSFolderPrivacy after slecetd value",payloadForDMSFolderPrivacy);

          //   try {
          //     const addedItem = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacy);
          //     console.log("Item added to the list DMSFolderPrivacy after selected value ",addedItem);
          //   } catch (error) {
          //     console.log("Erroe in adding items in the DMSFolderPrivacy after selected value",error);
          //   }
            
          // })

      // }
    // new code end
    const getUniqueRequestNo = async () => {
      const counterItem = await sp.web.lists.getByTitle('DMSFolderCounterList').items.getById(1)();
      console.log("Counter Item 0", counterItem);
      console.log("Counter Item 1", counterItem.FolderCount);
      let FolderCount = counterItem.FolderCount;
    
      // Increment the counter
      FolderCount++;
    
      // Generate the new RequestNo
      const newRequestNo = `Folder${String(FolderCount).padStart(2, '0')}`;
    
      // Update the counter in the CounterList
      await sp.web.lists.getByTitle('DMSFolderCounterList').items.getById(1).update({
        FolderCount: FolderCount
      });
    
      return newRequestNo;
    };
    const newRequestNo = await getUniqueRequestNo();

      if(OthProps.IsFolderDeligationUser === "true"){
        
        const payloadForFolderDelegation={
          SiteTitle:OthProps.Entity,
          CurrentUser:currentUserEmailRef.current,
          Processname:'New Folder Request',
          RequestNo:newRequestNo,
          Status:'Pending',
          SubmitStatus:'Submitted'
        }
        
        if(OthProps.DocumentLibrary === ""){
          (payloadForFolderDelegation as any).DocumentLibraryName=folderName.trim();
          //  (payloadForFolderDelegation as any).FolderPath=`/sites/IntranetUAT/${OthProps.Entity}/${folderName}`;
           (payloadForFolderDelegation as any).FolderPath=`${locationPath}/${OthProps.Entity}/${folderName.trim()}`;
          (payloadForFolderDelegation as any).IsLibrary=true;
          // (payloadForFolderDelegation as any).IsActive=false;
          if(folderPrivacy === "private"){
            (payloadForFolderDelegation as any).IsPrivate=true;
          }else if(folderPrivacy === "public"){
            (payloadForFolderDelegation as any).IsPrivate=false;
          }
          // if(OthProps.IsFolderDeligationUser === "true"){
          //   (payloadForFolderDelegation as any).IsFolderDeligation=true;
          // }
          if(approvalOption === "Yes"){
            (payloadForFolderDelegation as any).IsApproval=true;
          }else if(approvalOption === "No"){
            (payloadForFolderDelegation as any).IsApproval=false;
          }
        }else{
          (payloadForFolderDelegation as any).DocumentLibraryName=OthProps.DocumentLibrary;
          (payloadForFolderDelegation as any).FolderPath=`${OthProps.folderpath}/${folderName.trim()}`;
          (payloadForFolderDelegation as any).IsFolder=true;
          // (payloadForFolderDelegation as any).IsActive=true;
  
          if(OthProps.Folder ===  ""){
              (payloadForFolderDelegation as any).FolderName=folderName.trim();
          }else{
              (payloadForFolderDelegation as any).FolderName=folderName.trim();
              (payloadForFolderDelegation as any).ParentFolderId=OthProps.Folder;
              
          }
          if(folderPrivacy === "private"){
            (payloadForFolderDelegation as any).IsPrivate=true;
          }else if(folderPrivacy === "public"){
            (payloadForFolderDelegation as any).IsPrivate=false;
          }
  
          // if(OthProps.IsFolderDeligationUser === "true"){
          //   (payloadForFolderMaster as any).IsFolderDeligation=true;
          // }
        }
  
        if(OthProps.Department !== ""){
          (payloadForFolderDelegation as any).Department=OthProps.Department
        }
        if(OthProps.Devision !== ""){
          (payloadForFolderDelegation as any).Devision=OthProps.Devision
        }


        try {
          await sp.web.lists.getByTitle('DMSFolderDeligationMaster').items.add(payloadForFolderDelegation);
          console.log("Item added successfully in the DMSFolderDeligationMaster list");
        } catch (error) {
          console.log("Error in adding item in DMSFolderDeligationMaster list",error);
        }
      }
      // Clear form on successful submission
      Swal.fire({
        title: "Success",
        text: "Your request to create a folder has been submitted successfully. The folder will appear shortly as we complete the setup process.",
        icon: "success",
        // showCancelButton: true,
        confirmButtonText: 'OK',
      }).then((result) => {
        if (result.isConfirmed) {
          location.reload(); // This will reload the page
          onReturnToMain()
        }
        if(result.isDismissed){
          location.reload();
          onReturnToMain()
        }
      });
      
      //  setTimeout(() => {
      //     Swal.close(); // Close the pop-up
      //     onReturnToMain(); // Call onReturnToMain if needed
      //   }, 3000); // 3000 milliseconds = 3 seconds
      
      clearForm();
    }
  };
  // Handle form reset (Cancel button click)
  const clearForm = () => {
    setFolderName("");
    setFolderPrivacy("");
    setFolderOverview("");
    // setSelectField("");
    setApprover("");
    setErrors({});
    // setFormFields([{ id:0, fieldName: '', selectField: ''}]);
    // setRows([{ id: 0, selectionType: "One", approvedUserList: [] }])
  };

  // Handle radio button change for folder privacy

  const [showDiv, setShowDiv] = useState(false)
  const handlePrivacyChange = (e: any) => {

    setFolderPrivacy(e.target.value);
    setShowDiv(e.target.value === "private")
  };

  const createBreadCrumb=()=>{
    console.log("Props",OthProps)
    let path = OthProps.Entity;
    if(OthProps.Devision) {
      path += ` > ${OthProps.Devision}`;
    }
    if(OthProps.Department) {
      path += ` > ${OthProps.Department}`;
    }
    if (OthProps.DocumentLibrary !== "") {

      let nameArray:any=OthProps.folderpath.replace(`${locationPath}/`, "").split("/");
      console.log("nameArray bread crumb",nameArray);
      // console.log("nameArray",nameArray);
      nameArray.forEach((item:any,index:any)=>{
        console.log("Item of bread crumb",item);
        if(index !==0 ){
          path +=` > ${item}`
        }
      })
      nameArray = null
    }
    const breadCrumbElement=document.getElementById("breadCrumb")
    breadCrumbElement.innerText=""
    breadCrumbElement.innerText=`This Folder will create under: ${path}`;
    console.log("Bread Crumb Structure bread crumb",path);
  }

  return (
    <>
      <button className="BackButton me-0 mb-3"
         onClick={()=>{location.reload() ;onReturnToMain()}}
      >
 
        Back
      </button>
      <div className="mt-3">
      <p id="breadCrumb"></p>
        <div className="card cardborder p-31" style={{
          
        }}>
          <form>
            <div className="row mt-0">
              <h3 className="header-title text-dark font-16 mb-1">Basic Information</h3>
              <p className="subheader font-14 mb-3">Specify Basic Information and create folder  </p>
              <div className="col-12 col-md-6 mb-3">
                <div className="form-group">
                  <label htmlFor="folderName" className="headerfont" style={{ display: "flex", alignItems: "center" }}>
                    Folder Name<span className="text-danger" style={{ marginLeft: "4px" }}>*</span>
                  </label>
                  <input
                    type="text"
                    className="form-control fieldmargin"
                    id="folderName"
                    placeholder="Enter project name"
                    value={folderName}
                    onChange={(e) => setFolderName(e.target.value)}
                  />
                  {errors.folderName && (
                    <span className="text-danger">{errors.folderName}</span>
                  )}
                </div>
              </div>
              {/* {togglefolderPrivacy &&  ( */}
                  <div className="col-12 col-md-3 mb-3" id="folderPrivacy" style={{
                     
                  }}>
                        <div className="form-group">
                          <label htmlFor="folderPrivacy" className="headerfont" style={{ display: "flex", alignItems: "center", width:"max-content"}}>
                            Folder Privacy<span className="text-danger" style={{ marginLeft: "4px" }}>*</span>
                          </label>
                        <div>
                        <div className="form-check form-check-inline fieldmargin">
                          <input
                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                              handlePrivacyChange(e);
                              // handleToggleApproval();
                              // handlePermissionToggle(true);
                            }}
                            className="form-check-input"
                            type="radio"
                            name="folderPrivacy"
                            id="private"
                            value="private"
                            checked={folderPrivacy === "private"}
                          />
                          <label className="form-check-label" htmlFor="private">
                            Private
                          </label>
                        </div>
                        <div className="form-check form-check-inline">
                          <input
                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                              handlePrivacyChange(e);
                              // handleToggleRemove();
                              // handlePermissionToggle(false);
                            }}
                            className="form-check-input"
                            type="radio"
                            name="folderPrivacy"
                            id="public"
                            value="public"
                            checked={folderPrivacy === "public"}
                          />
                          <label className="form-check-label" htmlFor="public">
                            Public
                          </label>
                        </div>
                      </div>
                      {errors.folderPrivacy && (
                        <span className="text-danger">{errors.folderPrivacy}</span>
                      )}
                        </div>
                  </div>
              {/* )} */}

              {togglefolderPrivacy &&  (
              <div className="col-12 col-md-3 mb-3" id="approvalOption" style={{
                  
              }}>
                <div className="form-group">
                          <label htmlFor="approvalOption" className="headerfont" style={{ display: "flex", alignItems: "center" }}>
                            Approval<span className="text-danger" style={{ marginLeft: "4px" }}>*</span>
                          </label>
                        <div>
                        <div className="form-check form-check-inline fieldmargin">
                          <input
                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                              // handleDeleteOption(e);
                              handleToggleApproval(e);
                            }}
                            className="form-check-input"
                            type="radio"
                            name="approvalOption"
                            id="Yes"
                            value="Yes"
                            checked={approvalOption === "Yes"}
                          />
                          <label className="form-check-label" htmlFor="Yes">
                            Yes
                          </label>
                        </div>
                        <div className="form-check form-check-inline">
                          <input
                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                              // handleDeleteOption(e);
                              handleToggleRemove(e);
                            }}
                            className="form-check-input"
                            type="radio"
                            name="approvalOption"
                            id="No"
                            value="No"
                            checked={approvalOption === "No"}
                          />
                          <label className="form-check-label" htmlFor="No">
                              No
                          </label>
                        </div>
                      </div>
                      {errors.approvalOption && (
                        <span className="text-danger">{errors.approvalOption}</span>
                      )}
                </div>
              </div>
             )}
            </div>

            <div className="form-group mt-3">
                  <label htmlFor="folderOverview" className="headerfont" style={{ display: "flex", alignItems: "center" }}>
                    Folder Overview<span className="text-danger" style={{ marginLeft: "4px" }}>*</span>
                  </label>
                  <textarea style={{height:'70px'}}
                    className="form-control fieldmargin multilinetextWidth"
                    id="folderOverview"
                    placeholder="Enter some brief about project"
                    value={folderOverview}
                    onChange={(e) => setFolderOverview(e.target.value)}
                  />
                  {errors.folderOverview && (
                    <span className="text-danger">{errors.folderOverview}</span>
                  )}
            </div>


        
          </form>
        </div>
    

       
      </div>
      {/* this is meta column fields */}
      {OthProps.DocumentLibrary === "" && (
        <div className="card cardborder p-31 mt-3">
        <div className="">
        {toggleaddFieldsButton && ( 
              <div className="row mt-0" id="addFieldsButton">
                <div className="col-md-10  w90">
                <h3 className="header-title text-dark font-16 mb-1">List of tags</h3>
                <p className="subheader font-14 mb-3">Specify sub folder and create list of tags to be prepared and submitted by team members.</p>
                </div>
               
                <div style={{position:'relative'}} className="col-md-2">
                <div className="mb-3">
                  <div className="col-12 d-flex justify-content-end">
                      <a onClick={handleAddFields}>
                      <img 
                            className="bi bi-plus"
                            src={require("../assets/plus.png")}
                            alt="add"
                            style={{ width: "50px", top:'5px', left:'auto', right:'0px', marginLeft:'16px', position:'absolute', height: "50px" }}
                          />
                     
                      </a>
                  </div>
                  </div>
                </div>
                  
              </div>
          )}
  <table className="mtbalenew mtbalenewn createc">
    <thead>
      <tr>
        <th>  Field Name</th>
        <th>    Select Field Type</th>
        <th style={{minWidth:'40px',maxWidth:'40px'}}> Action</th>
      </tr>
    </thead>
    <tbody>

          {togglecolumneDetails && formFields.map((formField) => (
            
      <tr  key={formField.id} id="columnDetail">
        <td>
          <div className="form-group">
            {/* <label htmlFor={`fieldName-${formField.id}`} className="headerfont">
              Field Name
            </label> */}
            <input
              type="text"
              className="form-control fieldmargin"
              id={`fieldName-${formField.id}`}
              name="fieldName"
              placeholder="Enter field name"
              value={formField.fieldName}
              onChange={(e) => handleInputChange(formField.id, e)}
            />
            {/* <span className="text-danger">{errors.fieldName}</span> */}
            {errors1[formField.id]?.fieldName && (
              <span className="text-danger">{errors1[formField.id].fieldName}</span>
            )}
          </div>
        </td>

        <td>
          <div className="form-group">
            {/* <label htmlFor={`selectField-${formField.id}`} className="headerfont">
              Select Field Type
            </label> */}
            <select
              className="form-control"
              id={`selectField-${formField.id}`}
              name="selectField"
              value={formField.selectField}
              onChange={(e) => handleSelectedType(formField.id, e)}
            >
              <option value="">Open this select menu</option>
              <option value="Single Line of Text">Single Line of Text</option>
              <option value="Multiple Line of Text">Multiple Line of Text</option>
              <option value="Yes or No">Yes or No</option>
              <option value="Date & Time">Date & Time</option>
              <option value="Number">Number</option>
            </select>
            {/* <span className="text-danger">{errors.selectField}</span> */}
            
                {errors1[formField.id]?.selectField && (
              <span className="text-danger">{errors1[formField.id].selectField}</span>
            )}
          </div>
        </td>
        <td style={{minWidth:'40px',maxWidth:'40px',textAlign:'center'}}>
        <div >
        {formField.id === 0 ? (
                null
              ) : (
                <div style={{justifyContent:'center'}} className="d-flex align-items-end">
                  <a
                    onClick={(e) => handleRemoveField(formField.id, e)}
                    style={{
                      width: "50px",
                   
                      cursor: "pointer",
                    }}
                  >
                    <img style={{marginTop:'0px'}}
                      className="fas fa-trash"
                      src={require("../assets/del.png")}
                      alt="delete"
                    />
                  </a>
                </div>
              )}

          </div>
          </td>
      </tr>
      
            ))}
            </tbody>
              </table>

        </div>
      </div>
      )}
       
      
      {toggleApproval ? (
        <div className="card cardborder p-31 mt-3">
          <div className="" style={{
            
        }}>
            {/* <h5 className="mb-1 Permissionsectionstyle">
              <strong>Approval Hierarchy</strong>
            </h5> */}
            <div className="row">
              <div className="col-sm-10 w90">
              <h3 className="header-title text-dark font-16 mb-1">Approval Hierarchy</h3>
     
     <p className="subheader font-14 mb-3">
       Define approval hierarchy for the documents submitted by Team
       members in this folder.
     </p>

              </div>

              <div className="col-sm-2">
              <div style={{height:'0px', position:'relative'}} className="mb-0">
              <div className="col-12 d-flex justify-content-end">
                <a onClick={handleAddRow}>
                  <img
                    className="bi bi-plus"
                    src={require("../assets/plus.png")}
                    alt="add"
                    style={{ width: "50px", top:'0px', position:'absolute', right:'0px', left:'auto', height: "50px" }}
                  />
                </a>
              </div>
            </div>

              </div>
            </div>
            

           
            <div style={{clear:'both'}} className="row mb-2 approvalheirarcystyle">
            <table className="mtbalenew mtbalenewn createc">
    <thead>
      <tr>
        <th> Level</th>
        <th> Approver</th>
        <th style={{textAlign:'center', minWidth:'70px', maxWidth:'70px'}}> Select</th>
        <th style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}> Action</th>
      </tr>
    </thead>
    <tbody>


              {/* <div className="col-12 col-md-4">
                <label
                  htmlFor="level"
                  className="form-label approvalhierarcyfont"
                >
                  Level
                </label>
              </div> */}
              {/* <div className="col-12 col-md-5">
                <label
                  htmlFor="approver"
                  className="form-label approvalhierarcyfont"
                >
                  Approver
                </label>
              </div> */}
               {rows.map((row) => (
              <tr className="approvalheirarchyfield" key={row.id}>
                <td>
                  <input style={{height:'36px'}}
                    type="text"
                    className="form-control"
                    id={`level-${row.id}`}
                    value={`Level ${row.id + 1}`}
                    disabled
                  />
                </td>
                <td>
                  {/* start */}
                  <Select
                    isMulti
                    options={users}
                    // value={Approver}
                    onChange={(selected: any) =>
                      handleUserSelect(selected, row.id)
                    }
                    placeholder="Enter names or email addresses..."
                    noOptionsMessage={() => "No User Found..."}
                  />
                  {/* {errors.selectedUsers && (
                    <span className="text-danger">{errors.selectedUsers}</span>
                  )} */}
                  {errorsForUserSelection[row.id]?.userSelect && (
                      <span className="text-danger">{errorsForUserSelection[row.id].userSelect}</span>
                  )}
                  {/* end */}
                </td>
                {/* start */}
                <td style={{textAlign:'center', minWidth:'70px', maxWidth:'70px'}}>
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
                    <label
                      className="form-check-label"
                      htmlFor={`all-${row.id}`}
                    >
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
                    <label
                      className="form-check-label"
                      htmlFor={`one-${row.id}`}
                    >
                      One
                    </label>
                  </div>
                </div>
                </td>
                {/* end */}
                <td style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}>
                {row.id === 0 ? (
                 null
                ) : (
                  <div style={{justifyContent:'center'}} className="d-flex align-items-end">
                    <a
                      onClick={(e) => handleRemoveRow(row.id, e)}
                      style={{
                        width: "50px",
                       
                        cursor: "pointer",
                      }}
                    >
                      <img
                        className="fas fa-trash"
                        src={require("../assets/del.png")}
                        alt="delete"
                      />
                    </a>
                  </div>
                )}
                </td>

               
              </tr>
            ))}
            </tbody>
            </table>
            </div>
           

          </div>
        </div>
      ) : null
      }
      
      {/* {permission && ( */}
           {showDiv &&   <div className="card cardborder p-31 mt-3">
                <div className="" style={{
               
                  }}>
                      {/* <h5 className="mb-3 Permissionsectionstyle">
                          <strong>Permission</strong>
                      </h5> */}
                                  
                      <div className="row">
                        <div className="col-md-10 w90">

                        <h3 className="header-title text-dark font-16 mb-1">Permission</h3>
     
     <p className="subheader font-14 mb-3">
       Define Permission for the documents submitted by Team
       members in this folder.
     </p>
           
                        </div>
                        
                        <div className="col-md-2">
                        <div style={{position:'relative'}} className="mb-3">
                        <div className="col-12  d-flex justify-content-end">
                          <a onClick={handleAddRowForPermission}>
                            <img 
                              className="bi bi-plus"
                              src={require("../assets/plus.png")}
                              alt="add"
                              style={{ width: "50px", top:'0px', left:'auto', right:'0px', marginLeft:'10px', position:'absolute', height: "50px" }}
                            />
                          </a>
                        </div>
                      </div>
                        </div>

                      </div>
                      <table className="mtbalenew mtbalenewn createc">
    <thead>
      <tr>
        <th> Name</th>
        <th> Permission</th>
      
        <th style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}> Action</th>
      </tr>
    </thead>
    <tbody>
                      {rowsForPermission.map((rowForPermission)=>(
                          <tr className="approvalheirarcystyle" key={rowForPermission.id}>
                              <td className="">
                                  <Select
                                      isMulti
                                      options={siteUsers}
                                      onChange={(selected: any) =>
                                        handleUserSelectForPermission(selected,rowForPermission.id)
                                      }
                                      placeholder="Enter names or email addresses..."
                                      noOptionsMessage={() => "No User Found..."}
                                     
                                  />
                                  {errorsForPermissionSelection[rowForPermission.id]?.userSelect && (
                                    <span className="text-danger">{errorsForPermissionSelection[rowForPermission.id].userSelect}</span>
                                  )}
                              </td>
                              <td className="" 
                              
                              >
                                  <Select
                                      options={permissionArray}
                                      onChange={(selected: any) =>
                                        handlePermissionSelect(selected,rowForPermission.id)
                                      }
                                      placeholder="Select Permission"
                                      noOptionsMessage={() => "No Such Permission Find"}
                                  />
                                   {errorsForPermissionSelection[rowForPermission.id]?.permissionSelect && (
                                    <span className="text-danger">{errorsForPermissionSelection[rowForPermission.id].permissionSelect}</span>
                                  )}
                              </td>
                            <td style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}>
                            {rowForPermission.id === 0 ? (
                                null
                                ) : (
                                  <div style={{justifyContent:'center'}} className="d-flex align-items-end">
                                    <a
                                      onClick={(e) => handleRemoveRowForPermission(rowForPermission.id, e)}
                                      style={{
                                        width: "50px",
                                       
                                        cursor: "pointer",
                                      }}
                                    >
                                      <img
                                        className="fas fa-trash"
                                        src={require("../assets/del.png")}
                                        alt="delete"
                                      />
                                    </a>
                                  </div>
                                )}

                            </td>
                             
                          </tr>
                      ))}
                      </tbody>
                      </table>

                </div>

                <div>
           
                </div>
        </div> }
        <div className="d-flex mt-3 justify-content-center buttonstyle">
        <button
          className="btn btn-create me-2 mt-0 btncolorCreate"
          onClick={handleCreate}
          id="CreateFolderInsideSharePoint"
        >
          <img
            className="bi"
            src={require("../assets/checkmark2.png")}
            alt="Create"
          />
          Create
        </button>
        <button className="btn btn-cancel btncolorcancel" onClick={clearForm}>
          <img
            className="bi"
            src={require("../assets/cross.png")}
            alt="Cancel"
          />
          Cancel
        </button>
                </div>
      {/* ) */}
      {/* } */}
      <br/>
      
    </>
  );
};

export default CreateFolder;

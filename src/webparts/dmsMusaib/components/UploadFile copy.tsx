// @ts-ignore
import * as React from "react";
import { useEffect , useState , useRef} from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import './uploadfilecss'
import * as XLSX from 'xlsx';
import './uploadfilecss'
import Swal from 'sweetalert2';
let IsApproval : any
let status :any;
let buttontext :any = 'Submit'
interface UploadFileProps {
  currentfolderpath: { [key: string]: string };
  onReturnToMain: () => void;
  // myRequest: (event: React.MouseEvent<HTMLButtonElement>) => void;
}

let previewURLN;
const submitButton=document.createElement('button');
submitButton.textContent= buttontext
submitButton.id="submitBtn";
submitButton.type="submit";


// const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
// if(submitBtn){

//   submitBtn.disabled = true;
// }

const UploadFile: React.FC<UploadFileProps> = ({ currentfolderpath , onReturnToMain  }) => {
  const sp: SPFI = getSP();
  let locationPath=window.location.pathname.match(/\/sites\/[^\/]+/)[0];
  // check whether folder is private or public and save state
  const checkfolderprivace = async() =>{
    const folderItems = await sp.web.lists.getByTitle("DMSPreviewFormMaster")
    .items.filter(`DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' and SiteName eq '${currentfolderpath.Entity}' and IsDocumentLibrary eq 1`).select("IsApproval","IsPrivate")();
    console.log("folderItems",folderItems);
    IsApproval=folderItems[0].IsApproval;

    console.log('currentfolderpath' , currentfolderpath)
  }
  checkfolderprivace();


const [data, setData] = useState({
  Entity: '',
  Entityurl: '',
  siteID: '',
  Devision: '',
  Department: '',
  DocumentLibrary: '',
  Folder: '',
  folderpath: '',
});

const [state, setState] = useState({});

const currentUserEmailRef = useRef('');

const getcurrentuseremail = async()=>{
  const userdata = await sp.web.currentUser();
  currentUserEmailRef.current = userdata.Email;
  // console.log(currentUserEmailRef.current, "currentuser")
 }

useEffect(() => {
  getcurrentuseremail()
  setData({...data , ...currentfolderpath});

}, []);


console.log(data, "data"  )
console.log(data.Entity, "entity"  )
const SubsiteID = data.siteID
const currentPath =data.folderpath; 
const documentLibraryName  = data.DocumentLibrary;
console.log("documentLibraryName" , documentLibraryName)


  // const [libraryName, setLibraryName] = useState("Shared Documents"); // Set your library name here

  // let selectedFile:any=null;


  // const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  //   const file = event.target.files![0];
  //   // if (file) {
  //   //   // selectedFile=file;
  //   //   uploadFile(file);
  //   // }

  
  //   if (file) {
  //     const allowedExtensions = ['doc', 'docx', 'ppt', 'pdf', 'xls', 'xlsx', 'txt', 'png', 'jpg', 'jpeg'];
  //     const nonAlphaNumericForEntity = file.name.replace(/[^a-zA-Z0-9 -]/g, '');
  //     const fileExtension = file.name.split('.').pop()?.toLowerCase();
  
  //     if (!fileExtension || !allowedExtensions.includes(fileExtension)) {
  //       // alert('Invalid file type. Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are allowed.');
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file type',
  //         text: 'Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are allowed.',
  //       });
  //       event.target.value = ''; // Clear the input
  //     } else if(nonAlphaNumericForEntity !== file.name) {
  //       // alert('Special charaters are not allowed.');
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file name',
  //         text: 'File names can only contain letters, numbers, spaces, and hyphens.',
  //       });
  //       event.target.value = ''; // Clear the input
  //     } else {
  //       uploadFile(file);
  //     }
  //   }
  // };
  // const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  //   const file = event.target.files![0];
  
  //   if (file) {
  //     const allowedExtensions = ['doc', 'docx', 'ppt', 'pdf', 'xls', 'xlsx', 'txt', 'png', 'jpg', 'jpeg'];
  //     const fileExtension = file.name.split('.').pop()?.toLowerCase();
  //     const invalidCharacters = /[^a-zA-Z0-9 -]/g; // Allowed: letters, numbers, spaces, and hyphens
  
  //     // Check for invalid characters in the file name
  //     if (invalidCharacters.test(file.name)) {
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file name',
  //         text: 'File names can only contain letters, numbers, spaces, and hyphens.',
  //       });
  //       event.target.value = ''; // Clear the input
  //     } 
  //     // Check for invalid file extensions
  //     else if (!fileExtension || !allowedExtensions.includes(fileExtension)) {
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file type',
  //         text: 'Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are allowed.',
  //       });
  //       event.target.value = ''; // Clear the input
  //     } 
  //     // If all validations pass, upload the file
  //     else {
  //       uploadFile(file);
  //     }
  //   }
  // };
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  const file = event.target.files![0];

  if (file) {
    const allowedExtensions = ['doc', 'docx', 'ppt', 'pdf', 'xls', 'xlsx', 'txt', 'png', 'jpg', 'jpeg'];
    const fileExtension = file.name.split('.').pop()?.toLowerCase(); // Extract the file extension
    const baseFileName = file.name.substring(0, file.name.lastIndexOf('.')); // Extract the file name before the extension
    const invalidCharacters = /[^a-zA-Z0-9 -]/g; // Allowed: letters, numbers, spaces, and hyphens

    // Check for invalid file extensions
    if (!fileExtension || !allowedExtensions.includes(fileExtension)) {
      Swal.fire({
        icon: 'error',
        title: 'Invalid file type',
        text: 'Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are allowed.',
      });
      event.target.value = ''; // Clear the input
    }
    // Check for invalid characters in the file name (excluding extension)
    else if (invalidCharacters.test(baseFileName)) {
      Swal.fire({
        icon: 'error',
        title: 'Invalid file name',
        text: 'File names can only contain letters, numbers, spaces, and hyphens.',
      });
      event.target.value = ''; // Clear the input
    }
    // If all validations pass, upload the file
    else {
      uploadFile(file);
    }
  }
};

  
  
  
  
  
  const uploadFile = async (file: File) => {
    try {
      
      const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
      const uploadResult = await folder.files.addChunked(file.name, file);
      console.log("File uploaded successfully", uploadResult);

      // Generate the preview URL dynamically
      const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
      const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
      submitButton.disabled = false; // Disable the button
     
    
       previewFile(previewUrl);
    } catch (error) {
      console.error("Error uploading file:", error);
    }
  };

  const generatePreviewUrl = async (serverRelativeUrl: string) => {
    // Encode the file name and construct the preview URL
    const encodedFilePath = encodeURIComponent(serverRelativeUrl);
    
    // Example: 

    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;

    //  const previewUrl = `${siteUrl}/sites/edcspfx/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    //  const previewUrl = `${siteUrl}/sites/edcspfx/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
      const previewUrl = `${siteUrl}${locationPath}/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    console.log("Generated Preview URL:", previewUrl);
   if(previewUrl){
    console.log("enter herr")
    const deletebut = document.getElementById('closeCommand') as HTMLElement
    if(deletebut){
      console.log(" here " , deletebut)
    }
   }
    return previewUrl;
  };


// const previewFile = async (previewUrl: string) => {
//     try {
//       console.log("Previewing file at URL:", previewUrl);
//       const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
//       const spinner = document.getElementById("spinner") as HTMLElement;
  
//       // Show the spinner and hide the iframe initially
//       spinner.style.display = "block";
//       iframe.style.display = "none";
//       iframe.src = previewUrl;
  
//       // Add an onload event listener to the iframe
//       iframe.onload = () => {
//         console.log("Iframe has loaded");
  
//         const checkAndHideButton = () => {
//           try {
//             const iframeDocument = iframe.contentDocument || iframe.contentWindow?.document;
//             if (iframeDocument) {
//               const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
//               const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
//               if(excelToolbar){
//                 excelToolbar.style.display= "none"
//               }
//               if (button) {
//                 console.log("Hiding the OneUpCommandBar element");
//                 button.style.display = "none";
  

//                 spinner.style.display = "none";
//                 iframe.style.display = "block"; 


//               } else {
//                 console.log("OneUpCommandBar not found, rechecking...");
//               }
              
//               const helpbutton = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement; 
//               if(helpbutton){
//                 helpbutton.style.display = "none"
//               }
//             }
//           } catch (error) {
//             console.error("Error accessing iframe content:", error);
//           }
  

//           setTimeout(checkAndHideButton, 100);
//         };
  

//         checkAndHideButton();
//       };
//     } catch (error) {
//       console.error("Error previewing file:", error);
//     }

//   };

  // const entity=data.Entity;

  // const previewFile = async (previewUrl: string) => {
  //   try {
  //     console.log("Previewing file at URL:", previewUrl);
  //     const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
  //     const spinner = document.getElementById("spinner") as HTMLElement;
      
  //     // Show the spinner and hide the iframe initially
  //     spinner.style.display = "block";
  //     iframe.style.display = "none";
  //     iframe.src = previewUrl;
      
  //     // Add an onload event listener to the iframe
  //     iframe.onload = () => {
  //       console.log("Iframe has loaded");
  
  //       const checkAndHideButton = () => {
  //         try {
  //           const iframeDocument = iframe.contentDocument || iframe.contentWindow?.document;
  //           if (iframeDocument) {
  //             const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
  //             const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
  //             if (excelToolbar) {
  //               excelToolbar.style.display = "none";
  //             }
  
  //             // Hide OneUpCommandBar
  //             if (button) {
  //               console.log("Hiding the OneUpCommandBar element");
  //               button.style.display = "none";
  //               spinner.style.display = "none";
  //               iframe.style.display = "block";
  //             } else {
  //               console.log("OneUpCommandBar not found, rechecking...");
  //             }
  //           }
  //         } catch (error) {
  //           console.error("Error accessing iframe content:", error);
  //         }
  
  //         setTimeout(checkAndHideButton, 100);
  //       };
  
  //       checkAndHideButton();
  //     };
  //   } catch (error) {
  //     console.error("Error previewing file:", error);
  //   }
  
  //   // Add fullscreen button overlay
  //   const iframeContainer = document.getElementById("iframeContainer");
  //   const fullscreenButton = document.createElement("button");
  //   fullscreenButton.textContent = "Go Fullscreen";
  //   fullscreenButton.style.position = "absolute";
  //   fullscreenButton.style.top = "20px";
  //   fullscreenButton.style.right = "20px";
  //   fullscreenButton.style.padding = "10px";
  //   fullscreenButton.style.fontSize = "16px";
  //   fullscreenButton.style.backgroundColor = "rgba(0, 0, 0, 0.5)";
  //   fullscreenButton.style.color = "white";
  //   fullscreenButton.style.border = "none";
  //   fullscreenButton.style.cursor = "pointer";
    
  //   // Append the button to iframe container
  //   iframeContainer?.appendChild(fullscreenButton);
  
  //   fullscreenButton.addEventListener("click", () => {
  //     // Fullscreen logic
  //     iframeContainer!.style.position = "fixed";
  //     iframeContainer!.style.top = "0";
  //     iframeContainer!.style.left = "0";
  //     iframeContainer!.style.width = "100%";
  //     iframeContainer!.style.height = "100%";
  //     iframeContainer!.style.zIndex = "9999"; // Ensuring it stays on top
  //     iframeContainer!.style.backgroundColor = "rgba(0, 0, 0, 0.8)"; // Optional dark background
  
  //     // Hide fullscreen button after fullscreen is activated
  //     fullscreenButton.style.display = "none";
  //   });
  // };
  const previewFile = async (previewUrl: string) => {
    try {
      console.log("Previewing file at URL:", previewUrl);
      const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
      const spinner = document.getElementById("spinner") as HTMLElement;
  
      // Show the spinner and hide the iframe initially
      spinner.style.display = "block";
      iframe.style.display = "none";
      
      // Set iframe src only once to prevent reloading
      if (!iframe.src || iframe.src !== previewUrl) {
        iframe.src = previewUrl;
      }
  
      // Add an onload event listener to the iframe
      iframe.onload = () => {
        console.log("Iframe has loaded");
  
        const checkAndHideButton = () => {
          try {
            const iframeDocument = iframe.contentDocument || iframe.contentWindow?.document;
            if (iframeDocument) {
              const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
              const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
              if (excelToolbar) {
                excelToolbar.style.display = "none";
              }
  
              // Hide OneUpCommandBar
              if (button) {
                console.log("Hiding the OneUpCommandBar element");
                button.style.display = "none";
                spinner.style.display = "none";
                iframe.style.display = "block";
              } else {
                console.log("OneUpCommandBar not found, rechecking...");
              }
            }
          } catch (error) {
            console.error("Error accessing iframe content:", error);
          }
  
          setTimeout(checkAndHideButton, 100);
        };
  
        checkAndHideButton();
      };
    } catch (error) {
      console.error("Error previewing file:", error);
    }
  
    // Add fullscreen and minimize button overlay
    const iframeContainer = document.getElementById("iframeContainer");
    const fullscreenButton = document.createElement("button");
    fullscreenButton.textContent = "Go Fullscreen";
    fullscreenButton.style.position = "absolute";
    fullscreenButton.style.top = "20px";
    fullscreenButton.style.right = "20px";
    fullscreenButton.style.padding = "10px";
    fullscreenButton.style.fontSize = "16px";
    fullscreenButton.style.backgroundColor = "rgba(0, 0, 0, 0.5)";
    fullscreenButton.style.color = "white";
    fullscreenButton.style.border = "none";
    fullscreenButton.style.cursor = "pointer";
    
    // Minimize button
    const minimizeButton = document.createElement("button");
    minimizeButton.textContent = "Minimize";
    minimizeButton.style.position = "absolute";
    minimizeButton.style.top = "20px";
    minimizeButton.style.left = "20px";
    minimizeButton.style.padding = "10px";
    minimizeButton.style.fontSize = "16px";
    minimizeButton.style.backgroundColor = "rgba(0, 0, 0, 0.5)";
    minimizeButton.style.color = "white";
    minimizeButton.style.border = "none";
    minimizeButton.style.cursor = "pointer";
    minimizeButton.style.display = "none";  // Initially hidden
  
    // Append both buttons to iframe container
    iframeContainer?.appendChild(fullscreenButton);
    iframeContainer?.appendChild(minimizeButton);
  
    fullscreenButton.addEventListener("click", (event) => {
      event.preventDefault();
      // Fullscreen logic - modify only the iframe container's style
      iframeContainer!.style.position = "fixed";
      iframeContainer!.style.top = "0";
      iframeContainer!.style.left = "0";
      iframeContainer!.style.width = "100%";
      iframeContainer!.style.height = "100%";
      iframeContainer!.style.zIndex = "9999";  // Ensuring it stays on top
      iframeContainer!.style.backgroundColor = "rgba(0, 0, 0, 0.8)";  // Optional dark background
  
      // Hide fullscreen button and show minimize button after fullscreen is activated
      fullscreenButton.style.display = "none";
      minimizeButton.style.display = "block";
    });
  
    minimizeButton.addEventListener("click", (event) => {
      event.preventDefault();
      
      // Minimize logic - revert iframe container to its original size
      iframeContainer!.style.position = "relative";
      iframeContainer!.style.width = "80%"; // Or whatever size you want
      iframeContainer!.style.height = "80%"; // Or whatever size you want
      iframeContainer!.style.top = "50%";
      iframeContainer!.style.left = "50%";
      iframeContainer!.style.transform = "translate(-50%, -50%)";  // Center it on the page
      iframeContainer!.style.backgroundColor = "transparent";  // Remove the overlay
      
      // Hide minimize button and show fullscreen button again
      minimizeButton.style.display = "none";
      fullscreenButton.style.display = "block";
    });
  };
  
  
  
React.useEffect(()=>{
  const  loadFormOptions = async ()=> {
    try {
      // const testidsub = await sp.site.openWebById(data.siteID);
      
      // console.log("current Entity URL",testidsub.data.Url
      //   )
      
      // const fields = await testidsub.web.lists.getByTitle(`${documentLibraryName}`).fields.filter("Hidden eq false and ReadOnlyField eq false")();
      // console.log("Fields of document Library",fields);


      // const forms = await sp.web.lists.getByTitle("DMSPreviewformfields").items.select('*','SiteTitle/Title', 'SiteTitle/SiteURL').expand('SiteTitle')
      // .filter(`SiteTitle/Title eq '${propsDeatils.currentEntity}' `)();
      // console.log(forms, "forms");

      // start
      
      const documentLibraryFields=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName","ColumnType","IsRequired","IsRename")
      .filter(
            `SiteName eq '${currentfolderpath.Entity}' 
            and DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' 
            and AddorRemoveThisColumn eq  'Add To Library' and IsInProgress eq 0`)();

      console.log("Document Library Fields",documentLibraryFields);
      // end

      const formSelector = document.getElementById("formSelector");
     
      const uploadFileDiv=document.createElement('div');

      const createElement=(fieldName:string,type:string,required:boolean,IsRename:string)=>{
            let fName=fieldName
            if(IsRename!== null){
              fName=IsRename
            }
            const inputContainer = document.createElement("div"); 
            inputContainer.className = "input-container";
    
            // Create and set label
            const label = document.createElement("label");
            label.setAttribute("htmlFor", fieldName);
            // label.textContent = fieldName;
            label.textContent = fName;
            inputContainer.appendChild(label);
    
            let inputElement: HTMLInputElement | null = null;
    
            // Dynamically create the input field based on FieldType
            let modifiedType = type.replace(/\s+/g, '').toLowerCase();
            console.log("modifiedType",modifiedType);

            if (
                modifiedType === "singlelineoftext"
                || 
                modifiedType === "multiplelineoftext" 
                || 
                modifiedType === 'text'
            ){
              inputElement = document.createElement("input");
              inputElement.type = "text";
            } else if (
              modifiedType === "number"
            ) {
              inputElement = document.createElement("input");
              inputElement.type = "number";
            } else if (
              modifiedType === "date&time"
            ) {
              inputElement = document.createElement("input");
              inputElement.type = "date";
            } else if (
              modifiedType === "yesorno"
            ) {
              inputElement = document.createElement("input");
              inputElement.type = "checkbox";
            }

            if (inputElement) {
              inputElement.className="dynamic-input";
              inputElement.id = fieldName;
              // inputElement.required = required.toLowerCase() === "yes"; 
              inputElement.required=required;
              inputContainer.appendChild(inputElement); 
              formSelector.appendChild(inputContainer); 
            }

            return;
      }


      // start
      documentLibraryFields.forEach((field)=>{
        createElement(field.ColumnName,field.ColumnType,field.IsRequired,field.IsRename);
        })
      // end


      // properties of upload file div
      // uploadFileDiv.className="uploadfile";
      uploadFileDiv.className="input-container";

      // input for upload file
      const uploadFileInput=document.createElement('input');
      uploadFileInput.className="dynamic-input";
      uploadFileInput.type="file";
      uploadFileInput.id="fileInput";
      uploadFileInput.addEventListener('change', (event:any) => handleFileChange(event))

      // Set Label For upload file
      const label = document.createElement("label");
      label.setAttribute("htmlFor", 'fileInput');
      label.textContent = 'Upload File';

      uploadFileDiv.appendChild(label);
      uploadFileDiv.appendChild(uploadFileInput);
      formSelector.appendChild(uploadFileDiv);

      // Submit Button property
      
      submitButton.addEventListener('click',handleSubmit)
  
      formSelector.appendChild(submitButton);
      
    } catch (error) {
      console.error("Error loading form options:", error);
    }
      }  

      loadFormOptions();
},[])
  


// const handleSubmit = async (event: any) => {

//   event.preventDefault();
//   console.log("Button clicked");

//   const formSelector = document.getElementById("formSelector") as HTMLFormElement;
//   if (!formSelector.checkValidity()) {
//       formSelector.reportValidity(); // Show validation errors
//       return;
//   }

//   // Prepare the payload for SharePoint dynamically
//   const inputs = document.querySelectorAll('.dynamic-input');
//   // console.log("inputs",inputs)

  
//   const payload: any = {};

//   inputs.forEach((input) => {
//       const inputElement = input as HTMLInputElement;
//       const fieldName = inputElement.id;
//       if (!fieldName) return; // Skip if field name is invalid

//       if (inputElement.type === "checkbox") {
//           // console.log("fieldName",fieldName.includes(' '));
//           payload[fieldName] = inputElement.checked;
//       } else if (inputElement.type !== "file") {
//           if(inputElement.value === ""){
//              console.log("skip");
//           }else{
//             // if(fieldName.includes(' '))
//             // console.log("fieldName",fieldName.includes(' '));
//             payload[fieldName] = inputElement.value;
//           }
          
//       }
//   });

//   const fileInput = document.getElementById('fileInput') as HTMLInputElement;
//   const selectedFile = fileInput?.files?.[0]; 

//   if (!selectedFile) {
//       console.error("No file selected.");

//       return;
//   }

//   try {
//       console.log("Payload:", payload);
//       console.log("SiteID:", currentfolderpath.siteID);

//       const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
//       if (!testidsub) throw new Error("Subsite not found.");

//       const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
//       console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

//       const files = await documentLibraryInWhichWeUploadTheFile.files();
//       const fileExists = files.some(file => file.Name === selectedFile.name);
  
//       if (fileExists) {
//          Swal.fire({
//            icon: 'error',
//            title: 'File already exists',
//            text: 'The file you are trying to upload already exists in the document library. Please choose a different file name.',
//          })
//          return
//       }
  
      
//       const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
//       console.log("File uploaded successfully", uploadResult.data.Name);
//       const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
//       submitBtn.disabled = true; 
//       submitBtn.innerText = "Submitting..."; // Optional: Change button text to indicate progress
    
//       const listItem = await uploadResult.file.getItem();
//       console.log("ListItems ",listItem);
      
//       const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
//       const siteUrl = window.location.origin;
//       const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
//       console.log(encodedFilePath , "encodedFilePath")
//         const previewUrl = `${siteUrl}/sites/edcspfx/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
//       //  const previewUrl = `${siteUrl}/sites/edcspfx/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
//         //  const previewUrl = `${siteUrl}/sites/IntranetUAT/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
      
//       console.log("Generated Preview URL:", previewUrl);
//       if (!listItem) throw new Error("List item not found for the uploaded file.");


//       if(IsApproval === true){
  
//         status="Pending";
//       }else if(IsApproval === false){

//         status="Auto Approved";
//       }
//       (payload as any).Status=status;
//       await listItem.update(payload);
//       console.log("File metadata updated successfully with:", payload);
     
      
//       // alert(`status,${status}`);
//       const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
//           FileName: String(uploadResult.data.Name),
//           FileSize: String(uploadResult.data.Length),
//           FileVersion: String(uploadResult.data.MajorVersion),
//           CurrentFolderPath: String(currentfolderpath.folderpath),
//           FileUID: String(uploadResult.data.UniqueId),
//           CurrentUser: String(currentUserEmailRef.current),
//           SiteID: String(currentfolderpath.siteID),
//           Status: status,
//           FilePreviewURL : String(previewUrl),
//           DocumentLibraryName:String(currentfolderpath.DocumentLibrary),
//           SiteName : String(currentfolderpath.Entity),
//           MyRequest: true,
//           RequestNo: `DMS-${uploadResult.data.UniqueId}`
//       });
//       console.log(newItem, "New item added FileMaster");

      
//       if(IsApproval === true){
//         const AddIteminDMSFileApprovalList = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
//           SiteName : String(currentfolderpath.Entity),  
//            DocumentLibraryName : String(currentfolderpath.DocumentLibrary),
//            RequestedBy  : String(currentUserEmailRef.current),
//            FileName: String(uploadResult.data.Name),
//            FileUID: String(uploadResult.data.UniqueId),
//            FilePreviewUrl: String(previewUrl),
//            Status: String('Pending'),
//            FolderPath : String(currentfolderpath.folderpath),
//            ApproveAction : String('Submitted'),
//            ApprovedLevel : 1,
//            RequestNo: `DMS-${uploadResult.data.UniqueId}`
//       })
//       }


//     // console.log(AddIteminDMSFileApprovalList, "New item added to DMSFileApprovalList");

//     if(newItem ){
//       Deletemedia()
//       setTimeout(() => {
//         location.reload()
//         onReturnToMain(); // Call onReturnToMain after 3 seconds
//     }, 3000); // 3000 milliseconds = 3 seconds
//      }

    

//   }catch (error) {
//       console.error("Error during submission:", error);
//   }


// };
const getUniqueRequestNo = async () => {
  const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
  console.log("Counter Item 0", counterItem);
  console.log("Counter Item 1", counterItem.FileCount);
  let fileCounter = counterItem.FileCount;

  // Increment the counter
  fileCounter++;

  // Generate the new RequestNo
  const newRequestNo = `File${String(fileCounter).padStart(2, '0')}`;

  // Update the counter in the CounterList
  await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
    FileCount: fileCounter
  });

  return newRequestNo;
};

const handleSubmit = async (event: any) => {
 
  event.preventDefault();
  console.log("Button clicked");
 
  const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  if (!formSelector.checkValidity()) {
      formSelector.reportValidity(); // Show validation errors
      return;
  }
 
  // Prepare the payload for SharePoint dynamically
  const inputs = document.querySelectorAll('.dynamic-input');
  // console.log("inputs",inputs)
 
 
  const payload: any = {};
 
  inputs.forEach((input) => {
      const inputElement = input as HTMLInputElement;
      const fieldName = inputElement.id;
      if (!fieldName) return; // Skip if field name is invalid
 
      if (inputElement.type === "checkbox") {
          // console.log("fieldName",fieldName.includes(' '));
          payload[fieldName] = inputElement.checked;
      } else if (inputElement.type !== "file") {
          if(inputElement.value === ""){
             console.log("skip");
          }else{
            // if(fieldName.includes(' '))
            // console.log("fieldName",fieldName.includes(' '));
            payload[fieldName] = inputElement.value;
          }
         
      }
  });
 
  const fileInput = document.getElementById('fileInput') as HTMLInputElement;
  const selectedFile = fileInput?.files?.[0];
 
  if (!selectedFile) {
      console.error("No file selected.");
 
      return;
  }
 
  try {
      console.log("Payload:", payload);
      console.log("SiteID:", currentfolderpath.siteID);
 
      const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
      if (!testidsub) throw new Error("Subsite not found.");
 
      const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
      console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);
 
      const files = await documentLibraryInWhichWeUploadTheFile.files();
      // const fileExists = files.some(file => file.Name === selectedFile.name);
 
      // if (fileExists) {
      //    Swal.fire({
      //      icon: 'error',
      //      title: 'File already exists',
      //      text: 'The file you are trying to upload already exists in the document library. Please choose a different file name.',
      //    })
      //    return
      // }
      const originalFileName = selectedFile.name;
      const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
      const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      console.log("originalFileName",originalFileName);
      console.log("fileExtension",fileExtension);
      console.log("baseFileName",baseFileName);
      let uniqueFileName = originalFileName;
      let counter = 1;
 
      while (files.some(file => file.Name === uniqueFileName)) {
        uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
        counter++;
      }
      console.log(`Unique file name generated: ${uniqueFileName}`);
      // const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
      const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, selectedFile);
      console.log("File uploaded successfully", uploadResult.data.Name);
      const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
      submitBtn.disabled = true;
      submitBtn.innerText = "Submitting..."; // Optional: Change button text to indicate progress
   
      const listItem = await uploadResult.file.getItem();
      console.log("ListItems ",listItem);
     
      const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
      const siteUrl = window.location.origin;
      const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
      console.log(encodedFilePath , "encodedFilePath")
          const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
     
      console.log("Generated Preview URL:", previewUrl);
      if (!listItem) throw new Error("List item not found for the uploaded file.");
 
 
      if(IsApproval === true){
 
        status="Pending";
      }else if(IsApproval === false){
 
        status="Auto Approved";
      }
      (payload as any).Status=status;
      await listItem.update(payload);
      console.log("File metadata updated successfully with:", payload);
     
      const newRequestNo = await getUniqueRequestNo();
      // alert(`status,${status}`);
      const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
          FileName: String(uploadResult.data.Name),
          FileSize: String(uploadResult.data.Length),
          FileVersion: String(uploadResult.data.MajorVersion),
          CurrentFolderPath: String(currentfolderpath.folderpath),
          FileUID: String(uploadResult.data.UniqueId),
          CurrentUser: String(currentUserEmailRef.current),
          SiteID: String(currentfolderpath.siteID),
          Status: status,
          FilePreviewURL : String(previewUrl),
          DocumentLibraryName:String(currentfolderpath.DocumentLibrary),
          SiteName : String(currentfolderpath.Entity),
          MyRequest: true,
          Processname : 'New File Request',
          // RequestNo: `DMS-${uploadResult.data.UniqueId}`
          RequestNo: newRequestNo
      });
      console.log(newItem, "New item added FileMaster");
 
     
      if(IsApproval === true){
        const AddIteminDMSFileApprovalList = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
          SiteName : String(currentfolderpath.Entity),  
           DocumentLibraryName : String(currentfolderpath.DocumentLibrary),
           RequestedBy  : String(currentUserEmailRef.current),
           FileName: String(uploadResult.data.Name),
           FileUID: String(uploadResult.data.UniqueId),
           FilePreviewUrl: String(previewUrl),
           Status: String('Pending'),
           FolderPath : String(currentfolderpath.folderpath),
           ApproveAction : String('Submitted'),
           ApprovedLevel : 1,
           RequestNo: newRequestNo,
           Processname : 'New File Request',
      })
      }
 
 
    // console.log(AddIteminDMSFileApprovalList, "New item added to DMSFileApprovalList");
 
    if(newItem ){
      Deletemedia()
      setTimeout(() => {
        location.reload()
        onReturnToMain(); // Call onReturnToMain after 3 seconds
    }, 3000); // 3000 milliseconds = 3 seconds
     }
 
   
 
  }catch (error) {
      console.error("Error during submission:", error);
  }
 
 
};

const Deletemedia = () => {
 
  Swal.fire({
    title: "Success",
    text: "File uploaded successfully",
    icon: "success"
  });


 setTimeout(() => {
    Swal.close(); // Close the pop-up
    onReturnToMain(); // Call onReturnToMain if needed
  }, 3000); // 3000 milliseconds = 3 seconds

}

 

// const handleSubmit=async (event:any)=>{
          
//           event.preventDefault();
//           console.log("Button clicked");


//           const inputs = document.querySelectorAll('.dynamic-input');
//           const formSelector = document.getElementById("formSelector") as HTMLFormElement;
//           if (!formSelector.checkValidity()) {
//               formSelector.reportValidity(); // Show validation errors
//               return;
//             }

//           // Prepare the payload for SharePoint dynamically
//           const payload: any = {};
//           inputs.forEach((input) => {
//               const inputElement = input as HTMLInputElement;
//               const fieldName = inputElement.id; 

//               // Based on input type, store the correct value
//               if (inputElement.type === "checkbox") {
//                   payload[fieldName] = inputElement.checked;
//               }else if(fieldName === 'fileInput'){
//                   console.log("skip");
//               }else{
//                   payload[fieldName] = inputElement.value;
//               }

//           });
  


  

//           try {
//                   console.log("payload",payload);
//                   console.log("SiteID",data.siteID);
//                   const testidsub = await sp.site.openWebById(data.siteID);
//                   // console.log("subsite context",testidsub)
          
//                   const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(data.folderpath)
//                   console.log("Current Path",documentLibraryInWhichWeUploadTheFile)
                  
//                   const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
//                   console.log("File uploaded successfully", uploadResult.data.Name);
                  
//                   await new Promise((resolve) => {
//                     console.log(uploadResult , "uploadResult")
//                     setTimeout(resolve, 1000)
//                   });

//                   // Fetch the associated list item for the uploaded file
//                   const listItem = await uploadResult.file.getItem();
//                   console.log("List item before upJdate:", listItem);

//                   // Update the metadata on the file's list item
//                   await listItem.update(payload);
//                   console.log("File metadata updated successfully with:", payload);
//                   console.log(SubsiteID , "SubsiteID")
                  
//                   const newItem = await sp.web.lists.getByTitle(`DMS${data.Entity}FileMaster`).items.add({
              
//                     FileName: String(uploadResult.data.Name),          // Example: Set the Title field
//                     FileSize: String(uploadResult.data.Length),    // Example: Set the FileSize field
//                     FileVersion: String(uploadResult.data.MajorVersion),// Example: Set the FileVersion field
//                     CurrentFolderPath : String(currentPath),
//                     FileUID: String(uploadResult.data.UniqueId),// Example: Set the FileVersion field
//                     CurrentUser : String(currentUserEmailRef.current),
//                     SiteID : String(SubsiteID),
//                     Status:String("Pending")
//             });
//                     console.log(newItem, "Today")
//           } catch (error) {
//                   console.log("Error From Adding Field Name",error);
//           } 
// }

    // const ArgPoc=()=>{
      
    // }
    
    return (
      <>
          <button className='BackButton me-3 mb-3' 
          // onClick={(event) => {
          //       onReturnToMain();
          //       // myRequest(event);
          // }}
          onClick={()=>{location.reload() ;onReturnToMain()}}
          > Back 
          </button>
          <div className="container mt-3 UploadFileCont">
              <div className='main-containeruploadfile'>
                      <div className='column column1 p-3'>
                          <form id='formSelector'>
                              <h1>Upload file</h1>
                              {/* <div className="uploadfile">
                                  <input type="file" id="fileInput" onChange={handleFileChange} />
                              </div> */}
                          </form>
                      </div>
                      <div className='column column2 p-3'>
                          <h1>File Preview</h1>
                          {/* <div id="iframeContainer" style={{ position: 'relative' }}>
    <iframe id="filePreview" style={{ display: 'none' }} width="100%" height="100%"></iframe>
    <div id="spinner">Loading...</div>
</div> */}
<div id="iframeContainer" style={{ position: 'relative' }}>
    <iframe id="filePreview" style={{ display: 'none' }} width="100%" height="100%"></iframe>
    <div id="spinner">Loading...</div>
</div>

                          {/* <div id="spinner" style={{display: "none"}}>Loading...</div>
                          <iframe id="filePreview" width="100%" height="400"></iframe> */}
                      </div>
              </div>
          </div>
      </>
      // <div className="uploadfile">
      //   <input type="file" id="fileInput" onChange={handleFileChange} />
      //   <div id="spinner" style={{display: "none"}}>Loading...</div>
      //   <iframe id="filePreview" width="100%" height="500"></iframe>

      //   <div id="fileMetadata"></div>
      // </div>
    );
  }
export default UploadFile;
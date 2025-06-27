// @ts-ignore
import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import './uploadfilecss'
import * as XLSX from 'xlsx';
import './uploadfilecss'
import Swal from 'sweetalert2';
let showbulkupload: any;
let IsApproval: any
let status: any;
let buttontext: any = 'Submit'
interface UploadFileProps {
  currentfolderpath: { [key: string]: string };
  onReturnToMain: () => void;
  // myRequest: (event: React.MouseEvent<HTMLButtonElement>) => void;
}

let previewURLN;
const submitButton = document.createElement('button');
submitButton.textContent = buttontext
submitButton.id = "submitBtn";
submitButton.type = "submit";
submitButton.style.display = 'none';

const UploadFile: React.FC<UploadFileProps> = ({ currentfolderpath, onReturnToMain }) => {
  const sp: SPFI = getSP();
  let locationPath = window.location.pathname.match(/\/sites\/[^\/]+/)[0];
  // check whether folder is private or public and save state

  const [showBulkUpload, setShowBulkUpload] = useState<boolean | null>(null);
  const [isFinalUploading, setIsFinalUploading] = useState(false);
  const [isChecked, setIsChecked] = useState(false);
  const checkfolderprivace = async () => {
    const folderItems = await sp.web.lists.getByTitle("DMSPreviewFormMaster")
      .items.filter(`DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' and SiteName eq '${currentfolderpath.Entity}' and IsDocumentLibrary eq 1`).select("IsApproval", "IsPrivate")();
    console.log("folderItems", folderItems);
    showbulkupload = folderItems[0].IsApproval;
    setShowBulkUpload(folderItems[0].IsApproval)

    IsApproval = folderItems[0].IsApproval;

    console.log('currentfolderpath', currentfolderpath)
  }
  useEffect(() => {
    checkfolderprivace();
  }, []);


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

  const getcurrentuseremail = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    // console.log(currentUserEmailRef.current, "currentuser")
  }

  useEffect(() => {
    getcurrentuseremail()
    setData({ ...data, ...currentfolderpath });

  }, []);


  console.log(data, "data")
  console.log(data.Entity, "entity")
  const SubsiteID = data.siteID
  const currentPath = data.folderpath;
  const documentLibraryName = data.DocumentLibrary;
  console.log("documentLibraryName", documentLibraryName)


  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setIsUploading(true);
    const file = event.target.files![0];

    // upload file with validation like name and extension type
    // if (file) {
    //   const allowedExtensions = ['doc', 'docx', 'ppt', 'pdf', 'xls', 'xlsx', 'txt', 'png', 'jpg', 'jpeg'];
    //   const fileExtension = file.name.split('.').pop()?.toLowerCase(); // Extract the file extension
    //   const baseFileName = file.name.substring(0, file.name.lastIndexOf('.')); // Extract the file name before the extension
    //   const invalidCharacters = /[^a-zA-Z0-9 -]/g; // Allowed: letters, numbers, spaces, and hyphens

    //   // Check for invalid file extensions
    //   if (!fileExtension || !allowedExtensions.includes(fileExtension)) {
    //     Swal.fire({
    //       icon: 'error',
    //       title: 'Invalid file type',
    //       text: 'Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are allowed.',
    //     });
    //     event.target.value = ''; // Clear the input
    //   }
    //   // Check for invalid characters in the file name (excluding extension)
    //   else if (invalidCharacters.test(baseFileName)) {
    //     Swal.fire({
    //       icon: 'error',
    //       title: 'Invalid file name',
    //       text: 'File names can only contain letters, numbers, spaces, and hyphens.',
    //     });
    //     event.target.value = ''; // Clear the input
    //   }
    //   // If all validations pass, upload the file
    //   else {
    //     uploadFile(file);
    //   }
    // }

    if (file) {
      // Directly upload the file without any validation
      uploadFile(file);
    }
  };
  // const uploadFile = async (file: File) => {
  //   try {

  //     const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
  //     const uploadResult = await folder.files.addChunked(file.name, file);
  //     console.log("File uploaded successfully", uploadResult);

  //     // Generate the preview URL dynamically
  //     const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
  //     const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
  //     submitButton.disabled = false; // Disable the button


  //      previewFile(previewUrl);
  //   } catch (error) {
  //     console.error("Error uploading file:", error);
  //   }
  // };

  const uploadFile = async (file: File) => {
    try {
      const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
      const files = await folder.files();

      const originalFileName = file.name;
      const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
      const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      let uniqueFileName = originalFileName;
      let counter = 1;

      while (files.some(f => f.Name === uniqueFileName)) {
        uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
        counter++;
      }

      // Encode the file name to handle special characters
      const encodedFileName = encodeURIComponent(uniqueFileName);

      const uploadResult = await folder.files.addChunked(encodedFileName, file);
      console.log("File uploaded successfully", uploadResult);

      // Generate the preview URL dynamically
      const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
      const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
      submitBtn.disabled = false; // Enable the button
      setIsUploading(false);
      previewFile(previewUrl, "singleUpload");
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

      const previewUrl = `${siteUrl}${locationPath}/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    console.log("Generated Preview URL:", previewUrl);
    if (previewUrl) {
      console.log("enter herr")
      const deletebut = document.getElementById('closeCommand') as HTMLElement
      if (deletebut) {
        console.log(" here ", deletebut)
      }
    }
    return previewUrl;
  };


  const previewFile = async (previewUrl: string, flag: string) => {
    try {
      console.log("Previewing file at URL:", previewUrl);
      const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
      const spinner = document.getElementById("spinner") as HTMLElement;
      const submitButton = document.getElementById("submitBtn") as HTMLButtonElement;
      // Show the spinner and hide the iframe initially
      spinner.style.display = "block";
      iframe.style.display = "none";
      iframe.src = previewUrl;
      submitButton.style.display = 'none'
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
                excelToolbar.style.display = "none"
              }
              if (button) {
                console.log("Hiding the OneUpCommandBar element");
                button.style.display = "none";


                spinner.style.display = "none";
                iframe.style.display = "block";

                // Ensure submit button is shown only once in case of single upload
                if (flag === "singleUpload" && submitButton && submitButton.style.display !== "block") {
                  submitButton.style.display = "block";
                }

              } else {
                console.log("OneUpCommandBar not found, rechecking...");
              }

              const helpbutton = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
              if (helpbutton) {
                helpbutton.style.display = "none"
              }
            }
          } catch (error) {
            console.error("Error accessing iframe content:", error);
          }

          // Stop rechecking once the preview is fully loaded
          if (spinner.style.display === "none") return;
          setTimeout(checkAndHideButton, 100);
        };


        checkAndHideButton();
      };
    } catch (error) {
      console.error("Error previewing file:", error);
    }

  };


  //  handle bulk file


  const [uploadedFiles, setUploadedFiles] = useState<{ name: string; url: string, file: File }[]>([]); // Store uploaded files for preview
  const [isUploading, setIsUploading] = useState(false);

  // const handlebulkFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  //   const files = event.target.files; // Get all selected files

  //   if (!files || files.length === 0) return; // Exit if no files are selected

  //   const allowedExtensions = ['doc', 'docx', 'ppt', 'pdf', 'xls', 'xlsx', 'txt', 'png', 'jpg', 'jpeg'];
  //   const invalidCharacters = /[^a-zA-Z0-9 -]/g; // Allowed: letters, numbers, spaces, and hyphens

  //   const validFiles: File[] = []; // Store valid files

  //   for (let i = 0; i < files.length; i++) {
  //     const file = files[i];
  //     const fileExtension = file.name.split('.').pop()?.toLowerCase(); // Extract file extension
  //     const baseFileName = file.name.substring(0, file.name.lastIndexOf('.')); // Extract filename before extension

  //     // Validate file extension
  //     if (!fileExtension || !allowedExtensions.includes(fileExtension)) {
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file type',
  //         text: `File "${file.name}" is not allowed. Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are accepted. and special character is not allowed`,
  //       });
  //       continue; // Skip this file and move to the next
  //     }

  //     // Validate file name
  //     if (invalidCharacters.test(baseFileName)) {
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file name',
  //         text: `File "${file.name}" contains invalid characters. Only letters, numbers, spaces, and hyphens are allowed.`,
  //       });
  //       continue; // Skip this file
  //     }

  //     // If the file passes all checks, add it to the validFiles array
  //     validFiles.push(file);
  //   }

  //   // If there are valid files, proceed with upload
  //   if (validFiles.length > 0) {
  //     if (validFiles.length === 1) {
  //       uploadFile(validFiles[0]); // Upload single file
  //     } else {
  //       const fileList = new DataTransfer();
  //       validFiles.forEach((file) => fileList.items.add(file));
  //       bulkUploadFile(fileList.files); 
  //       // Upload multiple files
  //       //  bulkUploadFile(validFiles); // Upload multiple files
  //     }
  //   }

  //   // Clear input field to allow re-selection of the same files
  //   event.target.value = '';
  // };
  // const handlebulkFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  //   const files = event.target.files; // Get all selected files

  //   if (!files || files.length === 0) return; // Exit if no files are selected

  //   const allowedExtensions = ['doc', 'docx', 'ppt', 'pdf', 'xls', 'xlsx', 'txt', 'png', 'jpg', 'jpeg'];
  //   const invalidCharacters = /[^a-zA-Z0-9 -]/g; // Allowed: letters, numbers, spaces, and hyphens

  //   const validFiles: File[] = []; // Store valid files

  //   for (let i = 0; i < files.length; i++) {
  //     const file = files[i];
  //     const fileExtension = file.name.split('.').pop()?.toLowerCase(); // Extract file extension
  //     const baseFileName = file.name.substring(0, file.name.lastIndexOf('.')); // Extract filename before extension

  //     // Validate file extension
  //     if (!fileExtension || !allowedExtensions.includes(fileExtension)) {
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file type',
  //         text: `File "${file.name}" is not allowed. Only DOC, DOCX, PPT, PDF, XLS, XLSX, TXT, PNG, JPG, JPEG are accepted. and special character is not allowed`,
  //       });
  //       continue; // Skip this file and move to the next
  //     }

  //     // Validate file name
  //     if (invalidCharacters.test(baseFileName)) {
  //       Swal.fire({
  //         icon: 'error',
  //         title: 'Invalid file name',
  //         text: `File "${file.name}" contains invalid characters. Only letters, numbers, spaces, and hyphens are allowed.`,
  //       });
  //       continue; // Skip this file
  //     }

  //     // If the file passes all checks, add it to the validFiles array
  //     validFiles.push(file);
  //   }

  //   // If there are valid files, proceed with upload
  //   if (validFiles.length > 0) {
  //     if (validFiles.length === 1) {
  //       uploadFile(validFiles[0]); // Upload single file
  //     } else {
  //       const fileList = new DataTransfer();
  //       validFiles.forEach((file) => fileList.items.add(file));
  //       bulkUploadFile(fileList.files); 
  //       // Upload multiple files
  //       //  bulkUploadFile(validFiles); // Upload multiple files
  //     }
  //   }

  //   // Clear input field to allow re-selection of the same files
  //   event.target.value = '';
  // };

  const handlebulkFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setIsUploading(true);

    const files = event.target.files; // Get all selected files

    if (!files || files.length === 0) return; // Exit if no files are selected

    const validFiles: File[] = []; // Store valid files

    for (let i = 0; i < files.length; i++) {
      validFiles.push(files[i]); // Add all selected files without validation
    }

    // If there are valid files, proceed with upload
    if (validFiles.length > 0) {
      if (validFiles.length === 1) {
        const fileList = new DataTransfer();
        validFiles.forEach((file) => fileList.items.add(file));
        // uploadFile(validFiles[0]); // Upload single file
        bulkUploadFile(fileList.files);
      } else {
        const fileList = new DataTransfer();
        validFiles.forEach((file) => fileList.items.add(file));
        bulkUploadFile(fileList.files);
        // Upload multiple files
        //  bulkUploadFile(validFiles); // Upload multiple files
      }
    }

    // Clear input field to allow re-selection of the same files
    event.target.value = '';
  };


  const bulkUploadFile = async (files: FileList) => {
    console.log("Bulk uploading files:", files);

    const uploadedFilesList: { name: string; url: string, file: File }[] = [];

    for (let i = 0; i < files.length; i++) {
      try {
        const file = files[i];
        console.log(`Uploading file.name: ${file.name}`);
        console.log(`Uploading file: ${file.name}`);
        const folder = sp.web.getFolderByServerRelativePath("DMSOrphanDocs");

        const uploadResult = await folder.files.addChunked(file.name, file);
        console.log(`File ${file.name} uploaded successfully`, uploadResult);

        // Generate the preview URL dynamically
        const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);

        uploadedFilesList.push({ name: file.name, url: previewUrl, file: file });
      } catch (error) {
        console.error(`Error uploading file ${files[i].name}:`, error);
      }
    }

    // setUploadedFiles(uploadedFilesList); // Update state with all uploaded files
    setUploadedFiles((prevFiles) => [...prevFiles, ...uploadedFilesList]);
    setIsUploading(false);
    const submitButton = document.getElementById("submitBtn2") as HTMLButtonElement;
    if (submitButton) submitButton.style.display = "block";
  };


  const handlePreview = (previewUrl: string) => {
    previewFile(previewUrl, "bulkUpload"); // Call your preview function
  };





  React.useEffect(() => {
    const loadFormOptions = async () => {
      try {
        const documentLibraryFields = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "ColumnType", "IsRequired", "IsRename")
          .filter(
            `SiteName eq '${currentfolderpath.Entity}' 
            and DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' 
            and AddorRemoveThisColumn eq  'Add To Library' and IsInProgress eq 0`)();

        console.log("Document Library Fields", documentLibraryFields);
        // end

        const formSelector = document.getElementById("formSelector");

        const uploadFileDiv = document.createElement('div');

        const createElement = (fieldName: string, type: string, required: boolean, IsRename: string) => {
          let fName = fieldName
          if (IsRename !== null) {
            fName = IsRename
          }
          const inputContainer = document.createElement("div");
          inputContainer.className = "input-container";

          // Create and set label
          const label = document.createElement("label");
          label.setAttribute("htmlFor", fieldName);
          // label.textContent = fieldName;
          // label.textContent = fName;
          // Append an asterisk if the field is required 
          if (required) {
            const asterisk = document.createElement("span");
            asterisk.textContent = " *";
            asterisk.style.color = "red";
            label.textContent = fName;
            label.appendChild(asterisk);
          } else {
            label.textContent = fName;
          }
          inputContainer.appendChild(label);

          let inputElement: HTMLInputElement | null = null;

          // Dynamically create the input field based on FieldType
          let modifiedType = type?.replace(/\s+/g, '').toLowerCase();
          console.log("modifiedType", modifiedType);

          if (
            modifiedType === "singlelineoftext"
            ||
            modifiedType === "multiplelineoftext"
            ||
            modifiedType === 'text'
          ) {
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
            inputElement.className = "dynamic-input";
            inputElement.id = fieldName;
            // inputElement.required = required.toLowerCase() === "yes"; 
            inputElement.required = required;
            inputContainer.appendChild(inputElement);
            formSelector.appendChild(inputContainer);
          }

          return;
        }


        // start
        documentLibraryFields.forEach((field) => {
          createElement(field.ColumnName, field.ColumnType, field.IsRequired, field.IsRename);
        })
        // end


        // properties of upload file div
        // uploadFileDiv.className="uploadfile";
        uploadFileDiv.className = "input-container";

        // input for upload file
        const uploadFileInput = document.createElement('input');
        uploadFileInput.className = "dynamic-input";
        uploadFileInput.type = "file";
        uploadFileInput.id = "fileInput";
        uploadFileInput.addEventListener('change', (event: any) => handleFileChange(event))

        // Set Label For upload file
        const label = document.createElement("label");
        label.setAttribute("htmlFor", 'fileInput');
        label.textContent = 'Upload File';

        // Add a red asterisk if the field is required
        const asterisk = document.createElement("span");
        asterisk.textContent = " *";
        asterisk.style.color = "red";
        label.appendChild(asterisk);
        uploadFileDiv.appendChild(label);
        uploadFileDiv.appendChild(uploadFileInput);
        formSelector.appendChild(uploadFileDiv);

        // Submit Button property

        submitButton.addEventListener('click', handleSubmit)

        formSelector.appendChild(submitButton);

      } catch (error) {
        console.error("Error loading form options:", error);
      }
    }

    loadFormOptions();
  }, [])

  // const getUniqueRequestNo = async () => {
  //   const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
  //   console.log("Counter Item 0", counterItem);
  //   console.log("Counter Item 1", counterItem.FileCount);
  //   let fileCounter = counterItem.FileCount;

  //   // Increment the counter
  //   fileCounter++;

  //   // Generate the new RequestNo
  //   const newRequestNo = `File${String(fileCounter).padStart(2, '0')}`;

  //   // Update the counter in the CounterList
  //   await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
  //     FileCount: fileCounter
  //   });

  //   return newRequestNo;
  // };

  // const handleSubmit = async (event: any) => {
  //     // alert(` isChecked ${isChecked}`);

  //     event.preventDefault();
  //     console.log("Button clicked");

  //     const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  //     if (!formSelector.checkValidity()) {
  //         formSelector.reportValidity(); // Show validation errors
  //         return;
  //     }

  //     // Prepare the payload for SharePoint dynamically
  //     const inputs = document.querySelectorAll('.dynamic-input');
  //     // console.log("inputs",inputs)


  //     const payload: any = {};

  //     inputs.forEach((input) => {
  //         const inputElement = input as HTMLInputElement;
  //         const fieldName = inputElement.id;
  //         if (!fieldName) return; // Skip if field name is invalid

  //         if (inputElement.type === "checkbox") {
  //             // console.log("fieldName",fieldName.includes(' '));
  //             payload[fieldName] = inputElement.checked;
  //         } else if (inputElement.type !== "file") {
  //             if(inputElement.value === ""){
  //                console.log("skip");
  //             }else{
  //               // if(fieldName.includes(' '))
  //               // console.log("fieldName",fieldName.includes(' '));
  //               payload[fieldName] = inputElement.value;
  //             }

  //         }
  //     });

  //     const fileInput = document.getElementById('fileInput') as HTMLInputElement;
  //     const selectedFile = fileInput?.files?.[0];

  //     if (!selectedFile) {
  //         console.error("No file selected.");

  //         return;
  //     }

  //     try {
  //         console.log("Payload:", payload);
  //         console.log("SiteID:", currentfolderpath.siteID);

  //         const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //         if (!testidsub) throw new Error("Subsite not found.");

  //         const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //         console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

  //         const files = await documentLibraryInWhichWeUploadTheFile.files();

  //         const originalFileName = selectedFile.name;
  //         const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //         const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //         console.log("originalFileName",originalFileName);
  //         console.log("fileExtension",fileExtension);
  //         console.log("baseFileName",baseFileName);
  //         let uniqueFileName = originalFileName;
  //         let counter = 1;

  //         while (files.some(file => file.Name === uniqueFileName)) {
  //           uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //           counter++;
  //         }
  //         console.log(`Unique file name generated: ${uniqueFileName}`);
  //         // const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
  //         console.log("unique file name in single", uniqueFileName);
  //         console.log("selectedFile in single", selectedFile);
  //         const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, selectedFile);
  //         console.log("File uploaded successfully", uploadResult.data.Name);
  //         const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
  //         submitBtn.disabled = true;
  //         submitBtn.innerText = "Submitting..."; // Optional: Change button text to indicate progress

  //         const listItem = await uploadResult.file.getItem();
  //         console.log("ListItems ",listItem);

  //         const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //         const siteUrl = window.location.origin;
  //         const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
  //         console.log(encodedFilePath , "encodedFilePath")
  //           // const previewUrl = `${siteUrl}/sites/ededms/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
   //            const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

  //         console.log("Generated Preview URL:", previewUrl);
  //         if (!listItem) throw new Error("List item not found for the uploaded file.");


  //         if(IsApproval === true){
  //           //  alert(`check tru ${IsApproval}`);
  //           status="Pending";
  //         }else if(IsApproval === false){
  //           // alert(`check fas ${IsApproval}`);
  //           status="Auto Approved";
  //         }
  //         (payload as any).Status=status;
  //         await listItem.update(payload);
  //         console.log("File metadata updated successfully with:", payload);

  //         const newRequestNo = await getUniqueRequestNo();
  //         // alert(`status,${status}`);
  //         const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //             FileName: String(uploadResult.data.Name),
  //             FileSize: String(uploadResult.data.Length),
  //             FileVersion: String(uploadResult.data.MajorVersion),
  //             CurrentFolderPath: String(currentfolderpath.folderpath),
  //             FileUID: String(uploadResult.data.UniqueId),
  //             CurrentUser: String(currentUserEmailRef.current),
  //             SiteID: String(currentfolderpath.siteID),
  //             Status: status,
  //             FilePreviewURL : String(previewUrl),
  //             DocumentLibraryName:String(currentfolderpath.DocumentLibrary),
  //             SiteName : String(currentfolderpath.Entity),
  //             MyRequest: true,
  //             Processname : 'New File Request',
  //             // RequestNo: `DMS-${uploadResult.data.UniqueId}`
  //             RequestNo: newRequestNo
  //         });
  //         console.log(newItem, "New item added FileMaster");


  //         if(IsApproval === true){
  //           const AddIteminDMSFileApprovalList = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
  //             SiteName : String(currentfolderpath.Entity),  
  //              DocumentLibraryName : String(currentfolderpath.DocumentLibrary),
  //              RequestedBy  : String(currentUserEmailRef.current),
  //              FileName: String(uploadResult.data.Name),
  //              FileUID: String(uploadResult.data.UniqueId),
  //              FilePreviewUrl: String(previewUrl),
  //              Status: String('Pending'),
  //              FolderPath : String(currentfolderpath.folderpath),
  //              ApproveAction : String('Submitted'),
  //              ApprovedLevel : 1,
  //              RequestNo: newRequestNo,
  //              Processname : 'New File Request',
  //         })
  //         }

  //       if(newItem ){
  //         Deletemedia()
  //         setTimeout(() => {
  //           location.reload()
  //           onReturnToMain();
  //       }, 3000);
  //        }



  //     }catch (error) {
  //         console.error("Error during submission:", error);
  //     }




  // };
  const handleSubmit = async (event: any) => {
    event.preventDefault();
    console.log("Button clicked");

    const formSelector = document.getElementById("formSelector") as HTMLFormElement;
    if (!formSelector.checkValidity()) {
      formSelector.reportValidity(); // Show validation errors
      return;
    }
    const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
    submitBtn.disabled = true;
    submitBtn.innerText = "Submitting...";

    const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
    const spinner = document.getElementById("spinner") as HTMLElement;
    spinner.style.display = "none";
    iframe.style.display = "none";
    setIsFinalUploading(true);
    // Prepare the payload for SharePoint dynamically
    const inputs = document.querySelectorAll('.dynamic-input');
    const payload: any = {};

    inputs.forEach((input) => {
      const inputElement = input as HTMLInputElement;
      const fieldName = inputElement.id;
      if (!fieldName) return; // Skip if field name is invalid

      if (inputElement.type === "checkbox") {
        payload[fieldName] = inputElement.checked;
      } else if (inputElement.type !== "file") {
        if (inputElement.value !== "") {
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

      const originalFileName = selectedFile.name;
      const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
      const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      let uniqueFileName = originalFileName;
      let counter = 1;

      while (files.some(file => file.Name === uniqueFileName)) {
        uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
        counter++;
      }
      console.log(`Unique file name generated: ${uniqueFileName}`);

      const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, selectedFile);
      console.log("File uploaded successfully", uploadResult.data.Name);

      // const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
      // submitBtn.disabled = true;
      // submitBtn.innerText = "Submitting..."; // Optional: Change button text to indicate progress

      const listItem = await uploadResult.file.getItem();
      console.log("ListItems ", listItem);

      const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
      const siteUrl = window.location.origin;
      const encodeSharePointURL = (url: string) => {
        return encodeURIComponent(url)
          .replace(/'/g, "%27")
          .replace(/-/g, "%2D")
          .replace(/_/g, "%5F")
          .replace(/\./g, "%2E")
          .replace(/!/g, "%21")
          .replace(/\*/g, "%2A")
          .replace(/\(/g, "%28")
          .replace(/\)/g, "%29")
          .replace(/~/g, "%7E")
          .replace(/@/g, "%40")
          .replace(/\$/g, "%24")
          .replace(/,/g, "%2C")
          .replace(/;/g, "%3B")
          .replace(/:/g, "%3A")
          .replace(/\+/g, "%2B")
          .replace(/=/g, "%3D")
          .replace(/\?/g, "%3F")
          .replace(/\//g, "%2F")
          .replace(/#/g, "%23")
          .replace(/&/g, "%26");
      };

      // Usage
      const encodedFilePath = encodeSharePointURL(uploadResult.data.ServerRelativeUrl);
      const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeSharePointURL(parentFolder)}`;

      // Encode the file path for the preview URL
      // const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl).replace(/'/g, "%27");
      // const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

      console.log("Generated Preview URL:", previewUrl);
      if (!listItem) throw new Error("List item not found for the uploaded file.");

      let status = "";
      if (IsApproval === true) {
        status = "Pending";
      } else if (IsApproval === false) {
        status = "Auto Approved";
      }
      (payload as any).Status = status;
      await listItem.update(payload);
      console.log("File metadata updated successfully with:", payload);

      const newRequestNo = await getUniqueRequestNo();
      const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
        FileName: String(uploadResult.data.Name),
        FileSize: String(uploadResult.data.Length),
        FileVersion: String(uploadResult.data.MajorVersion),
        CurrentFolderPath: String(currentfolderpath.folderpath),
        FileUID: String(uploadResult.data.UniqueId),
        CurrentUser: String(currentUserEmailRef.current),
        SiteID: String(currentfolderpath.siteID),
        Status: status,
        FilePreviewURL: String(previewUrl),
        DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
        SiteName: String(currentfolderpath.Entity),
        MyRequest: true,
        Processname: 'New File Request',
         RequestNo: newRequestNo,
      });
      console.log(newItem, "New item added FileMaster");

      if (IsApproval === true) {
        await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
          SiteName: String(currentfolderpath.Entity),
          DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
          RequestedBy: String(currentUserEmailRef.current),
          FileName: String(uploadResult.data.Name),
          FileUID: String(uploadResult.data.UniqueId),
          FilePreviewUrl: String(previewUrl),
          Status: String('Pending'),
          FolderPath: String(currentfolderpath.folderpath),
          ApproveAction: String('Submitted'),
          ApprovedLevel: 1,
          RequestNo: newRequestNo,
          Processname: 'New File Request',
          CurrentLevel : 1
          });
      }

      if (newItem) {
        setIsFinalUploading(false);
        Deletemedia();
        setTimeout(() => {
          location.reload();
          onReturnToMain();
        }, 3000);
      }
    } catch (error) {
      console.error("Error during submission:", error);
    }
  };

  // const getUniqueRequestNo = async (incrementBy: number) => {
  //   try {
  //     const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
  //     console.log("Counter Item:", counterItem);

  //     let fileCounter = counterItem.FileCount;

  //     // Generate request numbers for all files
  //     const requestNumbers = [];
  //     for (let i = 0; i < incrementBy; i++) {
  //       fileCounter++;
  //       requestNumbers.push(`File${String(fileCounter).padStart(2, '0')}`);
  //     }

  //     // Update counter only once
  //     await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
  //       FileCount: fileCounter
  //     });

  //     return requestNumbers;
  //   } catch (error) {
  //     console.error("Error updating request number:", error);
  //     return [];
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
  const handleSubmitBulk = async (event: any) => {
    event.preventDefault();
    console.log("Bulk upload button clicked");
    const submitBtn = document.getElementById("submitBtn2") as HTMLButtonElement;
    submitBtn.disabled = true;
    submitBtn.innerText = "Submitting...";

    const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
    const spinner = document.getElementById("spinner") as HTMLElement;
    spinner.style.display = "none";
    iframe.style.display = "none";
    setIsFinalUploading(true);
    try {
      console.log("SiteID:", currentfolderpath.siteID);

      const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
      if (!testidsub) throw new Error("Subsite not found.");

      const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
      console.log("Current Path:", documentLibrary);

      const files = await documentLibrary.files();
      const existingFileNames = new Set(files.map(file => file.Name)); // Faster lookup
      const siteUrl = window.location.origin;

      // Function to generate a unique filename
      const generateUniqueFileName = (originalFileName: string) => {
        const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
        const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
        let uniqueFileName = originalFileName;
        let counter = 1;

        while (existingFileNames.has(uniqueFileName)) {
          uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
          counter++;
        }
        existingFileNames.add(uniqueFileName); // Add new name to the set
        return uniqueFileName;
      };

      // Upload all files in parallel
      const uploadPromises = uploadedFiles.map(async ({ name, file }) => {
        try {
          const uniqueFileName = generateUniqueFileName(name);
          console.log(`Uploading file: ${uniqueFileName}`);

          const uploadResult = await documentLibrary.files.addChunked(uniqueFileName, file);
          console.log(`${name} uploaded successfully`);

          const listItem = await uploadResult.file.getItem();
          if (!listItem) throw new Error("List item not found for the uploaded file.");

          const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
          const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
          const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

          // Update metadata
          await listItem.update({ Status: "Auto Approved" });

          // Prepare new file entry
          // const newRequestNo = await getUniqueRequestNo(1);
          return {
            FileName: String(uploadResult.data.Name),
            FileSize: String(uploadResult.data.Length),
            FileVersion: String(uploadResult.data.MajorVersion),
            CurrentFolderPath: String(currentfolderpath.folderpath),
            FileUID: String(uploadResult.data.UniqueId),
            CurrentUser: String(currentUserEmailRef.current),
            SiteID: String(currentfolderpath.siteID),
            Status: "Auto Approved",
            FilePreviewURL: String(previewUrl),
            DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
            SiteName: String(currentfolderpath.Entity),
            MyRequest: true,
            Processname: 'New File Request',
            // RequestNo: newRequestNo,
          };
        } catch (error) {
          console.error(`Error uploading file ${name}:`, error);
          return null;
        }
      });

      // Wait for all uploads to finish
      const uploadedItems = await Promise.all(uploadPromises);
      const validItems = uploadedItems.filter(item => item !== null);
      console.log(validItems, "validItemsvalidItems")
      console.log(typeof (validItems), " here i am checking validItemsvalidItems")
      // Batch insert all successfully uploaded files into FileMaster list
      for (const item of validItems) {
        await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add(item);
      }

      // Cleanup and refresh UI
      setIsFinalUploading(false);
      Deletemedia();
      setTimeout(() => {
        location.reload();
        onReturnToMain();
      }, 3000);

    } catch (error) {
      console.error("Error in bulk upload:", error);
    }
  };


  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");

  //   try {
  //     console.log("SiteID:", currentfolderpath.siteID);

  //     const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //     if (!testidsub) throw new Error("Subsite not found.");

  //     const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //     console.log("Current Path:", documentLibrary);

  //     const files = await documentLibrary.files();
  //     const existingFileNames = new Set(files.map(file => file.Name)); // Faster lookup
  //     const siteUrl = window.location.origin;

  //     // Function to generate a unique filename
  //     const generateUniqueFileName = (originalFileName: string) => {
  //       const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //       const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //       let uniqueFileName = originalFileName;
  //       let counter = 1;

  //       while (existingFileNames.has(uniqueFileName)) {
  //         uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //         counter++;
  //       }
  //       existingFileNames.add(uniqueFileName); // Add new name to the set
  //       return uniqueFileName;
  //     };

  //     // Upload all files in parallel
  //     const uploadPromises = uploadedFiles.map(async ({ name, file }) => {
  //       try {
  //         const uniqueFileName = generateUniqueFileName(name);
  //         console.log(`Uploading file: ${uniqueFileName}`);

  //         const uploadResult = await documentLibrary.files.addChunked(uniqueFileName, file);
  //         console.log(`${name} uploaded successfully`);

  //         const listItem = await uploadResult.file.getItem();
  //         if (!listItem) throw new Error("List item not found for the uploaded file.");

  //         const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //         const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
  //         const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

  //         // Update metadata
  //         await listItem.update({ Status: "Auto Approved" });

  //         // Prepare new file entry
  //         const newRequestNo = await getUniqueRequestNo();
  //         return {
  //           FileName: String(uploadResult.data.Name),
  //           FileSize: String(uploadResult.data.Length),
  //           FileVersion: String(uploadResult.data.MajorVersion),
  //           CurrentFolderPath: String(currentfolderpath.folderpath),
  //           FileUID: String(uploadResult.data.UniqueId),
  //           CurrentUser: String(currentUserEmailRef.current),
  //           SiteID: String(currentfolderpath.siteID),
  //           Status: "Auto Approved",
  //           FilePreviewURL: String(previewUrl),
  //           DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //           SiteName: String(currentfolderpath.Entity),
  //           MyRequest: true,
  //           Processname: 'New File Request',
  //           RequestNo: newRequestNo,
  //         };
  //       } catch (error) {
  //         console.error(`Error uploading file ${name}:`, error);
  //         return null;
  //       }
  //     });

  //     // Wait for all uploads to finish
  //     const uploadedItems = await Promise.all(uploadPromises);
  //     const validItems = uploadedItems.filter(item => item !== null);

  //     // ✅ Add each item one by one (Fixing the error)
  //     for (const item of validItems) {
  //       await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add(item);
  //     }
  //     console.log("All files added to FileMaster successfully.");

  //     // Cleanup and refresh UI
  //     Deletemedia();
  //     // setTimeout(() => {
  //     //   location.reload();
  //     //   onReturnToMain();
  //     // }, 3000);

  //   } catch (error) {
  //     console.error("Error in bulk upload:", error);
  //   }
  // };


  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");
  //    try {
  //     console.log("SiteID:", currentfolderpath.siteID);

  //     const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //     if (!testidsub) throw new Error("Subsite not found.");

  //     // Get the folder in the static document library
  //     // const documentLibrary = testidsub.web.getFolderByServerRelativePath(documentLibraryName);
  //     const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //     console.log("Current Path:", documentLibrary);
  //     console.log("uploadedFiles",uploadedFiles)

  //     const files = await documentLibrary.files();
  //     const siteUrl = window.location.origin;
  //     // Loop through each file and upload it
  //     for (let i = 0; i < uploadedFiles.length; i++) {
  //       const { name, url,file } = uploadedFiles[i];

  //       try {
  //         console.log(`Uploading file: ${name}`);
  //         console.log(`Uploading file url: ${url}`);
  //         const originalFileName = name;
  //         const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //         const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //         let uniqueFileName = originalFileName;
  //         let counter = 1;

  //         while (files.some(file => file.Name === uniqueFileName)) {
  //           uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //           counter++;
  //         }
  //         console.log(`Unique file name generated: ${uniqueFileName}`);

  //         // Fetch the file from the URL and create a File object
  //         // const fileBlob = await fetch(url).then((res) => res.blob());
  //         // Fetch the file and convert it to a Blob

  //         // const response = await fetch(url);
  //         // if (!response.ok) {
  //         //     throw new Error(`Failed to fetch file: ${response.statusText}`);
  //         // }

  //         // const fileBlob = await response.blob();
  //         // console.log("Fetched Blob:", fileBlob);

  //         // // Check if Blob is valid
  //         // if (!fileBlob || fileBlob.size === 0) {
  //         //     throw new Error("File Blob is empty or corrupted.");
  //         // }

  //         // // Convert Blob to File
  //         // const fileToUpload = new File([fileBlob], name, { type: fileBlob.type });
  //         // console.log("File to upload:", fileToUpload);

  //         // const fileToUpload = new File([fileBlob], name, { type: fileBlob.type });
  //         // console.log("File to upload in loop:", fileToUpload);

  //         // Upload the file using the add method (or addChunked if needed)
  //         // console.log("File to upload in bulk :", fileToUpload);
  //         // console.log("File name to upload in bulk :", name);

  //         // const uploadResult = await documentLibrary.files.addChunked(name, fileToUpload);
  //         // console.log(`${name} uploaded successfully to ${documentLibraryName}`, uploadResult);

  //         // Generate the preview URL for the uploaded file
  //         // const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
  //         // console.log("Generated Preview URL:", previewUrl);

  //         // Optionally, update metadata (e.g., file status)
  //                 // Generate preview URL
  //         // Upload the file using addChunked (use add for small files)
  //         const uploadResult = await documentLibrary.files.addChunked(uniqueFileName, file);
  //         console.log(`${name} uploaded successfully`, uploadResult);

  //         const listItem = await uploadResult.file.getItem();
  //         if (!listItem) throw new Error("List item not found for the uploaded file.");

  //         const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //         const encodeSharePointURL = (url: string) => {
  //           return encodeURIComponent(url)
  //               .replace(/'/g, "%27")
  //               .replace(/-/g, "%2D")
  //               .replace(/_/g, "%5F")
  //               .replace(/\./g, "%2E")
  //               .replace(/!/g, "%21")
  //               .replace(/\*/g, "%2A")
  //               .replace(/\(/g, "%28")
  //               .replace(/\)/g, "%29")
  //               .replace(/~/g, "%7E")
  //               .replace(/@/g, "%40")
  //               .replace(/\$/g, "%24")
  //               .replace(/,/g, "%2C")
  //               .replace(/;/g, "%3B")
  //               .replace(/:/g, "%3A")
  //               .replace(/\+/g, "%2B")
  //               .replace(/=/g, "%3D")
  //               .replace(/\?/g, "%3F")
  //               .replace(/\//g, "%2F")
  //               .replace(/#/g, "%23")
  //               .replace(/&/g, "%26");
  //       };


  //         const encodedFilePath = encodeSharePointURL(uploadResult.data.ServerRelativeUrl);
  //         const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
  //         console.log("Generated Preview URL:", previewUrl);
  //         if (listItem) {
  //           const status = "Auto Approved";
  //           await listItem.update({ Status: status });
  //           console.log("File metadata updated successfully.");
  //         }
  //         const newRequestNo = await getUniqueRequestNo();
  //         const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //             FileName: String(uploadResult.data.Name),
  //             FileSize: String(uploadResult.data.Length),
  //             FileVersion: String(uploadResult.data.MajorVersion),
  //             CurrentFolderPath: String(currentfolderpath.folderpath),
  //             FileUID: String(uploadResult.data.UniqueId),
  //             CurrentUser: String(currentUserEmailRef.current),
  //             SiteID: String(currentfolderpath.siteID),
  //             Status: String("Auto Approved"),
  //             FilePreviewURL: String(previewUrl),
  //             DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //             SiteName: String(currentfolderpath.Entity),
  //             MyRequest: true,
  //             Processname: 'New File Request',
  //             RequestNo: newRequestNo,
  //           });
  //           console.log(newItem, "New item added to FileMaster");
  //           if(newItem ){
  //             Deletemedia()
  //             setTimeout(() => {
  //               location.reload()
  //               onReturnToMain();
  //           }, 3000);
  //            }
  // 		}catch(error){
  //             console.log("Error in uploading the file infolder",error);

  // 		}

  // 		}
  //         Deletemedia();
  //    } catch (error) {
  //     console.log("Error in bulk upload",error);
  //    }
  //   // try {
  //   //   // Get the SharePoint subsite
  //   //   const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //   //   if (!testidsub) throw new Error("Subsite not found.");

  //   //   // Get the target document library folder
  //   //   const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //   //   console.log("Current Path:", documentLibrary);

  //   //   // Upload files sequentially (one by one)
  //   //   for (const { name, url } of uploadedFiles) {
  //   //     try {
  //   //       console.log(`Uploading file: ${name}`);

  //   //       // Fetch the file directly from the URL
  //   //       const response = await fetch(url);
  //   //       if (!response.ok) throw new Error(`Failed to fetch file: ${name}`);

  //   //       // Use the response blob directly for upload
  //   //       const blob = await response.blob();

  //   //       // Upload the file using addChunked
  //   //       const uploadResult = await documentLibrary.files.addChunked(name, blob);
  //   //       console.log(`File ${name} uploaded successfully`, uploadResult);

  //   //       // Generate the preview URL dynamically
  //   //       const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
  //   //       console.log("Generated Preview URL:", previewUrl);

  //   //       // Update metadata (optional)
  //   //       const status = "Auto Approved";
  //   //       const listItem = await uploadResult.file.getItem();
  //   //       if (listItem) {
  //   //         await listItem.update({ Status: status });
  //   //         console.log("File metadata updated successfully.");
  //   //       }

  //   //       // Add entry to the FileMaster list (optional)
  //   //       const newRequestNo = await getUniqueRequestNo();
  //   //       await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //   //         FileName: String(uploadResult.data.Name),
  //   //         FileSize: String(uploadResult.data.Length),
  //   //         FileVersion: String(uploadResult.data.MajorVersion),
  //   //         CurrentFolderPath: String(currentfolderpath.folderpath),
  //   //         FileUID: String(uploadResult.data.UniqueId),
  //   //         CurrentUser: String(currentUserEmailRef.current),
  //   //         SiteID: String(currentfolderpath.siteID),
  //   //         Status: status,
  //   //         FilePreviewURL: String(previewUrl),
  //   //         DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //   //         SiteName: String(currentfolderpath.Entity),
  //   //         MyRequest: true,
  //   //         Processname: 'New File Request',
  //   //         RequestNo: newRequestNo
  //   //       });
  //   //       console.log("File added to FileMaster successfully.");
  //   //     } catch (error) {
  //   //       console.error(`Error uploading file ${name}:`, error);
  //   //     }
  //   //   }

  //   //   console.log("Bulk upload process completed.");

  //   //   // Cleanup media after upload
  //   //   Deletemedia();

  //   //   // Reload the page once all files are processed
  //   //   setTimeout(() => {
  //   //     location.reload();
  //   //     onReturnToMain();
  //   //   }, 3000);

  //   // } catch (error) {
  //   //   console.error("Error during bulk submission:", error);
  //   // }
  // };
  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");

  //   try {
  //     // Get the SharePoint subsite
  //     const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //     if (!testidsub) throw new Error("Subsite not found.");

  //     // Get the target document library folder
  //     const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //     console.log("Current Path:", documentLibrary);

  //     // Upload files sequentially (one by one)
  //     for (const { name, url } of uploadedFiles) {
  //       try {
  //         console.log(`Uploading file: ${name}`);

  //         // Use the actual file object from uploadedFiles
  //         const file = new File([await fetch(url).then(res => res.blob())], name);

  //         // Upload the file using addChunked
  //         const uploadResult = await documentLibrary.files.addChunked(name, file);
  //         console.log(`File ${name} uploaded successfully`, uploadResult);

  //         // Generate the preview URL dynamically
  //         const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
  //         console.log("Generated Preview URL:", previewUrl);

  //         // Update metadata (optional)
  //         const status = "Auto Approved";
  //         const listItem = await uploadResult.file.getItem();
  //         if (listItem) {
  //           await listItem.update({ Status: status });
  //           console.log("File metadata updated successfully.");
  //         }

  //         // Add entry to the FileMaster list (optional)
  //         const newRequestNo = await getUniqueRequestNo();
  //         await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //           FileName: String(uploadResult.data.Name),
  //           FileSize: String(uploadResult.data.Length),
  //           FileVersion: String(uploadResult.data.MajorVersion),
  //           CurrentFolderPath: String(currentfolderpath.folderpath),
  //           FileUID: String(uploadResult.data.UniqueId),
  //           CurrentUser: String(currentUserEmailRef.current),
  //           SiteID: String(currentfolderpath.siteID),
  //           Status: status,
  //           FilePreviewURL: String(previewUrl),
  //           DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //           SiteName: String(currentfolderpath.Entity),
  //           MyRequest: true,
  //           Processname: 'New File Request',
  //           RequestNo: newRequestNo
  //         });
  //         console.log("File added to FileMaster successfully.");
  //       } catch (error) {
  //         console.error(`Error uploading file ${name}:`, error);
  //       }
  //     }

  //     console.log("Bulk upload process completed.");

  //     // Cleanup media after upload
  //     Deletemedia();

  //     // Reload the page once all files are processed
  //     setTimeout(() => {
  //       location.reload();
  //       onReturnToMain();
  //     }, 3000);

  //   } catch (error) {
  //     console.error("Error during bulk submission:", error);
  //   }
  // };


  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");

  //   try {
  //     console.log("SiteID:", currentfolderpath.siteID);

  //     const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //     if (!testidsub) throw new Error("Subsite not found.");

  //     const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //     console.log("Current Path:", documentLibrary);

  //     const existingFiles = await documentLibrary.files();

  //     // Upload all files in sequence (not parallel)
  //     for (const file of uploadedFiles) {
  //       try {
  //         console.log(`Uploading file: ${file.name}`);

  //         // Fetch file blob
  //         const response = await fetch(file.url);
  //         if (!response.ok) throw new Error(`Failed to fetch file: ${file.name}`);
  //         const blob = await response.blob();

  //         // Generate a unique file name to avoid conflicts
  //         const { name: originalFileName } = file;
  //         const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //         const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //         let uniqueFileName = originalFileName;
  //         let counter = 1;

  //         while (existingFiles.some(f => f.Name === uniqueFileName)) {
  //           uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //           counter++;
  //         }

  //         console.log(`Final filename: ${uniqueFileName}`);

  //         // Upload the file using addChunked()
  //         const uploadResult = await documentLibrary.files.addChunked(uniqueFileName, blob);
  //         console.log("Upload successful:", uploadResult.data.Name);

  //         // Ensure the file is available before updating metadata
  //         await new Promise(resolve => setTimeout(resolve, 1500));

  //         // Get the uploaded file item
  //         const listItem = await uploadResult.file.getItem();
  //         if (!listItem) {
  //           console.error("List item not found for uploaded file.");
  //           continue;
  //         }

  //         // Generate Preview URL
  //         const siteUrl = window.location.origin;
  //         const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
  //         const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //         const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

  //         console.log("Generated Preview URL:", previewUrl);

  //         // Update metadata with preview URL
  //         const status = "Auto Approved";
  //         await listItem.update({
  //           Status: status,

  //         });

  //         console.log("File metadata updated successfully.");

  //         // Create an entry in the FileMaster list
  //         const newRequestNo = await getUniqueRequestNo();
  //         await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //           FileName: String(uploadResult.data.Name),
  //           FileSize: String(uploadResult.data.Length),
  //           FileVersion: String(uploadResult.data.MajorVersion),
  //           CurrentFolderPath: String(currentfolderpath.folderpath),
  //           FileUID: String(uploadResult.data.UniqueId),
  //           CurrentUser: String(currentUserEmailRef.current),
  //           SiteID: String(currentfolderpath.siteID),
  //           Status: status,
  //           FilePreviewURL: String(previewUrl),
  //           DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //           SiteName: String(currentfolderpath.Entity),
  //           MyRequest: true,
  //           Processname: 'New File Request',
  //           RequestNo: newRequestNo
  //         });

  //         console.log("File added to FileMaster successfully.");
  //       } catch (error) {
  //         console.error(`Error uploading file ${file.name}:`, error);
  //       }
  //     }

  //     console.log("Bulk upload process completed.");

  //     // Cleanup media after upload
  //     Deletemedia();

  //     // Reload the page once all files are processed
  //     setTimeout(() => {
  //       location.reload();
  //       onReturnToMain();
  //     }, 3000);

  //   } catch (error) {
  //     console.error("Error during bulk submission:", error);
  //   }
  // };

  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");

  //   try {
  //     console.log("SiteID:", currentfolderpath.siteID);

  //     const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //     if (!testidsub) throw new Error("Subsite not found.");

  //     const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //     console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

  //     let files = await documentLibraryInWhichWeUploadTheFile.files();

  //     // Upload all files in parallel
  //     const uploadPromises = uploadedFiles.map(async (file) => {
  //       try {
  //         const response = await fetch(file.url);
  //         if (!response.ok) throw new Error(`Failed to fetch file: ${file.name}`);
  //         const blob = await response.blob();

  //         const originalFileName = file.name;
  //         const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //         const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //         console.log("originalFileName", originalFileName);
  //         console.log("fileExtension", fileExtension);
  //         console.log("baseFileName", baseFileName);

  //         let uniqueFileName = originalFileName;
  //         let counter = 1;

  //         while (files.some(f => f.Name === uniqueFileName)) {
  //           uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //           counter++;
  //         }
  //         console.log(`Unique file name generated: ${uniqueFileName}`);

  //         const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, blob);
  //         console.log("File uploaded successfully", uploadResult.data.Name);

  //         // Fetch the latest files list again to ensure correct metadata
  //         files = await documentLibraryInWhichWeUploadTheFile.files();

  //         // Add delay to ensure file availability before retrieving metadata
  //         await new Promise(resolve => setTimeout(resolve, 1000));

  //         const listItem = await uploadResult.file.getItem();
  //         console.log("ListItems ", listItem);

  //         if (!listItem) throw new Error("List item not found for the uploaded file.");

  //         // Break role inheritance to ensure correct permissions
  //         await listItem.breakRoleInheritance(false);
  //         console.log("Permissions set correctly");

  //         // Construct preview URL
  //         const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //         const siteUrl = window.location.origin;
  //         const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
  //         const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
  //         console.log("Generated Preview URL:", previewUrl);

  //         const status = "Auto Approved"; // Set status to Auto Approved for bulk upload

  //         // Update metadata
  //         const payload: any = { Status: status };
  //         await listItem.update(payload);
  //         console.log("File metadata updated successfully with:", payload);

  //         const newRequestNo = await getUniqueRequestNo();

  //         const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //           FileName: String(uploadResult.data.Name),
  //           FileSize: String(uploadResult.data.Length),
  //           FileVersion: String(uploadResult.data.MajorVersion),
  //           CurrentFolderPath: String(currentfolderpath.folderpath),
  //           FileUID: String(uploadResult.data.UniqueId),
  //           CurrentUser: String(currentUserEmailRef.current),
  //           SiteID: String(currentfolderpath.siteID),
  //           Status: status,
  //           FilePreviewURL: String(previewUrl),
  //           DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //           SiteName: String(currentfolderpath.Entity),
  //           MyRequest: true,
  //           Processname: 'New File Request',
  //           RequestNo: newRequestNo
  //         });
  //         console.log(newItem, "New item added to FileMaster");

  //         return newItem;
  //       } catch (error) {
  //         console.error(`Error uploading file ${file.name}:`, error);
  //         return null; // Continue with other files even if one fails
  //       }
  //     });

  //     // Wait for all uploads to complete
  //     const results = await Promise.all(uploadPromises);

  //     // Check if any uploads failed
  //     const failedUploads = results.filter(result => result === null).length;
  //     if (failedUploads > 0) {
  //       console.warn(`${failedUploads} files failed to upload.`);
  //     } else {
  //       console.log("All files uploaded successfully.");
  //     }

  //     // Reload the page only once after all files are processed
  //     if (failedUploads === 0) {
  //       Deletemedia();
  //       setTimeout(() => {
  //         location.reload();
  //         onReturnToMain();
  //       }, 3000);
  //     }
  //   } catch (error) {
  //     console.error("Error during bulk submission:", error);
  //   }
  // };
  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");

  //   try {
  //     console.log("SiteID:", currentfolderpath.siteID);

  //     const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //     if (!testidsub) throw new Error("Subsite not found.");

  //     const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //     console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

  //     const files = await documentLibraryInWhichWeUploadTheFile.files();

  //     // Upload all files in parallel
  //     const uploadPromises = uploadedFiles.map(async (file) => {
  //       try {
  //         const response = await fetch(file.url);
  //         if (!response.ok) throw new Error(`Failed to fetch file: ${file.name}`);
  //         const blob = await response.blob();

  //         const originalFileName = file.name;
  //         const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //         const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //         console.log("originalFileName", originalFileName);
  //         console.log("fileExtension", fileExtension);
  //         console.log("baseFileName", baseFileName);

  //         let uniqueFileName = originalFileName;
  //         let counter = 1;

  //         while (files.some(f => f.Name === uniqueFileName)) {
  //           uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //           counter++;
  //         }
  //         console.log(`Unique file name generated: ${uniqueFileName}`);

  //         const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, blob);
  //         console.log("File uploaded successfully", uploadResult.data.Name);

  //         const listItem = await uploadResult.file.getItem();
  //         console.log("ListItems ", listItem);

  //         const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //         const siteUrl = window.location.origin;
  //         const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
  //         console.log(encodedFilePath, "encodedFilePath");

  //         const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
  //         console.log("Generated Preview URL:", previewUrl);

  //         if (!listItem) throw new Error("List item not found for the uploaded file.");

  //         const status = "Auto Approved"; // Set status to Auto Approved for bulk upload

  //         const payload: any = {
  //           Status: status
  //         };

  //         await listItem.update(payload);
  //         console.log("File metadata updated successfully with:", payload);

  //         const newRequestNo = await getUniqueRequestNo();

  //         const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //           FileName: String(uploadResult.data.Name),
  //           FileSize: String(uploadResult.data.Length),
  //           FileVersion: String(uploadResult.data.MajorVersion),
  //           CurrentFolderPath: String(currentfolderpath.folderpath),
  //           FileUID: String(uploadResult.data.UniqueId),
  //           CurrentUser: String(currentUserEmailRef.current),
  //           SiteID: String(currentfolderpath.siteID),
  //           Status: status,
  //           FilePreviewURL: String(previewUrl),
  //           DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //           SiteName: String(currentfolderpath.Entity),
  //           MyRequest: true,
  //           Processname: 'New File Request',
  //           RequestNo: newRequestNo
  //         });
  //         console.log(newItem, "New item added to FileMaster");

  //         return newItem;
  //       } catch (error) {
  //         console.error(`Error uploading file ${file.name}:`, error);
  //         return null; // Continue with other files even if one fails
  //       }
  //     });

  //     // Wait for all uploads to complete
  //     const results = await Promise.all(uploadPromises);

  //     // Check if all uploads were successful
  //     const failedUploads = results.filter(result => result === null).length;
  //     if (failedUploads > 0) {
  //       console.warn(`${failedUploads} files failed to upload.`);
  //     } else {
  //       console.log("All files uploaded successfully.");
  //     }

  //     // Reload the page only once after all files are processed
  //     if(failedUploads ){
  //       Deletemedia()
  //       setTimeout(() => {
  //         location.reload()
  //         onReturnToMain();
  //     }, 3000);
  //      }
  //   } catch (error) {
  //     console.error("Error during bulk submission:", error);
  //   }
  // };
  // const handleSubmitBulk = async (event: any) => {
  //   event.preventDefault();
  //   console.log("Bulk upload button clicked");

  //   // const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  //   // if (!formSelector.checkValidity()) {
  //   //     formSelector.reportValidity(); // Show validation errors
  //   //     return;
  //   // }

  //   // Assuming uploadedFiles is an array of objects with name and url properties
  //   // const uploadedFiles = [
  //   //     { name: "file1.pdf", url: "https://example.com/file1.pdf" },
  //   //     { name: "file2.docx", url: "https://example.com/file2.docx" },
  //   //     // Add more files as needed
  //   // ];

  //   try {
  //       console.log("SiteID:", currentfolderpath.siteID);

  //       const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
  //       if (!testidsub) throw new Error("Subsite not found.");

  //       const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
  //       console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

  //       const files = await documentLibraryInWhichWeUploadTheFile.files();

  //       for (const file of uploadedFiles) {
  //           const response = await fetch(file.url);
  //           const blob = await response.blob();

  //           const originalFileName = file.name;
  //           const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //           const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //           console.log("originalFileName", originalFileName);
  //           console.log("fileExtension", fileExtension);
  //           console.log("baseFileName", baseFileName);
  //           let uniqueFileName = originalFileName;
  //           let counter = 1;

  //           while (files.some(f => f.Name === uniqueFileName)) {
  //               uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //               counter++;
  //           }
  //           console.log(`Unique file name generated: ${uniqueFileName}`);

  //           const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, blob);
  //           console.log("File uploaded successfully", uploadResult.data.Name);

  //           const listItem = await uploadResult.file.getItem();
  //           console.log("ListItems ", listItem);

  //           const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
  //           const siteUrl = window.location.origin;
  //           const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
  //           console.log(encodedFilePath, "encodedFilePath");

  //           const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
  //           console.log("Generated Preview URL:", previewUrl);

  //           if (!listItem) throw new Error("List item not found for the uploaded file.");

  //           const status = "Auto Approved"; // Set status to Auto Approved for bulk upload

  //           const payload: any = {
  //               Status: status
  //           };

  //           await listItem.update(payload);
  //           console.log("File metadata updated successfully with:", payload);

  //           const newRequestNo = await getUniqueRequestNo();

  //           const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
  //               FileName: String(uploadResult.data.Name),
  //               FileSize: String(uploadResult.data.Length),
  //               FileVersion: String(uploadResult.data.MajorVersion),
  //               CurrentFolderPath: String(currentfolderpath.folderpath),
  //               FileUID: String(uploadResult.data.UniqueId),
  //               CurrentUser: String(currentUserEmailRef.current),
  //               SiteID: String(currentfolderpath.siteID),
  //               Status: status,
  //               FilePreviewURL: String(previewUrl),
  //               DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
  //               SiteName: String(currentfolderpath.Entity),
  //               MyRequest: true,
  //               Processname: 'New File Request',
  //               RequestNo: newRequestNo
  //           });
  //           console.log(newItem, "New item added to FileMaster");

  //           if (newItem) {
  //               setTimeout(() => {
  //                   location.reload();
  //                   onReturnToMain();
  //               }, 3000);
  //           }
  //       }
  //   } catch (error) {
  //       console.error("Error during bulk submission:", error);
  //   }
  // };


  const Deletemedia = () => {

    Swal.fire({
      title: "Success",
      text: "File uploaded successfully",
      icon: "success"
    });


    setTimeout(() => {
      Swal.close();
      onReturnToMain();
    }, 3000);

  }



  const handleToggle = () => {
    // alert(`isChecked before toggle: ${isChecked}`);/
    console.log("isChecked before toggle:", isChecked);
    setIsChecked((prev) => {
      const newCheckedState = !prev; // Get the updated state
      console.log(`New isChecked state: ${newCheckedState}`);
      // Get all elements with the class "input-container"
      const inputContainers = document.getElementsByClassName("input-container");

      // Hide or show input fields based on the new state
      for (let i = 0; i < inputContainers.length; i++) {
        (inputContainers[i] as HTMLElement).style.display = newCheckedState ? "none" : "block";
      }
      const getsubmitbutton = document.getElementById("submitBtn") as HTMLButtonElement;
      // getsubmitbutton.style.display = newCheckedState ? "none" : "block";
      const fileInput = document.getElementById('fileInput') as HTMLInputElement;
      const selectedFile = fileInput?.files?.[0];
      getsubmitbutton.style.display = newCheckedState ? "none" : !selectedFile ? 'none' : 'block';


      return newCheckedState; // Update the state
    });
  };
  const handleRemove = (fileIndex: any) => {
    event.preventDefault();
    const updatedFiles = uploadedFiles.filter((_, index) => index !== fileIndex);
    setUploadedFiles(updatedFiles); // Update state
    console.log(uploadedFiles, "uploadedFiles in remove")
  };

  useEffect(() => {
    const getsubmitbuttonbulk = document.getElementById("submitBtn2") as HTMLButtonElement;
    if (getsubmitbuttonbulk) {
      console.log("isChecked-isChecked", isChecked)
      getsubmitbuttonbulk.style.display = isChecked ? uploadedFiles.length > 0 ? 'block' : 'none' : 'none';
    }
  }, [isChecked]);

  return (
    <>
      <button className='BackButton me-3 mb-3'
        onClick={() => { location.reload(); onReturnToMain() }}
      > Back
      </button>
      <div className="container mt-3 UploadFileCont">
        <div className='main-containeruploadfile'>
          <div className='column column1 p-3'>
            <form id='formSelector'>
              <h1>Upload file</h1>
              {/* <label className="switch">
                              <input type="checkbox"/>
                              <span className="slider round"></span>
                            </label> */}
              {showBulkUpload === true && (
                <p>File uploaded to this folder will require approval. After submission, a request will be sent for review. The file will only be visible and accessible once it has been approved by the designated approvers.
                </p>
              )}
              <div>
                {showBulkUpload === false && ( // Show only if IsApproval is false
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }} className="mt-3">
                    <p className="mb-0 text-dark">Bulk upload:</p>
                    <div style={{ display: 'flex', gap: '5px', alignItems: 'center' }}>
                      <label className="switch">
                        <input type="checkbox" checked={isChecked} onChange={handleToggle} />
                        <span className="slider round"></span>
                      </label>
                      <p className="mb-0 text-dark fw-bold">{isChecked ? "ON" : "OFF"}</p>
                    </div>

                  </div>


                )}

              </div>
              <div>
                {isChecked && (
                  <div className="input-container mt-3">
                    {/* <label htmlFor="Uplaod bulk">Bulk upload</label> */}
                    <label htmlFor="bulkfile" style={{ fontWeight: "bold" }}>
                      Upload File <span style={{ color: "red" }}>*</span>
                    </label>
                    <input type="file" name="bulkfile" id="bulkfile" multiple onChange={(e) => handlebulkFileChange(e)} />
                    <ul className="newbulnup">
                      {uploadedFiles.map((file, index) => (
                        <li key={index}>
                          <div>     {index + 1}.</div>
                          <div className="font-14" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', padding: '0px 5px', fontWeight: '500' }}>  <a style={{ color: '#f37421' }} href="#" onClick={() => handlePreview(file.url)}>
                            {file.name}
                          </a>
                          </div>
                          <div>   <a href="" onClick={() => handleRemove(index)} >
                            <img src={require("../assets/del.png")} className="fas fa-trash" alt="delete" />
                          </a>
                          </div>
                        </li>
                      ))}
                    </ul>
                    <div style={{ display: 'flex', justifyContent: 'right' }}>
                      <button style={{ display: 'none', width: '130px' }} id="submitBtn2" type="submit" onClick={handleSubmitBulk}>Bulk Submit</button>
                    </div>
                  </div>
                )}
              </div>
            </form>
          </div>
          <div className='column column2 p-3'>
            <h1>File Preview</h1>
            {isUploading && (
              <>
                <div id="spinner" style={{ display: "block", textAlign: "center" }}>
                  <div>
                    <img
                      src={require("../../../CustomAsset/argloader.gif")}
                      className="alignrightl"
                      alt="Loading..."
                    />
                  </div>
                  <span>Please wait, we are preparing your files for upload... </span>{" "}
                </div>
              </>
            )}
            {isFinalUploading && (
              <>
                <div id="spinner" style={{ display: "block", textAlign: "center" }}>
                  <div>
                    <img
                      src={require("../../../CustomAsset/argloader.gif")}
                      className="alignrightl"
                      alt="Loading..."
                    />
                  </div>
                  <span>Uploading items... This may take a moment. </span>{" "}
                </div>
              </>
            )}
            <div id="spinner" style={{ display: "none", textAlign: "center" }}>
              <div>
                <img
                  src={require("../../../CustomAsset/argloader.gif")}
                  className="alignrightl"
                  alt="Loading..."
                />
              </div>
              <span>Loading </span>{" "}
            </div>
            <iframe id="filePreview" width="100%" height="400"></iframe>
          </div>
        </div>
      </div>
    </>
  );
}
export default UploadFile;
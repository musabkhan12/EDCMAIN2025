import React, { useEffect } from 'react';

let mypreviewurl:any
let mypreviewurl2:any
const PreviewFile = ({ fileUrl }: { fileUrl: string }) => {
  useEffect(() => { 
  mypreviewurl2 =   localStorage.getItem('previewurladvancesearch')

  },[])

  const currentUrl = window.location.href;

  // Parse the full URL
  const url = new URL(currentUrl);
  
  // Get the value of the "Previewfile" parameter
  const previewFileEncoded = url.searchParams.get("Previewfile");
  mypreviewurl = previewFileEncoded

  // Decode it back to readable form
  // const previewFileUrl = decodeURIComponent(previewFileEncoded);
  
  console.log("Preview File URL:", previewFileEncoded);
  
    return (
      <div>
  <iframe src={mypreviewurl2} style={{ width: '100%', height: '600px', border: 'none' }} />
      {/* <iframe src={'https://edcadae.sharepoint.com/sites/EDeDMS/HR%20Department%20Testing/Policy/Forms/AllItems.aspx?id=%2Fsites%2FEDeDMS%2FHR%20Department%20Testing%2FPolicy%2FEDC%20Status%2Exlsx&parent=%2Fsites%2FEDeDMS%2FHR%20Department%20Testing%2FPolicy'} style={{ width: '100%', height: '600px', border: 'none' }} />
      <iframe src={'https://edcadae.sharepoint.com/sites/EDeDMS/HR%20Department%20Testing/Policy/Forms/AllItems.aspx?id=/sites/EDeDMS/HR%2520Department%2520Testing/Policy/EDC%2520Status.xlsx&parent=/sites/EDeDMS/HR%2520Department%2520Testing/Policy'} style={{ width: '100%', height: '600px', border: 'none' }} /> */}
      </div>
    
    );
  };
  

export default PreviewFile;
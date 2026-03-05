// import React, { useEffect } from 'react';

// let mypreviewurl:any
// let mypreviewurl2:any
// const PreviewFile = ({ fileUrl }: { fileUrl: string }) => {
//   useEffect(() => { 
//   mypreviewurl2 =   localStorage.getItem('previewurladvancesearch')

//   },[])

//   const currentUrl = window.location.href;

//   // Parse the full URL
//   const url = new URL(currentUrl);
  
//   // Get the value of the "Previewfile" parameter
//   const previewFileEncoded = url.searchParams.get("Previewfile");
//   mypreviewurl = previewFileEncoded

//   // Decode it back to readable form
//   // const previewFileUrl = decodeURIComponent(previewFileEncoded);
  
//   console.log("Preview File URL:", previewFileEncoded);
  
//     return (
//       <div>
//   <iframe src={mypreviewurl2} style={{ width: '100%', height: '600px', border: 'none' }} />
//       {/* <iframe src={'https://edcadae.sharepoint.com/sites/ED/HR%20Department%20Testing/Policy/Forms/AllItems.aspx?id=%2Fsites%2FEDeDMS%2FHR%20Department%20Testing%2FPolicy%2FEDC%20Status%2Exlsx&parent=%2Fsites%2FEDeDMS%2FHR%20Department%20Testing%2FPolicy'} style={{ width: '100%', height: '600px', border: 'none' }} />
//       <iframe src={'https://edcadae.sharepoint.com/sites/ED/HR%20Department%20Testing/Policy/Forms/AllItems.aspx?id=/sites/ED/HR%2520Department%2520Testing/Policy/EDC%2520Status.xlsx&parent=/sites/ED/HR%2520Department%2520Testing/Policy'} style={{ width: '100%', height: '600px', border: 'none' }} /> */}
//       </div>
    
//     );
//   };
  

// export default PreviewFile;

import React, { useEffect, useRef, useState } from 'react';
import { Spinner } from 'react-bootstrap';

const PreviewFile = () => {
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const [fileUrl, setFileUrl] = useState<string | null>(null);
  const [showLoader, setShowLoader] = useState(true);
  const [showIframe, setShowIframe] = useState(false);

  useEffect(() => {
    // Get URL from localStorage and/or URL params
    const urlFromStorage = localStorage.getItem('previewurladvancesearch');
    const currentUrl = new URL(window.location.href);
    const previewFileEncoded = currentUrl.searchParams.get("Previewfile");
    
    setFileUrl(urlFromStorage || previewFileEncoded || null);
  }, []);

  useEffect(() => {
    if (!fileUrl || !iframeRef.current) return;

    const hideElementsAndShowIframe = () => {
      try {
        const iframeDoc = iframeRef.current?.contentDocument || 
                         iframeRef.current?.contentWindow?.document;
        if (!iframeDoc) return;

        // Hide unwanted elements
        const commandBar = iframeDoc.getElementById("OneUpCommandBar");
        if (commandBar) {
          commandBar.style.display = "none";
        }

        const excelToolbar = iframeDoc.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar");
        if (excelToolbar) {
          excelToolbar.style.display = "none";
        }

        const userProfile = iframeDoc.getElementById("presenceCommand");
        if (userProfile) {
          userProfile.style.display = 'none';
        }

        // Show the iframe
        setShowIframe(true);
        setShowLoader(false);

      } catch (error) {
        console.error("Error accessing iframe content:", error);
        // If error occurs, still show the iframe
        setShowIframe(true);
        setShowLoader(false);
      }
    };

    // Set a small delay to ensure iframe has started loading
    const timer = setTimeout(() => {
      iframeRef.current?.addEventListener('load', hideElementsAndShowIframe);
    }, 100);

    return () => {
      clearTimeout(timer);
      iframeRef.current?.removeEventListener('load', hideElementsAndShowIframe);
    };
  }, [fileUrl]);

  if (!fileUrl) {
    return <div className="text-center p-5">No file selected for preview</div>;
  }

  return (
    <div style={{ position: 'relative', width: '100%', height: '600px' }}>
      {showLoader && (
        <div className="d-flex justify-content-center align-items-center" 
             style={{
               position: 'absolute',
               top: 0,
               left: 0,
               right: 0,
               bottom: 0,
               backgroundColor: 'rgba(255,255,255,0.8)',
               zIndex: 10
             }}>
          <Spinner animation="border" role="status">
            <span className="visually-hidden">Loading...</span>
          </Spinner>
          <span className="ms-2">Preparing file preview...</span>
        </div>
      )}
     {/* thsi was my previous working iframe just issue was pdf was not previewing */}
      {/* <iframe
        ref={iframeRef}
        src={fileUrl}
        style={{ 
          width: '100%', 
          height: '100%', 
          border: 'none',
          display: showIframe ? 'block' : 'none'
        }}
        title="File Preview"
        // sandbox="allow-same-origin allow-scripts allow-popups allow-forms"
      /> */}
      <iframe
  ref={iframeRef}
  src={fileUrl}
  style={{ 
    width: '100%', 
    height: '100%', 
    border: 'none',
    display: showIframe ? 'block' : 'none'
  }}
  title="File Preview"
  onLoad={() => {
    setShowIframe(true);
    setShowLoader(false);
  }}
/>
    </div>
  );
};

export default PreviewFile;
import Swal from 'sweetalert2';
export const getAllDocumentCode = async (_sp, dept) => {
  let arr = [];
  let sts = "Approved";

  // await _sp.web.lists.getByTitle("ChangeRequestList").items.filter(`Status eq '${sts}' and Department/ADDepartmentName eq '${dept}'`)
  //   .select("*,Department/ID,Department/Department,Location/ID,Location/Location,Custodian/ID,Custodian/Custodian,DocumentType/ID,DocumentType/DocumentType,AmendmentType/ID,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,ChangeRequestType/ID,Author/ID,Author/Title,TemplateType/TemplateTypeName,TemplateTypeId")
  //   .expand("TemplateType,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author,Department")
  //   .orderBy("ID", false)()
  await _sp.web.lists.getByTitle("ChangeRequestList").items.filter(`Status eq '${sts}'`)
    .select("*,Department/ID,Department/Department,Location/ID,Location/Location,Custodian/ID,Custodian/Custodian,DocumentType/ID,DocumentType/DocumentType,AmendmentType/ID,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,ChangeRequestType/ID,Author/ID,Author/Title,TemplateType/TemplateTypeName,TemplateTypeId,PreparedBy/ID,PreparedBy/Title")
    .expand("TemplateType,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author,Department,PreparedBy")
    .orderBy("ID", false)()
    .then(async (res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode with the greatest ID
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.DocumentCode] || acc[item.DocumentCode].ID < item.ID) {
          acc[item.DocumentCode] = item;
        }
        return acc;
      }, {});

      // Filter out documents with DocumentCancellationStatus as 'Yes'
      Object.keys(latestDocuments).forEach((key) => {
        if (latestDocuments[key].DocumentCancellationStatus === 'Yes') {
          delete latestDocuments[key];
        }
      });

      const documentIds = Object.values(latestDocuments).map((doc) => doc.ID);
      const chunkSize = 50; // Adjust chunk size as needed
      let digitalSignedDocs = [];

      for (let i = 0; i < documentIds.length; i += chunkSize) {
        const chunk = documentIds.slice(i, i + chunkSize).join(" or ListItemID/ID eq ");
        const chunkResults = await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items
          .filter(`ListItemID/ID eq ${chunk}`)
          .select("ID, ListItemID/ID")
          .expand("ListItemID")
          .getAll();
        digitalSignedDocs = digitalSignedDocs.concat(chunkResults);
      }

      const digitalSignedDocsMap = digitalSignedDocs.reduce((acc, item) => {
        acc[item.ListItemID.ID] = item;
        return acc;
      }, {});

      arr = Object.values(latestDocuments).map((doc) => {
        const digitalSignedDoc = digitalSignedDocsMap[doc.ID];
        const DigiSignId = digitalSignedDoc ? [digitalSignedDoc.ID] : [];
        return {
          ...doc,
          AttachmentId: digitalSignedDoc ? [] : doc.AttachmentId,
          AttachmentDigitalSignatureId: digitalSignedDoc ? DigiSignId : []
        };
      });
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
// const latestDocuments = res.reduce((acc, item) => {
//   if (!acc[item.DocumentCode]) {
//     const hasCancellationStatusYes = res.some(doc => doc.DocumentCode === item.DocumentCode && doc.DocumentCancellationStatus === 'Yes');
//     if (!hasCancellationStatusYes) {
//       acc[item.DocumentCode] = item;
//     }
//   }
//   // acc[item.DocumentCode] = item;
//   return acc;
// }, {});

// const filterCondition = documentIds.map(id => `ListItemID/ID eq ${id}`).join(" or ");
// const digitalSignedDocs = await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items
//   .filter(filterCondition)
//   .select("ID, ListItemID/ID")
//   .expand("ListItemID")
//   .top(4999)
//   .getAll();
export const getAllRequestType = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("RequestTypeMaster").items
    .select("*,Author/ID,Author/Title")
    .expand("Author")
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.RequestType]) {
          acc[item.RequestType] = item;
        }
        return acc;
      }, {});

      arr = Object.values(latestDocuments);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
export const getAllAmendmentType = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("AmendmentTypeMaster").items
    .select("*,Author/ID,Author/Title")
    .expand("Author")
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.AmendmentType]) {
          acc[item.AmendmentType] = item;
        }
        return acc;
      }, {});

      arr = Object.values(latestDocuments);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
export const getAllClassificationMaster = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ClassificationMaster").items
    .select("*,Author/ID,Author/Title")
    .expand("Author")
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.Classification]) {
          acc[item.Classification] = item;
        }
        return acc;
      }, {});

      arr = Object.values(latestDocuments);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};

export const addItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestDocumentCancellationList').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};
export const addItemChangeRequestList = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestList').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};
export const addItem2 = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestReasonDocumentCancellationList').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};
export const updateItemAuditreport = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditReportList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const updateItem = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestDocumentCancellationList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const updateItemMainList = async (itemData, _sp, id, listName) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle(listName).items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const addItemChangeRequestReasonlist = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestReasonList').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};

export const updateItemchangeRequestReasonlist = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestReasonList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const updateItemChangeRequestList = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const updateItem2 = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestReasonDocumentCancellationList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const updateItemChangeRequestReasonList = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ChangeRequestReasonList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const updateApprovalItem = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ProcessApprovalList').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};

export const getItemByID = async (_sp, id) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ChangeRequestDocumentCancellationList").items.getById(id)
    .select("*,TemplateTypeId,Department/ID,Department/Department,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title,Location/Location,Custodian/Custodian,DocumentType/DocumentType,AmendmentType/AmendmentType,Classification/Classification,TemplateType/TemplateTypeName,PreparedBy/ID,PreparedBy/Title,DocumentCancellationBy/Id,DocumentCancellationBy/Title,DocumentCancellationBy/EMail").expand("DocumentCancellationBy,PreparedBy,TemplateType,Department,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author")()
    .then((res) => {
      console.log(res, ' let arrs=[]');

      arr.push(res)
      // arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}
export const getItemByIDCR = async (_sp, id) => {

  let arr = []
  let arrs = []
  let bannerimg = []

  await _sp.web.lists.getByTitle("ChangeRequestList").items.getById(id)
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Location/Location,Custodian/Custodian,DocumentType/DocumentType,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,ChangeRequestType/ID,Author/ID,Author/Title").expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author")()
    .then((res) => {
      console.log(res, ' let arrs=[]');

      arr.push(res)
      // arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}
export const getItemByID2 = async (sp, ChangeRequestID) => {
  debugger
  let arr = []
  let sampleDataArray = []
  arr = await sp.web.lists.getByTitle("ChangeRequestReasonDocumentCancellationList").items.select("*,ChangeRequestDCID/ID").expand("ChangeRequestDCID").filter(`ChangeRequestDCID/ID eq ${ChangeRequestID}`).getAll();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}
export const getItemByIDChangeRequest = async (sp, ChangeRequestID) => {
  debugger
  let arr = []
  let sampleDataArray = []
  arr = await sp.web.lists.getByTitle("ChangeRequestReasonList").items.select("*,ChangeRequestID/ID").expand("ChangeRequestID").filter(`ChangeRequestID/ID eq ${ChangeRequestID}`).getAll();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}
export const GetQueryString = (string) =>
  new URLSearchParams(window.location.search).get(string);

export const getApprovalByID = async (_sp, id, processName) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName) {
        arr = res;
      }
      // .filter(`AssignedTo/Id eq ${currentUser.Id} and ProcessName eq ${processName}`)

      //  arr.push(res)

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}
export const getApprovalByID2 = async (_sp, id, processName) => {

  let arr;
  let arrs = []
  let bannerimg = []
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 && res.CurrentUserRole !== "OES") {
        arr = false;
      }
      else {
        arr = true;
      }
      // .filter(`AssignedTo/Id eq ${currentUser.Id} and ProcessName eq ${processName}`)

      //  arr.push(res)

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}

export const getDraftApprovalByID = async (_sp, id, processName) => {

  let arr = [];
  let val = "Yes"
  let Sts = "Save as draft";
  let sts = "Pending"
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`ListItemId eq '${id}' and AssignedTo/ID  eq '${currentUser.Id}' and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).top(1)()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      // if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 && res.CurrentUserRole !=="OES" ){
      //   arr = false;
      // }
      // else{
      //   arr = true;
      // }
      // .filter(`AssignedTo/Id eq ${currentUser.Id} and ProcessName eq ${processName}`)

      //  arr.push(res)

      arr = res;

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}

export const getDataRoles = async (_sp) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ApproverRoleMaster").items
    .select("*")()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr.push(res)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}

export const addApprovalItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ProcessApprovalList').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};

export const addAllProcessItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('AllProcessApprovalLevelList').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};

export const UpdateAllProcessItem = async (itemData, _sp, id) => {

  let resultArr = []
  try {

    const newItem = await _sp.web.lists.getByTitle('AllProcessApprovalLevelList').items.getById(id).update(itemData);
    // console.log('Item  successfully:', newItem);
    resultArr = newItem

  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};

export const getAllProcessData = async (_sp, MainId, processName, docCode) => {

  let arr;

  // const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("AllProcessApprovalLevelList").items
    .select("*,Author/ID,Author/Title,Approvers/Id,Approvers/Title,ApproverRole/Id").expand("Author,Approvers,ApproverRole").filter(`MainListID eq '${MainId}' and ProcessName eq '${processName}' and RequestId eq '${docCode}'`)()
    .then((res) => {
      //   res.map((item) => ({

      // }));
      arr = res;

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}


export const getRequesterID = async (_sp) => {

  var reqId;
  await _sp.web.lists.getByTitle("RequesterRoleMaster").items
    .select("*").filter("Role eq 'Initiator' and IsActive eq 'Yes'")()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId = res[0].Id
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}

export const getRequestTypeID = async (_sp) => {

  var reqId;
  await _sp.web.lists.getByTitle("RequestTypeMaster").items
    .select("*").filter("FormName eq 'Document Cancellation' and IsActive eq 'Yes'")()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId = res[0].Id
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}

export const getFormNameID = async (_sp, formname) => {

  var reqId;
  await _sp.web.lists.getByTitle("FormNameMaster").items
    .select("*").filter("FormName eq '" + formname + "'").top(1)()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId = res[0].Id
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}
export const getListNameID = async (_sp, formname) => {

  var reqId;
  await _sp.web.lists.getByTitle("ListNameMaster").items
    .select("*").filter("ListName eq '" + formname + "'").top(1)()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId = res[0].Id
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}

export const getDocumentLinkByID = async (_sp, itemId) => {

  var reqId;
  await _sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(itemId)
    .select("*,FileRef, FileLeafRef")()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId = res
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}
export const getDocumentLinkByIDSigned = async (_sp, itemId) => {

  var reqId;
  await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items.getById(itemId)
    .select("*,FileRef, FileLeafRef")()
    .then((res) => {
      console.log(res, ' let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId = res
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}

// export const getGeneratedTemplateDoc = async (_sp, itemId) => {
//   let results = [];
//   // for (let itemId of AttachmentIds) {
//     await _sp.web.lists.getByTitle("DocumentCancellationGeneratedTemplateDoc").items
//       .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
//       .then((res) => {
//         console.log(res, ' let arrs=[]');
//         results = res;
//       })
//       .catch((error) => {
//         console.log("Error fetching data: ", error);
//       });
//   // }
//   console.log(results, 'results');
//   return results;
// }

export const getGeneratedTemplateDoc = async (_sp, itemId, DocCode) => {
  let results = [];
  try {
    const res = await _sp.web.lists.getByTitle("DocumentCancellationDigitalSignedDocs").items
      .select("*,FileRef, FileLeafRef")
      .filter(`ListItemID/ID eq ${itemId}`)
      .orderBy("ID", false)
      .top(1)();

    if (res && res.length > 0) {
      results = res;
    } else {
      const res2 = await _sp.web.lists.getByTitle("DocumentCancellationGeneratedTemplateDoc").items
        .select("*,FileRef, FileLeafRef")
        .filter(`ListItemID/ID eq ${itemId}`)
        .orderBy("ID", false)
        .top(1)();

      results = res2 && res2.length > 0 ? res2 : [];

    }
  } catch (error) {
    console.log("Error fetching data: ", error);
  }
  console.log(results, 'results');
  return results;
};

export const getGeneratedTemplateDoc2 = async (_sp, itemId, ChangeReqID) => {
  let results = null;
  try {
    const res = await _sp.web.lists.getByTitle("DocumentCancellationAttachDigitalSignedDocs").items
      .select("*,FileRef, FileLeafRef")
      .filter(`ListItemID/ID eq ${itemId}`)
      .orderBy("ID", false)
      .top(1)();

    if (res && res.length > 0) {
      results = res[0];
      console.log(results, res, 'results res1');
    } else {
      const res2 = await _sp.web.lists.getByTitle("DocumentCancellationDocs").items
        .select("*,FileRef, FileLeafRef")
        .filter(`ListItemID/ID eq ${itemId}`)
        .orderBy("ID", false)
        .top(1)();


      if (res2 && res2.length > 0) {
        results = res2 && res2.length > 0 ? res2[0] : null;
        console.log(results, res2, 'results res2');
      }
      // ///////
      else if (res2.length == 0) {
        try {
          const res3 = await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items
            .select("*,FileRef, FileLeafRef")
            .filter(`ListItemID/ID eq ${ChangeReqID}`)
            .orderBy("ID", false)
            .top(1)();

          if (res3 && res3.length > 0) {
            results = res3[0];
            console.log(results, res3, 'results res3');
          }
          else {



            // const latestApprovedItem = await _sp.web.lists.getByTitle("ChangeRequestList").items
            //   .filter(`Status eq 'Approved' and DocumentCode eq '${DocumentCode}'`)
            //   .orderBy("ID", false)
            //   .top(1)()
            const latestApprovedItem = await _sp.web.lists.getByTitle("ChangeRequestList").items
            .getById(ChangeReqID).select("*")()   
  
              .then(async (res) => {
                if (res) {
                  // return res[0];
                  console.log(results, res, 'results res4');
                  // if (res && res.length > 0) { // when change req item is fetched
                  // results = res[0];
                  var AttachmentID = res?.AttachmentId[0];
                  if (AttachmentID) {
                    console.log(AttachmentID, 'AttachmentID');
                    await _sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(AttachmentID)
                      .select("*,FileRef, FileLeafRef")()
                      .then((resp) => {
                        console.log(resp, ' let arrs=[]');
                        results = resp;

                        console.log(results, 'results res4');

                      })
                      .catch((error) => {
                        console.log("Error fetching data: ", error);
                      });

                  }
                  else {
                    results = null;
                    console.log(results, 'results res4');

                  }


                  // }
                } else {
                  // return null;
                  results = null;
                }
              })
              .catch((error) => {
                console.log("Error fetching latest approved item: ", error);
                // return null;
                results = null;
              });


            // const res4 = await _sp.web.lists.getByTitle("DocumentCancellationDocs").items
            //   .select("*,FileRef, FileLeafRef")
            //   .filter(`ListItemID/ID eq ${itemId}`)
            //   .orderBy("ID", false)
            //   .top(1)();

            // results = res4 && res4.length > 0 ? res4[0] : null;

          }
        } catch (error) {
          console.log("Error fetching data: ", error);
        }

      }
      // ///////
    }
  } catch (error) {
    console.log("Error fetching data: ", error);
  }
  console.log(results, 'results');
  return results;
};


export const updateDigitalsign = async (listname, _sp, id) => {
  let resultArr = []
  try {
    console.log("iddddd", id);
    // const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
    //   .filter(`ListName eq '${listname}' and ListItemID eq ${id}`)
    //   .top(1)
    //   ().then(async (res) => {
    const postPayload2 = {
      DocSignedStatus: "Yes"
    }

    const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items.getById(id).update(postPayload2);
    console.log('Item added successfully:', newItem);



    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};

export const getdigitalsignaturerequestbyID = async (listname, _sp, id) => {
  let arr = []
  try {
    console.log("iddddd", id);
    const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
      .filter(`ListName eq '${listname}' and ListItemID eq ${id} and DocSignedStatus eq 'No'`)
      .top(1)
      ()
      .then((res) => {
        console.log(res, ' let arrs=[]');

        arr = res
        // arr = res;
      })
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    arr = null
  }
  return arr;
};

export const getUserDepartment = async (_sp, dept) => {


  let deptName = "";
  await _sp.web.lists.getByTitle("DepartmentMasterList").items
    .select("*,ToUsers/Title,CCUsers/Title").expand("ToUsers,CCUsers").filter(`Active eq 'Yes' and ADDepartmentName eq '${dept}'`)()
    .then((res) => {

      if (res && res.length > 0) {
        deptName = res[0].Department;
      }

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return deptName;
}


export const getLatestChangeRequestTemplateType = async (_sp, List) => {
  let arr = [];

  // const spCache = spfi(_self._sp).using(Caching({ store: "session" }));
  // const listItems = await sp.web.lists.getByTitle("AuditProgramTypeMaster").items();
  await _sp.web.lists.getByTitle("ChangeRequestList").items.filter(`(TemplateType/TemplateTypeValue eq '${List}') and Status eq 'Approved'`).orderBy("ID", false).top(1)()
    .then((res) => {
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}


export const CheckIfAlreadyactionTaken = async (_sp, id) => {
  try {
    const currentUser = await _sp.web.currentUser();

    const item = await _sp.web.lists
      .getByTitle("ProcessApprovalList")
      .items
      .getById(id)
      .select("Id", "ActionTakenById", "ActionTakenOn", "AssignedTo/Id", "ProcessName")
      .expand("AssignedTo")();

    const isUnprocessed = (!item.ActionTakenById || item.ActionTakenById == null) && (!item.ActionTakenOn || item.ActionTakenOn == null);

    // Optional: further check if it's assigned to current user and matches processName


    if (isUnprocessed) {
      return true;
    }

    return false;
  } catch (error) {
    console.error("Error in CheckIfAlreadyactionTaken:", error);
    return false;
  }
};
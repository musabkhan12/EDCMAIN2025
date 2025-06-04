import Swal from 'sweetalert2';
export const getAllDocumentCode = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title")
    .expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author")
    .filter("Status eq 'Approved'")
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log("eeee", res);
      debugger
      // Filter only latest entry for each unique DocumentCode
      // const latestDocuments = res.reduce((acc, item) => {
      //   if (!acc[item.DocumentCode]) {
      //     acc[item.DocumentCode] = item;
      //   }
      //   return acc;
      // }, {});

      // arr = Object.values(latestDocuments);
      var latestByDocumentCode = {};

      for (var i = 0; i < res.length; i++) {
        var item = res[i];
        var docCode = item.DocumentCode;

        // Store the first (latest) item per unique DocumentCode
        if (!latestByDocumentCode[docCode]) {
          latestByDocumentCode[docCode] = item;
        }
      }

      arr = Object.values(latestByDocumentCode);

      console.log("arrarr fetching data: ", arr);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
export const getchangerequesttemp = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Department/ID,Department/Department,Location/ID,Location/Location,Custodian/ID,Custodian/Custodian,DocumentType/ID,DocumentType/DocumentType,AmendmentType/ID,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,ChangeRequestType/ID,Author/ID,Author/Title,TemplateType/TemplateTypeName,TemplateTypeId")
    .expand("TemplateType,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author,Department")
    .filter(`TemplateType/TemplateTypeName eq 'Change Request' and Status eq 'Approved'`)
    .orderBy("ID", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      // const latestDocuments = res.reduce((acc, item) => {
      //   if (!acc[item.DocumentCode]) {
      //     acc[item.DocumentCode] = item;
      //   }
      //   return acc;
      // }, {});
      let SnoArr = [];
      if (res.length > 0) {
        SnoArr.push({
          DocumentCode: res[0].DocumentCode,
          IssueDate: res[0].IssueDate,
          RevisionDate: res[0].RevisionDate,
          SerialNo: Number(res[0].SerialNumber),
          IssueNo: Number(res[0].IssueNumber),
          RevisionNo: Number(res[0].RevisionNumber)
        })
      }
      console.log("resresr serialnumber", res, SnoArr);
      arr = SnoArr
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
export const getDocumentCodeselected = async (_sp, locId, custoId, doctypeId) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title")
    .expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author")
    .filter(`LocationId eq '${locId}' and CustodianId eq '${custoId}' and DocumentTypeId eq '${doctypeId}' and Status ne 'Save as draft' and Status ne 'Rejected'`)
    .orderBy("SerialNumber", false).top(1)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      // const latestDocuments = res.reduce((acc, item) => {
      //   if (!acc[item.DocumentCode]) {
      //     acc[item.DocumentCode] = item;
      //   }
      //   return acc;
      // }, {});
      let SnoArr = [];
      if (res.length > 0) {
        SnoArr.push({
          SerialNo: Number(res[0].SerialNumber),
          IssueNo: Number(res[0].IssueNumber),
          RevisionNo: Number(res[0].RevisionNumber)
        })
      }
      console.log("resresr serialnumber", res, SnoArr);
      arr = SnoArr
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
export const getDocumentCodeselectedApproved = async (_sp, locId, custoId, doctypeId) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title")
    .expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author")
    .filter(`LocationId eq '${locId}' and CustodianId eq '${custoId}' and DocumentTypeId eq '${doctypeId}' and Status eq 'Approved'`)
    .orderBy("ID", true).top(1)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);
      let SnoArr = [];
      if (res.length > 0) {
        SnoArr.push({
          SerialNo: Number(res[0].SerialNumber),
          IssueNo: Number(res[0].IssueNumber),
          RevisionNo: Number(res[0].RevisionNumber),
          IssueDate:res[0].IssueDate
        })
      }
      console.log("resresr serialnumber", res, SnoArr);
      arr = SnoArr
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
// export const getAllDepartment = async (_sp) => {
 
//   let arr = []
//   let arrs = []
//   let bannerimg = []
//   await _sp.web.lists.getByTitle("DepartmentMasterList").items
//   .select("*").filter("Active eq 'Yes'")()
//     .then((res) => {
//       // console.log(res, ' let arrs=[]');
//       arr = res.map((item) => ({
//           value: item.Id,
//           label: item.Department,
//           Department:item.Department,
       
//     }));
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   // console.log(arr, 'arr');
//   return arr;
// }
export const getAllDepartment = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("DepartmentMasterList").items
    .select("*,Author/ID,Author/Title")
    .expand("Author")
    .orderBy("Modified", false)
    .filter("Active eq 'Yes'")() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.ADDepartmentName]) {
          acc[item.ADDepartmentName] = item;
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
export const getAllTemplateType = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("TemplateTypeMaster").items
    .select("*,Author/ID,Author/Title")
    .expand("Author")
    .orderBy("Modified", false)
    .filter("IsActive eq 'Yes'")() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.TemplateTypeName]) {
          acc[item.TemplateTypeName] = item;
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

export const getAllRequestType = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("RequestTypeMaster").items
    .select("*,Author/ID,Author/Title")
    .expand("Author")
    .filter("FormName eq 'Change Request'")
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log("optrequest", res);

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
export const getGeneratedTemplateDocCR = async (_sp, itemId) => {
  let results = [];
  // for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("ChangeRequestGeneratedTemplateDoc").items
      .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
      .then((res) => {
        console.log(res, ' let arrs=[]');
        results = res;
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
  // }
  console.log(results, 'results');
  return results;
}
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

// export const addItemChangeRequestList = async (itemData, _sp) => {
//   let resultArr = [];
//   try {
//     // Convert date fields to ISO 8601 format
//     const formattedItemData = {
//       ...itemData,
//       RequestDate: itemData.RequestDate ? new Date(itemData.RequestDate).toISOString() : null,
//       DueDate: itemData.DueDate ? new Date(itemData.DueDate).toISOString() : null,
//     };

//     console.log("Formatted Item Data:", formattedItemData);

//     const newItem = await _sp.web.lists.getByTitle("ChangeRequestList").items.add(formattedItemData);

//     console.log("Item added successfully:", newItem);
//     resultArr = newItem;
//   } catch (error) {
//     console.log("Error adding item:", error);
//     resultArr = null;
//     Swal.fire("Cancelled", "", "error");
//   }
//   return resultArr;
// };




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
  debugger
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
    debugger
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
export const getallProcessApprovalitems = async (_sp, id) => {
  let resultArr = [];
  let newItem = [];
  try {
    const newitemnew = await _sp.web.lists.getByTitle('AllProcessApprovalLevelList').items
      .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,Approvers/Id,Approvers/Title")
      .expand("Author,RequesterName,Approvers")
      .filter(`MainListID eq ${id}`).orderBy("Level", false)
      ().then(async (res) => {
        if(res.length > 0){
          newItem = await _sp.web.lists.getByTitle('ProcessApprovalList').items
          .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title")
          .expand("Author,RequesterName,AssignedTo")
          .filter(`ListItemId eq ${id} and Status eq 'Pending' and Level eq ${res[0].Level}`)
          ();
        console.log('all process itemss: of highest level', newItem);
        }
        
      })
    console.log("newitemnewnewitemnew", newitemnew);
    
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
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title").expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author")()
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
    .select("*,Department/ID,Department/Department,Location/ID,Location/Location,Custodian/ID,Custodian/Custodian,DocumentType/ID,DocumentType/DocumentType,AmendmentType/ID,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,RequestType/ID,RequestType/RequestType,RequestType/RequestCode,ChangeRequestType/ID,Author/ID,Author/Title,TemplateType/TemplateTypeName,TemplateTypeId")
    .expand("TemplateType,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,RequestType,Author,Department")
    ()
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
export const getItemByIDCRlatest = async (_sp, id) => {

  let arr = []
  let arrs = []
  let bannerimg = []

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Department/ID,Department/Department,Location/ID,Location/Location,Custodian/ID,Custodian/Custodian,DocumentType/ID,DocumentType/DocumentType,AmendmentType/ID,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,RequestType/ID,RequestType/RequestType,RequestType/RequestCode,ChangeRequestType/ID,Author/ID,Author/Title,TemplateType/TemplateTypeName,TemplateType/TemplateTypeValue,TemplateTypeId")
    .expand("TemplateType,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,RequestType,Author,Department")
    .filter(`TemplateType/TemplateTypeValue eq 'Change Request'`)
    .orderBy("ID", false) // Order by ID descending to get latest first
    .top(1)
    ()
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
      console.log(res, 'ghghghghgh let arrs=[]');
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
export const getDocumentLinkByIDarr = async (_sp, itemId) => {

  let reqId = [];
  await _sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(itemId)
    .select("*,FileRef, FileLeafRef")()
    .then((res) => {
      console.log(res, 'file let arrs=[]');


      //  arr =(res[0].Id)
      // arr = res;
      reqId.push(res)
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr arrrr');
  return reqId;
}
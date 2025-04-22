import Swal from 'sweetalert2';
export const getLatestChangeRequestTemplateType = async (_sp, List) => {
  let arr = [];
  // var List ="Annual Audit Program"
  // const spCache = spfi(_self._sp).using(Caching({ store: "session" }));
  // const listItems = await sp.web.lists.getByTitle("AuditProgramTypeMaster").items();
  await _sp.web.lists.getByTitle("ChangeRequestList").items.filter(`TemplateType/TemplateTypeName eq '${List}' and Status eq 'Approved'`).orderBy("ID", false).top(1)()
    .then((res) => {
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}
export const getGeneratedTemplateDocCR = async (_sp, itemId) => {
  let results = [];
  // for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("AnnualAuditReportCheckListGeneratedTemplateDoc").items
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
export const getAllApprovedAuditplan = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("AnnualAuditPlanList").items
    .select("*,Author/ID,Author/Title,AuditPlanType/AuditPlanType,AuditPlanType/ID,To/ID,To/Title,Cc/ID,Cc/Title,From/ID,From/Title,From/EMail")
    .expand("Author,AuditPlanType,To,Cc,From")
    .filter(`Status eq 'Approved'`)
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log("rerererepo", res);

      // Filter only latest entry for each unique DocumentCode
      const latestDocuments = res.reduce((acc, item) => {
        if (!acc[item.ID]) {
          acc[item.ID] = item;
        }
        return acc;
      }, {});
      
      //arr = res;
      arr = Object.values(latestDocuments);
      console.log("accccc", arr);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};

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

export const getItemByID = async (_sp, id) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("AnnualAuditReportList").items.getById(id)
    .select("*,Author/ID,Author/Title,ApprovedAuditPlan/MemoNumber,ApprovedAuditPlan/ID,Department/ID,Department/Department,AnnualAuditPlanDocumentLink/ID,Attachment/ID").expand("Author,Department,ApprovedAuditPlan,AnnualAuditPlanDocumentLink,Attachment")()
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

export const getAllProcessData = async (_sp, MainId, processName, docCode) => {

  let arr;
  // .select("*,Author/ID,Author/Title,Approvers/Id,Approvers/Title,ApproverRole/Id").expand("Author,Approvers,ApproverRole").filter(`MainListID eq '${MainId}' and ProcessName eq '${processName}' and RequestId eq '${docCode}'`)()

  // const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("AllProcessApprovalLevelList").items
    .select("*,Author/ID,Author/Title,Approvers/Id,Approvers/Title,ApproverRole/Id").expand("Author,Approvers,ApproverRole").filter(`MainListID eq '${MainId}' and ProcessName eq '${processName}'`)()
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
        if (!acc[item.Department]) {
          acc[item.Department] = item;
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
export const getAllDepartment1 = async (_sp) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("DepartmentMasterList").items
    .select("*").filter("Active eq 'Yes'")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      arr = res.map((item) => ({
        value: item.Id,
        label: item.Department,
        Department: item.Department,

      }));


      //  arr.push(res)
      // arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}

export const addItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditReportList').items.add(itemData);

    // console.log('Item added successfully:', newItem);
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

export const updateItem2 = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditReportCheckList').items.getById(id).update(itemData);
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

export const getAllAuditType = async (_sp) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("AuditPlanTypeMaster").items
    .select("*").filter("IsActive eq 'Yes'")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      //   arr = res.map((item) => ({
      //       value: item.Id,
      //       label: item.Department,
      //       Department:item.Department,

      // }));


      arr.push(res)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}

export const addItem2 = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditReportCheckList').items.add(itemData);

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
      if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0) {
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


export const getItemByID2 = async (sp, AuditID) => {
  debugger
  let arr = []
  let sampleDataArray = []
  arr = await sp.web.lists.getByTitle("AnnualAuditReportCheckList").items.select("*,AnnualAuditReportID/ID,Sharewith/ID,Sharewith/Title").expand("AnnualAuditReportID,Sharewith").filter(`AnnualAuditReportID/ID eq ${AuditID}`).getAll();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}
export const getItemfromChecklistMaster = async (sp) => {
  debugger
  let arr = []
  let sampleDataArray = []
  arr = await sp.web.lists.getByTitle("AnnualAuditReportCheckListMaster").items.select("*,Sharewith/ID,Sharewith/Title").expand("Sharewith").getAll();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}

export const getDraftApprovalByID = async (_sp, id, processName) => {

  let arr = [];
  let val = "Yes"
  let Sts = "Save as draft";
  let sts = "Pending"
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`AssignedTo/ID  eq '${currentUser.Id}' and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).top(1)()
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

export const uploadAllFiles = async (files, sp, docLib) => {
  // Ensure files is an array
  const filesArray = Array.isArray(files) ? files : [files];
  console.log(filesArray, "filesArray")
  debugger
  //alert(`Files: ${JSON.stringify(filesArray)}`);
  console.log(filesArray, "Files Array");

  // Proceed with mapping only if filesArray is valid
  const uploadPromises = filesArray.map(file =>
    uploadFileToLibrary(file, sp, docLib)
  );

  const uploadResults = await Promise.all(uploadPromises);
  return uploadResults.flat(); // Flatten the results if each upload returns an array
};

export const uploadFileToLibrary = async (file, sp, docLib) => {
  let arrFIleData = [];
  let fileSize = 0;
  const folder = sp.web.getFolderByServerRelativePath('/sites/edcspfx/AnnualAuditReportDocs');
  try {
    // await sp.web.lists.getByTitle(docLib).rootFolder
    const result = folder.files.addChunked(file.name, file, (progress, data) => {
      console.log(progress, data);
      fileSize = progress.fileSize;
    }, true);
    const item = await sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("*", "ID", "AuthorId", "Modified");
    console.log(item.Id, 'itemitem');
    let arr = {
      ID: item.Id,
      Createdby: item.AuthorId,
      Modified: item.Modified,
      fileUrl: result.data.ServerRelativeUrl,
      fileSize: fileSize,
      fileType: file.type,
      fileName: file.name,
    }
    arrFIleData.push(arr);
    console.log(arrFIleData, 'arrFIleData');
    return arrFIleData;
  } catch (error) {
    console.log("Error uploading file:", error);
    return null;
  }
};
export const getDocumentLinkByIDPlan = async (_sp, AttachmentIds) => {
  let results = [];
  for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("AnnualAuditPlanDocs").items.getById(itemId)
      .select("*,FileRef, FileLeafRef")()
      .then((res) => {
        console.log(res, ' let arrs=[] report plan doc');
        results.push(res);
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
  }
  console.log(results, 'results');
  return results;
}
export const getDocumentLinkByID = async (_sp, AttachmentIds) => {
  let results = [];
  for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("AnnualAuditReportDocs").items.getById(itemId)
      .select("*,FileRef, FileLeafRef")()
      .then((res) => {
        console.log(res, ' let arrs=[] report docs');
        results.push(res);
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
  }
  console.log(results, 'results');
  return results;
}

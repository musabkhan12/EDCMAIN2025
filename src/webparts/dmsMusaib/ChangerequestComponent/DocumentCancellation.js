import Swal from 'sweetalert2';
export const getAllDocumentCode = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title,PreparedBy/ID,PreparedBy/Title")
    .expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author,PreparedBy")
    .filter("Status eq 'Approved' and SignedDocs eq 'Yes'")
    .orderBy("ID", false).top(5000)() // Order by Modified descending to get latest first
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

      arr = Object.values(latestByDocumentCode).filter(
        item => item.DocumentCancellationStatus !== 'Yes'
      );
      for (const item of res) {
        const docCode = item.DocumentCode;
        if (docCode && !latestByDocumentCode[docCode]) {
          latestByDocumentCode[docCode] = item;
        }
      }

      const uniqueLatestItems = Object.values(latestByDocumentCode);
      console.log("Filtered latest by DocumentCode: ", uniqueLatestItems);
      console.log("arrarr fetching data: ", arr);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
// export const getchangerequestnotes = async (_sp) => {
//   let arr = [];

//   await _sp.web.lists.getByTitle("ChangeRequestNotes").items
//     .select("*")
//     .expand("")
//     .filter("IsActive eq 'Yes'")
//     .orderBy("Modified", false).top(5000)() // Order by Modified descending to get latest first
//     .then((res) => {
//       console.log("eeee", res);
//       debugger
//       if (res.length > 0) {
//         arr = res;
//       }
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// };
export const CheckifDocumentisApproved = async (_sp, DocCode) => {
  let arr = true;
  debugger
  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Author/ID,Author/Title,RequestType/ID,RequestType/RequestType")
    .expand("Author,RequestType")
    .filter(`DocumentCode eq '${DocCode}' and RequestType/RequestType eq 'Change in Existing Documented Information' and (Status eq 'Save as draft' or Status eq 'Pending' or Status eq 'Rework' or (Status eq 'Approved' and SignedDocs eq 'No'))`)
    .orderBy("ID", false).top(5000)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log("eeee", res);
      debugger

      if (res.length > 0) {
        arr = false;
      } else {
        arr = true;
      }
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
    .select("ChangeRequestType/ID,TemplateType/TemplateTypeName,TemplateTypeId,DocumentCode,IssueDate,RevisionDate,SerialNumber,IssueNumber,RevisionNumber")
    .expand("TemplateType,ChangeRequestType")
    .filter(`TemplateType/TemplateTypeName eq 'Change Request' and Status eq 'Approved'`)
    .orderBy("ID", false).top(5000)() // Order by Modified descending to get latest first
    .then((res) => {
      // TemplateType/TemplateTypeName eq 'Change Request' and 
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
export const getChangeRequestTypeMaster = async (_sp, RequestType) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestTypeMaster").items
    .select("*", "ID", "ChangeRequestType", "RequestType/RequestType", "RequestType/ID")
    .expand("RequestType")
    .filter(`RequestType/RequestType eq '${RequestType}'`)
    .orderBy("ID", false)
    .top(5000)()
    .then((res) => {
      if (res.length > 0) {
        arr = res.map((item) => ({
          id: item.ID,
          name: item.ChangeRequestType,
          requestType: item.RequestType?.RequestType,
        }));
      }
    })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });

  return arr;
};


export const getDocumentCodeselected = async (_sp, locId, custoId, doctypeId) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title,PreparedBy/ID,PreparedBy/Title")
    .expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author,PreparedBy")
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
export const getDocumentCodeselectedApproved = async (_sp, doccode, locId, custoId, doctypeId) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestList").items
    .select("*,Location/ID,Custodian/ID,DocumentType/ID,AmendmentType/ID,Classification/ID,ChangeRequestType/ID,Author/ID,Author/Title,PreparedBy/ID,PreparedBy/Title")
    .expand("DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,Author,PreparedBy")
    .filter(`DocumentCode eq '${doccode}' and Status eq 'Approved'`)
    .orderBy("ID", true).top(1)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);
      let SnoArr = [];
      if (res.length > 0) {
        SnoArr.push({
          SerialNo: Number(res[0].SerialNumber),
          IssueNo: Number(res[0].IssueNumber),
          RevisionNo: Number(res[0].RevisionNumber),
          IssueDate: res[0].IssueDate
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

  await _sp.web.lists.getByTitle("ProcessDepartmentMasterList").items
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
  await _sp.web.lists.getByTitle("ChangeRequestDigitalSignedDocs").items
    .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      if (res.length > 0) {
        results = res;
      } else {
        results = [];
      }

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  await _sp.web.lists.getByTitle("ChangeRequestGeneratedTemplateDoc").items
    .select("*,FileRef, FileLeafRef").filter(`ListItemID/ID eq ${itemId}`)()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      if (results.length == 0) {
        results = res;
      }

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

export const updateDigitalsign = async (listname, _sp, id, formitemid) => {
  let resultArr = []
  try {
    console.log("iddddd", id);
    // const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
    //   .filter(`ListName eq '${listname}' and ListItemID eq ${id}`)
    //   .top(1)
    //   ().then(async (res) => {
    let newItem;
    const postPayload2 = {
      DocSignedStatus: "Yes"
    }
    const newItem1 = await getdigitalsignaturerequestbyID("ChangeRequestList", _sp, Number(formitemid))
      .then(async (res) => {
        console.log("eeee digi doc", res);
        for (var i = 0; i < res.length; i++) {
          newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items.getById(res[i].ID).update(postPayload2);
          console.log('Item added successfully:', newItem);

        }
      })
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
      .top(100)
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
export const getdigitalsignaturerequestbyIDYes = async (listname, _sp, id) => {
  let arr = [];
  let Norecrodsexist = "No";
  try {
    console.log("iddddd", id);
    const newItem = await _sp.web.lists.getByTitle('DigitalSignatureRequestList').items
      .filter(`ListName eq '${listname}' and ListItemID eq ${id}`)
      .top(100)
      ()
      .then((res) => {
        console.log(res, ' let arrs=[]');
        if (res.length > 0) {

          for (let i = 0; i < res.length; i++) {
            if (res[i].DocSignedStatus === "No" || res[i].DestinationIDUpdated === "No") {
              Norecrodsexist = "Yes";
              break; // Exit the loop early since we found a match
            }
          }

          arr = res
        } else {
          Norecrodsexist = "NoRecord";
        }

        // arr = res;
      })
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    arr = null
  }
  console.log("NorecrodsexistNorecrodsexist", Norecrodsexist);
  return Norecrodsexist;
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
  debugger
  let resultArr = [];
  let newItem = [];
  try {
    const newitemnew = await _sp.web.lists.getByTitle('AllProcessApprovalLevelList').items
      .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,Approvers/Id,Approvers/Title")
      .expand("Author,RequesterName,Approvers")
      .filter(`MainListID eq ${id}`).orderBy("Level", false)
      ().then(async (res) => {
        if (res.length > 0) {
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
export const getallProcessApprovalitemsLevel = async (_sp, id) => {
  debugger
  let resultArr = "";
  let newItem = "";
  try {
    const newitemnew = await _sp.web.lists.getByTitle('AllProcessApprovalLevelList').items
      .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,Approvers/Id,Approvers/Title")
      .expand("Author,RequesterName,Approvers")
      .filter(`MainListID eq ${id}`).orderBy("Level", false)
      ().then(async (res) => {
        if (res.length > 0) {
          newItem = res[0].Level
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
//
export const getDelegateduser = async (_sp, ApproverId) => {
  debugger
  let resultArr;
  let newItem = "";
  const todayISO = new Date().toISOString().split('T')[0];
  const currentUser = await _sp.web.currentUser();
  let isDelegated = false;
  try {

    const newitemnew = await _sp.web.lists.getByTitle('ARGDelegateList').items
      .select("*,Author/ID,Author/Title,ActingFor/Id,ActingFor/Title,ActingFor/EMail,DelegateName/Id,DelegateName/Title,DelegateName/EMail")
      .expand("Author,ActingFor,DelegateName")
      .filter(`DelegateNameId eq ${ApproverId} and ActingForId eq ${currentUser.Id} and StartDate le datetime'${todayISO}T00:00:00Z' and EndDate ge datetime'${todayISO}T00:00:00Z'`)
      ()
      .then(delegates => {
        if (delegates.length > 0) {
          newItem = delegates[0]?.ActingFor;
          isDelegated = true;
          arr = res;
        }
      });
    console.log("Actingforrr", newitemnew);

    resultArr = newitemnew
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return isDelegated;
};
//////////
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
    .select("*,Department/ID,Department/Department,Location/ID,Location/Location,Custodian/ID,Custodian/Custodian,DocumentType/ID,DocumentType/DocumentType,AmendmentType/ID,AmendmentType/AmendmentType,Classification/ID,Classification/Classification,RequestType/ID,RequestType/RequestType,RequestType/RequestCode,ChangeRequestType/ID,Author/ID,Author/Title,TemplateType/TemplateTypeName,TemplateTypeId,PreparedBy/ID,PreparedBy/Title,PreparedBy/EMail")
    .expand("TemplateType,DocumentType,Custodian,Classification,AmendmentType,Location,ChangeRequestType,RequestType,Author,Department,PreparedBy")
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
    .select("ChangeRequestType/ID,TemplateType/TemplateTypeName,TemplateTypeId,DocumentCode,IssueDate,RevisionDate,SerialNumber,IssueNumber,RevisionNumber")
    .expand("TemplateType,ChangeRequestType")
    .filter(`TemplateType/TemplateTypeValue eq 'Change Request' and Status eq 'Approved'`)
    .orderBy("ID", false) // Order by ID descending to get latest first
    .top(1)
    ()
    .then((res) => {
      console.log(res, ' let arrs=[]');

      arr = res
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

  let arr = [];
  let arrs = [];
  let bannerimg = [];
  let isDelegated = false;
  const todayISO = new Date().toISOString().split('T')[0];
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
    .then(async (res) => {
      console.log(res, 'ghghghghgh let arrs=[]');
      if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName) {
        arr = res;
      } else {
        const newitemnew = await _sp.web.lists.getByTitle('ARGDelegateList').items
          .select("*,Author/ID,Author/Title,ActingFor/Id,ActingFor/Title,ActingFor/EMail,DelegateName/Id,DelegateName/Title,DelegateName/EMail")
          .expand("Author,ActingFor,DelegateName")
          .filter(`DelegateNameId eq ${res.AssignedTo.Id} and ActingForId eq ${currentUser.Id} and StartDate le datetime'${todayISO}T00:00:00Z' and EndDate ge datetime'${todayISO}T00:00:00Z'`)
          ()
          .then(delegates => {
            if (delegates.length > 0) {
              //newItem = delegates[0]?.ActingFor;
              isDelegated = true;
              arr = res;
            }
          });
      }
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}
export const CheckIfAlreadyactionTaken = async (_sp, id, processName) => {
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

export const getApprovalByID2 = async (_sp, id, processName) => {

  let arr;
  let arrs = []
  let bannerimg = [];
  let isDelegated = false;
  const todayISO = new Date().toISOString().split('T')[0];
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
    .then(async (res) => {
      console.log(res, ' let arrs=[]');
      if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 && res.CurrentUserRole !== "OES") {
        arr = false;
      }
      else {
        const newitemnew = await _sp.web.lists.getByTitle('ARGDelegateList').items
          .select("ActingFor/Id,ActingFor/Title,ActingFor/EMail,DelegateName/Id,DelegateName/Title,DelegateName/EMail,StartDate,EndDate")
          .expand("ActingFor,DelegateName")
          .filter(`DelegateNameId eq ${res.AssignedTo.Id} and ActingForId eq ${currentUser.Id} and StartDate le datetime'${todayISO}T00:00:00Z' and EndDate ge datetime'${todayISO}T00:00:00Z'`)
          ()
          .then(delegates => {
            if (delegates.length > 0) {
              // newItem = delegates[0]?.ActingFor;
              isDelegated = true;
            }
          });
        if (isDelegated) {
          arr = false;
        } else {
          arr = true;
        }

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
  debugger
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
export const getchangerequestnotes = async (_sp) => {
  let arr = [];

  await _sp.web.lists.getByTitle("ChangeRequestNotes").items
    .select("*")
    .expand("")
    .filter("IsActive eq 'Yes'")
    .orderBy("Modified", false).top(5000)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log("eeee", res);
      debugger
      if (res.length > 0) {
        arr = res;
      }
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
};
export const getDocumentLinkByID = async (_sp, itemId, listid) => {
  debugger
  var reqId;
  // await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items
  //   .select("*,FileRef, FileLeafRef")
  //   .expand()
  //   .filter(`ListItemIDId eq ${listid}`)
  //   ()
  //   .then(async (res) => {
  //     console.log(res, ' let arrs=[] ghghgh');
  //     if (res.length > 0) {
  //       reqId = res[0]
  //     } else {
  //       await _sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(itemId)
  //         .select("*,FileRef, FileLeafRef")()
  //         .then((reschdoc) => {
  //           console.log(reschdoc, ' let arrs=[] reschdoc');


  //           //  arr =(res[0].Id)
  //           // arr = res;
  //           reqId = reschdoc
  //         })
  //     }
  //   })
  //   .catch((error) => {
  //     console.log("Error fetching data: ", error);
  //   });
  // console.log(reqId, 'arr');
  // return reqId;
  let arrs = [];

  try {
    // 1. Get from ChangeRequestAttachDigitalSignedDocs
    const signedDocs = await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items
      .select("*,FileRef,FileLeafRef")
      .filter(`ListItemIDId eq ${listid}`)();

    if (signedDocs && signedDocs.length > 0) {
      arrs.push(...signedDocs); // Add all signed docs
    }

    // 2. Get from ChangeRequestDocs by itemId
    const changeRequestDoc = await _sp.web.lists.getByTitle("ChangeRequestDocs").items
      .getById(itemId)
      //.select("*,FileRef,FileLeafRef")
      .select("Title,FileName,Created,FileLeafRef,FileRef,ID")
      ();

    if (changeRequestDoc) {
      arrs.push(changeRequestDoc); // Add regular doc
    }

    console.log("Combined Documents:", arrs);
  } catch (error) {
    console.error("Error fetching documents:", error);
  }

  console.log(arrs, 'arrssss');
  //return reqId;
  return arrs;
}
export const getDocumentLinkByIDarr = async (_sp, itemId, listid) => {

  let reqId = [];
  await _sp.web.lists.getByTitle("ChangeRequestAttachDigitalSignedDocs").items
    .select("*,FileRef, FileLeafRef")
    .expand()
    .filter(`ListItemIDId eq ${listid}`)
    ()
    .then(async (res) => {
      console.log(res, ' let arrs=[] ghghgh att');


      //  arr =(res[0].Id)
      // arr = res;
      if (res.length > 0) {
        reqId = res
      } else {
        await _sp.web.lists.getByTitle("ChangeRequestDocs").items.getById(itemId)
          //.select("*,FileRef, FileLeafRef")
          .select("*,Title,FileName,Created,FileLeafRef,FileRef,ID")
          ()
          .then((res) => {
            console.log(res, 'file leatttttt arrs=[]');


            //  arr =(res[0].Id)
            // arr = res;
            reqId.push(res)
          })
      }
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr arrrr');
  return reqId;
}
export const getTemplatelink = async (_sp) => {

  let reqId = [];
  await _sp.web.lists.getByTitle("ChangeRequestTemplate").items
    .select("*,FileRef, FileLeafRef").orderBy("SNo", true)()
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
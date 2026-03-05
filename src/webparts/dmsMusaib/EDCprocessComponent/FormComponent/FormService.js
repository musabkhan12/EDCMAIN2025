import Swal from 'sweetalert2';
export const getDataRoles = async (_sp) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ApproverRoleMaster").items
    // .select("*,ToUsers")()
    .select("*,ToUsers/Title").expand("ToUsers").filter("IsActive eq 'Yes'")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');      

      // arr = res.map((item) => ({
      //   value: item.Id,
      //   label: item.Department,
      //   Department: item.Department,
      //   DepartmentCode: item.DepartmentCode,
      //   ADDepartmentName: item.ADDepartmentName,
      //   ToUsers: item.ToUsersId || [],
      //   CCUsers: item.CCUsersId || [],
      //   ToUsersTitle: item.ToUsers || [],
      //   CCUsersTitle: item.CCUsers || [],

      // }));
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
  await _sp.web.lists.getByTitle("AnnualAuditProgram").items.getById(id)
  .select("*,RecommendationType/RecommendationTypeValue,Author/ID,Author/Title,AuditType/Title,AuditType/ID,To/ID,To/Title,Cc/ID,Cc/Title,From/ID,From/Title,From/EMail,CCDepartments/ID,CCDepartments/Department,CCDepartments/DepartmentCode,ToDepartments/ID,ToDepartments/Role,AuditProgramType/Title,AuditProgramType/ID").expand("RecommendationType,AuditProgramType,ToDepartments,CCDepartments,Author,AuditType,To,Cc,From")()

   // .select("*,RecommendationType/RecommendationTypeValue,Author/ID,Author/Title,AuditType/Title,AuditType/ID,To/ID,To/Title,Cc/ID,Cc/Title,From/ID,From/Title,From/EMail,CCDepartments/ID,CCDepartments/Department,CCDepartments/DepartmentCode,ToDepartments/ID,ToDepartments/Department,ToDepartments/DepartmentCode,AuditProgramType/Title,AuditProgramType/ID").expand("RecommendationType,AuditProgramType,ToDepartments,CCDepartments,Author,AuditType,To,Cc,From")()
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

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ProcessDepartmentMasterList").items
    .select("*,ToUsers/Title,CCUsers/Title").expand("ToUsers,CCUsers").filter("Active eq 'Yes'")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      arr = res.map((item) => ({
        value: item.Id,
        label: item.Department,
        Department: item.Department,
        DepartmentCode: item.DepartmentCode,
        ADDepartmentName: item.ADDepartmentName,
        ToUsers: item.ToUsersId || [],
        CCUsers: item.CCUsersId || [],
        ToUsersTitle: item.ToUsers || [],
        CCUsersTitle: item.CCUsers || [],

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


export const getAuditTypes = async (_sp) => {
  let arr = []
  // const spCache = spfi(_self._sp).using(Caching({ store: "session" }));
  // const listItems = await sp.web.lists.getByTitle("AuditProgramTypeMaster").items();
  await _sp.web.lists.getByTitle("AuditProgramTypeMaster").items.filter("IsActive eq 1")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      //     arr = res.map((item) => ({
      //         value: item.Id,
      //         label: item.Title,


      //   }));


      //  arr.push(res)
      arr = res;
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
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditProgram').items.add(itemData);

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
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditProgram').items.getById(id).update(itemData);
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
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditProgramRecommendationList').items.getById(id).update(itemData);
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
  await _sp.web.lists.getByTitle("AuditTypeMaster").items
    .select("*").filter("IsActive eq 1")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      //   arr = res.map((item) => ({
      //       value: item.Id,
      //       label: item.Department,
      //       Department:item.Department,

      // }));


      //  arr.push(res)
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
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditProgramRecommendationList').items.add(itemData);

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
      reqId = res[0]
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(reqId, 'arr');
  return reqId;
}

// export const getApprovalByID = async (_sp, id,processName) => {

//   let arr = []
//   let arrs = []
//   let bannerimg = []
//   const currentUser = await _sp.web.currentUser();
//   await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
//   .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
//     .then((res) => {
//       console.log(res, ' let arrs=[]');
//       if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName ){
//         arr = res;
//       }
//       // .filter(`AssignedTo/Id eq ${currentUser.Id} and ProcessName eq ${processName}`)

//       //  arr.push(res)

//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   console.log(arr, 'arr');
//   return arr;
// }
export const getApprovalByID = async (_sp, id, processName) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
    .then(async (res) => {
      // console.log(res, ' let arrs=[]');
      // if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName) {
      //   arr = res;
      // }





      // working *********

      // Check if the current user is the assigned user or a delegate
      // Fetch the delegate list to see if the current user is acting as a delegate
      const today = new Date().toISOString();
      // const today = new Date().toISOString().split('T')[0];
      await _sp.web.lists.getByTitle("ARGDelegateList")
        .items
        .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
        .expand("Author,DelegateName,ActingFor")
        .filter(`DelegateName/ID eq '${res.AssignedTo?.Id}' and ActingFor/ID eq '${currentUser.Id}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
        .orderBy("Created", false).top(1)()
        .then((result) => {
          if (result.length > 0) {
            // If the current user is a delegate, check if they are acting for the assigned user
            // and if the process name matches
            if (res && (res.AssignedTo.Id == currentUser.Id || res.AssignedTo.Id == result[0].DelegateNameId) && res.ProcessName === processName) {
              arr = res;
            }
            // if (res && (res.AssignedTo.Id == currentUser.Id || res.ActingFor.Id == currentUser.Id) && res.ProcessName === processName) {
            //   arr = res;
            // }


          }
          else {
            if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName) {
              arr = res;
            }

          }


        })
        .catch((error) => {
          console.log("Error fetching data: ", error);
        });

      // working *********




    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}

// export const getApprovalByID2 = async (_sp, id,processName) => {

//   let arr;
//   let arrs = []
//   let bannerimg = []
//   const currentUser = await _sp.web.currentUser();
//   await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
//   .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
//     .then((res) => {
//       console.log(res, ' let arrs=[]');
//       if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 ){
//         arr = false;
//       }
//       else{
//         arr = true;
//       }
//       // .filter(`AssignedTo/Id eq ${currentUser.Id} and ProcessName eq ${processName}`)

//       //  arr.push(res)

//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   console.log(arr, 'arr');
//   return arr;
// }
export const getApprovalByID2 = async (_sp, id, processName) => {

  let arr = true;;
  let arrs = []
  let bannerimg = []
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
    .then(async (res) => {
      console.log(res, ' let arrs=[]');
      // if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 ){
      //   arr = false;
      // }
      // else{
      //   arr = true;
      // }
      const today = new Date().toISOString();
      // const today = new Date().toISOString().split('T')[0];
      await _sp.web.lists.getByTitle("ARGDelegateList")
        .items
        .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
        .expand("Author,DelegateName,ActingFor")
        .filter(`DelegateName/ID eq '${res.AssignedTo?.Id}' and ActingFor/ID eq '${currentUser.Id}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
        .orderBy("Created", false).top(1)()
        .then((result) => {
          if (result.length > 0) {
            // If the current user is a delegate, check if they are acting for the assigned user
            // and if the process name matches
            if (res && (res.AssignedTo.Id == currentUser.Id || res.AssignedTo.Id == result[0].DelegateNameId) && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0) {
              arr = false;
            }
            //  if (res && (res.AssignedTo.Id == currentUser.Id || res.ActingFor.Id == currentUser.Id) && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0) {
            //   arr = false;
            // }

          }
          else {
            if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0) {
              arr = false;
            }

          }


        })
        .catch((error) => {
          console.log("Error fetching data: ", error);
        });


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
  //  var listname = "AnnualAuditPlanRecommendationList";
  //  var listname = "MemorandumRecommendationList"
  //  arr = await sp.web.lists.getByTitle(`${listname}`).items.select("*,Memorandum/ID,Auditor/ID,Auditor/Title").expand("Memorandum,Auditor").filter(`Memorandum/ID eq ${AuditID}`).getAll();

  arr = await sp.web.lists.getByTitle("AnnualAuditProgramRecommendationList").items.select("*,AnnualAuditProgram/ID,Auditor/ID,Auditor/Title,Auditors/ID,Auditors/Role").expand("AnnualAuditProgram,Auditor,Auditors").filter(`AnnualAuditProgram/ID eq ${AuditID}`).getAll();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}


//   export const getDraftApprovalByID = async (_sp, id,processName) => {

//   let arr =[];
//   let val = "Yes"
//   let Sts = "Save as draft";
//   let sts ="Pending"
//   const currentUser = await _sp.web.currentUser();
//   await _sp.web.lists.getByTitle("ProcessApprovalList").items
//   .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`ListItemId eq '${id}' and AssignedTo/ID  eq '${currentUser.Id}' and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).top(1)()
//     .then((res) => {
//       console.log(res, ' let arrs=[]');
//       // if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 && res.CurrentUserRole !=="OES" ){
//       //   arr = false;
//       // }
//       // else{
//       //   arr = true;
//       // }
//       // .filter(`AssignedTo/Id eq ${currentUser.Id} and ProcessName eq ${processName}`)

//       //  arr.push(res)

//       arr =res;

//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   console.log(arr, 'arr');
//   return arr;
// }
export const getDraftApprovalByID = async (_sp, id, processName) => {

  let arr = [];
  let val = "Yes"
  let Sts = "Save as draft";
  let sts = "Pending"
  const currentUser = await _sp.web.currentUser();
  const today = new Date().toISOString();


  await _sp.web.lists.getByTitle("ProcessApprovalList").items
    // .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`ListItemId eq '${id}' and AssignedTo/ID  eq '${currentUser.Id}' and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).orderBy("Created", false).top(1)()
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`ListItemId eq '${id}' and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).orderBy("Created", false).top(1)()

    .then(async (res) => {
      console.log(res, ' let arrs=[]');

      arr = res;

      await _sp.web.lists.getByTitle("ARGDelegateList")
        .items
        .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
        .expand("Author,DelegateName,ActingFor")
        .filter(`DelegateName/ID eq '${res.AssignedTo?.Id}' and ActingFor/ID eq '${currentUser.Id}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
        .orderBy("Created", false).top(1)()
        .then(async (result) => {
          if (result.length > 0) {

            if (res && (res.AssignedTo.Id == currentUser.Id || res.AssignedTo.Id == result[0].DelegateNameId) && res.ProcessName === processName && (res?.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0) {
              arr = res;
            }

          }
          else {

            if (res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0) {
              arr = res;
            }


          }


        })
        .catch((error) => {
          console.log("Error fetching data: ", error);
        });

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });


  // // //// working *********
  // // const today = new Date().toISOString().split('T')[0];
  // await _sp.web.lists.getByTitle("ARGDelegateList")
  //   .items
  //   .select("*,Author/ID,Author/Title,Author/EMail,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
  //   .expand("Author,DelegateName,ActingFor")
  //   .filter(`DelegateName/ID eq '${currentUser.Id}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)
  //   .orderBy("Created", false).top(1)()
  //   .then(async (result) => {
  //     if (result.length > 0) {



  //       await _sp.web.lists.getByTitle("ProcessApprovalList").items
  //         .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`ListItemId eq '${id}' and (AssignedTo/ID  eq '${result[0].ActingForId}' or AssignedTo/ID  eq '${currentUser.Id}') and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).top(1)()
  //         .then((res) => {
  //           console.log(res, ' let arrs=[]');

  //           arr = res;

  //         })
  //         .catch((error) => {
  //           console.log("Error fetching data: ", error);
  //         });


  //     }
  //     else {

  //       await _sp.web.lists.getByTitle("ProcessApprovalList").items
  //         .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/ID,AssignedTo/Title").expand("Author,RequesterName,AssignedTo").filter(`ListItemId eq '${id}' and AssignedTo/ID  eq '${currentUser.Id}' and ProcessName eq '${processName}' and IsInitiator eq '${val}' and (Status eq '${Sts}' or Status eq '${sts}')`).top(1)()
  //         .then((res) => {
  //           console.log(res, ' let arrs=[]');

  //           arr = res;

  //         })
  //         .catch((error) => {
  //           console.log("Error fetching data: ", error);
  //         });


  //     }


  //   })
  //   .catch((error) => {
  //     console.log("Error fetching data: ", error);
  //   });

  // // //// working *********


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
  const folder = sp.web.getFolderByServerRelativePath('/sites/ED/AnnualAuditProgramDocs');
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

export const getDocumentLinkByID = async (_sp, AttachmentIds) => {
  let results = [];
  for (let itemId of AttachmentIds) {
    await _sp.web.lists.getByTitle("AnnualAuditProgramDocs").items.getById(itemId)
      .select("*,FileRef, FileLeafRef")()
      .then((res) => {
        console.log(res, ' let arrs=[]');
        results.push(res);
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
  }
  console.log(results, 'results');
  return results;
}

// export const getGeneratedTemplateDoc = async (_sp, itemId) => {
//   let results = [];
//   // for (let itemId of AttachmentIds) {
//     await _sp.web.lists.getByTitle("AnnualAuditProgramGeneratedTemplateDoc").items
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

export const getGeneratedTemplateDoc = async (_sp, itemId) => {
  let results = [];
  try {
    const res = await _sp.web.lists.getByTitle("AnnualAuditProgramDigitalSignedDocs").items
      .select("*,FileRef, FileLeafRef")
      .filter(`ListItemID/ID eq ${itemId}`)
      .orderBy("ID", false)
      .top(1)();

    if (res && res.length > 0) {
      results = res;
    } else {
      const res2 = await _sp.web.lists.getByTitle("AnnualAuditProgramGeneratedTemplateDoc").items
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


export const getRecommendationTypes = async (_sp) => {
  let arr = []
  // const spCache = spfi(_self._sp).using(Caching({ store: "session" }));
  // const listItems = await sp.web.lists.getByTitle("AuditProgramTypeMaster").items();
  await _sp.web.lists.getByTitle("RecommendationTypeMaster").items()
    .then((res) => {

      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}

export const getLatestChangeRequestTemplateType = async (_sp, List) => {
  let arr = [];
  // var List ="Annual Audit Program"
  // const spCache = spfi(_self._sp).using(Caching({ store: "session" }));
  // const listItems = await sp.web.lists.getByTitle("AuditProgramTypeMaster").items();
  await _sp.web.lists.getByTitle("ChangeRequestList").items.filter(`TemplateType/TemplateTypeValue eq '${List}' and Status eq 'Approved'`).orderBy("ID", false).top(1)()
    .then((res) => {
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}
export const getAllMemoNumberList = async (_sp) => {
  let arr = [];
  let sts = "Approved";

  await _sp.web.lists.getByTitle("Memorandum").items.filter(`Status eq '${sts}'`)
    .select("*,Department/ID,Department/Department,From/ID,From/Title,From/EMail,AuditType/Title,AuditType/ID,To/ID,To/Title,Cc/ID,Cc/Title,CCDepartments/ID,CCDepartments/Department,CCDepartments/DepartmentCode,ToDepartments/ID,ToDepartments/Department,ToDepartments/DepartmentCode,RecommendationType/RecommendationTypeValue,RecommendationType/ID")
    .expand("From,Department,AuditType,To,Cc,CCDepartments,ToDepartments,RecommendationType")
    .orderBy("Modified", false)() // Order by Modified descending to get latest first
    .then((res) => {
      console.log(res);

      // Filter only latest entry for each unique DocumentCode
      // const latestDocuments = res.reduce((acc, item) => {
      //   if (!acc[item.DocumentCode]) {
      //     acc[item.DocumentCode] = item;
      //   }
      //   return acc;
      // }, {});

      // arr = Object.values(latestDocuments);
      arr = res;
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
export const addMemoNumber = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('MemoNumberLogic').items.add(itemData);

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


export const addYearlyList = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('AnnualAuditProgramYearlyList').items.add(itemData);

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

export const UpdatYearlyList = async (itemData, _sp, id) => {

  let resultArr = []
  try {

    const newItem = await _sp.web.lists.getByTitle('AnnualAuditProgramYearlyList').items.getById(id).update(itemData);
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


export const getYearlyItemByID = async (sp, AuditID) => {
  debugger
  let arr = []
  //  var listname = "AnnualAuditPlanRecommendationList";
  //  var listname = "MemorandumRecommendationList"
  //  arr = await sp.web.lists.getByTitle(`${listname}`).items.select("*,Memorandum/ID,Auditor/ID,Auditor/Title").expand("Memorandum,Auditor").filter(`Memorandum/ID eq ${AuditID}`).getAll();

  arr = await sp.web.lists.getByTitle("AnnualAuditProgramYearlyList").items.select("*,Department/Department,Department/ID,AnnualAuditProgramID/ID,Auditor/ID,Auditor/Role,Custodian/Custodian,Custodian/ID,Shift/Shift,Shift/ID,Location/Location,Location/ID").expand("Location,Shift,Custodian,Department,AnnualAuditProgramID,Auditor").filter(`AnnualAuditProgramID/ID eq ${AuditID}`).getAll();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}


export const getAuditProgDepartment = async (_sp) => {

  let arr = [];
  await _sp.web.lists.getByTitle("AuditProgramDepartmentMaster").items
    .select("*,Department,Location")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      arr = res.map((item) => ({
        value: item.Id,
        label: item.Department,
        Department: item.Department,
        Location: item.Location,

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

export const getAuditProgCustodian = async (_sp) => {

  let arr = [];
  await _sp.web.lists.getByTitle("AuditProgramCustodianMaster").items
    .select("Custodian,ID,Id")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      arr = res.map((item) => ({
        value: item.Id,
        label: item.Custodian,

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

export const getAuditProgShift = async (_sp) => {

  let arr = [];
  await _sp.web.lists.getByTitle("AuditProgramShiftMaster").items
    .select("Shift,ID,Id")()
    .then((res) => {
      // console.log(res, ' let arrs=[]');
      arr = res.map((item) => ({
        value: item.Id,
        label: item.Shift,

      }));

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}



export const fetchLocations = async (_sp) => {

  let arr = [];
  await _sp.web.lists.getByTitle("AuditProgramLocationMaster").items
    .top(5000)()
    .then((res) => {
      res.sort((a, b) => a.Location.localeCompare(b.Location));
      arr = res.map((item) => ({
        label: item.Location,
        value: item.ID

      }));

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  // console.log(arr, 'arr');
  return arr;
}
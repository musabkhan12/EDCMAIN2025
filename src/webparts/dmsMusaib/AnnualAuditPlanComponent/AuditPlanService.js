import Swal from 'sweetalert2';
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
    await _sp.web.lists.getByTitle("AnnualAuditPlanList").items.getById(id)
    .select("*,Author/ID,Author/Title,AuditPlanType/AuditPlanType,AuditPlanType/ID,To/ID,To/Title,Cc/ID,Cc/Title,From/ID,From/Title,From/EMail").expand("Author,AuditPlanType,To,Cc,From")()
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

  export const getAllProcessData = async (_sp, MainId,processName,docCode) => {
 
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
  
  export const UpdateAllProcessItem = async (itemData, _sp,id) => {
   
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
    await _sp.web.lists.getByTitle("DepartmentMasterList").items
    .select("*").filter("Active eq 'Yes'")()
      .then((res) => {
        // console.log(res, ' let arrs=[]');
        arr = res.map((item) => ({
            value: item.Id,
            label: item.Department,
            Department:item.Department,
         
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
      const newItem = await _sp.web.lists.getByTitle('AnnualAuditPlanList').items.add(itemData);
   
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
      const newItem = await _sp.web.lists.getByTitle('AnnualAuditPlanList').items.getById(id).update(itemData);
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
      const newItem = await _sp.web.lists.getByTitle('AnnualAuditPlanRecommendationList').items.getById(id).update(itemData);
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
      const newItem = await _sp.web.lists.getByTitle('AnnualAuditPlanRecommendationList').items.add(itemData);
   
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
  export const getListNameID = async (_sp,formname) => {
 
    var reqId;
    await _sp.web.lists.getByTitle("ListNameMaster").items
    .select("*").filter("ListName eq '"+formname+"'").top(1)()
      .then((res) => {
        console.log(res, ' let arrs=[]');
       
  
        //  arr =(res[0].Id)
        // arr = res;
        reqId=res[0].Id
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    console.log(reqId, 'arr');
    return reqId;
  }

  export const getFormNameID = async (_sp,formname) => {
 
    var reqId;
    await _sp.web.lists.getByTitle("FormNameMaster").items
    .select("*").filter("FormName eq '"+formname+"'").top(1)()
      .then((res) => {
        console.log(res, ' let arrs=[]');
       
  
        //  arr =(res[0].Id)
        // arr = res;
        reqId=res[0].Id
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    console.log(reqId, 'arr');
    return reqId;
  }

  export const getApprovalByID = async (_sp, id,processName) => {
 
    let arr = []
    let arrs = []
    let bannerimg = []
    const currentUser = await _sp.web.currentUser();
    await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
      .then((res) => {
        console.log(res, ' let arrs=[]');
        if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName ){
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

  export const getApprovalByID2 = async (_sp, id,processName) => {
 
    let arr;
    let arrs = []
    let bannerimg = []
    const currentUser = await _sp.web.currentUser();
    await _sp.web.lists.getByTitle("ProcessApprovalList").items.getById(id)
    .select("*,Author/ID,Author/Title,RequesterName/Id,RequesterName/Title,AssignedTo/Id,AssignedTo/Title").expand("Author,RequesterName,AssignedTo")()
      .then((res) => {
        console.log(res, ' let arrs=[]');
        if(res && res.AssignedTo.Id == currentUser.Id && res.ProcessName === processName && (res.Status == "Pending" || res?.Status === "Save as draft") && res.Level === 0 ){
          arr = false;
        }
        else{
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
    arr = await sp.web.lists.getByTitle("AnnualAuditPlanRecommendationList").items.select("*,AnnualAuditPlanID/ID,Auditor/ID,Auditor/Title").expand("AnnualAuditPlanID,Auditor").filter(`AnnualAuditPlanID/ID eq ${AuditID}`).getAll();
    // .then((res) => {
    //   arr = res
    //   console.log(arr, 'arr');
    // })
    return arr
  }

 
import * as React from 'react';
import type { IAnnualAuditReportProps } from './IAnnualAuditReportProps';
import { escape } from '@microsoft/sp-lodash-subset';

import Provider from '../../../GlobalContext/provider';

import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';

import HorizontalNavbar from '../../horizontalNavBar/components/HorizontalNavBar';

import { getSP } from '../../dmsMusaib/loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import UserContext from "../../../GlobalContext/context";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import "./annualauditReport.scss";
import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import Select from "react-select";
import Swal from 'sweetalert2';
import { FormSubmissionMode } from '../../../Shared/Interfaces';
import { decryptId } from '../../../APISearvice/CryptoService';
import { WorkflowAction } from '../../../CustomJSComponents/WorkflowAction/WorkflowAction';
import { WorkflowAuditHistory } from '../../../CustomJSComponents/WorkflowAuditHistory/WorkflowAuditHistory';
import { CONTENTTYPE_AuditReport, CONTENTTYPE_AuditReportNew, CONTENTTYPE_AuditReportTemp, LIST_TITLE_AuditReport, SITE_URL, Tenant_URL } from '../../../Shared/Constants';
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import { Modal } from 'react-bootstrap';
import { faDownload, faEye } from '@fortawesome/free-solid-svg-icons';
import CustomBreadcrumb from './CustomBreadcrumb/CustomBreadcrumb';
import { addAllProcessItem, addItem, addItem2, addItemNC, getAllApprovedAuditplan, getAllAuditType, getAllDepartment, getAllDepartment1, getAllProcessData, getApprovalByID, getApprovalByID2, getDataRoles, getDocumentLinkByID, getDocumentLinkByIDPlan, getDraftApprovalByID, getFormNameID, getGeneratedTemplateDocAuditplan, getGeneratedTemplateDocCR, getItemByID, getItemByID2, getItemfromChecklistMaster, getItemsAuditReportNC, getItemsAuditReportObs, getLatestChangeRequestTemplateType, getListNameID, getNCNumberbyID, UpdateAllProcessItem, updateApprovalItem, updateItem, updateItem2, updateItemNC, uploadAllFiles } from './AuditReportService';
import { TextField } from '@fluentui/react';
import { DatePicker, isMac } from 'office-ui-fabric-react';
import moment from 'moment';
//React Time picker
// import TimePicker from 'react-time-picker';
// import 'react-time-picker/dist/TimePicker.css';
// import 'react-clock/dist/Clock.css';
//MUI time picker
import { LocalizationProvider, TimePicker } from '@mui/x-date-pickers';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';
import { getAuditProgShift } from '../EDCprocessComponent/FormComponent/FormService';
import FileViewer from '../ChangerequestComponent/fileviewer';
import { CheckIfAlreadyactionTaken } from '../ChangerequestComponent/DocumentCancellation';
//SYnc time picker
// import { TimePickerComponent } from '@syncfusion/ej2-react-calendars';

// import '@syncfusion/ej2-base/styles/material.css';
// import '@syncfusion/ej2-react-calendars/styles/material.css';

// import { TimePicker } from 'antd';
// import dayjs from 'dayjs';
// import 'antd/dist/reset.css'; // Or 'antd/dist/antd.css' if using antd < 5

// let myloader = '../../'
let newfileupload: any
let newfilepreview: any;
let filechanged: boolean = false;
let ncrow: any = [];
let EnableNC: boolean = false;
let maxncseq: number = 0;
let maxobsseq: number = 0;
let ncseq: number = 0;
let obsseq: number = 0;
let departmentChanged: boolean = false;
let Showfile: boolean = false;
interface ForwardTo {
    id: number;
    role: number;
    level: number;
    approvers: any[]; // Or a more specific type like `string[]` or `SPUser[]`
    approvalType: string;
    Responsibility?: string;
    IsSignatureRequired?: boolean;
    responsibilityerror?: boolean;
}
interface ncNumber {
    Id: number,
    id: number,
    ncnumberNC: string,
    observationnumberObs: string,
    reportcode: string,
    ncsequenceNC: number,
    observationsequenceObs: number,
    nctype: string,
    DepartmentNCId: number;
    descriptionNC: string
}
const AnnualAuditReportContext = ({ props }: any) => {
    const sp: SPFI = getSP();
    const elementRef = React.useRef<HTMLDivElement>(null);
    const siteUrl = props.siteUrl;
    const { useHide }: any = React.useContext(UserContext);
    const [InputDisabled, setInputDisabled] = React.useState(false);
    const [EnableNCObs, setEnableNCObs] = React.useState(false);
    const selectedTextDiv = document.getElementById('selectedText');

    selectedTextDiv.style.display = 'none';
    const [ShowModalTemplateDoc, setShowModalTemplateDoc] = React.useState(false);
    const [AuditProgShift, setAuditProgShift] = React.useState([]);
    const [FilesArr, setFilesArr] = React.useState<any>([]);
    const [FilesArr1, setFilesArr1] = React.useState<any>([]);
    const [FilesArrDoclink, setFilesArrDoclink] = React.useState<any>([]);
    const [Loading, setLoading] = React.useState(false);
    const [FormLoading, setFormLoading] = React.useState(false);
    const [showForwardapproval, setshowForwardapproval] = React.useState(true);
    const [FormItemId, setFormItemId] = React.useState(null);
    const [editID, setEditID] = React.useState(null);
    const [editItemID, setEditItemID] = React.useState(null);
    const [MainEditItem, setMainEditItem] = React.useState(null);
    const [rows, setRows] = React.useState<any>([]);
    const [UserRoles, setUserRoles] = React.useState<any>([]);
    const [rows1, setRows1] = React.useState<any>([]);
    const [currentUser, setCurrentUser] = React.useState(null);
    const [selectedOption, setSelectedOption] = React.useState(null);
    const [selectedPeople, setSelectedPeople] = React.useState(null);
    const [selectedRole, setSelectedRole] = React.useState(null);
    const [ValidDRecomm, setValidDRecomm] = React.useState(true);
    const [ValidDraft, setValidDraft] = React.useState(true);
    const [ValidAudit, setValidAudit] = React.useState(true);
    const [ValidSubmit, setValidSubmit] = React.useState(true);
    const [ValidCancelReason, setValidCancelReason] = React.useState(true);
    const [forwardToValidationErrors, setForwardToValidationErrors] = React.useState<{ [level: number]: { role: boolean; approvers: boolean; approvalType: boolean, responsibility: boolean } }>({});
    const [ValidForwardTo, setValidForwardTo] = React.useState(true);
    const [RequesterRoleId, setRequesterRoleId] = React.useState(null);
    const [RequestTypeId, setRequestTypeId] = React.useState(null);
    const [FormNameId, setFormNameId] = React.useState(null);
    const [AuditPlanType, setAuditPlanType] = React.useState([]);
    const [editForm, setEditForm] = React.useState(false);
    const [modeValue, setmode] = React.useState("");
    const [showviewdownload, setshowviewdownload] = React.useState(true);
    const [currentUserDept, setcurrentUserDept] = React.useState(null);
    const [selectUserDept, setselectUserDept] = React.useState(null);
    const [selectshift, setselectshift] = React.useState(null);
    const [isDepartmentReady, setIsDepartmentReady] = React.useState(false);
    const [reportCode, setreportCode] = React.useState("");
    const [sequencesLoaded, setSequencesLoaded] = React.useState(false);
    const [selectAuditplan, setselectAuditplan] = React.useState(null);
    const [doccode, setdoccode] = React.useState("");
    const [auditplandate, setauditplandate] = React.useState("");
    const [AllDept, setAllDept] = React.useState([]);
    const [RowErrors, setRowErrors] = React.useState<any[]>([]);
    const [DocumentLink, setDocumentLink] = React.useState(null);
    const [DraftApprovalItem, setDraftApprovalItem] = React.useState(null);
    // const [cancellReason, setcancellReason] = React.useState([{ id: 0, description: "", reason: "" }]);
    // const [RecommendRows, setRecommendRows] = React.useState([]);
    const [showModal, setShowModal] = React.useState(false);
    const [ShowModalpre, setShowModalpre] = React.useState(false);
    //error
    const [recommendationRows, setRecommendationRows] = React.useState([
        { id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }
    ]);
    const skipNCUpdateRef = React.useRef(false);
    const [redirecturl, setredirecturl] = React.useState(null);
    const [NCNumberrows, setNCNumberrows] = React.useState<ncNumber[]>([]);
    // { id: 1, ncnumberNC: "", observationnumberObs: "", reportcode: "", ncsequenceNC: 0, observationsequenceObs: 0, nctype: "", descriptionNC: "" }
    const [NCNumberrowsEdit, setNCNumberrowsEdit] = React.useState<ncNumber[]>([]);
    const [recommendationRowsEdit, setRecommendationRowsEdit] = React.useState([]);
    const [dateerr, setdateerr] = React.useState(false);
    const [isoreferenceerr, setisoreferenceerr] = React.useState(false);
    const [departmenterr, setdepartmenterr] = React.useState(false);
    const [fromdepartmenterr, setfromdepartmenterr] = React.useState(false);
    const [imsprocedureerr, setimsprocedureerr] = React.useState(false);
    const [inquirieserr, setinquirieserr] = React.useState(false);
    const [timeerr, settimeerr] = React.useState(false);
    const [auditorcommentserr, setauditorcommentserr] = React.useState(false);
    const [attachmenterr, setattachmenterr] = React.useState(false);
    const [sharewitherr, setsharewitherr] = React.useState(false);
    const [approvedauditplanerr, setapprovedauditplanerr] = React.useState(false);
    const [TemplateDoc, setTemplateDoc] = React.useState<any>([]);
    const [TemplateDocAudit, setTemplateDocAudit] = React.useState<any>([]);
    const [sharewithusers, setSharewithusers] = React.useState([]);
    const [ValidRemark, setValidRemark] = React.useState(true);
    const [checkboxValues, setCheckboxValues] = React.useState({
        ConformingPositiveFindings: false,
        FailureofIntentNonconformity: false,
        Observations: false,
        FailureofImplementation: false,
        OpportunitiesforImprovement: false,
        FailureofEffectiveness: false,
    });
    //const [checkboxNumberValues, setCheckboxNumberValues] = React.useState<{ [key: string]: number }>({});
    const [checkboxNumberValues, setCheckboxNumberValues] = React.useState({
        ConformingPositiveFindingsN: 0,
        FailureofIntentNonconformityN: 0,
        ObservationsN: 0,
        FailureofImplementationN: 0,
        OpportunitiesforImprovementN: 0,
        FailureofEffectivenessN: 0,
    });
    const maxNCSeqRef = React.useRef(0);
    const maxObsSeqRef = React.useRef(0);
    const hiddenDivRef = React.useRef(null);
    //error end
    const [formData, setFormData] = React.useState({
        RemarksInitiator: "",
        RequesterNameId: 0,
        RequesterName: "",
        RequesterDesignation: "",
        RequestDate: "",
        ObservationSequence: 0,
        NCSequence: 0,
        NCNo: "",
        ObservationNo: "",
        description: "",
        reportCode: "",
        memoNumber: "",
        approvedauditplanId: 0,
        deptId: 0,
        shiftId: 0,
        fromdeptId: 0,
        date: "",
        Auditplandate: "",
        documentCode: "",
        issueNo: "",
        revisionNo: "",
        Status: "",
        submitStatus: "",
        revisionDate: null,
        issueDate: null,
        referenceNo: "",
        documentLink: "",
        attachment: null,
        attachmentIds: null,
        AuditplanDocId: null,
        attachmentJson: null

    });
    const [selectCCUsers, setSelectCCUsers] = React.useState([]);
    const [ListNameId, setListNameId] = React.useState(null);

    const [selectToUsers, setSelectToUsers] = React.useState([]);
    // React.useEffect(() => {
    //     if (!selectedOption?.value) return;

    //     const fetchSequences = async () => {
    //         setIsDepartmentReady(false);

    //       try {
    //         const auditNCItems = await getItemsAuditReportNC(sp, selectedOption.value);
    //         const auditObsItems = await getItemsAuditReportObs(sp, selectedOption.value);

    //         let maxncseq = null;
    //         let maxobsseq = null;

    //         const updatedFormData = { ...formData };

    //         if (auditNCItems.length > 0) {
    //           maxncseq = auditNCItems[0].NCSequence;
    //           updatedFormData.NCSequence = maxncseq;
    //         }

    //         if (auditObsItems.length > 0) {
    //           maxobsseq = auditObsItems[0].ObservationSequence;
    //           updatedFormData.ObservationSequence = maxobsseq;
    //         }

    //         maxNCSeqRef.current = maxncseq;
    //         maxObsSeqRef.current = maxobsseq;

    //         setFormData(updatedFormData);

    //         setNCNumberrows([]); // clear rows on dept change
    //       } catch (error) {
    //         console.error("Error fetching audit sequences", error);
    //       } finally {
    //         setIsDepartmentReady(true);
    //       }
    //     };

    //     fetchSequences();
    //   }, [selectedOption]);
    const handleHiddenDivClick = () => {
        console.log('Hidden div clicked due to dropdown change:', selectUserDept, maxncseq, maxobsseq, maxNCSeqRef.current, maxObsSeqRef.current);
        // Do your action here
    };
    const handleDepartmentChange = async (selectedOption: any) => {
        debugger
        setSequencesLoaded(false);
        setIsDepartmentReady(false); // Prevent effects during update
        skipNCUpdateRef.current = true;
        departmentChanged = true;
        EnableNC = true;
        setEnableNCObs(true);

        const setAuditreportNC = await getItemsAuditReportNC(sp, selectedOption?.value);
        const setAuditreportObs = await getItemsAuditReportObs(sp, selectedOption?.value);
        if (setAuditreportNC.length > 0 || setAuditreportObs.length > 0) {
            maxncseq = setAuditreportNC?.length > 0 && setAuditreportNC[0].NCSequence;
            maxobsseq = setAuditreportObs?.length > 0 && setAuditreportObs[0].ObservationSequence;
            setFormData(prevData => ({
                ...prevData,
                NCSequence: setAuditreportNC?.length > 0 ? setAuditreportNC[0].NCSequence : 0,
                ObservationSequence: setAuditreportObs?.length > 0 ? setAuditreportObs[0].ObservationSequence : 0,
            }));
        }
        // if (setAuditreportObs.length > 0) {
        //     maxobsseq = setAuditreportObs[0].ObservationSequence;
        //     setFormData(prevData => ({
        //         ...prevData,
        //         ObservationSequence: setAuditreportObs[0].ObservationSequence,
        //     }));
        // };
        maxNCSeqRef.current = maxncseq;
        maxObsSeqRef.current = maxobsseq;
        setNCNumberrows([]);
        setSequencesLoaded(true);
        setCheckboxValues(prev => ({
            ...prev,
            FailureofIntentNonconformity: false,
            Observations: false,
            FailureofEffectiveness: false,
            FailureofImplementation: false,
            ConformingPositiveFindings: false,
            OpportunitiesforImprovement: false
        }));
        const updatedValues: Record<string, number> = { ...checkboxNumberValues };

        AuditFindingsOptions.forEach((checkbox) => {
            //const fieldKey = CheckboxFieldMap[checkbox.value];
            const fieldKeyN = CheckboxFieldMapN[checkbox.value];
            if (checkbox.value === 'Observations') {
                updatedValues[fieldKeyN] = 0;
            }

            if (checkbox.value === 'Failure of Intent / Nonconformity') {
                updatedValues[fieldKeyN] = 0;
            }
        });

        setCheckboxNumberValues(prev => ({
            ...prev,
            ...updatedValues
        }));

        setselectUserDept(selectedOption);
        debugger
        let reportcode: string = "";
        if ((formData.date != "") && (selectedOption != null || selectedOption != "")) {
            reportcode = selectedOption.departmentcode + "/" + moment(new Date(formData.date)).format("DD/MM/YYYY");
        }
        setreportCode(reportcode);
        setFormData({ ...formData, deptId: selectedOption?.value, reportCode: reportcode });
        setTimeout(() => {
            if (hiddenDivRef.current) {
                hiddenDivRef.current.click();
            }
            setIsDepartmentReady(true);
        }, 2000); // Let React batch state updates first
    };
    const handleShiftChange = async (selectedOption: any) => {
        setselectshift(selectedOption);
        setFormData({ ...formData, shiftId: selectedOption?.value });
    };
    const handleActualAuditDateChange = (date: any) => {

        debugger
        let reportcode: string = "";
        // if ((formData.date != "" || formData.date != null) && (formData.deptId != 0 || formData.deptId != null)) {
        //     reportcode = selectUserDept?.departmentcode + "/" + moment(new Date(e.target.value)).format("DD/MM/YYYY");
        // }
        if ((date !== "" && date !== null) && (formData.deptId !== 0 && formData.deptId !== null && selectUserDept != null)) {
            reportcode = (selectUserDept.length == 1 ? selectUserDept[0]?.departmentcode : selectUserDept?.departmentcode) + "/" + moment(new Date(date)).format("DD/MM/YYYY");
        }

        setreportCode(reportcode);
        //setFormData({ ...formData, date: new Date(date).toLocaleDateString("en-CA") });
        setFormData({ ...formData, date: new Date(date).toLocaleDateString("en-CA"), reportCode: reportcode })

    };
    const handleFromDepartmentChange = (selectedOption: any) => {
        setcurrentUserDept(selectedOption);

        setFormData({ ...formData, fromdeptId: selectedOption?.value });
    };
    // ////// Recommendation


    const handleAddRecommendationRow = () => {
        setRecommendationRows([...recommendationRows, { id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }]);
    };


    const handleNCNumberChange = (index: number, field: string, value: any) => {
        debugger
        let updatedRows;

        updatedRows = NCNumberrows.map((row, i) =>
            i === index ? { ...row, [field]: value } : row
        );
        setNCNumberrows(updatedRows);
    };

    const removeBlankNCNumberRows = () => {
        const cleanedRows = NCNumberrows.filter(
            row => row.ncnumberNC.trim() !== "" || row.observationnumberObs.trim() !== ""
        );
        setNCNumberrows(cleanedRows);
    };

    // const handleDeleteNCNumberRow = (nctype: string) => {
    //     const rowsOfType = NCNumberrows
    //         .map((row, index) => ({ row, index }))
    //         .filter(({ row }) => row.nctype === nctype);

    //     const last = rowsOfType.pop(); // Get the last row of the given type

    //     if (!last) return; // No matching row found

    //     const updatedRows = [...NCNumberrows];
    //     updatedRows.splice(last.index, 1); // Remove last matching row

    //     // Reassign sequence and number fields for the remaining rows of this type
    //     const reassigned = updatedRows.map((row) => {
    //         if (row.nctype !== nctype) return row; // Don't modify other types

    //         const typeRows = updatedRows.filter(r => r.nctype === nctype);
    //         const newIndex = typeRows.indexOf(row); // Stable ordering

    //         const newSequence = newIndex + 1;
    //         const padded = newSequence.toString().padStart(3, "0");

    //         return {
    //             ...row,
    //             ncsequenceNC: nctype === "NC Number" ? newSequence : row.ncsequenceNC,
    //             ncnumberNC: nctype === "NC Number" ? padded : row.ncnumberNC,
    //             observationsequenceObs: nctype === "Observation Number" ? newSequence : row.observationsequenceObs,
    //             observationnumberObs: nctype === "Observation Number" ? padded : row.observationnumberObs
    //         };
    //     });

    //     setNCNumberrows(reassigned);
    // };
    const handleDeleteNCNumberRow = (nctype: string) => {
        const lastIndex = [...NCNumberrows]
            .map((row, i) => ({ row, index: i }))
            .filter(({ row }) => row.nctype === nctype)
            .map(({ index }) => index)
            .pop(); // Get last matching index

        if (lastIndex === undefined) return; // No matching row found

        const updatedRows = [...NCNumberrows];
        updatedRows.splice(lastIndex, 1);
        setNCNumberrows(updatedRows);
        ncrow = updatedRows;
    };
    const handleaddncrows = (ncobstype: string) => {
        debugger
        ncrow = NCNumberrows;
        console.log("ncncncnnc", NCNumberrows, formData, maxncseq, maxobsseq);
        const type = ncobstype;
        //let NCNumberrows1: any;
        let maxid: number = 0;
        let ncrowsnew: any = ncrow;

        if (ncrowsnew.length > 0) {
            const ids = ncrowsnew.map((row: any) => row.id);
            maxid = Math.max(...ids);
        }
        const deptId = formData.deptId;
        // Find the max sequence number from existing rows of this type
        const existingRows = ncrow.filter((row: any) => row.DepartmentNCId === deptId && row.nctype === type);
        // const departmentRows = ncrowsnew.filter(
        //     (row) => row.DepartmentNCId === deptId && row.nctype === type
        // );
        let maxSequence = 0;
        if (editItemID) {
            if (existingRows.length > 0) {
                if (type === "NC Number") {
                    maxSequence = Math.max(...existingRows.map((r: any) => r.ncsequenceNC || 0));

                } else if (type === "Observation Number") {
                    maxSequence = Math.max(...existingRows.map((r: any) => r.observationsequenceObs || 0));
                }
            } else {
                if (type === "NC Number") {
                    maxSequence = Math.max(maxNCSeqRef.current, ...existingRows.map((r: any) => r.ncsequenceNC || 0));

                } else if (type === "Observation Number") {
                    maxSequence = Math.max(maxObsSeqRef.current, ...existingRows.map((r: any) => r.observationsequenceObs || 0));
                }
            }
        } else {
            if (type === "NC Number") {
                maxSequence = Math.max(maxNCSeqRef.current, ...existingRows.map((r: any) => r.ncsequenceNC || 0));

            } else if (type === "Observation Number") {
                maxSequence = Math.max(maxObsSeqRef.current, ...existingRows.map((r: any) => r.observationsequenceObs || 0));
            }
        }


        const nextSequenceNumber = maxSequence + 1;
        setNCNumberrows(prev => [...prev, {
            Id: 0,
            id: maxid + 1,
            ncnumberNC: type === "NC Number" ? nextSequenceNumber.toString().padStart(3, "0") : "",
            observationnumberObs:
                type === "Observation Number" ? nextSequenceNumber.toString().padStart(3, "0") : "",
            reportcode: "",
            ncsequenceNC: type === "NC Number" ? nextSequenceNumber : 0,
            observationsequenceObs: type === "Observation Number" ? nextSequenceNumber : 0,
            nctype: type,
            DepartmentNCId: formData.deptId,
            descriptionNC: ""
        }]);

    };


    // React.useEffect(() => {
    //     debugger
    //     if (checkboxValues["FailureofIntentNonconformity"] || checkboxValues["FailureofImplementation"] || checkboxValues["FailureofEffectiveness"]) {
    //         const hasNCRow = NCNumberrows.some(row => row.nctype === "NC Number");
    //         if (!hasNCRow) {
    //             handleaddncrows("NC Number");
    //         }
    //     }
    // }, [checkboxValues["FailureofIntentNonconformity"]]);
    React.useEffect(() => {
        //if (departmentChanged && !selectUserDept ) return;
        // if (skipNCUpdateRef.current) {
        //     skipNCUpdateRef.current = false; // Reset after skip
        //     return;
        // }
        //if (!isDepartmentReady) return; // Don't run during/after department switch
        if (
            checkboxValues["FailureofIntentNonconformity"] ||
            checkboxValues["FailureofImplementation"] ||
            checkboxValues["FailureofEffectiveness"]
        ) {
            const hasNCRow = NCNumberrows.some(row => row.nctype === "NC Number");
            if (!hasNCRow && sequencesLoaded) {
                handleaddncrows("NC Number");
            }
        }
    }, [
        checkboxValues["FailureofIntentNonconformity"],
        checkboxValues["FailureofImplementation"],
        checkboxValues["FailureofEffectiveness"],
        // NCNumberrows,
        // handleaddncrows
    ]);

    React.useEffect(() => {
        // if (skipNCUpdateRef.current) {
        //     skipNCUpdateRef.current = false; // Reset after skip
        //     return;
        // }
        //if (!isDepartmentReady) return; // Don't run during/after department switch
        //if (departmentChanged && !selectUserDept ) return;
        if (checkboxValues["Observations"]) {
            const hasObsRow = NCNumberrows.some(row => row.nctype === "Observation Number");
            if (!hasObsRow && sequencesLoaded) {
                handleaddncrows("Observation Number");
            }
        }
    }, [checkboxValues["Observations"]]);

    // React.useEffect(() => {
    //     if (!checkboxValues["FailureofIntentNonconformity"] || !checkboxValues["Observations"]) {
    //         removeBlankNCNumberRows();
    //     }
    // }, [checkboxValues]);

    const handleRecommendationChange = (index: number, field: string, value: any) => {
        debugger
        let updatedRows;
        if (field == "sharewith") {
            // const valuesOnly = value.map((option: any) => option.value);
            updatedRows = recommendationRows.map((row, i) =>
                i === index ? { ...row, [field]: value, sharewithIds: value?.value } : row
            );
        } else {
            updatedRows = recommendationRows.map((row, i) =>
                i === index ? { ...row, [field]: value } : row
            );
        }

        setRecommendationRows(updatedRows);
    };

    const cancelModalAction = (refresh?: boolean,) => {
        debugger
        setredirecturl(window.location.href);
        //setShowfileNew(false);
        setShowModalTemplateDoc(false);
        Showfile = false;
    }
    const handleNCDescriptionChange = (lvl: number, field: string, value: any) => {
        debugger
        console.log("NCNumberrows", NCNumberrows);
        setNCNumberrows((prev) =>
            prev.map((row) =>
                row.id === lvl ? { ...row, descriptionNC: value || "" } : row
            )
        );
        // const updatedRows = NCNumberrows.map((row, i) =>
        //     i === index ? { ...row, [field]: value } : row
        // );

        // setNCNumberrows(updatedRows);
    };

    const handleDeleteRecommendationRow = (index: number) => {
        const updatedRows = recommendationRows.filter((_, i) => i !== index);
        setRecommendationRows(updatedRows);
    };
    // //////


    // Handle change event
    const handleChange = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        const updatedArr = forwardToArr.map(row =>
            row.level === lvl ? { ...row, approvalType: event.target.value } : row
        );
        //   setApprovalType(event.target.value);
        setForwardToArr(updatedArr);
    };


    const [forwardToArr, setForwardToArr] = React.useState<ForwardTo[]>([
        {
            id: 0, role: 0, level: 1, approvers: [], approvalType: "One", Responsibility: "Signer", IsSignatureRequired: true,
            responsibilityerror: false,
        } // Default row
    ]);

    const Breadcrumb = [
        {
            MainComponent: "My Request",
            MainComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx`,
        },
        {
            ChildComponent: "IMS Audit Report and Checklist",
            ChildComponentURl: `${SITE_URL}/SitePages/EDCMAIN.aspx#/AnnualAuditReport`,
        },
    ];
    const [forwardToArrEdit, setForwardToArrEdit] = React.useState<ForwardTo[]>([]);
    const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);

    // const [remark, setRemark] = React.useState("");

    // Function to handle People Picker selection
    const onPeoplePickerChange = (items: any[]) => {
        setSelectedUsers(items);
    };

    const getNewFileName = async (originalFileName: string): Promise<string> => {
        const userId = currentUser.Id; // Or however you get the current user ID
        const date = new Date();
        const fileExtension = originalFileName.split('.').pop();
        const fileNameWithoutExt = originalFileName.substring(0, originalFileName.lastIndexOf('.'));

        // Format date components
        const components = [
            date.getFullYear(),
            (date.getMonth() + 1).toString().padStart(2, '0'),
            date.getDate().toString().padStart(2, '0'),
            date.getHours().toString().padStart(2, '0'),
            date.getMinutes().toString().padStart(2, '0'),
            date.getSeconds().toString().padStart(2, '0')
        ];

        const iddd = formData.reportCode.replace(/\//g, '_');

        return `${iddd}_${fileNameWithoutExt}_${components.join('')}.${fileExtension}`;

    };
    const ApiCallFunc = async () => {
        setLoading(true);
        setFormLoading(true);
        setTimeout(() => {
            setLoading(false);
            setFormLoading(false);
        }, 5000); // 5000ms = 5 seconds
        setAuditPlanType(await getAllAuditType(sp));
        var DepartmentArr = await getAllDepartment(sp);
        DepartmentArr.sort((a, b) => a.Department.localeCompare(b.Department));
        const optionsDepartment = DepartmentArr.map((item: any) => ({
            value: item.ID,
            label: item.Department,
            itemId: item.ID,
            department: item.Department,
            departmentcode: item.DepartmentCode
        }));
        setAllDept(optionsDepartment);
        var setAllDept1 = await getAllDepartment1(sp);

        var DocCodeArr = await getAllApprovedAuditplan(sp);
        console.log("DocCodeArrDocCodeArr", DocCodeArr);
        const options = DocCodeArr.map((item: any) => ({
            value: item.ID,
            label: item.MemoNumber,
            ApprovedAuditPlanId: item.ID,
            //IssueNumber: item.IssueNumber,
            ReferenceNumber: item.ReferenceNumber,
            //RevisionNumber: item.RevisionNumber,
            Date: item.Date,
            SubmiitedDate: item.SubmiitedDate,
            SubmitStatus: item.SubmitStatus,
            MemoNumber: item.MemoNumber,
            DepartmentId: item.DepartmentId,
            AttachmentId: item.AttachmentId,
            AttachmentJson: item.AttachmentJson,
            ID: item.ID,
            Status: item.Status,

        }));
        setRows(options);
        let shiftoptions = await getAuditProgShift(sp);
        setAuditProgShift(shiftoptions);
        const path1 = window.location.href;

        if (path1.includes("/view/") || path1.includes("/approve/")) {
            setFormLoading(true); ////
            setInputDisabled(true);
            setshowForwardapproval(false)
        }
        else {
            setInputDisabled(false);
            setshowForwardapproval(true)
        }
        if (path1.includes("/edit/")) {
            setFormLoading(true); ////
            setshowForwardapproval(true)
        }
        const path = window.location.href;
        const segments = path.split('/').filter(Boolean); // Remove empty elements

        // Check if "edit" or "view" exists in the URL
        const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");
        const rowData: any[] = await getItemfromChecklistMaster(sp) //baseUrl
        if (rowData.length > 0) {
            const initialRows = rowData.map((item: any) => ({
                id: item.Id,
                isoreference: item.ISOReference,
                imsprocedure: item.IMSProcedure,
                inquiries: item.Inquiries,
                auditorcomments: item.AuditorsComments,
                time: item.Time,
                sharewithIds: item.SharewithId,
                //endTime: "",
                sharewith: item.Sharewith ? { label: item.Sharewith.Title, value: item.Sharewith.ID } : null // Convert single object
            }));
            setRecommendationRows(initialRows);
            setRecommendationRowsEdit(initialRows);
        } else {
            setRecommendationRowsEdit([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
            setRecommendationRows([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
        }

        const Currusers: any = await getCurrentUser(sp, siteUrl);
        setCurrentUser(await getCurrentUser(sp, siteUrl));
        const userProfile = await sp.profiles.myProperties();

        const UserDept = userProfile.UserProfileProperties ? userProfile.UserProfileProperties[userProfile.UserProfileProperties.findIndex((obj: any) => obj.Key === "Department")].Value : "";
        //setselectUserDept(setAllDept1.filter(user => user.label === UserDept));
        let currentuserdepartment = UserDept == "IT" ? "Information Technology" : UserDept;
        setcurrentUserDept(setAllDept1.filter(user => user.label === currentuserdepartment))
        setFormData(prevData => ({
            ...prevData,
            RequesterNameId: Currusers?.Id || "",
            RequesterDesignation: userProfile?.Title || "",
            RequesterName: userProfile?.DisplayName || "",
            RequestDate: new Date().toLocaleDateString("en-CA"),
            // Department: UserDept
            //DepartmentId: optionsfilterdepart && optionsfilterdepart[0]?.value
            //RequestedDate: new Date().toISOString().split("T")[0] // Format as YYYY-MM-DD

        }));
        console.log("userrrrdeptt", UserDept);
        setFormData({ ...formData, fromdeptId: setAllDept1.filter(user => user.label === currentuserdepartment)[0]?.value });
        const AllUserRoles = await getDataRoles(sp);
        const setRolesValue = AllUserRoles.map((item: any) => ({
            value: item.Id,
            label: item.Role,

        }));

        setUserRoles(setRolesValue)

        const users = await sp.web.siteUsers();
        const people = users.filter(user => user.PrincipalType === PrincipalType.User);

        const Selectedoptions = people.map(item => ({
            value: item.Id,
            label: item.Title,
            UserName: item.Title,
            UserEmail: item.Email
        }));

        setRows1(Selectedoptions);

        setFormData(prevData => ({
            ...prevData,
            from: Currusers?.Id || 0,
            fromEmail: Currusers?.Email,

        }));
        let ChangeRequestTemplateType = await getLatestChangeRequestTemplateType(sp, CONTENTTYPE_AuditReportTemp);
        debugger
        if (ChangeRequestTemplateType.length > 0) {
            const template = ChangeRequestTemplateType[0];
            console.log("template", template);
            setFormData(prevData => ({
                ...prevData,
                documentCode: template.DocumentCode || "",
                revisionNo: template.RevisionNumber,
                issueNo: template.IssueNumber,
                referenceNo: template.ReferenceNumber || "",
                revisionDate: template.RevisionDate == null ? null : new Date(template.RevisionDate).toLocaleDateString("en-CA"),
                issueDate: template.IssueDate == null ? null : new Date(template.IssueDate).toLocaleDateString("en-CA"),
            }));
        }

        // const setAuditreportNC = await getItemsAuditReportNC(sp);
        // const setAuditreportObs = await getItemsAuditReportObs(sp);

        // if (setAuditreportNC.length > 0) {
        //     setFormData(prevData => ({
        //         ...prevData,
        //         NCSequence: setAuditreportNC[0].NCSequence
        //     }));
        // }
        // if (setAuditreportObs.length > 0) {
        //     setFormData(prevData => ({
        //         ...prevData,
        //         ObservationSequence: setAuditreportObs[0].ObservationSequence,
        //     }));
        // }
        let formitemid;
        //#region getdataByID
        if (sessionStorage.getItem("AuditReportId") != undefined) {
            const iD = sessionStorage.getItem("AuditReportId")
            let iDs = decryptId(iD)
            formitemid = Number(iDs);
            setFormItemId(Number(iDs))
        }
        else {


            const path = window.location.href;
            const segments = path.split('/').filter(Boolean); // Remove empty elements

            // Check if "edit" or "view" exists in the URL
            const paramIndex = segments.findIndex(seg => seg === "edit" || seg === "view" || seg === "approve");


            if (paramIndex !== -1 && segments[paramIndex + 1]) {
                setmode(segments[paramIndex])
                // mode = segments[paramIndex]; // Will be "edit" or "view"
                formitemid = segments[paramIndex + 1]; // Get the ID
                if (segments[paramIndex + 2] !== undefined) {
                    // var ProcessListItem={
                    //     Status:"",
                    //     Level:0,
                    //     CurrentUserRole:"",

                    // }

                    //  ProcessListItem =await getApprovalByID(sp, Number(segments[paramIndex + 2]),CONTENTTYPE_DocumentCancel);
                    // setInputDisabled((ProcessListItem.Status == "Pending" || ProcessListItem?.Status === "Save as draft") && ProcessListItem.Level === 0 && ProcessListItem.CurrentUserRole !=="OES")
                    setEditID(await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditReport, CONTENTTYPE_AuditReportNew));
                    // var ProcessItemId: any = await getApprovalByID(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditReport);
                    setInputDisabled(await getApprovalByID2(sp, Number(segments[paramIndex + 2]), CONTENTTYPE_AuditReport, CONTENTTYPE_AuditReportNew));
                }
                // else {

                //     setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditReport))

                // }
            }

            setDraftApprovalItem(await getDraftApprovalByID(sp, Number(formitemid), CONTENTTYPE_AuditReport, CONTENTTYPE_AuditReportNew))

        }
        // if (setAuditreportNC.length > 0) {
        //     setFormData(prevData => ({
        //         ...prevData,
        //         NCSequence: setAuditreportNC[0].NCSequence

        //     }));
        // }
        // if (setAuditreportObs.length > 0) {
        //     setFormData(prevData => ({
        //         ...prevData,
        //         ObservationSequence: setAuditreportObs[0].ObservationSequence,
        //     }));
        // }
        // formitemid =20;
        if (formitemid) {

            setEditItemID(Number(formitemid));

            const setBannerById = await getItemByID(sp, Number(formitemid))

            if (setBannerById.length > 0) {
                debugger
                setEditForm(true);
                setMainEditItem(setBannerById[0]);

                console.log("setBannerById[0]..", setBannerById);
                // const valuesOnly = selectedOptions.map((option: any) => option.value);
                // setFormData({ ...formData, [fieldName]: valuesOnly });

                setFormData(prevData => ({
                    ...prevData,
                    // memoNo: "",
                    reportCode: setBannerById[0].ReportCode,
                    NCNo: setBannerById[0].NCNumber,
                    ObservationNo: setBannerById[0].ObservationNumber,
                    RemarksInitiator: setBannerById[0].RemarksInitiator,
                    //ObservationSequence: setBannerById[0].ObservationSequence,
                    //NCSequence: setBannerById[0].NCSequence,
                    memoNumber: setBannerById[0].MemoNumber,
                    description: setBannerById[0].Description,
                    deptId: setBannerById[0].DepartmentAuditedId,
                    shiftId: setBannerById[0].ShiftId, // Default to 0 if ShiftId is null
                    fromdeptId: setBannerById[0].DepartmentId,
                    issueNo: setBannerById[0].IssueNumber,
                    revisionNo: setBannerById[0].RevisionNumber,
                    revisionDate: setBannerById[0].RevisionDate == null ? null : new Date(setBannerById[0].RevisionDate).toLocaleDateString("en-CA") || null,
                    issueDate: setBannerById[0].IssueDate == null ? null : new Date(setBannerById[0].IssueDate).toLocaleDateString("en-CA") || null,
                    date: new Date(setBannerById[0].Date).toLocaleDateString("en-CA"),
                    Auditplandate: setBannerById[0].AuditPlanDate && new Date(setBannerById[0].AuditPlanDate).toLocaleDateString("en-CA"),
                    approvedauditplanId: setBannerById[0].ApprovedAuditPlanId,
                    documentcode: setBannerById[0].DocumentCode,
                    //Description: setBannerById[0].description,
                    referenceNo: setBannerById[0].ReferenceNumber,
                    AnnualAuditPlanDocumentLinkId: setBannerById[0].AnnualAuditPlanDocumentLinkId,
                    AuditplanDocId: setBannerById[0].AuditplanDocId,
                    SubmiitedDate: setBannerById[0].SubmiitedDate,
                    submitStatus: setBannerById[0].SubmitStatus,
                    Status: setBannerById[0].Status,
                    documentname: setBannerById[0].DocumentName,
                    isrework: setBannerById[0].IsRework,
                    attachmentIds: setBannerById[0].AttachmentId || null,
                    attachmentJson: setBannerById[0].AttachmentJson || null
                }));
                setauditplandate(setBannerById[0].AuditPlanDate && new Date(setBannerById[0].AuditPlanDate).toLocaleDateString("en-CA"))
                setFormData(prevData => ({
                    ...prevData,
                    reportCode: setBannerById[0].ReportCode,
                    NCNo: setBannerById[0].NCNumber,
                    ObservationNo: setBannerById[0].ObservationNumber,
                    //ObservationSequence: setBannerById[0].ObservationSequence,
                    //NCSequence: setBannerById[0].NCSequence,
                    description: setBannerById[0].Description
                }));
                const newValues: any = {};
                for (const [label, field] of Object.entries(CheckboxFieldMap)) {
                    newValues[field] = setBannerById[0][field] === "Yes";
                }
                setCheckboxValues(newValues);
                const newValuesN: any = {};

                for (const [label, field] of Object.entries(CheckboxFieldMapN)) {
                    newValuesN[field] = Number(setBannerById[0][field]) || 0; // fallback to 0 if null/undefined
                }

                setCheckboxNumberValues(newValuesN);
                let sharewithuser = setBannerById[0].Sharewith?.map((approver: any) => ({
                    value: approver.ID,
                    label: approver.Title,
                    UserName: approver.Title,
                    //UserEmail: approver.EMail
                })) || [];
                setSharewithusers(sharewithuser);
                setdoccode(setBannerById[0].Title);
                debugger
                setselectUserDept(setAllDept1.filter(user => user.value === setBannerById[0].DepartmentAuditedId));
                if (setBannerById[0].DepartmentAuditedId > 0) {
                    let departmentselected = setAllDept1.filter(user => user.value === setBannerById[0].DepartmentAuditedId)
                    // handleDepartmentChange(departmentselected);
                    setSequencesLoaded(false);
                    setIsDepartmentReady(false); // Prevent effects during update
                    skipNCUpdateRef.current = true;
                    departmentChanged = true;
                    EnableNC = true;
                    setEnableNCObs(true);

                    const setAuditreportNC = await getItemsAuditReportNC(sp, setBannerById[0].DepartmentAuditedId);
                    const setAuditreportObs = await getItemsAuditReportObs(sp, setBannerById[0].DepartmentAuditedId);
                    if (setAuditreportNC.length > 0 || setAuditreportObs.length > 0) {
                        maxncseq = setAuditreportNC?.length > 0 && setAuditreportNC[0].NCSequence;
                        maxobsseq = setAuditreportObs?.length > 0 && setAuditreportObs[0].ObservationSequence;
                        setFormData(prevData => ({
                            ...prevData,
                            NCSequence: setAuditreportNC?.length > 0 ? setAuditreportNC[0].NCSequence : 0,
                            ObservationSequence: setAuditreportObs?.length > 0 ? setAuditreportObs[0].ObservationSequence : 0,
                        }));
                    };
                    maxNCSeqRef.current = maxncseq;
                    maxObsSeqRef.current = maxobsseq;
                    setNCNumberrows([]);
                    setSequencesLoaded(true);
                    setTimeout(() => {
                        if (hiddenDivRef.current) {
                            hiddenDivRef.current.click();
                        }
                        setIsDepartmentReady(true);
                    }, 2000);
                }
                setcurrentUserDept(setAllDept1.filter(user => user.value === setBannerById[0].DepartmentId));
                const selectedauditplan = options.filter((cust: { value: any; }) => cust.value === setBannerById[0].ApprovedAuditPlanId) || null;
                setselectAuditplan(selectedauditplan);
                setSelectedOption(selectedauditplan);
                setselectshift(shiftoptions.filter((x: { value: any; }) => x.value == setBannerById[0].ShiftId) || null);
                console.log(" setBannerById[0].AttachmentId if", setBannerById[0], selectedauditplan)
                if (setBannerById[0].AttachmentId) {
                    // setDocumentLink(await getDocumentLinkByID(sp, setBannerById[0].AttachmentId));
                    let arrn = await getDocumentLinkByID(sp, setBannerById[0].AttachmentId);
                    setFilesArr([...FilesArr, ...arrn]);
                    setFilesArr1([...FilesArr1, ...arrn]);

                }
                setTemplateDoc(await getGeneratedTemplateDocCR(sp, Number(formitemid)));

                if (setBannerById[0].AnnualAuditPlanDocumentLinkId.length > 0) {
                    let arrn = await getDocumentLinkByIDPlan(sp, setBannerById[0].AnnualAuditPlanDocumentLinkId);
                    console.log("arrrrrrn5ghgh6", arrn);
                    setFilesArrDoclink([...FilesArrDoclink, ...arrn]);
                    console.log("arrrrrrn5ghjghj6 after", arrn);
                }
                if (setBannerById[0].AuditplanDocId != null) {
                    // let arrn = await getAuditDocumentLinkByIDPlan(sp, setBannerById[0].AuditplanDocId);
                    // console.log("arrrrrrn audit doc", arrn);
                    setTemplateDocAudit(await getGeneratedTemplateDocAuditplan(sp, Number(setBannerById[0].AuditplanDocId)));
                    //setFilesArrDoclink([...FilesArrDoclink, ...arrn]);
                    console.log("arrrrrrn5ghjghj6 audit doc");
                }
                const ApprowData: any[] = await getAllProcessData(sp, Number(formitemid), CONTENTTYPE_AuditReport, setBannerById[0].ReferenceNumber, CONTENTTYPE_AuditReportNew)

                if (ApprowData.length > 0) {

                    const EditApprowData = ApprowData.map((item: any) => ({
                        id: item.ID,
                        role: item.ApproverRole?.Id || 0, // Assuming role comes from ApproverRole
                        level: item.Level || 1, // Default to 1 if missing
                        approvalType: item.LevelType,
                        Responsibility: item.Responsibility || "",
                        IsSignatureRequired: item.IsSignatureRequired == "Yes" ? true : false,
                        approvers: item.Approvers?.map((approver: any) => ({
                            value: approver.Id,
                            label: approver.Title,
                        })) || []

                    }));
                    setForwardToArr(EditApprowData);
                    setForwardToArrEdit(EditApprowData);

                }

                const rowDataNC: any[] = await getNCNumberbyID(sp, Number(formitemid)) //baseUrl
                if (rowDataNC.length > 0) {
                    const initialRows = rowDataNC.map((item: any) => ({
                        Id: item.Id,
                        id: Number(item.Title),
                        ncnumberNC: item.NCType == "NC Number" ? item.NCNumber : "",
                        observationnumberObs: item.NCType == "NC Number" ? "" : item.NCNumber,
                        reportcode: item.ReportCode,
                        ncsequenceNC: item.NCType == "NC Number" ? item.NCSequence : 0,
                        observationsequenceObs: item.NCType == "NC Number" ? 0 : item.NCSequence,
                        nctype: item.NCType,
                        DepartmentNCId: item.DepartmentId,
                        descriptionNC: item.Description
                    }));
                    setNCNumberrows(initialRows);
                    setNCNumberrowsEdit(initialRows);
                } else {
                    setNCNumberrowsEdit([])
                    setNCNumberrows([])
                }
                const rowData1: any[] = await getItemByID2(sp, Number(formitemid));
                const rowData: any[] = await getItemfromChecklistMaster(sp) //baseUrl
                if ((segments[paramIndex] == "edit" || segments[paramIndex] == "approve" || segments[paramIndex] == "view") && rowData1.length > 0) {
                    const initialRows1 = rowData1.map((item: any) => ({
                        id: item.Id,
                        isoreference: item.ISOReference,
                        imsprocedure: item.IMSProcedure,
                        inquiries: item.Inquiries,
                        auditorcomments: item.AuditorsComments,
                        time: item.Time,
                        sharewithIds: item.SharewithId,
                        //endTime: "",
                        sharewith: item.Sharewith ? { label: item.Sharewith.Title, value: item.Sharewith.ID } : null // Convert single object
                    }));
                    setRecommendationRows(initialRows1);
                    setRecommendationRowsEdit(initialRows1);
                } else {
                    if (rowData.length > 0 && rowData1.length == 0) {
                        const initialRows = rowData.map((item: any) => ({
                            id: item.Id,
                            isoreference: item.ISOReference,
                            imsprocedure: item.IMSProcedure,
                            inquiries: item.Inquiries,
                            auditorcomments: item.AuditorsComments,
                            time: item.Time,
                            sharewithIds: item.SharewithId,
                            endTime: "",
                            sharewith: item.Sharewith ? { label: item.Sharewith.Title, value: item.Sharewith.ID } : null // Convert single object
                        }));
                        setRecommendationRows(initialRows);
                        setRecommendationRowsEdit(initialRows);
                    } else {
                        setRecommendationRowsEdit([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
                        setRecommendationRows([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
                    }
                }



            }





        }
        setFormLoading(false);

        // setRequesterRoleId(await getRequesterID(sp))
        setFormNameId(await getFormNameID(sp, CONTENTTYPE_AuditReport))
        setListNameId(await getListNameID(sp, LIST_TITLE_AuditReport))

        //}
        //#endregion


    };

    const onSelectsharewith = (selectedOptions: any) => {
        // const newSelections = selectedOptions || [];
        // const allOptions = [...sharewithusers, ...newSelections];

        // const uniqueOptions = allOptions.filter(
        //     (option, index, self) =>
        //         index === self.findIndex((o) => o.value === option.value)
        // );

        // setSharewithusers(uniqueOptions);
        const uniqueOptions = (selectedOptions || []).filter(
            (option: any, index: any, self: any) =>
                index === self.findIndex((o: any) => o.value === option.value)
        );
        setSharewithusers(uniqueOptions);

    };

    const onSelectApprovers = (selectedOptions: any, lvl: number) => {
        setForwardToArr((prev) =>
            prev.map((row) =>
                row.level === lvl ? { ...row, approvers: selectedOptions || [] } : row
            )
        );
    };



    const onSelectRole = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        const updatedArr = forwardToArr.map(row =>
            row.level === lvl ? { ...row, role: Number(event.target.value) } : row
        );
        setForwardToArr(updatedArr);
        setSelectedRole(forwardToArr.map(r => r.role).filter(role => role))
    };


    const handleAddRow = () => {
        setForwardToArr((prev) => [
            ...prev,
            { id: 0, role: 0, level: prev.length + 1, approvers: [], approvalType: "One", Responsibility: "Signer", IsSignatureRequired: true }
        ]);
    };
    const handleChangeResp = (event: React.ChangeEvent<HTMLSelectElement>, lvl: number) => {
        event.preventDefault();
        const updatedArr = forwardToArr.map(row =>
            // row.level === lvl ? { ...row, Responsibility: event.target.value} : row

            row.level === lvl ? { ...row, Responsibility: event.target.value, IsSignatureRequired: event.target.value === "Signer" ? true : false} : row
        );
        //   setApprovalType(event.target.value);
        setForwardToArr(updatedArr);
    };
    const handleDeleteRow = (index: number) => {
        const updatedRows = forwardToArr
            .filter((_, i) => i !== index) // Remove selected row
            .map((row, newIndex) => ({ ...row, level: newIndex + 1 })); // Reassign levels

        setForwardToArr([...updatedRows]); // Ensure a new array reference
    };



    React.useEffect(() => {

        ApiCallFunc();

    }, [useHide]);

    const handleCancel = () => {
        // window.location.reload();
        window.history.back();
        // window.location.reload();
        setTimeout(() => {
            location.reload();
        }, 100);

    }

    const OpenFile = (obj: any, sts: string) => {
        debugger
        const fileUrl = `${Tenant_URL}${obj?.FileRef != "" ? obj.FileRef : obj.fileUrl}`;
        if (sts == "Open") {
            Showfile = true;
        }
        if (sts == "Open") {
            if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
                const viewerUrlppt = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview`;
                setredirecturl(viewerUrlppt);
            } else {
                setredirecturl(fileUrl);
            }

        } else if (sts == "Download") {
            const link = document.createElement("a");
            link.href = fileUrl;
            link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        }
    }
    const onChange = (name: string, value: string) => {
        debugger
        setFormData((prevData) => ({
            ...prevData,
            [name]: value,
            RemarksInitiator: value
        }));
    };
    const OpenFileTemplate = (obj: any, sts: string) => {
        debugger
        setShowModalTemplateDoc(true);
        if (sts == "Open") {
            Showfile = true;
        }
        console.log("ttrtrtrtt", obj)
        const fileUrl = `${Tenant_URL}${obj?.FileRef != "" ? obj.FileRef : obj.fileUrl}`;
        if (sts == "Open") {
            if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
                const viewerUrl = `${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=embedview`;

                //window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj?.FileRef != "" ? obj.FileRef : obj.fileUrl)}&action=view`);
                setredirecturl(viewerUrl);
            } else {
                setredirecturl(fileUrl);
                //window.open(fileUrl, "_blank"); // Open PDF and other files normally
            }

        } else if (sts == "Download") {
            const link = document.createElement("a");
            link.href = fileUrl;
            link.setAttribute("download", obj?.FileLeafRef != "" ? cleanFileName(obj.FileLeafRef) : cleanFileName(obj.name)); // Suggests a filename for download
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        }

    }
    // const OpenFile = (obj: any, sts: string) => {
    //     debugger
    //     console.log("obbbj", obj)
    //     const fileUrl = `${Tenant_URL}${obj.FileRef}`;
    //     if (sts == "Open") {
    //         Showfile = true;
    //     }
    //     if (sts == "Open") {
    //         if (/\.(doc|docx|xls|xlsx|ppt|pptx|csv|docs)$/i.test(fileUrl)) {
    //             window.open(`${SITE_URL}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(obj.FileRef)}&action=default`, "_blank");
    //         } else {
    //             window.open(fileUrl, "_blank"); // Open PDF and other files normally
    //         }

    //     } else if (sts == "Download") {
    //         const link = document.createElement("a");
    //         link.href = fileUrl;
    //         link.setAttribute("download", obj.FileLeafRef); // Suggests a filename for download
    //         document.body.appendChild(link);
    //         link.click();
    //         document.body.removeChild(link);

    //     }



    // }

    const cleanFileName = (filename: string) => {
        // Match a 14-digit datetime suffix before the file extension
        const datetimePattern = /_\d{14}(?=\.[^.]+$)/;

        if (datetimePattern.test(filename)) {
            return filename.replace(datetimePattern, '');
        }

        return filename;
    }

    const cleanFileNameSave = async (filename: string) => {
        // Match a 14-digit datetime suffix before the file extension
        const datetimePattern = /_\d{14}(?=\.[^.]+$)/;

        if (datetimePattern.test(filename)) {
            return filename.replace(datetimePattern, '');
        }

        return filename;
    }

    const handleKeyDown = (e: React.KeyboardEvent<HTMLSelectElement>) => {
        if (e.key === 'Enter') {
            e.preventDefault(); // 🛑 Prevents page reload
        }
    };
    // const handleKeyDowntext = (e: React.KeyboardEvent<HTMLInputElement>) => {
    //     if (e.key === 'Enter') {
    //         e.preventDefault();
    //     }
    // };
    // const handleKeyDowntextarea = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    //     if (e.key === 'Enter' && !e.shiftKey) {
    //         e.preventDefault(); // Prevent form submission or unwanted behavior
    //         // Optional: do something when Enter is pressed (like submit or blur)
    //     }
    // };
    const onSelectDocCode = async (selectedList: any) => {
        debugger
        if (selectedList != null) {
            setTemplateDocAudit(await getGeneratedTemplateDocAuditplan(sp, Number(selectedList?.value)));
        }

        console.log(selectedList, "selectedList");
        if (selectedList != null) {
            setLoading(true);
            setFormData(prevData => ({
                ...prevData,
                //issueNo: selectedList.IssueNumber != "" || selectedList.IssueNumber != null ? selectedList.IssueNumber : 0,
                //referenceNo: selectedList.ReferenceNumber != "" || selectedList.ReferenceNumber != null ? selectedList.ReferenceNumber : 0,
                //revisionNo: selectedList.RevisionNumber != "" || selectedList.RevisionNumber != null ? selectedList.RevisionNumber : 0,
                deptId: selectedList?.DepartmentId,
                AuditplanDocId: Number(selectedList?.value),
                attachmentIds: selectedList?.AttachmentId,
                attachmentJson: selectedList?.AttachmentJson,
                ApprovedAuditPlanId: selectedList?.ID,
                Auditplandate: new Date(selectedList?.Date).toLocaleDateString("en-CA"),
                //documentCode: selectedList.MemoNumber
                // Format as YYYY-MM-DD
            }));
            setSelectedOption(selectedList);
            setFormData(prevData => ({
                ...prevData,
                deptId: selectedList?.DepartmentId
            }));
            setdoccode(selectedList?.MemoNumber);
            setauditplandate(new Date(selectedList?.Date).toLocaleDateString("en-CA"))
            //const rowData: any[] = await getItemByID2(sp, Number(selectedList.ID)) //baseUrl
            // const rowData: any[] = await getItemfromChecklistMaster(sp) //baseUrl
            // if (rowData.length > 0) {
            //     const initialRows = rowData.map((item: any) => ({
            //         id: item.Id,
            //         isoreference: item.ISOReference,
            //         imsprocedure: item.IMSProcedure,
            //         inquiries: item.Inquiries,
            //         auditorcomments: item.auditorcomments,
            //         time: item.Time,
            //         sharewithIds: item.sharewithId,
            //         endTime: "",
            //         sharewith: item.sharewith ? { label: item.sharewith.Title, value: item.sharewith.ID } : null // Convert single object
            //     }));
            //     setRecommendationRows(initialRows);
            //     setRecommendationRowsEdit(initialRows);
            // } else {
            //     setRecommendationRowsEdit([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
            //     setRecommendationRows([{ id: 0, isoreference: "", imsprocedure: "", inquiries: "", auditorcomments: "", time: "", sharewith: null, sharewithIds: null }])
            // }


            const selecteddepart = AllDept.filter((cust: { value: any; }) => cust.value === selectedList?.DepartmentId)[0] || null;
            //let reportcodeaudit = selecteddepart.departmentcode + "/" + moment(new Date()).format("DD/MM/YYYY");
            //setreportCode(reportcodeaudit);
            setFormData({ ...formData, deptId: selectedList?.DepartmentId });
            //setselectUserDept(selecteddepart);
            console.log("alllldept", AllDept, selecteddepart);
            const selectedauditplan = rows.filter((cust: { value: any; }) => cust.value === selectedList?.ID)[0] || null;
            setselectAuditplan(selectedauditplan);
            console.log("selectedList.AttachmentId[0]", selecteddepart, selectedList?.AttachmentId, selectedauditplan)
            if (selectedList.AttachmentId.length > 0) {
                let arrn = await getDocumentLinkByIDPlan(sp, selectedList?.AttachmentId);
                //let arraynew: any[];
                //arraynew.push(arrn)
                console.log("arrrrrrn56", arrn);
                setFilesArrDoclink([...FilesArrDoclink, ...arrn]);
                console.log("arrrrrrn56 after", arrn);
                //setAttachmentarr(arrn);
                // setDocumentLink(await getDocumentLinkByIDPlan(sp, selectedList.AttachmentId[0]))
            }
            else {
                setDocumentLink(null);
            }  // Set the selected users
            setLoading(false);
        } else {
            // 🔄 User cleared the select dropdown — reset fields
            setSelectedOption(null);
            setFormData(prev => ({
                ...prev,
                deptId: null,
                AuditplanDocId: null,
                attachmentIds: null,
                attachmentJson: null,
                ApprovedAuditPlanId: null,
                Auditplandate: ""
            }));
            setdoccode("");
            setauditplandate("");
            setFilesArrDoclink([]);
            setDocumentLink(null);
            setselectAuditplan(null);
            setTemplateDocAudit(null); // Optional, if needed
        }
    };

    const validateForm = async (fmode: FormSubmissionMode) => {
        const {
            approvedauditplanId,
            deptId,
            fromdeptId,
            date,
            documentCode,
            issueNo,
            revisionNo,
            referenceNo,
            documentLink,
            attachment,
            attachmentIds,
            attachmentJson } = formData;
        const validationErrors: { [level: number]: { role: boolean; approvers: boolean; approvalType: boolean, responsibility: boolean } } = {};
        // const { description } = richTextValues;
        console.log("formdataaaaa", formData);
        let valid = true;
        let validraft = true;
        let valid1 = true;
        let validRec = true;
        let validAudit = true;
        // let validateOverview:boolean = false;
        // let validatetitlelength = false;
        // let validateTitle = false;
        // setValidDraft(true);
        setdepartmenterr(false);
        setfromdepartmenterr(false);
        setdateerr(false);
        setattachmenterr(false);
        setisoreferenceerr(false);
        setimsprocedureerr(false);
        setinquirieserr(false);
        settimeerr(false);
        setauditorcommentserr(false);
        setsharewitherr(false);
        setapprovedauditplanerr(false);
        setValidSubmit(true);
        setValidCancelReason(true);
        setValidForwardTo(true);
        setValidAudit(true);
        setValidDraft(true);
        let errormsg = "";

        if (fmode == FormSubmissionMode.SUBMIT) {
            if (!selectedOption) {
                setapprovedauditplanerr(true);
                valid = false;
            }
            if (!deptId && !selectUserDept) {
                setdepartmenterr(true);
                //Swal.fire('Error', 'Title is required!', 'error');
                valid = false;
            }
            if ((fromdeptId == 0 || fromdeptId == undefined) && (!currentUserDept || currentUserDept.length == 0)) {
                setfromdepartmenterr(true);
                //Swal.fire('Error', 'Title is required!', 'error');
                valid = false;
            }
            if (!FilesArr.length) {
                setattachmenterr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                valid = false;
            }
            if (!date) {
                setdateerr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                valid = false;
            }
            if (date == "Invalid Date") {
                setdateerr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                valid = false;
            }
            //  else if (!auditPlanTypeId.length) {
            //     //Swal.fire('Error', 'Entity is required!', 'error');
            //     valid = false;
            // }
            if (!sharewithusers || sharewithusers.length === 0) {
                setsharewitherr(true);
            }
            if (!recommendationRows.length) {
                validRec = false;
            }

            if (recommendationRows.length > 0) {
                let rowErrors: any[] = [];
                //let validRec = true;

                const timeMap: Record<string, number[]> = {}; // To track which times are selected and by which rows

                recommendationRows.forEach((row: any, index: number) => {
                    let rowError: any = {};

                    if (!row.isoreference || row.isoreference.trim() === "") {
                        rowError.isoreference = true;
                    }
                    if (!row.imsprocedure || row.imsprocedure.trim() === "") {
                        rowError.imsprocedure = true;
                    }
                    if (!row.inquiries || row.inquiries.trim() === "") {
                        rowError.inquiries = true;
                    }
                    if (!row.time || row.time.trim() === "") {
                        rowError.time = true;
                    } else {
                        // Track which rows have selected the same time
                        const trimmedTime = row.time.trim();
                        if (!timeMap[trimmedTime]) {
                            timeMap[trimmedTime] = [];
                        }
                        timeMap[trimmedTime].push(index);
                    }
                    if (!row.auditorcomments || row.auditorcomments.trim() === "") {
                        rowError.auditorcomments = true;
                    }
                    // if (!row.sharewith || row.sharewith.length === 0) {
                    //     rowError.sharewith = true;
                    // }

                    // If there are errors in this row, store them
                    if (Object.keys(rowError).length > 0) {
                        rowErrors[index] = rowError;
                        validRec = false;
                    }
                });

                // Now handle duplicate time selections
                Object.keys(timeMap).forEach((timeKey) => {
                    const indices = timeMap[timeKey];
                    if (indices.length > 1) {
                        // Mark all rows with duplicate times as having an error
                        indices.forEach((idx) => {
                            if (!rowErrors[idx]) rowErrors[idx] = {};
                            rowErrors[idx].time = true;
                            validRec = false;
                        });
                    }
                });

                setRowErrors(rowErrors);

                if (!validRec) {
                    console.log("Validation failed. Highlight errors accordingly.");
                }
            }
            let isValidfor = true;
            if (!forwardToArr) {
                valid1 = false;
            }
            if (forwardToArr.length > 0 && forwardToArr.every((row: any) => row.role !== 0 && row.approvalType.trim() !== "" && row.approvers.length != 0) == false) {
                // const isValid = cancellReason.every((row:any) => row.description.trim() !== "" && row.reason.trim() !== "");
                valid1 = false;
            }
            forwardToArr.forEach(row => {
                const errors = {
                    role: row.role === null || row.role === 0,
                    approvers: !row.approvers || row.approvers.length === 0,
                    approvalType: row.approvalType.trim() === "",
                    responsibility: row.Responsibility.trim() === ""
                };

                if (errors.role || errors.approvers || errors.approvalType || errors.responsibility) {
                    validationErrors[row.level] = errors;
                    isValidfor = false;
                }
            });

            setForwardToValidationErrors(validationErrors);
            setValidSubmit(valid);
            setValidDRecomm(validRec);
            setValidAudit(validAudit);
            setValidForwardTo(valid1);

        }
        else {
            setdepartmenterr(false);
            setfromdepartmenterr(false);
            setdateerr(false);
            setattachmenterr(false);
            setisoreferenceerr(false);
            setimsprocedureerr(false);
            setinquirieserr(false);
            settimeerr(false);
            setauditorcommentserr(false);
            setsharewitherr(false);
            setapprovedauditplanerr(false);
            setRowErrors([]);
            setForwardToValidationErrors([]);
            if (!date) {
                setdateerr(true);
                //Swal.fire('Error', 'Title is required!', 'error');
                validraft = false;
            }
            if (date == "Invalid Date") {
                setdateerr(true);
                //Swal.fire('Error', 'Category is required!', 'error');
                validraft = false;
            }
            if (!selectedOption) {
                setapprovedauditplanerr(true);
                validraft = false;
            }
            if (!deptId && !selectUserDept) {
                setdepartmenterr(true);
                //Swal.fire('Error', 'Type is required!', 'error');
                validraft = false;
            }
            if ((fromdeptId == 0 || fromdeptId == undefined) && (!currentUserDept || currentUserDept.length == 0)) {
                setfromdepartmenterr(true);
                //Swal.fire('Error', 'Title is required!', 'error');
                validraft = false;
            }
            // else if (selectedOption == null || !selectedOption.value) {
            //     //Swal.fire('Error', 'Entity is required!', 'error');
            //     valid = false;
            // }


            setValidDraft(validraft);
            // setValidCancelReason(valid1);

        }

        // console.log("validateTitle", validateTitle,"errormsg", errormsg,"valid,", valid, ImagepostArr.length);
        // if (!valid && fmode == FormSubmissionMode.SUBMIT){
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');

        // }
        // else if (!valid1 && fmode == FormSubmissionMode.SUBMIT)
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields in description section.');
        // else if (!valid && fmode == FormSubmissionMode.DRAFT) {
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');
        // }
        // else if (!valid1 && fmode == FormSubmissionMode.DRAFT) {
        //     Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields in description section..');
        // }
        if (valid == false || valid1 == false || validRec == false || validraft == false) {
            Swal.fire(errormsg !== "" ? errormsg : 'Please fill all the mandatory fields.');

            return false
        }
        else {
            return true
        }
        // return valid;
    };
    // #region  Submit Form
    const convertForSharePoint = (values: Record<string, boolean>) => {
        const formatted: Record<string, string> = {};
        for (const key in values) {
            formatted[key] = values[key] ? "Yes" : ""; // or "true"/"false" if your list uses that
        }
        return formatted;
    };
    const convertCheckboxValuesForSharePoint = (values: Record<string, boolean>): Record<string, string> => {
        const formatted: Record<string, string> = {};
        for (const key in values) {
            formatted[key] = values[key] ? "Yes" : "";
        }
        return formatted;
    };

    // Final data object: "Yes"/"" for checkboxes, raw numbers for number fields


    const convertForSharePointNumber = (values: Record<string, number>): Record<string, number> => {
        const formatted: Record<string, number> = {};

        for (const key in values) {
            const value = values[key];
            formatted[key] = value; // Keep numbers as-is
        }

        return formatted;
    };
    // const convertForSharePointNumber = (values: Record<string|number, boolean>) => {
    //     const formatted: Record<string | number, string | number> = {};
    //     for (const key in values) {
    //         formatted[key] = values[key] ? "Yes" : ""; // or "true"/"false" if your list uses that
    //     }
    //     return formatted;
    // };
    const handleFormSubmit = async () => {
        debugger
        if (editID != null) {
            const IsactionTaken = await CheckIfAlreadyactionTaken(sp, editID?.Id);
            if (!IsactionTaken) {
                Swal.fire("Action has already been taken for this record.");
                return;
            }
        }
      
        setValidRemark(true);
        let postPayload = {};
        let postPayload2 = {};
        let postPayloadMemo = {};
        let postPayload2AuditReport: {};
        if (formData.Status === 'Rework' && formData.RemarksInitiator === "") {
            setValidRemark(false);
            Swal.fire('Please fill the mandatory fields', '', 'warning');
            return;
        }
        // Get max sequence/number from the updated array
        const maxNCSequence = Math.max(...NCNumberrows.map(r => r.ncsequenceNC || 0), 0);
        const maxObsSequence = Math.max(...NCNumberrows.map(r => r.observationsequenceObs || 0), 0);
        const maxNCNo = maxNCSequence.toString().padStart(3, "0");
        const maxObsNo = maxObsSequence.toString().padStart(3, "0");
        debugger
        console.log("maxNCSequencemaxNCSequence", maxNCSequence, maxObsSequence, "maxNCNo", maxNCNo, maxObsNo);
        console.log("updatedArr", NCNumberrows);
        // setFormData(prev => ({
        //     ...prev,
        //     NCSequence: maxNCSequence,
        //     ObservationSequence: maxObsSequence,
        //     NCNo: maxNCNo,
        //     ObservationNo: maxObsNo,
        // }));

        if (await validateForm(FormSubmissionMode.SUBMIT)) {
            if (editForm) {
                Swal.fire({
                    title: 'Do you want to submit this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    console.log(result)
                    if (result.isConfirmed) {
                        setLoading(true);
                        debugger
                        setFormData((prevData) => {
                            const updatedData = { ...prevData };
                            if (checkboxValues["FailureofIntentNonconformity"] ||
                                checkboxValues["FailureofImplementation"] ||
                                checkboxValues["FailureofEffectiveness"]) {
                                updatedData.NCSequence = (prevData.NCSequence || 0) + 1;
                            }
                            if (checkboxValues["Observations"]) {
                                updatedData.ObservationSequence = (prevData.ObservationSequence || 0) + 1;
                            }
                            return updatedData;
                        });
                        let newNCsequence: number = formData.NCSequence;
                        let newOnservationsequence: number = formData.ObservationSequence;
                        let newncnumber: string = formData.NCNo;
                        let newobservationNo: string = formData.ObservationNo;
                        if (checkboxValues["FailureofIntentNonconformity"] ||
                            checkboxValues["FailureofImplementation"] ||
                            checkboxValues["FailureofEffectiveness"]) {
                            newNCsequence = (formData.NCSequence || 0) + 1;
                        } else {
                            newncnumber = "";
                        }
                        if (checkboxValues["Observations"]) {
                            newOnservationsequence = (formData.ObservationSequence || 0) + 1;

                        } else {
                            newobservationNo = "";
                        }
                        console.log("updateddatatatata", formData.NCSequence, formData.ObservationSequence)
                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditReportDocs');
                        // const finalDataForSharePoint = {
                        //     ...convertCheckboxValuesForSharePoint(checkboxValues),
                        //     ...checkboxNumberValues
                        // };
                        const formattedData = convertForSharePoint(checkboxValues);
                        const formattedDataNumber = convertForSharePointNumber(checkboxNumberValues);

                        if (FilesArr.length > 0) {
                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    let filenamedraft = await cleanFileNameSave(file.name);
                                    const newFileName = await getNewFileName(filenamedraft);
                                    DocumentName = newFileName;
                                    const fileAddResult = await folder.files.addChunked(newFileName, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult rep)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        debugger
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {
                            ...formattedData,
                            ...formattedDataNumber,
                            MemoNumber: doccode,
                            // MemoSerialNumber:,
                            // IssueNumber:,
                            // RevisionNumber:,
                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            RequestDate: formData.RequestDate != "" ? new Date(formData.RequestDate).toISOString() : new Date().toISOString(),
                            ReportCode: formData.reportCode,
                            NCNumber: checkboxValues["FailureofIntentNonconformity"] ||
                                checkboxValues["FailureofImplementation"] ||
                                checkboxValues["FailureofEffectiveness"] ? maxNCNo : "",
                            ObservationNumber: checkboxValues["Observations"] ? maxObsNo : "",
                            ObservationSequence: maxObsSequence,
                            NCSequence: maxNCSequence,
                            // NCNumber: checkboxValues["FailureofIntentNonconformity"] ? formData.NCNo : "",
                            // ObservationNumber:  formData.ObservationNo : "",
                            // ObservationSequence: newOnservationsequence,
                            // NCSequence: newNCsequence,
                            Title: doccode,
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.fromdeptId,
                            ShiftId: formData.shiftId,
                            DepartmentAuditedId: formData.deptId,
                            Date: formData.date,
                            AuditPlanDate: auditplandate,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds != null ? formData.attachmentIds : [],
                            AuditplanDocId: selectedOption.value,
                            Description: formData.description,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            DocumentName: DocumentName,
                            IsRework: "No",
                            SharewithId: sharewithIds,
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            RemarksInitiator: formData.RemarksInitiator,


                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        debugger
                        const cleanedRows = NCNumberrows.filter(
                            row => row.ncnumberNC.trim() !== "" || row.observationnumberObs.trim() !== ""
                        );
                        for (const row of cleanedRows) {
                            // if ((row.ncnumberNC.trim() == "" && row.observationnumberObs.trim() == "") == false) {
                            const postPayloadNC = {
                                AnnualAuditReportListId: editItemID, // Assuming "Title" column exists
                                NCNumber: row.nctype == "NC Number" ? row.ncnumberNC : row.observationnumberObs,
                                NCSequence: row.nctype == "NC Number" ? row.ncsequenceNC : row.observationsequenceObs,
                                ReportCode: formData.reportCode,
                                NCType: row.nctype,
                                Description: row.descriptionNC,
                                DepartmentId: formData.deptId,
                                Title: row.id.toString()
                            }
                            if (row.Id) {
                                const postResultNC = await updateItemNC(postPayloadNC, sp, row.Id);
                                const postIdNCAdd = postResultNC?.data?.ID;
                            }
                            else {
                                const postResultNC = await addItemNC(postPayloadNC, sp);
                                const postIdNCupdate = postResultNC?.data?.ID;
                            }

                        }
                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditReportIDId: editItemID, // Assuming "Title" column exists
                                ISOReference: row.isoreference,
                                IMSProcedure: row.imsprocedure,
                                Inquiries: row.inquiries,
                                AuditorsComments: row.auditorcomments,
                                Time: row.time,
                                SharewithId: row.sharewithIds ? row.sharewithIds : 0
                                //SharewithId: row.sharewithIds,
                                // sharewith: row.sharewith ? { label: row.sharewith.Title, value: row.sharewith.ID } : null
                            }

                            if (row.id) {

                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {
                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;

                            }

                            // debugger
                            // if (!postId2) {
                            //     console.error("Post creation failed.");
                            //     return;
                            // }
                        }

                        let isValid = true;

                        if (forwardToArr.length) {
                            isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                                row.approvalType.trim() !== "" && row.Responsibility.trim() !== "");

                            // if (!isValid) {
                            //     // alert("Each row must have a role selected and at least one approver.");
                            //     valid = false;
                            // }



                            // if (!valid) {
                            //     Swal.fire('Please fill all the mandatory fields.');
                            //     // setValidSubmit(false);
                            //     setValidForwardTo(false);
                            //     return;
                            // }

                        }

                        if (isValid) {
                            for (const item of forwardToArr) {

                                const approversIds: any[] = [];
                                item.approvers.forEach((user: any) => {
                                    if (user?.value) {
                                        approversIds.push(user.value);
                                    }
                                });

                                let arr2 = {
                                    Title: currentUser.Title,


                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: item.approvers?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(editItemID),
                                    ContentTitle: doccode,
                                    RequestId: formData.reportCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "IMS Audit Report and Checklist",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    Responsibility: item.Responsibility || "",
                                    IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",
                                    // IsApprovalGenerated: "No"
                                    // RedirectionLink:,



                                }
                                if (item.id) {
                                    const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }
                        // *******************************???????????




                        // /////////*************************** */

                        const toDelete = forwardToArrEdit.filter(
                            (itemEdit) => !forwardToArr.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        const toDelete1 = recommendationRowsEdit.filter(
                            (itemEdit) => !recommendationRows.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete1) {
                            try {
                                await sp.web.lists.getByTitle("AnnualAuditReportCheckList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }
                        const toDeleteNC = NCNumberrowsEdit.filter(
                            (itemEdit) => !NCNumberrows.some(item => item.Id === itemEdit.Id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDeleteNC) {
                            try {
                                await sp.web.lists.getByTitle("AuditReportNCNumber").items.getById(item.Id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.Id}`, error);
                            }
                        }


                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            let arr2 = {
                                ActionTakenById: currentUser.Id,
                                ActionTakenOn: new Date().toISOString(),
                                // ActionTakenRoleId: formData.RequesterDesignation,
                                Status: "Approved",
                                Remark: formData.RemarksInitiator,

                            }
                            const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                            const postId = postResult?.data?.ID;

                        }


                        let boolval = false;

                        // if (boolval == true) {
                        setLoading(false);

                        sessionStorage.removeItem("DocumentCancelId");
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
                        // setTimeout(() => {

                        //     window.history.back();
                        //     // window.location.reload();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        //     // let url = window.location.href;
                        //     // let baseUrl = url.split("#")[0];
                        // }, 500);
                        // }
                    }

                })
            }
            else {
                Swal.fire({
                    title: 'Do you want to submit this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    //console.log("Form Submitted:", formValues, bannerImages, galleryImages, documents);
                    if (result.isConfirmed) {
                        setLoading(true);

                        debugger
                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditReportDocs');
                        const formattedData = convertForSharePoint(checkboxValues);
                        const formattedDataNumber = convertForSharePointNumber(checkboxNumberValues);
                        // const finalDataForSharePoint = {
                        //     ...convertCheckboxValuesForSharePoint(checkboxValues),
                        //     ...checkboxNumberValues
                        // };
                        setFormData((prevData) => {
                            const updatedData = { ...prevData };
                            if (checkboxValues["FailureofIntentNonconformity"] ||
                                checkboxValues["FailureofImplementation"] ||
                                checkboxValues["FailureofEffectiveness"]) {
                                updatedData.NCSequence = (prevData.NCSequence || 0) + 1;
                            }
                            if (checkboxValues["Observations"]) {
                                updatedData.ObservationSequence = (prevData.ObservationSequence || 0) + 1;
                            }
                            return updatedData;
                        });
                        let newNCsequence1: number = formData.NCSequence;
                        let newOnservationsequence1: number = formData.ObservationSequence;
                        if (checkboxValues["FailureofIntentNonconformity"] ||
                            checkboxValues["FailureofImplementation"] ||
                            checkboxValues["FailureofEffectiveness"]) {
                            newNCsequence1 = (formData.NCSequence || 0) + 1;
                        }
                        if (checkboxValues["Observations"]) {
                            newOnservationsequence1 = (formData.ObservationSequence || 0) + 1;
                        }
                        console.log("updateddatatatata", formData.NCSequence, formData.ObservationSequence)
                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    const newFileName = await getNewFileName(file.name);
                                    DocumentName = newFileName;
                                    const fileAddResult = await folder.files.addChunked(newFileName, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        debugger
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        let arr = {
                            ...formattedData,
                            ...formattedDataNumber,
                            ReportCode: formData.reportCode,
                            NCNumber: checkboxValues["FailureofIntentNonconformity"] ||
                                checkboxValues["FailureofImplementation"] ||
                                checkboxValues["FailureofEffectiveness"] ? maxNCNo : "",
                            ObservationNumber: checkboxValues["Observations"] ? maxObsNo : "",
                            ObservationSequence: maxObsSequence,
                            NCSequence: maxNCSequence,
                            MemoNumber: doccode,
                            Title: doccode,
                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            RequestDate: formData.RequestDate != "" ? new Date(formData.RequestDate).toISOString() : new Date().toISOString(),
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.fromdeptId,
                            ShiftId: formData.shiftId,
                            DepartmentAuditedId: formData.deptId,
                            Date: formData.date,
                            AuditPlanDate: auditplandate,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds != null ? formData.attachmentIds : [],
                            AuditplanDocId: selectedOption.value,
                            Description: formData.description,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "Yes",
                            Status: "Pending",
                            DocumentName: DocumentName,
                            IsRework: "No",
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            SharewithId: sharewithIds,
                            RemarksInitiator: formData.RemarksInitiator,

                        }

                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }
                        debugger
                        const cleanedRows = NCNumberrows.filter(
                            row => row.ncnumberNC.trim() !== "" || row.observationnumberObs.trim() !== ""
                        );
                        for (const row of cleanedRows) {
                            //if ((row.ncnumberNC.trim() == "" && row.observationnumberObs.trim() == "") == false) {
                            const postPayloadNC = {
                                AnnualAuditReportListId: postId, // Assuming "Title" column exists
                                NCNumber: row.nctype == "NC Number" ? row.ncnumberNC : row.observationnumberObs,
                                NCSequence: row.nctype == "NC Number" ? row.ncsequenceNC : row.observationsequenceObs,
                                ReportCode: formData.reportCode,
                                NCType: row.nctype,
                                Description: row.descriptionNC,
                                DepartmentId: formData.deptId,
                                Title: row.id.toString()
                            }

                            const postResultNC = await addItemNC(postPayloadNC, sp);
                            const postIdNCupdate = postResultNC?.data?.ID;


                        }
                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditReportIDId: postId, // Assuming "Title" column exists
                                ISOReference: row.isoreference,
                                IMSProcedure: row.imsprocedure,
                                Inquiries: row.inquiries,
                                AuditorsComments: row.auditorcomments,
                                Time: row.time,
                                SharewithId: row.sharewithIds ? row.sharewithIds : 0
                            }

                            const postResult2 = await addItem2(postPayload2, sp);
                            const postId2 = postResult2?.data?.ID;
                            // debugger
                            if (!postId2) {
                                console.error("Post creation failed.");
                                return;
                            }
                        }

                        let isValid = true;

                        if (forwardToArr.length) {
                            isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                                row.approvalType.trim() !== "" && row.Responsibility.trim() !== "");

                            // if (!isValid) {
                            //     // alert("Each row must have a role selected and at least one approver.");
                            //     valid = false;
                            // }



                            // if (!valid) {
                            //     Swal.fire('Please fill all the mandatory fields.');
                            //     // setValidSubmit(false);
                            //     setValidForwardTo(false);
                            //     return;
                            // }

                        }

                        if (isValid) {
                            for (const item of forwardToArr) {

                                const approversIds: any[] = [];
                                item.approvers.forEach((user: any) => {
                                    if (user?.value) {
                                        approversIds.push(user.value);
                                    }
                                });

                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: doccode,
                                    RequestId: formData.reportCode,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role,
                                    Level: Number(item.level),
                                    ApproversId: approversIds,
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "Yes",
                                    Maxlevel: item.approvers?.length,
                                    Responsibility: item.Responsibility || "",
                                    IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",
                                    // MainListID: String(editItemID),
                                    MainListID: String(postId),
                                    // RequestId: selectedOption.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "IMS Audit Report and Checklist",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No"
                                    // RedirectionLink:,



                                }
                                if (item.id) {
                                    const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }




                        let boolval;

                        // if (boolval == true) {
                        setLoading(false);
                        Swal.fire('Submitted successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
                        //Swal.fire('Submitted successfully.', '', 'success');
                        // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 500);
                        // }

                    }
                })

            }
        }

    }
    React.useEffect(() => {
        const updatedValues: Record<string, number> = { ...checkboxNumberValues };

        AuditFindingsOptions.forEach((checkbox) => {
            //const fieldKey = CheckboxFieldMap[checkbox.value];
            const fieldKeyN = CheckboxFieldMapN[checkbox.value];
            if (checkbox.value === 'Observations') {
                updatedValues[fieldKeyN] = NCNumberrows.filter(
                    row => (row.ncnumberNC || row.observationnumberObs) && row.nctype === "Observation Number"
                ).length;
            }

            if (checkbox.value === 'Failure of Intent / Nonconformity') {
                updatedValues[fieldKeyN] = NCNumberrows.filter(
                    row => (row.ncnumberNC || row.observationnumberObs) && row.nctype === "NC Number"
                ).length;
            }
        });

        setCheckboxNumberValues(prev => ({
            ...prev,
            ...updatedValues
        }));
    }, [NCNumberrows]);

    const handleSaveAsDraft = async () => {
        debugger

        const maxNCSequence = Math.max(...NCNumberrows.map(r => r.ncsequenceNC || 0), 0);
        const maxObsSequence = Math.max(...NCNumberrows.map(r => r.observationsequenceObs || 0), 0);
        const maxNCNo = maxNCSequence.toString().padStart(3, "0");
        const maxObsNo = maxObsSequence.toString().padStart(3, "0");
        debugger
        console.log("maxNCSequencemaxNCSequence", maxNCSequence, maxObsSequence, "maxNCNo", maxNCNo, maxObsNo);
        console.log("updatedArr", NCNumberrows);
        console.log("checkboxvalueee", checkboxValues);
        console.log("checkboxNumberValues", checkboxNumberValues);
        if (await validateForm(FormSubmissionMode.DRAFT)) {
            if (editForm) {
                Swal.fire({
                    title: 'Do you want to save this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    debugger
                    console.log(result)
                    if (result.isConfirmed) {
                        console.log("checkboxvalueee conf", checkboxValues);
                        setLoading(true);
                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditReportDocs');
                        const formattedData = convertForSharePoint(checkboxValues);
                        const formattedDataNumber = convertForSharePointNumber(checkboxNumberValues);
                        // const finalDataForSharePoint = {
                        //     ...convertCheckboxValuesForSharePoint(checkboxValues),
                        //     ...checkboxNumberValues
                        // };

                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    let filenamedraft = await cleanFileNameSave(file.name);
                                    const newFileName = await getNewFileName(filenamedraft);
                                    DocumentName = newFileName;
                                    const fileAddResult = await folder.files.addChunked(newFileName, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        // let TypeMasterData: any = await getAnnouncementandNewsTypeMaster(sp, Number(formData.Type))
                        let arr = {
                            ...formattedData,
                            ...formattedDataNumber,
                            ReportCode: formData.reportCode,
                            NCNumber: checkboxValues["FailureofIntentNonconformity"] ||
                                checkboxValues["FailureofImplementation"] ||
                                checkboxValues["FailureofEffectiveness"] ? maxNCNo : "",
                            ObservationNumber: checkboxValues["Observations"] ? maxObsNo : "",
                            ObservationSequence: maxObsSequence,
                            NCSequence: maxNCSequence,
                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            RequestDate: formData.RequestDate != "" && formData.RequestDate != null ? new Date(formData.RequestDate).toISOString() : new Date().toISOString(),
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.fromdeptId,
                            ShiftId: formData.shiftId,
                            DepartmentAuditedId: formData.deptId,
                            Date: formData.date,
                            AuditPlanDate: auditplandate,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            MemoNumber: doccode,
                            Title: doccode,
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds != null ? formData.attachmentIds : [],
                            AuditplanDocId: selectedOption.value,
                            Description: formData.description,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "No",
                            Status: formData.Status == "Rework" ? "Rework" : "Save as draft",
                            DocumentName: DocumentName,
                            IsRework: formData.Status == "Rework" ? "Yes" : "No",
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            SharewithId: sharewithIds,
                            RemarksInitiator: formData.RemarksInitiator,
                        }
                        const postResult = await updateItem(arr, sp, editItemID);
                        const postId = postResult?.data?.ID;
                        //  ////////////////////////////
                        debugger

                        const cleanedRows = NCNumberrows.filter(
                            row => row.ncnumberNC.trim() !== "" || row.observationnumberObs.trim() !== ""
                        );

                        for (const row of cleanedRows) {
                            //if ((row.ncnumberNC.trim() == "" && row.observationnumberObs.trim() == "") == false) {
                            const postPayloadNC = {
                                AnnualAuditReportListId: editItemID, // Assuming "Title" column exists
                                NCNumber: row.nctype == "NC Number" ? row.ncnumberNC : row.observationnumberObs,
                                NCSequence: row.nctype == "NC Number" ? row.ncsequenceNC : row.observationsequenceObs,
                                ReportCode: formData.reportCode,
                                NCType: row.nctype,
                                Description: row.descriptionNC,
                                DepartmentId: formData.deptId,
                                Title: row.id.toString()
                            }
                            if (row.Id) {
                                const postResultNC = await updateItemNC(postPayloadNC, sp, row.Id);
                                const postIdNCAdd = postResultNC?.data?.ID;
                            }
                            else {
                                if ((row.ncnumberNC.trim() == "") == false || (row.observationnumberObs.trim() == "") == false) {
                                    const postResultNC = await addItemNC(postPayloadNC, sp);
                                    const postIdNCupdate = postResultNC?.data?.ID;
                                }
                            }

                        }
                        for (const row of recommendationRows) {

                            const postPayload2 = {
                                AnnualAuditReportIDId: editItemID, // Assuming "Title" column exists
                                ISOReference: row.isoreference || "",
                                IMSProcedure: row.imsprocedure || "",
                                Inquiries: row.inquiries || "",
                                AuditorsComments: row.auditorcomments || "",
                                Time: row.time || "",
                                SharewithId: row.sharewithIds ? row.sharewithIds : 0,
                            }

                            if (row.id) {

                                const postResult2 = await updateItem2(postPayload2, sp, row.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {

                                if ((row.isoreference.trim() == "" && row.imsprocedure.trim() == "" && row.inquiries.trim() == "" && row.auditorcomments.trim() == "" && row.time.trim() == "" && (row.sharewith == null || row.sharewith.length == 0)) == false) {


                                    const postResult2 = await addItem2(postPayload2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                    if (!postId2) {
                                        console.error("Post creation failed.");
                                        return;
                                    }

                                }



                            }
                            // }


                            // debugger

                        }

                        let isValid = true;

                        // if (forwardToArr.length) {
                        //     isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                        //         row.approvalType.trim() !== "");


                        // }

                        // if (isValid) {
                        for (const item of forwardToArr) {

                            const approversIds: any[] = [];
                            item.approvers.forEach((user: any) => {
                                if (user?.value) {
                                    approversIds.push(user.value);
                                }
                            });

                            let arr2 = {
                                Title: currentUser.Title,
                                // ContentTitle: selectedOption.ReferenceNumber,
                                ContentTitle: doccode,
                                RequestId: formData.reportCode,
                                MainListNameId: ListNameId,
                                ApproverRoleId: item.role || 0,
                                Level: Number(item.level),
                                ApproversId: approversIds || [],
                                // LevelType: "One",
                                LevelType: item.approvalType,
                                SubmitStatus: "No",
                                Maxlevel: item.approvers?.length,
                                Responsibility: item.Responsibility || "",
                                IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",
                                // MainListID: String(editItemID),
                                MainListID: String(editItemID),
                                // RequestId: selectedOption.DocumentCode,
                                // RequestId:String(editID.Id),
                                RequesterNameId: currentUser.Id,
                                RequestedDate: new Date().toLocaleDateString("en-CA"),
                                RequesterRoleId: RequesterRoleId,
                                ProcessName: "IMS Audit Report and Checklist",
                                FormNameId: FormNameId,
                                ApprovalType: "Approval",
                                // IsApprovalGenerated: "No"
                                // RedirectionLink:,



                            }
                            if (item.id) {
                                const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                const postId2 = postResult2?.data?.ID;

                            }
                            else {
                                if (forwardToArr.every(row => row.role == 0 && row.approvers.length == 0 &&
                                    row.approvalType.trim() == "" && row.Responsibility.trim() == "") == false) {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }

                        // }



                        // /#############//////////////////

                        const toDelete = forwardToArrEdit.filter(
                            (itemEdit) => !forwardToArr.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete) {
                            try {
                                await sp.web.lists.getByTitle("AllProcessApprovalLevelList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }

                        const toDelete1 = recommendationRowsEdit.filter(
                            (itemEdit) => !recommendationRows.some(item => item.id === itemEdit.id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDelete1) {
                            try {
                                await sp.web.lists.getByTitle("AnnualAuditReportCheckList").items.getById(item.id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.id}`, error);
                            }
                        }
                        const toDeleteNC = NCNumberrowsEdit.filter(
                            (itemEdit) => !NCNumberrows.some(item => item.Id === itemEdit.Id) // Assuming ID is the unique key
                        );

                        // Delete each item from SharePoint
                        for (const item of toDeleteNC) {
                            try {
                                await sp.web.lists.getByTitle("AuditReportNCNumber").items.getById(item.Id).delete();
                                // console.log(`Deleted item with ID: ${item.ID}`);
                            } catch (error) {
                                console.error(`Error deleting item with ID: ${item.Id}`, error);
                            }
                        }

                        if (DraftApprovalItem != null && DraftApprovalItem != undefined && DraftApprovalItem.length > 0) {

                            let arr2 = {
                                ActionTakenById: currentUser.Id,
                                ActionTakenOn: new Date().toISOString(),
                                // ActionTakenRoleId: formData.RequesterDesignation,
                                //Status: "Save as draft",
                                // Remark: remark,

                            }
                            const postResult = await updateApprovalItem(arr2, sp, DraftApprovalItem[0].Id);
                            const postId = postResult?.data?.ID;

                        }


                        let boolval = false;

                        // if (boolval == true) {
                        setLoading(false);
                        sessionStorage.removeItem("DocumentCancelId")
                        Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });

                        // setTimeout(() => {

                        //     window.history.back();
                        //     // window.location.reload();
                        //     setTimeout(() => {
                        //         location.reload();
                        //     }, 100);
                        // }, 1000);
                        // }
                    }

                })
            }
            else {
                Swal.fire({
                    title: 'Do you want to save this request?',
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: 'warning'
                }
                ).then(async (result) => {
                    //console.log("Form Submitted:", formValues, bannerImages, galleryImages, documents);
                    if (result.isConfirmed) {
                        setLoading(true);
                        debugger

                        // let galleryIds: any[] = [];

                        let galleryArray: any[] = [];
                        let bannerImageArray: any = {};
                        let DocumentName: string = "";
                        let attachmentIds = [];
                        const folder = sp.web.getFolderByServerRelativePath('/sites/ededms/AnnualAuditReportDocs');
                        const formattedData = convertForSharePoint(checkboxValues);
                        const formattedDataNumber = convertForSharePointNumber(checkboxNumberValues);
                        // const finalDataForSharePoint = {
                        //     ...convertCheckboxValuesForSharePoint(checkboxValues),
                        //     ...checkboxNumberValues
                        // };
                        if (FilesArr.length > 0) {

                            for (const file of FilesArr) {
                                if (!file.ID) {
                                    //bannerImageArray = await uploadFile(file, sp, "ChangeRequestDocs", tenantUrl);
                                    const newFileName = await getNewFileName(file.name);
                                    DocumentName = newFileName;
                                    const fileAddResult = await folder.files.addChunked(newFileName, file);
                                    const fileNew = fileAddResult.file;
                                    const documentName = fileAddResult.data.Name;
                                    bannerImageArray = fileAddResult;
                                    galleryArray.push(bannerImageArray);
                                    // Get the item ID for the uploaded file
                                    const currentItemId = await fileNew.getItem<{ Id: number }>();
                                    const itemId = currentItemId.Id;
                                    // await currentItemId.update({
                                    //     FileName: documentName, // Assuming FileName is the internal name of the column
                                    // });
                                    console.log("JSON.stringify(fileAddResult)", JSON.stringify(fileAddResult))
                                    // Save the document ID for the attachment field in ChangeRequestList
                                    attachmentIds.push(itemId);


                                }
                                else {
                                    // const itemId = file.ID;
                                    attachmentIds.push(file.ID);
                                }



                            }
                        }
                        const sharewithIds: any[] = [];
                        sharewithusers.forEach((user: any) => {
                            if (user?.value) {
                                sharewithIds.push(user.value);
                            }
                        });
                        let arr = {

                            // MemoNumber:,
                            // MemoSerialNumber:,
                            // IssueNumber:,
                            // RevisionNumber:,
                            ...formattedData,
                            ...formattedDataNumber,
                            NCNumber: checkboxValues["FailureofIntentNonconformity"] ||
                                checkboxValues["FailureofImplementation"] ||
                                checkboxValues["FailureofEffectiveness"] ? maxNCNo : "",
                            ObservationNumber: checkboxValues["Observations"] ? maxObsNo : "",
                            RequesterNameId: formData.RequesterNameId,
                            RequesterDesignation: formData.RequesterDesignation,
                            RequestDate: new Date().toISOString(),
                            ObservationSequence: maxObsSequence,
                            NCSequence: maxNCSequence,
                            ReportCode: formData.reportCode,
                            MemoNumber: doccode,
                            Title: doccode,
                            ApprovedAuditPlanId: selectAuditplan.ID,
                            DepartmentId: formData.fromdeptId,
                            ShiftId: formData.shiftId,
                            DepartmentAuditedId: formData.deptId,
                            Date: formData.date,
                            AuditPlanDate: auditplandate,
                            DocumentCode: formData.documentCode,
                            IssueNumber: Number(formData.issueNo),
                            RevisionNumber: Number(formData.revisionNo),
                            ReferenceNumber: formData.referenceNo,
                            RevisionDate: formData.revisionDate == null ? null : new Date(formData.revisionDate).toLocaleDateString("en-CA"),
                            IssueDate: formData.issueDate == null ? null : new Date(formData.issueDate).toLocaleDateString("en-CA"),
                            //DocumentLink: "",
                            AnnualAuditPlanDocumentLinkId: formData.attachmentIds != null ? formData.attachmentIds : [],
                            AuditplanDocId: selectedOption.value,
                            Description: formData.description,
                            SubmiitedDate: new Date().toLocaleDateString("en-CA"),
                            SubmitStatus: "No",
                            Status: formData.Status == "Rework" ? "Rework" : "Save as draft",
                            DocumentName: DocumentName,
                            IsRework: formData.Status == "Rework" ? "Yes" : "No",
                            AttachmentId: attachmentIds || [],
                            AttachmentJson: JSON.stringify(bannerImageArray) || "",
                            SharewithId: sharewithIds,
                            RemarksInitiator: formData.RemarksInitiator,

                        }
                        // console.log(postPayload);

                        const postResult = await addItem(arr, sp);
                        const postId = postResult?.data?.ID;
                        // // debugger
                        if (!postId) {
                            console.error("Post creation failed.");
                            return;
                        }
                        debugger
                        const cleanedRows = NCNumberrows.filter(
                            row => row.ncnumberNC.trim() !== "" || row.observationnumberObs.trim() !== ""
                        );
                        for (const row of cleanedRows) {
                            //if ((row.ncnumberNC.trim() == "" && row.observationnumberObs.trim() == "") == false) {
                            debugger
                            const postPayloadNC = {
                                AnnualAuditReportListId: postId, // Assuming "Title" column exists
                                NCNumber: row.nctype == "NC Number" ? row.ncnumberNC : row.observationnumberObs,
                                NCSequence: row.nctype == "NC Number" ? row.ncsequenceNC : row.observationsequenceObs,
                                ReportCode: formData.reportCode,
                                NCType: row.nctype,
                                Description: row.descriptionNC,
                                DepartmentId: formData.deptId,
                                Title: row.id.toString()
                            }
                            const postResultNC = await addItemNC(postPayloadNC, sp);
                            const postId2 = postResultNC?.data?.ID;
                            // debugger
                            if (!postId2) {
                                console.error("Post creation failed.");
                                return;
                            }

                        }
                        for (const row of recommendationRows) {
                            if ((row.isoreference.trim() == "" && row.imsprocedure.trim() == "" && row.inquiries.trim() == "" && row.auditorcomments.trim() == "" &&
                                row.time.trim() == "" && (row.sharewith == null || row.sharewith.length == 0)) == false) {
                                // if ((row.section.trim() == "" && row.date.trim() == "" && row.startTime.trim() == "" && 
                                // (row.auditor == null || row.auditor.length == 0)) == false) {
                                debugger
                                const postPayload2 = {
                                    AnnualAuditReportIDId: postId, // Assuming "Title" column exists
                                    ISOReference: row.isoreference || "",
                                    IMSProcedure: row.imsprocedure || "",
                                    Inquiries: row.inquiries || "",
                                    AuditorsComments: row.auditorcomments || "",
                                    Time: row.time || "",
                                    SharewithId: row.sharewithIds ? row.sharewithIds : 0,

                                }

                                const postResult2 = await addItem2(postPayload2, sp);
                                const postId2 = postResult2?.data?.ID;
                                // debugger
                                if (!postId2) {
                                    console.error("Post creation failed.");
                                    return;
                                }
                            }


                        }

                        let isValid = true;

                        if (forwardToArr.length) {
                            isValid = forwardToArr.every(row => row.role !== 0 && row.approvers.length > 0 &&
                                row.approvalType.trim() !== "" && row.Responsibility.trim() !== "");
                        }

                        // if (isValid) {
                        for (const item of forwardToArr) {
                            if ((item.role == 0 && item.approvers.length == 0 &&
                                item.approvalType.trim() == "") == false) {



                                const approversIds: any[] = [];
                                item.approvers.forEach((user: any) => {
                                    if (user?.value) {
                                        approversIds.push(user.value);
                                    }
                                });

                                let arr2 = {
                                    Title: currentUser.Title,
                                    // ContentTitle: selectedOption.ReferenceNumber,
                                    ContentTitle: doccode,
                                    RequestId: formData.reportCode,
                                    MainListNameId: ListNameId,
                                    ApproverRoleId: item.role ? item.role : 0,
                                    Level: Number(item.level),
                                    ApproversId: approversIds || [],
                                    Responsibility: item.Responsibility || "",
                                    IsSignatureRequired: item.IsSignatureRequired ? "Yes" : "No",
                                    // LevelType: "One",
                                    LevelType: item.approvalType,
                                    SubmitStatus: "No",
                                    Maxlevel: item.approvers?.length,

                                    // MainListID: String(editItemID),
                                    MainListID: String(postId),
                                    // RequestId: selectedOption.DocumentCode,
                                    // RequestId:String(editID.Id),
                                    RequesterNameId: currentUser.Id,
                                    RequestedDate: new Date().toLocaleDateString("en-CA"),
                                    RequesterRoleId: RequesterRoleId,
                                    ProcessName: "IMS Audit Report and Checklist",
                                    FormNameId: FormNameId,
                                    ApprovalType: "Approval",
                                    IsApprovalGenerated: "No"
                                    // RedirectionLink:,



                                }
                                if (item.id) {
                                    const postResult2 = await UpdateAllProcessItem(arr2, sp, item.id);
                                    const postId2 = postResult2?.data?.ID;

                                }
                                else {
                                    const postResult2 = await addAllProcessItem(arr2, sp);
                                    const postId2 = postResult2?.data?.ID;
                                }

                            }

                        }

                        // }




                        setLoading(false);
                        Swal.fire('Saved successfully.', '', 'success').then(async (result) => {
                            if (result.isConfirmed) {
                                window.location.href = `https://edcadae.sharepoint.com/sites/ededms/SitePages/EDCMAIN.aspx`;
                            }
                        });
                        // sessionStorage.removeItem("bannerId")
                        // setTimeout(() => {
                        //     window.location.reload();
                        //     // window.history.back();
                        // }, 1000);
                    }
                })

            }
        }

    }

    const formatToDDMMYYYY = (dateInput: string | number | Date) => {
        const date = new Date(dateInput);
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0'); // months are zero-indexed
        const year = date.getFullYear();

        return `${day}-${month}-${year}`;
    }

    const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>, libraryName: string, docLib: string) => {
        event.preventDefault();
        setshowviewdownload(false);
        const allowedTypes = [
            "image/jpeg",
            "image/png",
            "image/gif",
            "image/bmp",
            "image/svg+xml",
            "application/pdf",
            "application/msword",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-powerpoint",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ];

        filechanged = true;
        newfileupload = true;
        let uloadBannerImageFiles: any[] = [];
        let uloadImageFiles: any[] = [];
        let uloadImageFiles1: any[] = [];


        if (event.target.files && event.target.files.length > 0) {
            const files = Array.from(event.target.files);
            (event.target as HTMLInputElement).value = '';

            if (files.length > 0) {

                for (const fn of files) {
                    // const file = files[0];
                    if (!allowedTypes.includes(fn.type)) {
                        Swal.fire({
                            icon: "error",
                            title: "Invalid File Type",
                            text: "Only images and document files are allowed.",
                        });
                        return;
                    }
                    const fileType = fn.type.split("/")[0]; // Extract file type (image, pdf, etc.)
                    const preview = URL.createObjectURL(fn);
                }
                setFilesArr([...FilesArr, ...files]);
            } else {
                Swal.fire("upload a document")
            }
        }
    };
    const AuditFindingsOptions = [
        { value: 'Conforming & Positive Findings', label: 'Conforming & Positive Findings' },
        { value: 'Failure of Intent / Nonconformity', label: 'Failure of Intent / Nonconformity' },
        { value: 'Observations', label: 'Observations' },
        { value: 'Failure of Implementation', label: 'Failure of Implementation' },
        { value: 'Opportunities for Improvement', label: 'Opportunities for Improvement' },
        { value: 'Failure of Effectiveness', label: 'Failure of Effectiveness' }

    ] as const;
    const CheckboxFieldMap = {
        "Conforming & Positive Findings": "ConformingPositiveFindings",
        "Failure of Intent / Nonconformity": "FailureofIntentNonconformity",
        "Observations": "Observations",
        "Failure of Implementation": "FailureofImplementation",
        "Opportunities for Improvement": "OpportunitiesforImprovement",
        "Failure of Effectiveness": "FailureofEffectiveness",
    } as const;
    const CheckboxFieldMapN = {
        "Conforming & Positive Findings": "ConformingPositiveFindingsN",
        "Failure of Intent / Nonconformity": "FailureofIntentNonconformityN",
        "Observations": "ObservationsN",
        "Failure of Implementation": "FailureofImplementationN",
        "Opportunities for Improvement": "OpportunitiesforImprovementN",
        "Failure of Effectiveness": "FailureofEffectivenessN",
    } as const;
    type CheckboxKey = keyof typeof CheckboxFieldMap;
    type CheckboxState = {
        [K in (typeof CheckboxFieldMap)[CheckboxKey]]: boolean;
    }
    type CheckboxLabel = keyof typeof CheckboxFieldMap;
    const renderCheckboxes = () => {
        console.log("formdataaaaaaaaa", formData, EnableNC, InputDisabled, formData?.Status !== "Rework", EnableNC, EnableNCObs);
        return (
            <div className="row">
                {AuditFindingsOptions.map((checkbox) => {
                    const fieldKey = CheckboxFieldMap[checkbox.value];
                    const fieldKeyN = CheckboxFieldMapN[checkbox.value];
                    const shouldDisableCheckbox =
                        (InputDisabled && formData?.Status !== "Rework") || !EnableNCObs || !isDepartmentReady;
                    return (
                        <div className="col-lg-6 mb-3" key={checkbox.value}>
                            <div className="d-flex justify-content-between align-items-center">
                                {/* Left side: checkbox and label */}
                                <div className="d-flex align-items-center">
                                    <input
                                        //onKeyDown={handleKeyDowntext}
                                        type="checkbox"
                                        className="form-check-input me-2"
                                        id={`checkbox-${fieldKey}`}
                                        disabled={shouldDisableCheckbox}
                                        checked={checkboxValues[fieldKey]}
                                        onChange={() => handleCheckboxChange(checkbox.value)}
                                    />
                                    <label className="form-check-label" htmlFor={`checkbox-${fieldKey}`}>
                                        {checkbox.label}
                                    </label>
                                </div>

                                {/* Right side: number input */}
                                <input
                                    //onKeyDown={handleKeyDowntext}
                                    type="number"
                                    className="form-control"
                                    style={{ maxWidth: '70px', minWidth: '60px' }}
                                    //value={checkboxNumberValues[fieldKeyN] ?? ''}
                                    value={
                                        checkboxNumberValues[fieldKeyN] === 0 ? '' : checkboxNumberValues[fieldKeyN] ?? ''
                                    }
                                    onChange={(e) => handleNumberChange(fieldKeyN, Number(e.target.value))}
                                    disabled={InputDisabled || checkbox.value === 'Observations' || checkbox.value === 'Failure of Intent / Nonconformity'}
                                    placeholder="0"
                                />
                            </div>
                        </div>

                    )
                })}
                <div className="col-lg-12">
                    <div className="mb-3">
                        <label htmlFor="Description" className="col-form-label">Conforming & Positive Findings</label>
                        <div>
                            <textarea
                                //onKeyDown={handleKeyDowntextarea}
                                title={formData.description}
                                id="simpleinput"
                                disabled={InputDisabled}
                                value={formData.description}
                                className={`form-control mb-0`}
                                onChange={(e) => setFormData({ ...formData, description: e.target.value })}
                            />

                        </div>
                    </div>
                </div>
                {/* {checkboxValues["FailureofIntentNonconformity"] &&
                    <><div className='row'>
                        <div className='col-sm-6'>
                            <h3 className='text-dark font-16 fw-bold mb-3'>NC Number</h3>
                        </div>
                        <div style={{ textAlign: 'right' }} className='col-sm-6'>
                            {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")}
                                onClick={() => handleaddncrows("NC Number")}></img>}
                            {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/minus.png")}
                                onClick={() => handleDeleteNCNumberRow("NC Number")}></img>}


                        </div>
                    </div>
                    <div style={{ display: 'grid', overflow: 'auto' }}>
                            <table id="tabRec" className='mtbalenew overhi'>
                                <thead>
                                    <tr><th>Type</th>
                                        <th>NC</th>
                                        {/* <th style={{ minWidth: '60px', maxWidth: '60px' }}>Action</th> */}
                {/* </tr>
                                </thead>

                                <tbody>
                                    {console.log("validNCNumberrows", NCNumberrows)}
                                    {NCNumberrows.filter(row => row.nctype === "NC Number").map((row, index) => (
                                        <tr key={index}>
                                            <td title={"NC Number"}
                                                style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                {/* <label htmlFor="approvalType">Approval Type: </label> */}
                {/* <select
                                                    id="approvalType"
                                                    value="NC Number"
                                                    className={`newse form-select`}
                                                    disabled
                                                >
                                                    <option value="NC Number">NC</option>
                                                </select>

                                            </td> */}
                {/* <td title={row?.ncnumberNC ? row?.ncnumberNC : row?.ncnumberNC}>
                                                <input
                                                    type="number"
                                                    className={`form-control ${(RowErrors[index]?.ncnumberNC) ? "border-on-error" : ""}`}
                                                    // className="form-control"
                                                    value={row.nctype === "NC Number"
                                                        ? row.ncnumberNC
                                                        : row.nctype === "Observation Number"
                                                            ? row.observationnumberObs
                                                            : ""}
                                                    onChange={(e) => handleNCNumberChange(index, row.nctype === "NC Number"
                                                        ? "ncnumberNC"
                                                        : row.nctype === "Observation Number"
                                                            ? "observationnumberObs"
                                                            : "", e.target.value)}
                                                    disabled={true} />
                                            </td> */}

                {/* {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td style={{ minWidth: '60px', maxWidth: '60px' }}>
                                                <img src={require("../assets/del.png")} onClick={() => handleDeleteNCNumberRow(index, row.nctype)} />

                                            </td>} */}
                {/* </tr>
                                    ))}
                                </tbody>

                            </table>
                        </div></> */}
                {/* }  */}
                {/* {checkboxValues["Observations"] &&
                    <><div className='row'>
                        <div className='col-sm-6'>
                            <h3 className='text-dark font-16 fw-bold mb-3'>Observation</h3>
                        </div>
                        <div style={{ textAlign: 'right' }} className='col-sm-6'>
                            {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")}
                                onClick={() => handleaddncrows("Observation Number")}></img>}
                            {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/minus.png")}
                                onClick={() => handleDeleteNCNumberRow("Observation Number")}></img>}


                        </div>
                    </div><div style={{ display: 'grid', overflow: 'auto' }}>
                            <table id="tabRec" className='mtbalenew overhi'>
                                <thead>
                                    <tr><th>Type</th>
                                        <th>Observation</th>
                                        {/* <th style={{ minWidth: '60px', maxWidth: '60px' }}>Action</th> */}
                {/* </tr>
                                </thead>

                                <tbody>
                                    {console.log("validNCNumberrows", NCNumberrows)}
                                    {NCNumberrows.filter(row => row.nctype === "Observation Number").map((row, index) => (
                                        <tr key={index}>
                                            <td style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                {/* <label htmlFor="approvalType">Approval Type: </label> */}
                {/* <select
                                                    id="approvalType"
                                                    value="Observation Number"
                                                    className={`newse form-select`}
                                                    disabled
                                                >
                                                    <option value="Observation Number">Observation Number</option>
                                                </select>


                                            </td>
                                            <td title={row?.ncnumberNC ? row?.ncnumberNC : row?.ncnumberNC}>
                                                <input */}
                {/* type="number"
                                                    className={`form-control ${(RowErrors[index]?.ncnumberNC) ? "border-on-error" : ""}`}
                                                    // className="form-control"
                                                    value={row.nctype === "NC Number"
                                                        ? row.ncnumberNC
                                                        : row.nctype === "Observation Number"
                                                            ? row.observationnumberObs
                                                            : ""}
                                                    onChange={(e) => handleNCNumberChange(index, row.nctype === "NC Number"
                                                        ? "ncnumberNC"
                                                        : row.nctype === "Observation Number"
                                                            ? "observationnumberObs"
                                                            : "", e.target.value)}
                                                    disabled={true} />
                                            </td>

                                            {/* {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td style={{ minWidth: '60px', maxWidth: '60px' }}>
                                                <img src={require("../assets/del.png")} onClick={() => handleDeleteNCNumberRow(index, row.nctype)} />

                                            </td>} */}
                {/* </tr>
                                    ))}
                                </tbody>

                            </table>
                        </div></> */}
                {/* }  */}
                <div style={{ gap: '20px' }}>
                    {/* NC Number Table */}
                    {(checkboxValues["FailureofIntentNonconformity"] ||
                        checkboxValues["FailureofImplementation"] ||
                        checkboxValues["FailureofEffectiveness"]) && (
                            // <div style={{ width: '100%' }}>

                            <> <div className='row'>
                                <div className='col-sm-6'>
                                    <h3 className='text-dark font-16 fw-bold mb-3'>NC Number</h3>
                                </div>
                                <div style={{ textAlign: 'right' }} className='col-sm-6'>
                                    {!InputDisabled && (
                                        <>
                                            <img
                                                style={{ width: '30px', cursor: 'pointer' }}
                                                className='mt-0'
                                                src={require("../assets/plus.png")}
                                                onClick={() => handleaddncrows("NC Number")}
                                            />
                                            <img
                                                style={{ width: '30px', cursor: 'pointer', marginLeft: '10px' }}
                                                className='mt-0'
                                                src={require("../assets/minus.png")}
                                                onClick={() => handleDeleteNCNumberRow("NC Number")}
                                            />
                                        </>
                                    )}
                                </div>
                            </div>
                                <div style={{ display: 'grid', overflow: 'auto' }}>
                                    <table id="tabRec" className='mtbalenew overhi'>
                                        <thead>
                                            <tr>
                                                <th style={{ minWidth: '60px', maxWidth: '60px' }}>Sr No</th>
                                                <th style={{ minWidth: '70px', maxWidth: '70px' }}>Category</th>
                                                <th style={{ minWidth: '220px', maxWidth: '220px' }}>Description</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {NCNumberrows
                                                .filter(row => (row.ncnumberNC || row.observationnumberObs) && row.nctype === "NC Number").map((row, index) => (
                                                    <tr key={index}>

                                                        <td style={{ minWidth: '60px', maxWidth: '60px' }} title={row.ncnumberNC}>
                                                            <input
                                                                //onKeyDown={handleKeyDowntext}
                                                                type="number"
                                                                className={`form-control ${(RowErrors[index]?.ncnumberNC) ? "border-on-error" : ""}`}
                                                                value={row.ncnumberNC}
                                                                onChange={(e) =>
                                                                    handleNCNumberChange(index, "ncnumberNC", e.target.value)
                                                                }
                                                                disabled={true}
                                                            />
                                                        </td>
                                                        <td style={{ minWidth: '70px', maxWidth: '70px' }} title={"NC"}>
                                                            <select
                                                                id="approvalType"
                                                                value="NC Number"
                                                                className="newse form-select"
                                                                disabled
                                                            >
                                                                <option value="NC Number">NC</option>
                                                            </select>
                                                        </td>
                                                        <td style={{ minWidth: '220px', maxWidth: '220px' }} title={row.descriptionNC}>
                                                            <textarea
                                                                //onKeyDown={handleKeyDowntextarea}
                                                                id="simpleinput"
                                                                className={`form-control`}
                                                                // className="form-control"
                                                                value={row.descriptionNC}
                                                                onChange={(e) => handleNCDescriptionChange(row.id, 'descriptionNC', e.target.value)}
                                                                disabled={InputDisabled}
                                                            />
                                                        </td>
                                                    </tr>
                                                ))}
                                        </tbody>
                                    </table>
                                </div>
                            </>
                            // {/* </div> */}
                        )}

                    {/* Observation Table */}
                    {checkboxValues["Observations"] && (
                        // <div style={{ width: '100%' }}>
                        <>
                            <div className='row mt-3'>
                                <div className='col-sm-6'>
                                    <h3 className='text-dark font-16 fw-bold mb-3'>Observation</h3>
                                </div>
                                <div style={{ textAlign: 'right' }} className='col-sm-6'>
                                    {!InputDisabled && (
                                        <>
                                            <img
                                                style={{ width: '30px', cursor: 'pointer' }}
                                                className='mt-0'
                                                src={require("../assets/plus.png")}
                                                onClick={() => handleaddncrows("Observation Number")}
                                            />
                                            <img
                                                style={{ width: '30px', cursor: 'pointer', marginLeft: '10px' }}
                                                className='mt-0'
                                                src={require("../assets/minus.png")}
                                                onClick={() => handleDeleteNCNumberRow("Observation Number")}
                                            />
                                        </>
                                    )}
                                </div>
                            </div>
                            <div style={{ display: 'grid', overflow: 'auto' }}>
                                <table id="tabRec" className='mtbalenew overhi'>
                                    <thead>
                                        <tr>
                                            <th style={{ minWidth: '60px', maxWidth: '60px' }}>Sr No</th>
                                            <th style={{ minWidth: '70px', maxWidth: '70px' }}>Category</th>
                                            <th style={{ minWidth: '220px', maxWidth: '220px' }}>Description</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {NCNumberrows.filter(row => (row.ncnumberNC || row.observationnumberObs) && row.nctype === "Observation Number").map((row, index) => (
                                            <tr key={index}>
                                                <td style={{ minWidth: '60px', maxWidth: '60px' }} title={row.observationnumberObs}>
                                                    <input
                                                        // onKeyDown={handleKeyDowntext}
                                                        type="number"
                                                        className={`form-control ${(RowErrors[index]?.observationnumberObs) ? "border-on-error" : ""}`}
                                                        value={row.observationnumberObs}
                                                        onChange={(e) =>
                                                            handleNCNumberChange(index, "observationnumberObs", e.target.value)
                                                        }
                                                        disabled={true}
                                                    />
                                                </td>
                                                <td style={{ minWidth: '70px', maxWidth: '70px' }} title={"Observation"}>
                                                    <select
                                                        id="approvalType"
                                                        value="Observation Number"
                                                        className="newse form-select"
                                                        disabled
                                                    >
                                                        <option value="Observation Number">Observation</option>
                                                    </select>
                                                </td>

                                                <td style={{ minWidth: '220px', maxWidth: '220px' }} title={row.descriptionNC}>
                                                    <textarea
                                                        //onKeyDown={handleKeyDowntextarea}
                                                        id="simpleinput"
                                                        className={`form-control`}
                                                        // className="form-control"
                                                        value={row.descriptionNC}
                                                        onChange={(e) => handleNCDescriptionChange(row.id, 'descriptionNC', e.target.value)}
                                                        disabled={InputDisabled}
                                                    />
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </>
                        // {/* </div> */}
                    )}
                </div>

                {/* {checkboxValues["FailureofIntentNonconformity"] &&
                    <div className="col-lg-6">
                        <div className="mb-3">
                            <label htmlFor="NCNo" className="col-form-label">NC Number</label>
                            <div >
                                <input
                                    type="text"
                                    className={`form-control`}
                                    // className="form-control"
                                    id="NCNo"
                                    value={formData.NCNo}
                                    //onChange={(e) => setFormData({ ...formData, issueNo: e.target.value })}
                                    disabled={true}
                                />
                            </div>
                        </div>
                    </div>
                }
                {checkboxValues["Observations"] &&
                    <div className="col-lg-6">
                        <div className="mb-3">
                            <label htmlFor="ObservationNo" className="col-form-label">Observation Number</label>
                            <div >
                                <input
                                    type="text"
                                    className={`form-control`}
                                    // className="form-control"
                                    id="ObservationNo"
                                    value={formData.ObservationNo}
                                    //onChange={(e) => setFormData({ ...formData, issueNo: e.target.value })}
                                    disabled={true}
                                />
                            </div>
                        </div>
                    </div>
                } */}

            </div>
        );
    };
    // const handleCheckboxChange = (label: CheckboxLabel) => {
    //     debugger
    //     setFormData(prevData => ({
    //         ...prevData,
    //         NCNo: ,
    //         ObservationNo: ,

    //     }));
    //     const columnKey = CheckboxFieldMap[label];
    //     setCheckboxValues((prevValues) => ({
    //         ...prevValues,
    //         [columnKey]: !prevValues[columnKey],
    //     }));
    // };
    // const handleCheckboxChange = (label: CheckboxLabel) => {
    //     const columnKey = CheckboxFieldMap[label];

    //     setCheckboxValues((prevValues) => {
    //         const isChecked = !prevValues[columnKey]; // new checkbox value (toggled)
    //         const updatedValues = {
    //             ...prevValues,
    //             [columnKey]: isChecked,
    //         };

    //         setFormData((prevData) => {
    //             const updatedFormData = { ...prevData };

    //             if (columnKey === "FailureofIntentNonconformity") {
    //                 if (isChecked) {
    //                     const newNCSeq = (prevData.NCSequence || 0) + 1;
    //                     updatedFormData.NCSequence = newNCSeq;
    //                     updatedFormData.NCNo = newNCSeq.toString().padStart(3, "0");
    //                 } else {
    //                     const newNCSeq = Math.max((prevData.NCSequence || 1) - 1, 0);
    //                     updatedFormData.NCSequence = newNCSeq;
    //                     updatedFormData.NCNo = ""; // clear NC number
    //                 }
    //             }

    //             if (columnKey === "Observations") {
    //                 if (isChecked) {
    //                     const newObsSeq = (prevData.ObservationSequence || 0) + 1;
    //                     updatedFormData.ObservationSequence = newObsSeq;
    //                     updatedFormData.ObservationNo = newObsSeq.toString().padStart(3, "0");
    //                 } else {
    //                     const newObsSeq = Math.max((prevData.ObservationSequence || 1) - 1, 0);
    //                     updatedFormData.ObservationSequence = newObsSeq;
    //                     updatedFormData.ObservationNo = ""; // clear observation number
    //                 }
    //             }

    //             return updatedFormData;
    //         });

    //         return updatedValues;
    //     });
    // };
    // const handleCheckboxChange = (label: CheckboxLabel) => {
    //     const columnKey = CheckboxFieldMap[label];

    //     setCheckboxValues((prevValues) => {
    //         const isChecked = !prevValues[columnKey];
    //         const updatedValues = {
    //             ...prevValues,
    //             [columnKey]: isChecked,
    //         };

    //         setFormData((prevData) => {
    //             const updatedFormData = { ...prevData };
    //             console.log("prevDataprevData", prevData, formData)
    //             // NC Logic
    //             if (columnKey === "FailureofIntentNonconformity") {
    //                 const baseNCSeq = formData.NCSequence || 0;
    //                 updatedFormData.NCNo = isChecked
    //                     ? (baseNCSeq + 1).toString().padStart(3, "0")
    //                     : baseNCSeq.toString().padStart(3, "0");

    //             }

    //             // Observation Logic
    //             if (columnKey === "Observations") {
    //                 const baseObsSeq = formData.ObservationSequence || 0;
    //                 updatedFormData.ObservationNo = isChecked
    //                     ? (baseObsSeq + 1).toString().padStart(3, "0")
    //                     : baseObsSeq.toString().padStart(3, "0");

    //             }

    //             return updatedFormData;
    //         });

    //         return updatedValues;
    //     });
    // };
    // const handleCheckboxChange = (label: CheckboxLabel) => {
    //     const columnKey = CheckboxFieldMap[label];

    //     setCheckboxValues((prevValues) => ({
    //         ...prevValues,
    //         [columnKey]: !prevValues[columnKey],
    //     }));
    // };
    const handleCheckboxChange = (label: CheckboxLabel) => {
        const columnKey = CheckboxFieldMap[label];
        console.log("ghgh", formData, maxncseq, maxobsseq);
        debugger
        setCheckboxValues((prevValues) => {
            const isChecked = !prevValues[columnKey];

            // If checkbox is being unchecked, remove corresponding rows
            if (!isChecked) {
                const typeToRemove =
                    columnKey === "Observations"
                        ? "Observation Number"
                        : (columnKey === "FailureofIntentNonconformity" || columnKey === "FailureofImplementation" || columnKey === "FailureofEffectiveness")
                            ? "NC Number"
                            : null;

                if (typeToRemove) {
                    setNCNumberrows((prevRows) =>
                        prevRows.filter((row) => row.nctype !== typeToRemove)
                    );
                }
            }

            return {
                ...prevValues,
                [columnKey]: isChecked,
            };
        });
    };
    const handleNumberChange = (fieldKey: string, value: number) => {
        setCheckboxNumberValues(prev => ({
            ...prev,
            [fieldKey]: value
        }));
    };

    const handleDelete = (index: number) => {
        setFilesArr((prevFiles: any[]) => prevFiles.filter((_file: any, i: number) => i !== index));
    };


    return (
        <div id="wrapper" ref={elementRef}>
            {/* <div
                className="app-menu"
                id="myHeader">
                <VerticalSideBar _context={sp} />
            </div> */}
            <div className="content-page">
                {/* <HorizontalNavbar _context={sp} siteUrl={siteUrl} /> */}
                {/* <div className="content" style={{ marginLeft: `${!useHide ? '0px' : '80px'}`, marginTop: '2.3rem' }}> */}
                <div>
                    <div className="">
                        <div className="row">
                            <div className="col-lg-6">
                                <CustomBreadcrumb Breadcrumb={Breadcrumb} />
                            </div>

                        </div>
                        <div className="row mt-0">

                            {/* <!-- Right Sidebar --> */}
                            <div className="col-12">
                                <div >
                                    <div>
                                        {/* <!-- Left sidebar --> */}


                                        {Loading ?

                                            <div className="loadernewadd mt-10">
                                                <div>
                                                    <img
                                                        src={require("../assets/edc-gif.gif")}
                                                        className="alignrightl"
                                                        alt="Loading..."
                                                    />
                                                </div>
                                                <span>Loading </span>{" "}
                                                <span>
                                                    <img
                                                        src={require("../assets/edcnew.gif")}
                                                        className="alignrightl"
                                                        alt="Loading..."
                                                    />
                                                </span>
                                            </div>
                                            :
                                            <div style={{ width: '100%' }} className="inbox-rightbar">

                                                <div className="card">
                                                    <div className="card-body">
                                                        {/* <h4 className="text-dark font-16 fw-bold mb-3">Memo Details</h4>
 */}
                                                        <div className="previewIcon">
                                                            <h4 style={{ textAlign: 'left' }} className="text-dark font-16 fw-bold mb-3">Requested By</h4>

                                                            {TemplateDoc && TemplateDoc.length > 0 && (
                                                                <span
                                                                    onClick={() => OpenFileTemplate(TemplateDoc[0], "Open")}
                                                                    style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                >
                                                                    <div className="btn btn-primary">
                                                                        <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../assets/noun-download-5006210.png")} ></img></div>
                                                                </span>
                                                            )}
                                                        </div>

                                                        <div style={{ clear: "both" }}></div>

                                                        <form className="form-horizontal">
                                                            <div className="row">
                                                                {console.log("sudit memo no", selectedOption, formData.approvedauditplanId, rows.filter((x: any) => x.value == formData.approvedauditplanId))}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="DocumentCode" className="col-form-label">Approved Memo No./Audit Plan:<span className="text-danger1">*</span>
                                                                        </label>
                                                                        <div
                                                                            title={selectedOption && selectedOption[0]?.label || selectedOption && selectedOption?.label}
                                                                            style={{ width: "100%" }}
                                                                        >
                                                                            <Select
                                                                                //onKeyDown={handleKeyDown}
                                                                                isClearable={true}
                                                                                options={rows}
                                                                                value={selectedOption}
                                                                                name="AuditPlan"

                                                                                isSearchable={true}
                                                                                className={`newse  ${(!ValidSubmit && approvedauditplanerr) ? "border-on-error" : ""} ${(!ValidDraft && approvedauditplanerr) ? "border-on-error" : ""}`}
                                                                                onChange={(selectedOption: any) => onSelectDocCode(selectedOption)}
                                                                                placeholder={"Search Audit plan"}
                                                                                isDisabled={InputDisabled}
                                                                            />
                                                                        </div>

                                                                    </div>
                                                                </div>
                                                                {console.log("from deptitplandate", currentUserDept, formData.fromdeptId, AllDept.filter((x: any) => x.value == formData.fromdeptId))}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="Department" className="col-form-label">From Department<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <div
                                                                                title={currentUserDept && currentUserDept[0]?.label || currentUserDept && currentUserDept?.label}
                                                                                style={{ width: "100%" }}
                                                                            >
                                                                                <Select
                                                                                    //onKeyDown={handleKeyDown}
                                                                                    isClearable={true}
                                                                                    options={AllDept}
                                                                                    isDisabled={InputDisabled}
                                                                                    value={currentUserDept}
                                                                                    name="deptId"
                                                                                    //className={`newse`}
                                                                                    className={`newse  ${(!ValidSubmit && fromdepartmenterr) ? "border-on-error" : ""} ${(!ValidDraft && fromdepartmenterr) ? "border-on-error" : ""}`}
                                                                                    // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                    // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                    // onChange={handleDepartmentChange}
                                                                                    onChange={(selectedOptions: any) => handleFromDepartmentChange(selectedOptions)}
                                                                                    placeholder="Select Department"
                                                                                />
                                                                            </div>

                                                                        </div>
                                                                    </div>
                                                                </div>


                                                                {console.log("doccodedoccodedoccode", formData, doccode, reportCode)}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Report Code<span className="text-danger1"> *</span></label>
                                                                        <div title={formData.reportCode || "Select a report code"}>
                                                                            <input
                                                                                //onKeyDown={handleKeyDowntext}
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control`}

                                                                                // className="form-control"
                                                                                id="documentCode"
                                                                                value={formData.reportCode}
                                                                            //onChange={(e) => setFormData({ ...formData, documentCode: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="memoNo" className="col-form-label">Document Code<span className="text-danger1"> *</span></label>
                                                                        <div className="" title={formData.documentCode || "Select a document code"}>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}

                                                                                // className="form-control"
                                                                                id="docCode"
                                                                                value={formData.documentCode}

                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="issueNo" className="col-form-label">Issue No<span className="text-danger1"> *</span></label>
                                                                        <div title={formData.issueNo || "Select a issue no"}>
                                                                            <input

                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issueNo"
                                                                                value={formData.issueNo}
                                                                                //onChange={(e) => setFormData({ ...formData, issueNo: e.target.value })}
                                                                                disabled={true}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {/* <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="issuedate" className="col-form-label">Issue Date<span className="text-danger1"> *</span></label>
                                                                        <div>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="issuedate"
                                                                                value={formData.issueDate == null ? "" : moment(formData.issueDate).format("MM/DD/YYYY")}
                                                                            //onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>*/}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisionNo" className="col-form-label">Revision No<span className="text-danger1"> *</span></label>
                                                                        <div title={formData.revisionNo}>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="revisionNo"
                                                                                value={formData.revisionNo}
                                                                            //onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {/*}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisiondate" className="col-form-label">Revision Date<span className="text-danger1"> *</span></label>
                                                                        <div>
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="revisionNo"
                                                                                value={formData.revisionDate == null ? "" : moment(formData.revisionDate).format("MM/DD/YYYY")}
                                                                            //onChange={(e) => setFormData({ ...formData, revisionNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div> 
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisionNo" className="col-form-label">Reference No<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <input
                                                                                disabled
                                                                                type="text"
                                                                                className={`form-control ${(!ValidSubmit) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="referenceNo"
                                                                                value={formData.referenceNo}
                                                                                onChange={(e) => setFormData({ ...formData, referenceNo: e.target.value })}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>*/}

                                                                {console.log("documentlinkkkkkkk", DocumentLink, FilesArrDoclink)}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="DocumentCode" className="form-label">Audit Plan Document:</label>
                                                                        {/* <input type="text" id="example-email" name="example-email" className="form-control" placeholder="Search Document Code" value={formData.DocumentCode} /> */}

                                                                        {/* <div className="text-dark mt-0"> <span >
                                                                            <a onClick={() => setShowModal(true)} ><FontAwesomeIcon icon={faPaperclip} />{DocumentLink && "1 file Attached"}</a>

                                                                        </span>
                                                                        </div> */}
                                                                        {/* {TemplateDocAudit && TemplateDocAudit.length > 0 && (
                                                                            <span
                                                                                onClick={() => OpenFile(TemplateDocAudit[0], "Open")}
                                                                                style={{ color: "blue", cursor: "pointer", margin: "10px" }}
                                                                            >
                                                                                <div className="btn btn-primary">
                                                                                    <img style={{ cursor: 'pointer' }} className='mt-0' src={require("../assets/noun-download-5006210.png")} ></img></div>
                                                                            </span>
                                                                        )} */}
                                                                        {TemplateDocAudit && TemplateDocAudit.length > 0 ?
                                                                            (<a style={{ fontSize: '0.875rem' }} onClick={() => setShowModalpre(true)}>
                                                                                <FontAwesomeIcon icon={faPaperclip} />{TemplateDocAudit.length} {TemplateDocAudit.length > 0 ? "files" : "file"} Attached
                                                                            </a>) : ""

                                                                        }
                                                                    </div>
                                                                </div>
                                                                {console.log("auditplandate", formData)}
                                                                <div className="col-lg-4">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="date" className="col-form-label">Audit Plan Date<span className="text-danger1"> *</span></label>
                                                                        <div title={moment(auditplandate).format('DD/MMM/YYYY')}>
                                                                            {/* <input

                                                                                type="date"
                                                                                className={`form-control`}
                                                                                // className="form-control"
                                                                                id="date"
                                                                                //formatDate={(date: any) => moment(date).format('DD/MMM/YYYY')}
                                                                                value={moment(auditplandate).format('DD/MMM/YYYY')}
                                                                                //onChange={(e) => setFormData({ ...formData, Auditplandate: new Date(e.target.value).toLocaleDateString("en-CA") })}
                                                                                disabled={true}
                                                                            /> */}
                                                                            <DatePicker id="date"
                                                                                value={
                                                                                    auditplandate
                                                                                        ? new Date(moment(auditplandate).format('YYYY-MM-DD'))
                                                                                        : null
                                                                                }
                                                                                // onSelectDate={(date: Date | null) => {
                                                                                //     handleActualAuditDateChange(date)
                                                                                // }}
                                                                                //className={`${(!ValidSubmit && dateerr) ? "textfield-error" : ""}${(!ValidDraft && dateerr) ? "textfield-error" : ""}`}

                                                                                minDate={new Date()}
                                                                                disabled={true}
                                                                                formatDate={(date: any) => moment(date).format('DD/MMM/YYYY')}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {console.log("deoartment auditeddd", selectUserDept, formData.deptId, AllDept.filter((x: any) => x.value == formData.deptId))}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="Department" className="col-form-label">Department Audited<span className="text-danger1"> *</span></label>
                                                                        <div >
                                                                            <div
                                                                                title={selectUserDept && selectUserDept[0]?.label || selectUserDept && selectUserDept?.label}
                                                                                style={{ width: "100%" }}
                                                                            >
                                                                                <Select
                                                                                    //onKeyDown={handleKeyDown}
                                                                                    isClearable={true}
                                                                                    options={AllDept}
                                                                                    isDisabled={InputDisabled}
                                                                                    value={selectUserDept}
                                                                                    name="deptId"
                                                                                    className={`newse  ${(!ValidSubmit && departmenterr) ? "border-on-error" : ""} ${(!ValidDraft && departmenterr) ? "border-on-error" : ""}`}
                                                                                    // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                    // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                    // onChange={handleDepartmentChange}
                                                                                    onChange={(selectedOptions: any) => handleDepartmentChange(selectedOptions)}
                                                                                    placeholder="Select Department"
                                                                                />
                                                                            </div>

                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="row mb-3">
                                                                        <label htmlFor="date" className="col-form-label">Actual Audit Date<span className="text-danger1"> *</span></label>
                                                                        <div title={moment(formData?.date).format('DD/MMM/YYYY')}>
                                                                            {/* <input
                                                                                type="date"
                                                                                className={`form-control ${(!ValidSubmit && dateerr) ? "border-on-error" : ""}${(!ValidDraft && dateerr) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                id="date"
                                                                                value={formData.date}
                                                                                onChange={(e) => handleActualAuditDateChange(e)}
                                                                                disabled={InputDisabled}
                                                                            /> */}
                                                                            <DatePicker id="date"
                                                                                value={
                                                                                    formData?.date
                                                                                        ? new Date(moment(formData?.date).format('YYYY-MM-DD'))
                                                                                        : null
                                                                                }
                                                                                onSelectDate={(date: Date | null) => {
                                                                                    handleActualAuditDateChange(date)
                                                                                    //setFormData({ ...formData, date: new Date(date).toLocaleDateString("en-CA") });
                                                                                    // setFormData({ ...formData, date: moment(date).format('DD/MMM/YYYY') });
                                                                                }}
                                                                                className={`${(!ValidSubmit && dateerr) ? "textfield-error" : ""}${(!ValidDraft && dateerr) ? "textfield-error" : ""}`}
                                                                                //className={`form-control ${(!ValidSubmit && dateerr) ? "textfield-error" : ""}${(!ValidDraft && dateerr) ? "textfield-error" : ""}`}
                                                                                // maxDate={new Date()}
                                                                                minDate={new Date()}
                                                                                disabled={InputDisabled}
                                                                                formatDate={(date: any) => moment(date).format('DD/MMM/YYYY')}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <div className="d-flex justify-content-between align-items-center">
                                                                            <label htmlFor="attachment" className="col-form-label">
                                                                                Audit Report Attachment<span className="text-danger1"> *</span>
                                                                            </label>

                                                                            {/* File count on right */}
                                                                            {FilesArr.length > 0 && (
                                                                                <a
                                                                                    style={{ fontSize: '0.875rem', cursor: 'pointer' }}
                                                                                    onClick={() => setShowModal(true)}
                                                                                >
                                                                                    <FontAwesomeIcon icon={faPaperclip} /> {FilesArr.length} {FilesArr.length > 1 ? "files" : "file"} attached
                                                                                </a>
                                                                            )}
                                                                        </div>

                                                                        <input
                                                                            type="file"
                                                                            className={`form-control ${(!ValidSubmit && attachmenterr) ? "border-on-error" : ""}`}
                                                                            id="attachment"
                                                                            accept=".jpg,.jpeg,.png,.gif,.bmp,.svg,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx"
                                                                            onChange={(e) => onFileChange(e, "Gallery", "AnnualAuditReportDocs")}
                                                                            disabled={InputDisabled}
                                                                            multiple
                                                                        />
                                                                    </div>

                                                                </div>
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="revisionNo" className="col-form-label">Share With<span className="text-danger1"> *</span></label>
                                                                        <div title={sharewithusers.map(user => user.label).join(', ')}>
                                                                            <Select
                                                                                //onKeyDown={handleKeyDown}
                                                                                isClearable={true}
                                                                                options={rows1}
                                                                                isMulti
                                                                                value={sharewithusers}
                                                                                name="share with"
                                                                                className={`newse ${(!ValidSubmit && sharewitherr) ? "border-on-error" : ""}`}
                                                                                // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                onChange={(selectedOptions: any) => onSelectsharewith(selectedOptions)}
                                                                                placeholder="Enter Share with"
                                                                                isDisabled={InputDisabled}
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {console.log("shifttttt ", selectshift, formData.shiftId, AuditProgShift.filter((x: any) => x.value == formData.shiftId))}
                                                                <div className="col-lg-4">
                                                                    <div className="mb-3">
                                                                        <label htmlFor="Department" className="col-form-label">Shift</label>
                                                                        <div >
                                                                            <div
                                                                                title={selectshift && selectshift?.label || selectshift && selectshift[0]?.label}
                                                                                style={{ width: "100%" }}
                                                                            >
                                                                                <Select
                                                                                    //onKeyDown={handleKeyDown}
                                                                                    isClearable={true}
                                                                                    options={AuditProgShift}
                                                                                    isDisabled={InputDisabled}
                                                                                    value={selectshift}
                                                                                    name="shiftId"
                                                                                    className={`newse`}
                                                                                    // onChange={(selectedOptions: any) => handleCCChange(selectedOptions, 'CC')}
                                                                                    // onChange={(e: any) => setFormData({ ...formData, deptId: e.value })}
                                                                                    // onChange={handleDepartmentChange}
                                                                                    onChange={(selectedOptions: any) => handleShiftChange(selectedOptions)}
                                                                                    placeholder="Select Shift"
                                                                                />
                                                                            </div>

                                                                        </div>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </form>
                                                    </div>
                                                </div>
                                                <section className='card card-body mt-2'>
                                                    <fieldset>
                                                        <div className='row'>
                                                            <div className='col-sm-6'>
                                                                <h3 className='text-dark font-16 fw-bold mb-3'>Summary of the Audit Findings</h3>
                                                            </div>
                                                        </div>

                                                        <div style={{ display: 'grid' }}>
                                                            <div className="row newlavl"> {renderCheckboxes()}</div>
                                                        </div>
                                                    </fieldset>
                                                </section>

                                                <section className='card card-body mt-2'>
                                                    <fieldset style={{ display: 'grid' }}>
                                                        <div className='row'>
                                                            <div className='col-sm-6'>
                                                                <h3 className='text-dark font-16 fw-bold mb-3'>Checklist</h3>
                                                            </div>
                                                            <div style={{ textAlign: 'right' }} className='col-sm-6'>
                                                                {!InputDisabled && <img style={{ width: '30px', cursor: 'pointer' }} className='mt-0' src={require("../assets/plus.png")} onClick={handleAddRecommendationRow}></img>}

                                                            </div>
                                                        </div>

                                                        <div style={{ display: 'grid', overflow: 'auto' }}>
                                                            <table id="tabRec" className='mtbalenewscrollnew4'>
                                                                <thead>
                                                                    <tr>
                                                                        <th style={{ minWidth: '200px', maxWidth: '200px' }} >ISO reference<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '220px', maxWidth: '220px' }} >IMS procedure<span className="text-danger1"> *</span></th>
                                                                        <th colSpan={2} style={{ minWidth: '300px', maxWidth: '300px' }} >Inquiries<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '300px', maxWidth: '300px' }} >Auditor’s
                                                                            Comments<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '160px', maxWidth: '160px' }} >Time<span className="text-danger1"> *</span></th>
                                                                        {/* <th>Share with<span className="text-danger1"> *</span></th> */}

                                                                        {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <th style={{ minWidth: '70px', maxWidth: '70px' }} >Action</th>}
                                                                    </tr>
                                                                </thead>

                                                                <tbody>
                                                                    {console.log("ValidDRecommValidDRecomm", ValidDRecomm)}
                                                                    {recommendationRows.map((row, index) => (
                                                                        <tr key={row.id}>
                                                                            <td title={row?.isoreference ? row?.isoreference : row?.isoreference} style={{ minWidth: '200px', maxWidth: '200px' }}>
                                                                                <input
                                                                                    //onKeyDown={handleKeyDowntext}
                                                                                    type="text"
                                                                                    className={`form-control ${(RowErrors[index]?.isoreference) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.isoreference}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'isoreference', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            {/* <td>
                                                                            <input
                                                                                type="text"
                                                                                className={`form-control ${(!ValidDRecomm && imsprocedureerr) ? "border-on-error" : ""}`}
                                                                                // className="form-control"
                                                                                value={row.imsprocedure}
                                                                                onChange={(e) => handleRecommendationChange(index, 'imsprocedure', e.target.value)}
                                                                                disabled={InputDisabled}
                                                                            />
                                                                        </td> */}
                                                                            <td title={row?.imsprocedure ? row?.imsprocedure : row?.imsprocedure} style={{ minWidth: '220px', maxWidth: '220px' }}>
                                                                                <input
                                                                                    //onKeyDown={handleKeyDowntext}
                                                                                    type="text"
                                                                                    className={`form-control ${RowErrors[index]?.imsprocedure ? "border-on-error" : ""}`}
                                                                                    value={row.imsprocedure}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'imsprocedure', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            <td title={row?.inquiries ? row?.inquiries : row?.inquiries} style={{ minWidth: '300px', maxWidth: '300px' }}>
                                                                                <textarea
                                                                                    //onKeyDown={handleKeyDowntextarea}
                                                                                    id="simpleinput"
                                                                                    className={`form-control ${(RowErrors[index]?.inquiries) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.inquiries}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'inquiries', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                                {/* <input
                                                                                    type="text"
                                                                                    className={`form-control ${(RowErrors[index]?.inquiries) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.inquiries}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'inquiries', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                /> */}
                                                                            </td>
                                                                            <td title={row?.auditorcomments ? row?.auditorcomments : row?.auditorcomments} style={{ minWidth: '300px', maxWidth: '300px' }}>
                                                                                <textarea
                                                                                    // onKeyDown={handleKeyDowntextarea}
                                                                                    id="simpleinput"
                                                                                    className={`form-control ${(RowErrors[index]?.auditorcomments) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.auditorcomments}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'auditorcomments', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            <td style={{ minWidth: '160px', maxWidth: '160px' }} id="timepicker"
                                                                                title={row?.time ? new Date(`1970-01-01T${row.time}`)
                                                                                    .toLocaleTimeString([], {
                                                                                        hour: '2-digit',
                                                                                        minute: '2-digit',
                                                                                        hour12: true
                                                                                    })
                                                                                    .replace(/am|pm/i, match => match.toUpperCase()) : "Select Time"}>
                                                                                {/* <input
                                                                                    type="time"
                                                                                    className={`form-control ${(RowErrors[index]?.time) ? "border-on-error" : ""}`}
                                                                                    // className="form-control"
                                                                                    value={row.time}
                                                                                    onChange={(e) => handleRecommendationChange(index, 'time', e.target.value)}
                                                                                    disabled={InputDisabled}
                                                                                /> */}
                                                                                {/* <TimePicker
                                                                                    className={`${RowErrors[index]?.time ? "border-on-error" : ""}`}
                                                                                    value={row.time}
                                                                                    onChange={(value:any) => handleRecommendationChange(index, 'time', value)}
                                                                                    disableClock={true}
                                                                                    format="HH:mm"
                                                                                    clearIcon={null}
                                                                                    disabled={InputDisabled}
                                                                                /> */}
                                                                                <LocalizationProvider dateAdapter={AdapterDateFns}>
                                                                                    <TimePicker
                                                                                        label="Select Time"
                                                                                        className={`form-control recommendClsErr `}
                                                                                        value={row.time ? new Date(`1970-01-01T${row.time}`) : null}
                                                                                        onChange={(newValue) => {
                                                                                            const formattedTime = newValue
                                                                                                ? newValue.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false })
                                                                                                : '';
                                                                                            handleRecommendationChange(index, 'time', formattedTime);
                                                                                        }}
                                                                                        disabled={InputDisabled}
                                                                                        slotProps={{
                                                                                            textField: {
                                                                                                fullWidth: true,
                                                                                                className: `form-control ${RowErrors[index]?.time ? 'ErrBorder-ErrColor' : ''}`,
                                                                                                sx: {
                                                                                                    '& .MuiInputBase-input.Mui-disabled': {
                                                                                                        color: 'black',
                                                                                                    },
                                                                                                    '& .Mui-disabled': {
                                                                                                        WebkitTextFillColor: 'black', // Safari
                                                                                                    },
                                                                                                    '& .MuiSvgIcon-root': {
                                                                                                        color: 'black !important', // Force icon color to black
                                                                                                    },

                                                                                                }
                                                                                            },
                                                                                            openPickerIcon: {
                                                                                                sx: {
                                                                                                    color: 'black !important',
                                                                                                    opacity: 1,
                                                                                                },
                                                                                            },
                                                                                        }}

                                                                                    />
                                                                                </LocalizationProvider>
                                                                                {/* <TimePickerComponent
                                                                                    value={row.time ? new Date(`1970-01-01T${row.time}`) : null}
                                                                                    change={(e) => {
                                                                                        const time = e?.value
                                                                                            ? e.value.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false })
                                                                                            : '';
                                                                                        handleRecommendationChange(index, 'time', time);
                                                                                    }}
                                                                                    format="HH:mm"
                                                                                    cssClass={RowErrors[index]?.time ? 'e-error' : ''}
                                                                                    enabled={!InputDisabled}
                                                                                /> */}
                                                                                {/* <TimePicker
                                                                                    value={row.time ? dayjs(row.time, 'HH:mm') : null}
                                                                                    onChange={(time) => {
                                                                                        const formattedTime = time ? time.format('HH:mm') : '';
                                                                                        handleRecommendationChange(index, 'time', formattedTime);
                                                                                    }}
                                                                                    format="HH:mm"
                                                                                    disabled={InputDisabled}
                                                                                    className={`form-control ${RowErrors[index]?.time ? 'border-on-error' : ''}`}
                                                                                    allowClear={false}
                                                                                /> */}
                                                                            </td>
                                                                            {/* <td title={row?.sharewith ? row?.sharewith : row?.sharewith}>
                                                                                <Select
                                                                                    options={rows1}
                                                                                    // isMulti
                                                                                    className={`${(RowErrors[index]?.sharewith) ? "border-on-error" : ""}`}
                                                                                    value={row.sharewith}
                                                                                    onChange={(selectedOptions: any) => handleRecommendationChange(index, 'sharewith', selectedOptions)}
                                                                                    placeholder="Select"
                                                                                    isDisabled={InputDisabled}
                                                                                />
                                                                            </td> */}
                                                                            {(modeValue === "" || modeValue === "edit" || InputDisabled != true) && <td style={{ minWidth: '70px', maxWidth: '70px' }} >
                                                                                <img src={require("../assets/del.png")} onClick={() => handleDeleteRecommendationRow(index)} />

                                                                            </td>
                                                                            }
                                                                        </tr>
                                                                    ))}
                                                                </tbody>

                                                            </table>
                                                        </div>
                                                    </fieldset>
                                                </section>

                                                {/* {modeValue === "approve" && editID != null && editID.Status === "Pending" && editID.CurrentUserRole !== "Initiator" && */}

                                                <div className="card mt-2" style={{ marginBottom: '17px' }}>
                                                    <div className="card-body">
                                                        <div className='row'>
                                                            <div className='col-sm-8'>
                                                                <h4 className="text-dark font-16 fw-bold mb-3 ">Approval Hierarchy</h4>
                                                                <label>Define the approval hierarchy to ensure requests are routed to the appropriate approvers.
                                                                </label>
                                                            </div>
                                                            <div className='col-sm-4'>
                                                                <div className="mt-0 mb-0 float-end text-right" style={{ textAlign: "right", paddingRight: "22px" }}>
                                                                    {!InputDisabled &&
                                                                        <img style={{ width: '34px' }} src={require("../assets/plus.png")} onClick={handleAddRow} className='' />
                                                                    }

                                                                </div>
                                                            </div>

                                                        </div>

                                                        <div style={{ overflow: 'inherit' }} className="table-responsive mt-3 pt-0">
                                                            <table style={{ overflow: 'inherit' }} className="mtbalenew  table-centered table-nowrap table-borderless mb-0 overhi" id="myTabl">
                                                                <thead >
                                                                    <tr>
                                                                        <th style={{ minWidth: "30px", maxWidth: "30px" }}>S.No</th>
                                                                        <th style={{ borderBottomLeftRadius: "0px", minWidth: '80px', maxWidth: '80px', }}>Role<span className="text-danger1"> *</span></th>
                                                                        <th style={{ borderBottomLeftRadius: "0px", minWidth: '86px', maxWidth: '86px' }}>Responsibility<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: "46px", maxWidth: "46px" }} title="Use your electronic digital signature to sign this document digitally">E-Sign?</th>
                                                                        <th style={{ minWidth: '40px', maxWidth: '40px' }} >Level</th>
                                                                        <th>Approver name<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '70px', maxWidth: '70px' }} >Approval criteria<span className="text-danger1"> *</span></th>
                                                                        <th style={{ minWidth: '40px', maxWidth: '40px' }}>Action</th>
                                                                    </tr>
                                                                </thead>
                                                                <tbody style={{ maxHeight: "8000000007px", overflow: 'inherit' }}>
                                                                    {forwardToArr.map((row, index) => (
                                                                        <tr>
                                                                            <td style={{ minWidth: "30px", maxWidth: "30px", overflow: 'inherit' }}> <div
                                                                                style={{ marginLeft: "5px" }}
                                                                                className="indexdesign"
                                                                            >
                                                                                {index + 1}</div>
                                                                            </td>
                                                                            <td title={UserRoles.find((opt: any) => opt.value === row.role)?.label || "Select role"} style={{ overflow: 'inherit', minWidth: '80px', maxWidth: '80px', }} className="ng-binding">
                                                                                <select
                                                                                    // className="form-select"
                                                                                    //className={`form-select newse ${(!ValidForwardTo) ? "border-on-error" : ""} `}
                                                                                    className={`form-select newse ${forwardToValidationErrors[row.level]?.role ? "border-on-error" : ""}`}
                                                                                    onChange={(e) => onSelectRole(e, row.level)} value={row.role} disabled={InputDisabled}>

                                                                                    <option value="" selected>Select Role</option>
                                                                                    {/* {UserRoles.map((role: any, index: number) => (
                                                                                    <option key={index} value={role.value}>{role.label}</option>
                                                                                ))} */}
                                                                                    {UserRoles.filter((role: any) =>
                                                                                        !forwardToArr.some((r) => r.role === role.value && r.level !== row.level) || role.value === row.role // Allow the current row's role
                                                                                    ).map((role: any, idx: number) => (
                                                                                        <option key={idx} value={role.value}>{role.label}</option>
                                                                                    ))}



                                                                                </select>

                                                                            </td>
                                                                            <td style={{ overflow: 'inherit', minWidth: '86px', maxWidth: '86px' }}>
                                                                                <div
                                                                                //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                                                                >
                                                                                    <select
                                                                                        id="responsibleId"
                                                                                        value={row.Responsibility}
                                                                                        onChange={(e) => handleChangeResp(e, row.level)}
                                                                                        className={`form-select newse ${forwardToValidationErrors[row.level]?.role ? "border-on-error" : ""}`}
                                                                                        disabled={InputDisabled}
                                                                                        title={row.Responsibility ? row.Responsibility : "Select"}
                                                                                    >
                                                                                        <option value="">Select </option>
                                                                                        <option value="Signer">Signer</option>
                                                                                        <option value="Reviewer">Reviewer</option>


                                                                                    </select>

                                                                                </div>
                                                                            </td>
                                                                            <td style={{ minWidth: "46px", maxWidth: "46px", overflow: 'inherit' }}>
                                                                                <div
                                                                                //  style={{ display: "flex", alignItems: 'center', gap: '8px' }}
                                                                                >
                                                                                    <input
                                                                                        type="checkbox"
                                                                                        checked={row.IsSignatureRequired}
                                                                                        disabled={row.Responsibility === "Signer" || row.Responsibility === "" || InputDisabled}
                                                                                        style={{ marginLeft: '17px', width: "15px" }}

                                                                                        onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                                                                                            const isChecked = e.target.checked;
                                                                                            const updatedArr = forwardToArr.map(row1 =>
                                                                                                row1.level === row.level ? { ...row1, IsSignatureRequired: isChecked } : row1
                                                                                            );
                                                                                            setForwardToArr(updatedArr);
                                                                                        }}
                                                                                    />

                                                                                </div>
                                                                            </td>
                                                                            <td style={{ minWidth: '40px', maxWidth: '40px', overflow: 'inherit' }}>Level {index + 1}</td>
                                                                            <td title={row.approvers.map(user => user.label).join(', ') || "Select Approvers"} style={{ overflow: 'inherit' }}>

                                                                                <Select
                                                                                    //onKeyDown={handleKeyDown}
                                                                                    isClearable={true}
                                                                                    options={rows1}
                                                                                    isMulti
                                                                                    value={row.approvers}
                                                                                    name="Approvers"
                                                                                    //className={`newse ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                    className={`newse ${forwardToValidationErrors[row.level]?.approvers ? "border-on-error" : ""}`}
                                                                                    // onChange={(selectedOption: any) => onSelect(selectedOption)}
                                                                                    onChange={(selectedOptions: any) => onSelectApprovers(selectedOptions, row.level)}
                                                                                    placeholder="Enter Approver Name"
                                                                                    isDisabled={InputDisabled}
                                                                                />
                                                                            </td>
                                                                            <td title={row.approvalType === "All" ? "Everyone" : row.approvalType === "One" ? "Anyone" : "Select"} style={{ overflow: 'inherit', minWidth: '70px', maxWidth: '70px', }}>
                                                                                {/* <label htmlFor="approvalType">Approval Type: </label> */}
                                                                                <select id="approvalType"
                                                                                    value={row.approvalType}
                                                                                    onChange={(e) => handleChange(e, row.level)}
                                                                                    //className={`newse form-select ${(!ValidForwardTo) ? "border-on-error" : ""}`}
                                                                                    className={`form-select newse ${forwardToValidationErrors[row.level]?.approvalType ? "border-on-error" : ""}`}
                                                                                    disabled={InputDisabled} >
                                                                                    <option value="">Select </option>
                                                                                    <option value="One">Anyone</option>
                                                                                    <option value="All">Everyone</option>
                                                                                </select>
                                                                            </td>
                                                                            <td style={{ minWidth: '40px', maxWidth: '40px', overflow: 'inherit' }}>

                                                                                {/* {editID.CurrentUserRole === "OES" ?  */}

                                                                                {!InputDisabled ? <img src={require("../assets/del.png")} onClick={() => handleDeleteRow(index)} /> : <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' />}
                                                                                {/* : <img src={require("../assets/recycle-bin.png")} className='sidebariconsmall' /> */}
                                                                                {/* } */}

                                                                            </td>
                                                                        </tr>

                                                                    ))}


                                                                </tbody>
                                                            </table>
                                                        </div>



                                                        {/* {editID.CurrentUserRole === "OES" &&  */}
                                                        {/* <div className="row mt-3">
                                                            <div className="col-12 text-center">
                                                                <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardApproval("Forward")} >
                                                                    <i className="fe-check-circle me-1"></i> Forward
                                                                </button>

                                                                <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={() => ForwardApproval("Rework")} >
                                                                    <i className="fe-corner-up-left me-1"></i> Rework
                                                                </button>

                                                                <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={() => ForwardApproval("Rejected")} >
                                                                    <i className="fe-x me-1"></i> Reject
                                                                </button>

                                                                <button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}>
                                                                    <i className="fe-x me-1"></i> Cancel
                                                                </button>

                                                            </div>
                                                        </div> */}
                                                        {/* } */}
                                                    </div>
                                                </div>
                                                {/* // } */}
                                                {formData.Status == "Rework" && formData.submitStatus == "No" && editID.CurrentUserRole == "Initiator" &&
                                                    <div className="card">

                                                        <div className="card-body">

                                                            <div className="row">
                                                                <div className="col-lg-12">

                                                                    <div className="mb-0" >

                                                                        <label htmlFor="example-textarea" className="form-label text-dark font-14">Remarks <span className="text-danger1"> *</span></label>

                                                                        <textarea style={{ height: '80px' }} className={`form-control ${(!ValidRemark) ? "border-on-error" : ""}`} id="example-textarea" rows={5} name="RemarksInitiator" value={formData.RemarksInitiator}

                                                                            onChange={(e) => onChange(e.target.name, e.target.value)}></textarea>

                                                                    </div>

                                                                </div>
                                                            </div>

                                                        </div>
                                                    </div>
                                                }

                                                {/* ////////////Approval card */}

                                                {
                                                    (InputDisabled && editID != null && modeValue === "approve" && editID.ApprovalType === "Approval" && editID.Status === "Pending") ? (
                                                        <WorkflowAction currentItem={editID} ctx={props.context} ContentType={CONTENTTYPE_AuditReport}
                                                            DisableApproval={false} DisableCancel={false}
                                                        />
                                                    ) : (<div></div>)
                                                }

                                                {/* ////////////Audit History card */}
                                                {/* {editID !== null && editID.length != 0 && modeValue === "approve" && */}
                                                {MainEditItem !== null && MainEditItem.length != 0 && MainEditItem?.Status !== "Save as draft" &&
                                                    <WorkflowAuditHistory ContentItemId={MainEditItem} ContentType={CONTENTTYPE_AuditReportNew} ctx={props.context} />
                                                }
                                                {/* ////////////Audit History card */}


                                                {/* /////////////////%%%%%%%%%%%%%%%%%%%%%%%% */}

                                                <div className="row mt-3">
                                                    <div className="col-12 text-center">


                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") ||
                                                            (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <div className="btn btn-primary waves-effect waves-light m-1" onClick={handleSaveAsDraft}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Save As Draft</div>
                                                        }

                                                        {((InputDisabled != true && editItemID == null && MainEditItem == null) || (modeValue === "" || modeValue === "edit") ||
                                                            (editID != null && editID.Level === 0 && editID.CurrentUserRole == "Initiator" && editID.IsInitiator == "Yes" && (editID?.Status === "Pending" || editID?.Status === "Save as draft"))) &&
                                                            <div className="btn btn-primary waves-effect waves-light m-1" onClick={handleFormSubmit}>
                                                                <img src={require('../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" />
                                                                Submit</div>
                                                        }

                                                        {/* {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Save as draft")}>  <img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Save As Draft</button>}


                                                        {((editID?.Status === "Pending" || editID?.Status === "Save as draft") && (editID.Level === 0 && editID.CurrentUserRole !== "OES" && editID.IsInitiator == "Yes")) && (modeValue === "approve") && <button type="button" className="btn btn-primary waves-effect waves-light m-1" onClick={() => ForwardInitiatorApproval("Approved")}><img src={require('../../../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} className='me-1' alt="Check" /> Submit</button>}
 */}

                                                        {console.log("eddddd", editID)}
                                                        {((modeValue === "" || modeValue === "edit" || modeValue === "view") || (editID !== null && editID?.ApprovalType !== "Approval") || (modeValue === "approve" && editID?.Status == "Approved")) &&
                                                            <div className="btn cancel-btn waves-effect waves-light m-1" onClick={handleCancel}> <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                                                className='me-1' alt="x" /> Cancel</div>
                                                        }

                                                    </div>
                                                </div>


                                                {/* ////////////Approval card */}


                                                {/* </div> */}

                                                {/* /////////// */}

                                                <Modal show={showModal} onHide={() => setShowModal(false)} size={Showfile ? "xl" : "lg"} className='filemodal'>
                                                    <Modal.Header closeButton>
                                                        <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-0'>Attachment Details Audit report attachment</h4>
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for IMS Audit Report and Checklist
                                                            </p>

                                                        </Modal.Title>


                                                    </Modal.Header>
                                                    <Modal.Body className="" id="style-5">

                                                        {/* {DocumentLink &&
                                                            (
                                                                <> */}
                                                        {Showfile ?

                                                            <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                                            :
                                                            <table className="mtbalenew">
                                                                <thead style={{ background: '#eef6f7' }}>
                                                                    <tr>
                                                                        <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                                                        <th style={{ minWidth: '200px', maxWidth: '200px' }}>File Name</th>

                                                                        {/* <th>File Link</th> */}
                                                                        <th style={{ minWidth: '60px', maxWidth: '60px' }} className='text-center'>Upload date</th>
                                                                        {
                                                                            ((
                                                                                (
                                                                                    (
                                                                                        modeValue !== null &&
                                                                                        modeValue !== "" &&
                                                                                        (modeValue === "edit" || modeValue === "view" || modeValue === "approve")
                                                                                    ) ||
                                                                                    (modeValue === "approve" && formData?.Status === "Rework")
                                                                                ) && showviewdownload
                                                                            ) ||
                                                                                (
                                                                                    (
                                                                                        modeValue === "edit" ||
                                                                                        modeValue === null ||
                                                                                        modeValue === "" ||
                                                                                        (modeValue === "approve" && formData?.Status === "Rework") ||
                                                                                        (modeValue == "approve" && editID?.IsRework == "Yes")
                                                                                    )
                                                                                )) &&
                                                                            <th style={{ minWidth: '50px', maxWidth: '50px' }}>Action</th>

                                                                        }

                                                                        {/* {(modeValue == "edit" || modeValue == null || modeValue == ""
                                                                        || (modeValue == "approve" && formData?.Status == "Rework")) &&
                                                                        <th style={{ textAlign: 'center' }}> Action </th>
                                                                    } */}
                                                                    </tr>
                                                                </thead>
                                                                <tbody>
                                                                    {console.log("filearrrr", FilesArr)}
                                                                    {FilesArr.length > 0 && (
                                                                        FilesArr.map((row: any, index: number) => (
                                                                            <tr>
                                                                                <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>
                                                                                {/* <td title={row.name || row.FileLeafRef}>{row.name || row.FileLeafRef}</td> */}
                                                                                <td style={{ minWidth: '200px', maxWidth: '200px' }} title={cleanFileName(row.name) || cleanFileName(row.FileLeafRef)}>
                                                                                    {cleanFileName(row.name) || cleanFileName(row.FileLeafRef)}</td>
                                                                                <td style={{ minWidth: '50px', maxWidth: '50px' }} title={row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                    day: "2-digit",
                                                                                    month: "short",
                                                                                    year: "numeric"
                                                                                }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                    day: "2-digit",
                                                                                    month: "short",
                                                                                    year: "numeric"
                                                                                }).replace(/ /g, "/")}>

                                                                                    {row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                        day: "2-digit",
                                                                                        month: "short",
                                                                                        year: "numeric"
                                                                                    }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                        day: "2-digit",
                                                                                        month: "short",
                                                                                        year: "numeric"
                                                                                    }).replace(/ /g, "/")}</td>

                                                                                {((
                                                                                    (
                                                                                        (
                                                                                            modeValue !== null &&
                                                                                            modeValue !== "" &&
                                                                                            (modeValue === "edit" || modeValue === "view" || modeValue === "approve")
                                                                                        ) ||
                                                                                        (modeValue === "approve" && formData?.Status === "Rework")
                                                                                    ) && showviewdownload
                                                                                ) ||
                                                                                    (
                                                                                        (
                                                                                            modeValue === "edit" ||
                                                                                            modeValue === null ||
                                                                                            modeValue === "" ||
                                                                                            (modeValue === "approve" && formData?.Status === "Rework") ||
                                                                                            (modeValue == "approve" && editID?.IsRework == "Yes")
                                                                                        )
                                                                                    )) &&
                                                                                    <td style={{ minWidth: '60px', maxWidth: '60px', textAlign: 'center' }}>
                                                                                        {((modeValue != null && modeValue != "" && modeValue == "edit" || modeValue == "view" || modeValue == "approve")
                                                                                            || (modeValue == "approve" && formData?.Status == "Rework")) && row.Id &&
                                                                                            <>
                                                                                                <span onClick={() => OpenFile(row && row, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                                    <FontAwesomeIcon title='Preview file' icon={faEye} /></span>
                                                                                                <span onClick={() => OpenFile(row && row, "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                                    <FontAwesomeIcon title='Download file' icon={faDownload} /></span>
                                                                                            </>
                                                                                        }
                                                                                        {(((modeValue == "edit" || modeValue == null || modeValue == ""
                                                                                            || (modeValue == "approve" && formData?.Status == "Rework")) && !showviewdownload) ||
                                                                                            (modeValue == "approve" && editID?.IsRework == "Yes") || modeValue == "edit") &&
                                                                                            <img src={require("../assets/del.png")} style={{ cursor: "pointer" }} onClick={() => handleDelete(index)} />
                                                                                        }
                                                                                    </td>
                                                                                }
                                                                            </tr>
                                                                        ))
                                                                    )}
                                                                </tbody>
                                                            </table>
                                                        }
                                                        {/* </>
                                                            )
                                                        } */}

                                                    </Modal.Body>

                                                </Modal>
                                                <Modal show={ShowModalpre} onHide={() => setShowModalpre(false)} size={Showfile ? "xl" : "lg"} className='filemodal'>
                                                    <Modal.Header closeButton>
                                                        <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-0'>Attachment Details</h4>
                                                            <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for IMS Audit Report and Checklist
                                                            </p>

                                                        </Modal.Title>


                                                    </Modal.Header>
                                                    <Modal.Body className="" id="style-5">

                                                        {/* {DocumentLink &&
                                                            (
                                                                <> */}
                                                        {Showfile ?

                                                            <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />
                                                            :
                                                            <table className="mtbalenew">
                                                                <thead style={{ background: '#eef6f7' }}>
                                                                    <tr>
                                                                        <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                                                        <th style={{ minWidth: '200px', maxWidth: '200px' }}>File Name</th>

                                                                        <th style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>Upload date</th>
                                                                        <th style={{ minWidth: '50px', maxWidth: '50px' }}>Action</th>
                                                                        {/* {(InputDisabled || !InputDisabled) && <th className='text-center'>Action</th>} */}
                                                                    </tr>
                                                                </thead>
                                                                <tbody>
                                                                    {TemplateDocAudit && TemplateDocAudit.length > 0 && (
                                                                        TemplateDocAudit.map((row: any, index: number) => (
                                                                            <tr>
                                                                                <td style={{ minWidth: '50px', maxWidth: '50px' }} className='text-center'>{index + 1}</td>
                                                                                {/* <td title={row.name || row.FileLeafRef}>{row.name || row.FileLeafRef}</td> */}
                                                                                <td style={{ minWidth: '200px', maxWidth: '200px' }} title={cleanFileName(row.name) || cleanFileName(row.FileLeafRef)}>
                                                                                    {cleanFileName(row.name) || cleanFileName(row.FileLeafRef)}</td>
                                                                                {/* <td style={{ textAlign: 'center' }} >
                                                                                       
                                                                                        <span onClick={() => OpenFile(DocumentLink, "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}><FontAwesomeIcon icon={faEye} /></span>
                                                                                         </td> */}
                                                                                {/* <td>{DocumentLink.Created
                                                                                        ? new Intl.DateTimeFormat('en-GB', {
                                                                                            day: '2-digit',
                                                                                            month: 'short',
                                                                                            year: 'numeric'
                                                                                        }).format(new Date(DocumentLink.Created)).replace(/ /g, "/")
                                                                                        : ""}</td> */}



                                                                                <td style={{ minWidth: '50px', maxWidth: '50px' }} title={row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                    day: "2-digit",
                                                                                    month: "short",
                                                                                    year: "numeric"
                                                                                }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                    day: "2-digit",
                                                                                    month: "short",
                                                                                    year: "numeric"
                                                                                }).replace(/ /g, "/")}>

                                                                                    {row.Created ? new Date(row.Created).toLocaleDateString("en-GB", {
                                                                                        day: "2-digit",
                                                                                        month: "short",
                                                                                        year: "numeric"
                                                                                    }).replace(/ /g, "/") : new Date().toLocaleDateString("en-GB", {
                                                                                        day: "2-digit",
                                                                                        month: "short",
                                                                                        year: "numeric"
                                                                                    }).replace(/ /g, "/")}</td>
                                                                                <td style={{ minWidth: '50px', maxWidth: '50px', textAlign: 'center' }} title={"Preview Document"}>
                                                                                    <span onClick={() => OpenFile(TemplateDocAudit && TemplateDocAudit[0], "Open")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                        <FontAwesomeIcon title='Preview file' icon={faEye} /></span>
                                                                                    <span onClick={() => OpenFile(TemplateDocAudit && TemplateDocAudit[0], "Download")} style={{ color: "blue", cursor: "pointer", margin: "10px" }}>
                                                                                        <FontAwesomeIcon icon={faDownload} /></span>
                                                                                </td>
                                                                                {/* {!InputDisabled && <td style={{ textAlign: 'center' }}> <img src={require("../assets/del.png")} style={{ cursor: "pointer" }} onClick={() => handleDelete(index)} /></td>} */}



                                                                            </tr>
                                                                        ))
                                                                    )}
                                                                </tbody>
                                                            </table>
                                                        }
                                                        {/* </>
                                                            )
                                                        } */}

                                                    </Modal.Body>

                                                </Modal>
                                                <Modal show={ShowModalTemplateDoc} onHide={() => setShowModalTemplateDoc(false)} size={Showfile ? "xl" : "lg"} className='newmobmodal'>

                                                    {/* </Modal.Header> */}
                                                    <Modal.Body className="" id="style-5">
                                                        <>
                                                            {Showfile &&

                                                                <FileViewer showfile={Showfile} docurl={redirecturl} cancelAction={cancelModalAction} />

                                                            }
                                                        </>
                                                    </Modal.Body>
                                                </Modal>
                                                {/* ///////////////// */}

                                            </div>



                                        }

                                        <div
                                            ref={hiddenDivRef}
                                            onClick={handleHiddenDivClick}
                                            style={{ display: 'none' }}
                                        >

                                        </div>
                                    </div>
                                </div>

                            </div>
                        </div>


                    </div>
                </div>
            </div>
        </div>
    );
}

const AnnualAuditReport: React.FC<IAnnualAuditReportProps> = (props) => (
    <Provider>
        <AnnualAuditReportContext props={props} />
    </Provider>
);



export default AnnualAuditReport;
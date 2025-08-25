import * as React from 'react';
import { useState, useEffect, useRef, useContext } from 'react';
import styles from './AdvancedSearch.module.scss';
import type { IAdvancedSearchProps } from './IAdvancedSearchProps';
import { escape } from '@microsoft/sp-lodash-subset';
import "bootstrap/dist/css/bootstrap.min.css";
import { GraphSearchHelper } from '../../../Shared/SearchHelper1';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { IDocumentDisplayFields } from '../../dmsMusaib/components/DMSSearch/Interfaces';
import { ISearchHitResource } from '../../../Shared/SearchHelperInterfaces';
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import UserContext from "../../../GlobalContext/context";
import '../../../CustomCss/mainCustom.scss';
import { SearchAggregation } from '@microsoft/microsoft-graph-types';
// import { DateRangeFilter } from './DateRangeFilter';
import { DMSEntitySearchTreeView, enumfieldtype } from './DMSEntitySearchFilter';
import Provider from '../../../GlobalContext/provider';
// import { CheckedDropdownWithSearch } from './CheckedDropDown';
// import { DMSEntitySearchDropDowns } from './DMSEntitySearchDropDownFilter';
import TagsComponent from './SearchResultsTags';
import  './AdvancedSearch.scss';
import { FilterCheckBox } from './FilterCheckBoxes';
import TreeView from './EnitySearchtree';
import { DMSSearchFilterCheckedDropDown } from './DMSSearchFilterCheckedDropDown';
import { fieldnamesmapping, GetFieldName, removeDuplicates } from './Common';
import {SearchResultsWithPagination,SearchResult} from './SearchResultsWithPagination';
// import { AnyPtrRecord } from 'dns';
import PreviewFile from './previewfile';
export const getUrlParameter = (name: string) => {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(name);
}

export interface IField {
    fieldname: string;
    fieldtype: enumfieldtype;
}

export enum enumPathType
{library, site}

const fieldsinit:IField[]=[
//     {
//     fieldname:'lastModifiedDateTime',
//     fieldtype:enumfieldtype.DateTime,    
//    }
//,
{
    fieldname:'LastModifiedTime',
    fieldtype:enumfieldtype.DateTime,    
},{
    fieldname:'FileType',
    fieldtype:enumfieldtype.SingleLineofText
}];
const AdvancedSearch: React.FC<IAdvancedSearchProps> = (props) => {

    const [previewFileUrl, setPreviewFileUrl] = useState<string | null>(null);

    const sp: SPFI = getSP(props.context);
    const [searchText, setSearchText] = useState<string>('');
    // const [searchFilter, setSearchFilter] = useState<string>('');
    const [searchFilter, setSearchFilter] = useState<Map<string,string>>(new Map<string,string>());
    // const [searchPath, setSearchPath] = useState<string>(props.context.pageContext.site.absoluteUrl);
    const [searchPath, setSearchPath] = useState<string>("site:"+props.context.pageContext.site.absoluteUrl);
    const [tsearchpaths, settSearchPath] = useState<Map<string,string[]>>();
    const [searchResult, setSearchResult] = useState<IDocumentDisplayFields[]>([]);
    const [searchQueryRefiners, setSearchQueryRefiners] = useState<string[]>([]);
    const [searchfields, setsearchfields] = useState<IField[]>(fieldsinit);
    const [searchRefiners, setSearchRefiners] = useState<SearchAggregation[]>([]);
    // const [searchRefinerFilters, setSearchRefinerFilters] = useState<string[]>([]);
    const [searchRefinerFilters, setSearchRefinerFilters] = useState<Map<string,string[]>>(new Map<string,string[]>());
    //const [startDate, setStartDate] = useState<string>('');
    const [startDate, setStartDate] = useState<Map<string,string>>(new Map<string,string>());
    //const [endDate, setEndDate] = useState<string>('');
    const [endDate, setEndDate] = useState<Map<string,string>>(new Map<string,string>());
    const elementRef = useRef<HTMLDivElement>(null);
    const useHide = useContext(UserContext);
    const [isLoading, setIsLoading] = useState<boolean>(false);

    //let _searchpaths:string[]=[]; 
    //let _searchpaths:Map<string,string[]>=new Map<string,string[]>(); 
    let rootpath=window.location.origin;


    useEffect(() => {
        const searchQueryFromUrl = getUrlParameter("searchquery");
        const searchPathFromUrl = getUrlParameter("searchpath");
        const _searchQuery = searchQueryFromUrl ? decodeURIComponent(searchQueryFromUrl) : getSearchFilter();
        const _searchPath = searchPathFromUrl ? decodeURIComponent(searchPathFromUrl) : searchPath;
        settSearchPath(new Map<string,string[]>()); 
        setSearchText(_searchQuery);
        runSearch(_searchQuery,_searchQuery, _searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
    }, []);
      let checkurl = window.location.href;
       console.log("checkurl", checkurl)
    // const runSearch = async (searchText: string, searchFilters: string, searchPath: string, refiners: string[] = [], refinerfilters: string[] = []) => {
    //     let qyerytext = `${searchText} IsDocument:True ${searchFilters} ${searchPath}`;
    //     let graphcl = await (props.context as BaseWebPartContext).msGraphClientFactory.getClient("3");
    //     let mssearch = new GraphSearchHelper(graphcl);
    //     let hitcont = await mssearch.searchAll(qyerytext, 500, refiners, refinerfilters);
    //     const searchres = (hitcont.hits) ? hitcont.hits : []
    //     const results: Partial<ISearchHitResource>[] = (hitcont.hits) ? hitcont.hits.map((hit) => {
    //         const resource: Partial<ISearchHitResource> = hit.resource as ISearchHitResource;
    //         return resource;
    //     }) : [];
    //     let resultsdoc: IDocumentDisplayFields[] = searchres.map(filehit => {
    //         let file: Partial<ISearchHitResource> = filehit.resource;
    //         let tRes: IDocumentDisplayFields = { Title: file.name, Size: file.size, Extension: file.name.split('.').pop(), Path: file.webUrl, Summary: filehit.summary,Properties:file.listItem.fields }
    //         return tRes;
    //     });
    //     console.log(`Rendering search result2: ${JSON.stringify(resultsdoc)}`);
    //     setSearchResult(removeDuplicates(resultsdoc,'Path'));
    //     setSearchRefiners(hitcont.aggregations);
    // }

    
const runSearch = async (searchText: string, searchFilters: string, searchPath: string, refiners: string[] = [], refinerfilters: string[] = []) => {
    setIsLoading(true); // Start loader

    let qyerytext = `${searchText} IsDocument:True ${searchFilters} ${searchPath}`;
    let graphcl = await (props.context as BaseWebPartContext).msGraphClientFactory.getClient("3");
    let mssearch = new GraphSearchHelper(graphcl);
    let hitcont = await mssearch.searchAll(qyerytext, 500, refiners, refinerfilters);
    const searchres = (hitcont.hits) ? hitcont.hits : []
    const results: Partial<ISearchHitResource>[] = (hitcont.hits) ? hitcont.hits.map((hit) => {
        const resource: Partial<ISearchHitResource> = hit.resource as ISearchHitResource;
        return resource;
    }) : [];
    
    // let resultsdoc: IDocumentDisplayFields[] = await Promise.all(searchres.map(async filehit => {
    //     let file: Partial<ISearchHitResource> = filehit.resource;
    //     console.log(filehit + "filehit")
    //     console.log(JSON.stringify(filehit) + "filehit")
    //     let previewUrl = file.webUrl; // fallback

    //     if (file.webUrl && file.name) {
    //         const webUrlObj = new URL(file.webUrl);
    //         const serverRelativePath = webUrlObj.pathname;
    //         const fileName = encodeURIComponent(file.name);
    //         const fileExtension = file.name.split('.').pop()?.toLowerCase();
    //         const officeExtensions = ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'];
    //         const officePrefixMap: { [key: string]: string } = {
    //             'docx': ':w:/r',
    //             'doc': ':w:/r',
    //             'xlsx': ':x:/r',
    //             'xls': ':x:/r',
    //             'pptx': ':p:/r',
    //             'ppt': ':p:/r',
    //         };

    //         if (officeExtensions.includes(fileExtension) && file.listItem?.id) {
    //             // Office files: Modern Office File Viewer URL
    //             const prefix = officePrefixMap[fileExtension] || ':x:/r';
    //             // Dynamically determine if the next segment is a subsite
    //             let siteSubsitePath = '';
    //             const webUrlPath = webUrlObj.pathname;
    //             const pathParts = webUrlPath.split('/');
    //             // Default to site root
    //             siteSubsitePath = pathParts.slice(0, 3).join('/'); // /sites/<site>
    //             if (pathParts[1] === 'sites' && pathParts.length > 3) {
    //                 // Check if the next segment is a subsite using SharePoint REST API
    //                 try {
    //                     // Use PnPjs or fetch to get subsites
    //                     const subsites = await sp.web.webs(); // PnPjs: returns array of subsites
    //                     const possibleSubsite = decodeURIComponent(pathParts[3]);
    //                     if (subsites.some(subsite => subsite.Title === possibleSubsite)) {
    //                         // /sites/<site>/<subsite>
    //                         siteSubsitePath = pathParts.slice(0, 4).join('/');
    //                     }
    //                 } catch (e) {
    //                     // fallback to site root
    //                 }
    //             }
    //             previewUrl = `${webUrlObj.origin}/${prefix}${siteSubsitePath}/_layouts/15/Doc.aspx?sourcedoc=%7B${file.listItem.id.toUpperCase()}%7D&file=${fileName}&action=default&mobileredirect=true&DefaultItemOpen=1`;
    //         } else {
    //             // PDFs and other files: AllItems.aspx URL (library view)
    //             // const idParam = encodeURIComponent(serverRelativePath);
    //             // const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf('/'));
    //             // const parentParam = encodeURIComponent(parentFolder);
    //             // const pathParts = serverRelativePath.split('/');
    //             // let libraryRoot = '';
    //             // if (pathParts[1] === 'sites' && pathParts.length > 4) {
    //             //     libraryRoot = pathParts.slice(0, 5).join('/');
    //             // } else if (pathParts[1] === 'teams' && pathParts.length > 4) {
    //             //     libraryRoot = pathParts.slice(0, 5).join('/');
    //             // } else {
    //             //     libraryRoot = pathParts.slice(0, 3).join('/');
    //             // }
    //             // console.log("libraryRoot for pdf", libraryRoot)
    //             // previewUrl = `${webUrlObj.origin}${libraryRoot}/Forms/AllItems.aspx?id=${idParam}&parent=${parentParam}`;
                
    //                 // PDFs and other files: AllItems.aspx URL (library view)
    //                 const idParam = encodeURIComponent(serverRelativePath);
    //                 const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf('/'));
    //                 const parentParam = encodeURIComponent(parentFolder);
    //                 const pathParts = serverRelativePath.split('/');
    //                 let libraryRoot = '';
                
    //                 if ((pathParts[1] === 'sites' || pathParts[1] === 'teams') && pathParts.length > 4) {
    //                     // Remove the file name if it exists (e.g., ends with .pdf, .js, etc.)
    //                     const fileName = pathParts[pathParts.length - 1];
    //                     const isFile = fileName.includes('.') && !fileName.endsWith('/');
    //                     const trimmedPathParts = isFile ? pathParts.slice(0, -1) : pathParts;
                
    //                     // Start checking for document library name after site/subsite
    //                     const baseIndex = 1; // "sites" or "teams"
    //                     let libraryIndex = baseIndex + 2; // start from 3rd index (siteName, maybe subsite, then doclib)
                
    //                     for (let i = baseIndex + 2; i < trimmedPathParts.length; i++) {
    //                         const part = trimmedPathParts[i];
    //                         // Stop at first folder name that includes '.' (indicating it's likely a file)
    //                         if (part.includes('.')) break;
    //                         libraryIndex = i;
    //                     }
                
    //                     // Join only up to the detected document library
    //                     libraryRoot = trimmedPathParts.slice(0, libraryIndex + 1).join('/');
    //                 } else {
    //                     // For short or unexpected paths, fallback
    //                     libraryRoot = pathParts.slice(0, 3).join('/');
    //                 }
    //                 const  idparam2= serverRelativePath;
    //                 console.log("libraryRoot for pdf", libraryRoot);
                
    //                 previewUrl = `${webUrlObj.origin}${libraryRoot}/Forms/AllItems.aspx?id=${serverRelativePath}&parent=${parentParam}`;
                
                
    //         }
    //     }

    //     let tRes: IDocumentDisplayFields = { 
    //         Title: file.name || '', 
    //         Size: file.size, 
    //         Extension: file.name ? file.name.split('.').pop() : '', 
    //         Path: previewUrl,
    //         Summary: filehit.summary,
    //         Properties: file.listItem?.fields 
    //     };
    //     return tRes;
    // }));

    // let resultsdoc: IDocumentDisplayFields[] = await Promise.all(searchres.map(async filehit => {
    //     let file: Partial<ISearchHitResource> = filehit.resource;
    //     let previewUrl = file.webUrl; // fallback
    
    //     if (file.webUrl && file.name) {
    //         const webUrlObj = new URL(file.webUrl);
    //         const serverRelativePath = webUrlObj.pathname;
            
    //         // Get parent folder path
    //         const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf('/'));
            
    //         // Split the path into parts
    //         const pathParts = serverRelativePath.split('/');
    //         let libraryRoot = '';
    
    //         if ((pathParts[1] === 'sites' || pathParts[1] === 'teams') && pathParts.length > 4) {
    //             // Remove the file name if it exists
    //             const fileName = pathParts[pathParts.length - 1];
    //             const isFile = fileName.includes('.') && !fileName.endsWith('/');
    //             const trimmedPathParts = isFile ? pathParts.slice(0, -1) : pathParts;
    
    //             // Start checking for document library name after site/subsite
    //             const baseIndex = 1; // "sites" or "teams"
    //             let libraryIndex = baseIndex + 2;
    
    //             // Check if pathParts[3] is a subsite
    //             if (pathParts.length > 4) {
    //                 try {
    //                     const subsites = await sp.web.webs();
    //                     const possibleSubsite = decodeURIComponent(pathParts[3]);
    //                     if (subsites.some(subsite => subsite.Title === possibleSubsite)) {
    //                         libraryIndex = baseIndex + 3;
    //                     }
    //                 } catch (e) {
    //                     console.error("Error checking subsites:", e);
    //                 }
    //             }
    
    //             for (let i = libraryIndex; i < trimmedPathParts.length; i++) {
    //                 const part = trimmedPathParts[i];
    //                 if (part.includes('.')) break;
    //                 libraryIndex = i;
    //             }
    
    //             libraryRoot = trimmedPathParts.slice(0, libraryIndex + 1).join('/');
    //         } else {
    //             libraryRoot = pathParts.slice(0, 3).join('/');
    //         }
    
    //         // CORRECT ENCODING FOR SHAREPOINT URL
    //         const encodeForSharePoint = (path: string): string => {
    //             // First decode any existing encodings to avoid double-encoding
    //             let decodedPath = decodeURIComponent(path);
    //             // Then encode spaces as %20 and other special characters
    //             let encoded = encodeURI(decodedPath);
    //             // Finally encode forward slashes as %2F
    //             encoded = encoded.replace(/\//g, '%2F');
    //             return encoded;
    //         };
    
    //         const idParam = encodeForSharePoint(serverRelativePath);
    //         const parentParam = encodeForSharePoint(parentFolder);
            
    //         // Construct the final URL
    //         previewUrl = `${webUrlObj.origin}${libraryRoot}/Forms/AllItems.aspx?id=${idParam}&parent=${parentParam}`;
            
    //         // Convert %20 back to spaces in the main path only (not query parameters)
    //         const [baseUrl, queryString] = previewUrl.split('?');
    //         const fixedBaseUrl = baseUrl.replace(/%20/g, ' ');
    //         previewUrl = queryString ? `${fixedBaseUrl}?${queryString}` : fixedBaseUrl;
    //     }
    
    //     let tRes: IDocumentDisplayFields = { 
    //         Title: file.name || '', 
    //         Size: file.size, 
    //         Extension: file.name ? file.name.split('.').pop() : '', 
    //         Path: previewUrl,
    //         Summary: filehit.summary,
    //         Properties: file.listItem?.fields 
    //     };
    //     return tRes;
    // }));
    // STEP 1: FETCH subsites once before looping
let subsiteTitles: string[] = [];
try {
    const subsites = await sp.web.webs();
    subsiteTitles = subsites.map(sub => decodeURIComponent(sub.Title));
} catch (e) {
    console.error("Error fetching subsites:", e);
}

// STEP 2: Use this array inside your map logic
let resultsdoc: IDocumentDisplayFields[] = await Promise.all(searchres.map(async filehit => {
    let file: Partial<ISearchHitResource> = filehit.resource;
    let previewUrl = file.webUrl;

    if (file.webUrl && file.name) {
        const webUrlObj = new URL(file.webUrl);
        const serverRelativePath = webUrlObj.pathname;
        const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf('/'));
        const pathParts = serverRelativePath.split('/');
        let libraryRoot = '';

        if ((pathParts[1] === 'sites' || pathParts[1] === 'teams') && pathParts.length > 4) {
            const fileName = pathParts[pathParts.length - 1];
            const isFile = fileName.includes('.') && !fileName.endsWith('/');
            const trimmedPathParts = isFile ? pathParts.slice(0, -1) : pathParts;

            const baseIndex = 1; // "sites" or "teams"
            let libraryIndex = baseIndex + 2;

            // NOW use the pre-fetched subsite titles
            const possibleSubsite = decodeURIComponent(pathParts[3]);
            if (subsiteTitles.includes(possibleSubsite)) {
                libraryIndex = baseIndex + 3;
            }

            for (let i = libraryIndex; i < trimmedPathParts.length; i++) {
                const part = trimmedPathParts[i];
                if (part.includes('.')) break;
                libraryIndex = i;
            }

            libraryRoot = trimmedPathParts.slice(0, libraryIndex + 1).join('/');
        } else {
            libraryRoot = pathParts.slice(0, 3).join('/');
        }

        // Same encoding logic
        const encodeForSharePoint = (path: string): string => {
            let decodedPath = decodeURIComponent(path);
            let encoded = encodeURI(decodedPath);
            encoded = encoded.replace(/\//g, '%2F');
            return encoded;
        };

        const idParam = encodeForSharePoint(serverRelativePath);
        const parentParam = encodeForSharePoint(parentFolder);
        previewUrl = `${webUrlObj.origin}${libraryRoot}/Forms/AllItems.aspx?id=${idParam}&parent=${parentParam}`;
        const [baseUrl, queryString] = previewUrl.split('?');
        const fixedBaseUrl = baseUrl.replace(/%20/g, ' ');
        previewUrl = queryString ? `${fixedBaseUrl}?${queryString}` : fixedBaseUrl;
    }

    return {
        Title: file.name || '',
        Size: file.size,
        Extension: file.name ? file.name.split('.').pop() : '',
        Path: previewUrl,
        Summary: filehit.summary,
        Properties: file.listItem?.fields
    };
}));

    console.log(`Rendering search result2: ${JSON.stringify(resultsdoc)}`);
    setSearchResult(removeDuplicates(resultsdoc,'Path'));
    setSearchRefiners(hitcont.aggregations);
    setIsLoading(false); // Stop loader

}
    const getSearchFilter=(searchFiltertemp?:Map<string,string>):string=>{
   
        if(!searchFiltertemp) searchFiltertemp=new Map([...searchFilter]);

        let refinefilter= (Array.from(searchFiltertemp.values()).length>0)?Array.from(searchFiltertemp.values()).join(" "):"";
        return refinefilter;  
    }

    const getSearchRefineFiltersArray=(searchRefinerFilterstemp?:Map<string,string[]>):string[]=>{
   
        if(!searchRefinerFilterstemp) searchRefinerFilterstemp=new Map([...searchRefinerFilters]);
        let refinefilterarray=
        Array.from(searchRefinerFilterstemp.entries()).map(([key, value]) => {
            return (value.length>1)?`${key}:or(${value.join(',')})`:`${key}:${value}`;
          });
        return refinefilterarray;  
    }

    const searchClickHandler: React.MouseEventHandler = (ev) => {
        ev.preventDefault();
        runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
    }

    const searchTextChangeHandler: React.ChangeEventHandler = (ev) => {
        setSearchText((ev.target as HTMLInputElement).value);
    }

    const handleDateTimeFilter = (filter: string) => {

        // setSearchFilter(searchFilter+" "+filter);
        // runSearch(searchText, filter, searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
    }

    const GetFieldType = (fieldname: string) => {
        try{
            console.log('fieldname',fieldname);
            console.log('searchfields',searchfields);
            return searchfields.filter(f => f.fieldname.toLowerCase() == fieldname.toLowerCase())[0].fieldtype;
        }
        catch(ex)
        {
            
            console.log("error",ex);
        }
    }
    

    const handleCheckboxChange = (refinerName: string, value: string, checked: boolean) => {
        console.log(refinerName);
    
        setSearchRefinerFilters((prevFilters) => {
            // Create a new Map to preserve immutability
            const updatedFilters = new Map([...prevFilters]);
    
            if (checked) {
                // If the checkbox is checked, add the value
                if (updatedFilters.has(refinerName)) {
                    const existingValues = [...(updatedFilters.get(refinerName) || [])];
                    updatedFilters.set(refinerName, [...existingValues, value]);
                } else {
                    updatedFilters.set(refinerName, [value]);
                }
            } else {
                // If the checkbox is unchecked, remove the value
                const existingValues = [...(updatedFilters.get(refinerName) || [])];
                const filteredValues = existingValues.filter((v) => v !== value);
    
                if (filteredValues.length > 0) {
                    updatedFilters.set(refinerName, filteredValues);
                } else {
                    updatedFilters.delete(refinerName);
                }
            }
    
            return updatedFilters;
        });
    };
    

    // Generate filter query based on input
    const generateFilter = (datefield: string,startdt?:string,enddt?:string): string => {
        let stdate=(startdt)?startdt:startDate.get(datefield);
        let enddate=(enddt)?enddt:endDate.get(datefield);
        if (stdate && enddate) {
            // return `LastModifiedTime:Range(${startDate}..${endDate})`;
            return `(${datefield}>=${stdate} AND ${datefield}<=${enddate})`;
        } else if (stdate) {
            return `${datefield}>=${stdate}`;
        } else if (enddate) {
            return `${datefield}<=${enddate}`;
        }
        return '';
    };

    const handleDateFilterChange = (datefield: string) => {
        // let filter = '';
        // if (startDate && endDate) {
        //     filter = `LastModifiedTime:Range(${startDate}..${endDate})`;
        // } else if (startDate) {
        //     filter = `LastModifiedTime>=${startDate}`;
        // } else if (endDate) {
        //     filter = `LastModifiedTime<=${endDate}`;
        // }

        // onFilterChange(filter); // Pass filter to parent
        // const filter = generateFilter(datefield);
        // setSearchFilter(filter);
        // runSearch(searchText, filter, searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
    };

    const handleApply = (datefield: string) => {
        // const filter = generateFilter(datefield);
        // let newfilter = searchFilter +" "+filter;
        // setSearchFilter(newfilter);
        //runSearch(searchText, filter, searchPath, searchQueryRefiners, searchRefinerFilters);
    };

    const handleApplyFilters = () => {
        // const filter = generateFilter(datefield);
        // setSearchFilter(filter);
        // runSearch(searchText, filter, searchPath, searchQueryRefiners, searchRefinerFilters);

        runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
    };

    const GetStat=()=>{
        return Array.from(searchRefinerFilters.keys()).length;
    }

    const ClearFilter=(filterfield:string)=>{

        const updatedFilters = new Map([...searchFilter]);
        updatedFilters.delete(filterfield);
        setSearchFilter(updatedFilters);
        const stdt=new Map([...startDate]);
        stdt.delete(filterfield);
        setStartDate(stdt);
        const endt=new Map([...endDate]);
        endt.delete(filterfield);
        setEndDate(endt);
        runSearch(searchText, getSearchFilter(updatedFilters), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());

    }

    const ClearRefinerFilter=(filterfield:string)=>{

        const updatedFilters = new Map([...searchRefinerFilters]);
        updatedFilters.delete(filterfield);
        setSearchRefinerFilters(updatedFilters);
        runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray(updatedFilters));

    }

    const handleKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
        if (event.key === "Enter") {
            event.preventDefault();
            console.log("Enter key pressed. Current value:", event.currentTarget.value);
            runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
            
        }
    };

    // const  UpdatedSearchPaths=(allpath:{selepath:string,typeofpath:enumPathType}[])=>
    // {
    //     allpath.forEach(a=>{        
    //         let patht;
    //         if(a.typeofpath==enumPathType.library) patht=rootpath+a.selepath;
    //         else patht=a.selepath
    //         _searchpaths.push(path);            
    //     });
    //     _searchpaths=getDeepestPaths(_searchpaths);
    //     let updatedpath=_searchpaths.map(d=>"path:"+d).join(' ');
    //     setSearchPath(updatedpath);
    //     return updatedpath;
    // }

    const getUpdatedSearchPath=(_srchpaths?:Map<string,string[]>)=>{

        let allpaths:string[]=[];
        let _searchpaths:Map<string,string[]>;
        if(_srchpaths) _searchpaths=_srchpaths;
        else _searchpaths=new Map([...tsearchpaths]);
       
        Array.from(_searchpaths.keys()).forEach(s=>{
            // if(_searchpaths.get(s).length==0)
            // {
            //     // allpaths.push(s);
            //     allpaths.push("site:"+s);
            // }
            // else
            // {
            //     // allpaths=allpaths.concat(_searchpaths.get(s));
            //     allpaths=allpaths.concat(_searchpaths.get(s).map(p=>`listid:${p}`));
            // }

            // allpaths.push("site:"+s);
            allpaths.push("site:"+"\""+s+"\"");
 
            // allpaths.push("path:"+s);
            allpaths=allpaths.concat(_searchpaths.get(s).map(p=>`listid:${p}`));

            // if(_searchpaths.get(s).length!=0)
            // {
            //    allpaths=allpaths.concat(_searchpaths.get(s).map(p=>`listid:${p}`));
            // }
        })

        // let updatedpath=(allpaths.length>0)?allpaths.map(d=>"path:"+d).join(' '):("path:"+props.context.pageContext.site.absoluteUrl);
        let updatedpath= (allpaths.length>0)?allpaths.join(" "):("site:"+props.context.pageContext.site.absoluteUrl);
        // let updatedpath= (allpaths.length>0)?allpaths.join(" "):("path:"+props.context.pageContext.site.absoluteUrl);
        setSearchPath(updatedpath);
        return updatedpath;

    }

    const getDeepestPaths=(paths: string[]): string[]=> {
        const getDepth = (path: string): number => {
            // Remove the protocol and domain, if present, and count the segments
            const cleanedPath = path.replace(/^https?:\/\/[^/]+/, "");
            return cleanedPath.split("/").filter(Boolean).length;
        };
    
        const groupedByRoot: Record<string, string[]> = {};
    
        // Group paths by their root (protocol + domain)
        paths.forEach((path) => {
            const root = path.match(/^https?:\/\/[^/]+/)?.[0] || ""; // Extract root
            if (!groupedByRoot[root]) {
                groupedByRoot[root] = [];
            }
            groupedByRoot[root].push(path);
        });
    
        const deepestPaths: string[] = [];
    
        // Find the deepest path in each group
        Object.values(groupedByRoot).forEach((group) => {
            let maxDepth = 0;
            let deepestGroup: string[] = [];
    
            group.forEach((path) => {
                const depth = getDepth(path);
                if (depth > maxDepth) {
                    maxDepth = depth;
                    deepestGroup = [path];
                } else if (depth === maxDepth) {
                    deepestGroup.push(path);
                }
            });
    
            deepestPaths.push(...deepestGroup);
        });
    
        return deepestPaths;
    }

    return (
        <div id="wrapper" ref={elementRef}>
            <div className="app-menu" id="myHeader">
                <VerticalSideBar _context={sp} />
            </div>

            <div className="content-page">
                <HorizontalNavbar _context={sp} siteUrl={props.siteUrl} context={props.context}  />
                {
    window.location.href.includes('.aspx?Previewfile') 
      ? <PreviewFile /> 
      :  <div className="content" style={{ marginLeft: `${!useHide ? '80px' : '230px'}`, marginTop: '1.5rem' }}>
      <section className='container-fluid'>
          <div className='row'>
              <header className="p-3 pb-0 col-12">
                  {/* <h1 style={{fontSize:'20px', fontWeight:'600'}} className='text-dark'>Search</h1> */}
                  <form>
                      <div className="input-group">
                          <input style={{ padding: '.75rem .75rem' }} type="text" className="form-control" placeholder="Search Files.."
                           onChange={searchTextChangeHandler} 
                           
                           onKeyDown={handleKeyPress} value={searchText}
                          


                          />
                            {/* Clear button */}
  {searchText && (
    <span
       onClick={(ev: React.MouseEvent) => {
             setSearchText('')
               runSearch('', getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
         ; // If you need to keep this
           // Clear the search text
      }}
      style={{
        position: 'absolute',
        right: '10.5rem',
        top: '50%',
        transform: 'translateY(-50%)',
        cursor: 'pointer',
        fontSize: '1.2rem',
        color: '#888'
      }}
    >
      &times;
    </span>
  )}

                          <button style={{ fontSize: '1rem' }} className="btn btn-primary" type="button" onClick={searchClickHandler}>Advance Search 1</button>
                      </div>
                  </form>
              </header>

              <main className="mt-3 p-0 pb-3 col-12">
                  <div className='row mt-0 p-0'>
                      <section style={{paddingLeft:'25px'}}className='col-sm-3 nwsera'>
                          <div  style={{background:'transparent', border:'0px solid #ccc',width:'95%',}}  className='card'>
                          <h5 style={{textAlign:'left', fontSize:'18px',borderBottom:'1px solid #ccc',  margin:'inherit'}} className='text-dark p-0 pt-0 mb-2 fw-bold'>Filter</h5> 
                          <DMSSearchFilterCheckedDropDown context={props.context} 
                          
                          onMultiFieldSelect={selfields => {
                              let selfld = [...searchQueryRefiners];
                              // console.log('fieldtype', fldtype);
                              // selfld.push(selfields.map(s=>s.field));
                              selfld=selfld.concat(selfields.map(s=>s.field));

                              setSearchQueryRefiners(selfld);
                              let flds = [...searchfields];
                              // flds.push({ fieldname: fld, fieldtype: fldtype });
                              flds=flds.concat(selfields.map(s=>({fieldname:s.field,fieldtype:s.fieldtype})));
                              setsearchfields(flds);
                              runSearch(searchText, getSearchFilter(), searchPath, selfld, getSearchRefineFiltersArray());
                          }} 
                          
                          // onMultiFieldSelect={selfields => {
                          //     let selfld = [...searchQueryRefiners];
                          //     let flds = [...searchfields];
                          
                          //     // Extract fields and field objects from selfields
                          //     const selectedFields = selfields.map(s => s.field);
                          //     const selectedFieldObjects = selfields.map(s => ({ fieldname: s.field, fieldtype: s.fieldtype }));
                          
                          //     // Add new fields to searchQueryRefiners
                          //     selfld = selfld.filter(field => selectedFields.includes(field));
                          //     selfld = selfld.concat(selectedFields.filter(field => !selfld.includes(field)));
                          
                          //     // Add new field objects to searchfields
                          //     flds = flds.filter(fld => selectedFields.includes(fld.fieldname));
                          //     flds = flds.concat(selectedFieldObjects.filter(fld => !flds.some(existingFld => existingFld.fieldname === fld.fieldname)));
                          
                          //     setSearchQueryRefiners(selfld);
                          //     setsearchfields(flds);
                          //     runSearch(searchText, getSearchFilter(), searchPath, selfld, getSearchRefineFiltersArray());
                          // }}

                          onSiteSelect={(sites)=>{
                            console.log("Selected istes",sites)
                            let _searchpaths=new Map([...tsearchpaths]);
                            
                            sites.forEach(s=>{
                               if(!_searchpaths.has(s.SiteURL))//if(!allkeys.some(s1=>s.SiteURL.toLowerCase()==s1.toLowerCase()))
                               {
                                  _searchpaths.set(s.SiteURL,[]);
                               }
                            })
                            let allkeys=Array.from(_searchpaths.keys());
                            let siteurls=Array.from(sites.map(s=>s.SiteURL));
                            allkeys.forEach(s=>{
                              if(!siteurls.some(p=>p==s))//if(!allkeys.some(s1=>s.SiteURL.toLowerCase()==s1.toLowerCase()))
                              {
                                 _searchpaths.delete(s);
                              }
                            })

                            settSearchPath(_searchpaths);

                            let newserachpath=getUpdatedSearchPath(_searchpaths);//UpdatedSearchPaths(s.map(s1=>({selepath:s1.SiteURL,typeofpath:enumPathType.site })))
                            runSearch(searchText, getSearchFilter(), newserachpath, searchfields.map(sf=>sf.fieldname), getSearchRefineFiltersArray());

                          }} 
                          
                          onLibrarySelect={(libs)=>{
                              let _searchpaths=new Map([...tsearchpaths]);
                              let allkeys=Array.from(_searchpaths.keys());
                              
                              libs.forEach(l=>{
                                  let liburl=rootpath+l.FolderPath;
                                  //let liburl=l.ListID;
                                  let libsite= allkeys.find(a=>liburl.toLowerCase().startsWith(a.toLowerCase()))
                                  if(libsite)
                                      {
                                          let oldlibs= _searchpaths.get(libsite);
                                          // if(!oldlibs.some(s=>s==liburl)) oldlibs.push(liburl);
                                          if(!oldlibs.some(s=>s==l.ListID)) oldlibs.push(l.ListID);
                                          _searchpaths.set(libsite,oldlibs);
                                      }
                              })

                          //    let alllibpaths= libs.map(l=>rootpath+l.FolderPath);
                            let alllibpaths= libs.map(l=>l.ListID);
                              allkeys.forEach(k=>{
                                 let tlibs= _searchpaths.get(k);   
                                 let newtlib=tlibs;                                         
                                 tlibs.forEach(tlib=>{
                                   if(!alllibpaths.some(al=>al==tlib))
                                   {
                                      newtlib.splice(newtlib.indexOf(tlib));
                                   }

                                 })
                                 _searchpaths.set(k,newtlib);
                              })
                              
                              settSearchPath(_searchpaths);
                              let newserachpath=getUpdatedSearchPath(_searchpaths);//UpdatedSearchPaths(s.map(s1=>({selepath:s1.FolderPath,typeofpath:enumPathType.library })))
                              runSearch(searchText, getSearchFilter(), newserachpath, searchfields.map(sf=>sf.fieldname), getSearchRefineFiltersArray());

                          }}/>
                          </div>
                          <div style={{width:'95%'}}>
                              <h5 style={{textAlign:'left', fontSize:'18px',borderBottom:'1px solid #ccc'}} className='text-dark  mb-2 fw-bold'>Refiners</h5>
                              <div className='row p-0'>
                                  {searchRefiners?.map(refiner => (
                                      <div  style={{border:'0px solid #1fb0e5', borderRadius:'0px', background:'transparent'}} key={refiner.field} className="card col-12 mb-3 pt-2">
                                          <h6 style={{textAlign:'left', fontSize:'16px', margin:'inherit'}} className='mt-0 mb-1'>{GetFieldName(refiner.field)}</h6>

                                          <div className="form-check">
                                              {
                                                  (GetFieldType(refiner.field) == enumfieldtype.DateTime) ?
                                                      (<div className='col-12'>
                                                          <div className='col'><label htmlFor="startDate" className="form-label mb-1 mt-1">Start Date</label></div>
                                                          <div className="col">
                                                          <input
                                                              type="date"
                                                              id="startDate"
                                                              className="form-control"
                                                              value={startDate.get(refiner.field) || ""}
                                                              onChange={(e) => {
                                                                  // Create a new Map to ensure immutability
                                                                  const updatedStartDate = new Map([...startDate]);
                                                                  updatedStartDate.set(refiner.field, e.target.value);
                                                                  setStartDate(updatedStartDate);
                                                                  let dtfilter=generateFilter(refiner.field,e.target.value);

                                                                  const updatedSearchFilter = new Map([...searchFilter]);
                                                                  updatedSearchFilter.set(refiner.field, dtfilter);
                                                                  setSearchFilter(updatedSearchFilter);
                                                                  // handleFilterChange();
                                                              }}
                                                          />
                                                              
                                                          </div>
                                                          <div className="col">  <label htmlFor="endDate" className="form-label mb-1 mt-1">End Date</label> </div>
                                                          <div className="col">

                                                          <input
                                                              type="date"
                                                              id="endDate"
                                                              className="form-control"
                                                              value={endDate.get(refiner.field) || ""}
                                                              onChange={(e) => {
                                                                  // Create a new Map to ensure immutability
                                                                  const updatedEndDate = new Map([...endDate]);
                                                                  updatedEndDate.set(refiner.field, e.target.value);
                                                                  setEndDate(updatedEndDate);
                                                                  let dtfilter=generateFilter(refiner.field,null,e.target.value);

                                                                  const updatedSearchFilter = new Map([...searchFilter]);
                                                                  updatedSearchFilter.set(refiner.field, dtfilter);
                                                                  setSearchFilter(updatedSearchFilter);
                                                              }}
                                                          />
                                                          </div>
                                                          {/* <div className="col mt-1 mb-3">
                                                              <button type='button' className="btn btn-primary float-end mt-2 mb-3" onClick={() => ClearFilter(refiner.field)}>
                                                                  Clear
                                                              </button>
                                                          </div> */}
                                                      </div>)
                                                      :
                                                      <>
                                                      <FilterCheckBox refiner={refiner} handleCheckboxChange={handleCheckboxChange} searchRefinerFilters={searchRefinerFilters} />
                                                      <div style={{clear:'both', position:'relative'}} className="col mt-1 mb-3">
                                                          <button type='button' className="btn btn-secondary newbuttons float-end mt-2 mb-3 me-2" onClick={() => ClearRefinerFilter(refiner.field)}>
                                                              Clear
                                                          </button>
                                                      </div>
                                                      </>
                                                      
                                                              
                                              }
                                          </div>
                                      </div>
                                  ))}
                                  {(searchRefiners && searchRefiners.length>0)?<div className="col">
                                      <button type='button' className="btn btn-success neewd" onClick={() => handleApplyFilters()}>
                                          Apply Filters
                                      </button>
                                  </div>:<></>}
                              </div>
                          </div>
                      </section>
                      <section className="col-sm-9">
                          {/* {searchResult.map(res => (
                              <div className="col mt-1 mb-1 search-result">
                                  <div className="card h-100">
                                      <div className="card-body">
                                          <h5 className="card-title">{res.Title}</h5>
                                          <p className="card-text">{res.Summary}</p>
                                      </div>
                                      <TagsComponent tags={Object.entries(res.Properties).map(([key, value]) =>`${fieldnamesmapping[key]?fieldnamesmapping[key]:key}: ${value}` )}/>
                                  </div>
                              </div>
                              // <div>{res.Title}</div>
                          ))} */}
                          {

                        //   (searchResult && searchResult.length>0)?
                        //   <SearchResultsWithPagination searchResult={searchResult.map(s=>({Title:s.Title,Summary:s.Summary,Properties:s.Properties,Path:s.Path} as SearchResult))} fieldnamesmapping={fieldnamesmapping} fieldtypemappings={searchfields}   setPreviewFileUrl={setPreviewFileUrl} />:
                        //   <div className="loadnewarg" >
                        //   <span>Loading </span>{" "}
                        //   <span>
                        //     <img style={{ width: '35px' }}
                        //       src={require("../assets/EDCLoader.gif")}
                        //       className="alignrightl"
                        //       alt="Loading..."
                        //     />
                        //   </span>
                        // </div>
                        
                            isLoading ? (
                              <div className="loadnewarg">
                                <span>Loading </span>{" "}
                                <span>
                                  <img
                                    style={{ width: '35px' }}
                                    src={require("../assets/EDCLoader.gif")}
                                    className="alignrightl"
                                    alt="Loading..."
                                  />
                                </span>
                              </div>
                            ) : searchResult && searchResult.length > 0 ? (
                                <SearchResultsWithPagination searchResult={searchResult.map(s=>({Title:s.Title,Summary:s.Summary,Properties:s.Properties,Path:s.Path} as SearchResult))} fieldnamesmapping={fieldnamesmapping} fieldtypemappings={searchfields}   setPreviewFileUrl={setPreviewFileUrl} />
                            ) : (
                              <div className="no-data-message">No data found</div>
                            )
                          
                          
                          }
                      </section>
                  </div>
              </main>
          </div>
      </section>
  </div>
  }
               
            </div>
        </div>
    );
}

const DMSMain: React.FC<IAdvancedSearchProps> = (props) => {
    return (
        <Provider>
            <AdvancedSearch {...props} />
        </Provider>
    );
}

export default DMSMain;

// previous working code 
// import * as React from 'react';
// import { useState, useEffect, useRef, useContext } from 'react';
// import styles from './AdvancedSearch.module.scss';
// import type { IAdvancedSearchProps } from './IAdvancedSearchProps';
// import { escape } from '@microsoft/sp-lodash-subset';
// import "bootstrap/dist/css/bootstrap.min.css";
// import { GraphSearchHelper } from '../../../Shared/SearchHelper1';
// import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
// import { IDocumentDisplayFields } from '../../dmsMusaib/components/DMSSearch/Interfaces';
// import { ISearchHitResource } from '../../../Shared/SearchHelperInterfaces';
// import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
// import { getSP } from "../loc/pnpjsConfig";
// import { SPFI } from "@pnp/sp";
// import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
// import UserContext from "../../../GlobalContext/context";
// import '../../../CustomCss/mainCustom.scss';
// import { SearchAggregation } from '@microsoft/microsoft-graph-types';
// // import { DateRangeFilter } from './DateRangeFilter';
// import { DMSEntitySearchTreeView, enumfieldtype } from './DMSEntitySearchFilter';
// import Provider from '../../../GlobalContext/provider';
// // import { CheckedDropdownWithSearch } from './CheckedDropDown';
// // import { DMSEntitySearchDropDowns } from './DMSEntitySearchDropDownFilter';
// import TagsComponent from './SearchResultsTags';
// import  './AdvancedSearch.scss';
// import { FilterCheckBox } from './FilterCheckBoxes';
// import TreeView from './EnitySearchtree';
// import { DMSSearchFilterCheckedDropDown } from './DMSSearchFilterCheckedDropDown';
// import { fieldnamesmapping, GetFieldName, removeDuplicates } from './Common';
// import {SearchResultsWithPagination,SearchResult} from './SearchResultsWithPagination';
// // import { AnyPtrRecord } from 'dns';

// export const getUrlParameter = (name: string) => {
//     const urlParams = new URLSearchParams(window.location.search);
//     return urlParams.get(name);
// }

// export interface IField {
//     fieldname: string;
//     fieldtype: enumfieldtype;
// }

// export enum enumPathType
// {library, site}

// const fieldsinit:IField[]=[
// //     {
// //     fieldname:'lastModifiedDateTime',
// //     fieldtype:enumfieldtype.DateTime,    
// //    }
// //,
// {
//     fieldname:'LastModifiedTime',
//     fieldtype:enumfieldtype.DateTime,    
// },{
//     fieldname:'FileType',
//     fieldtype:enumfieldtype.SingleLineofText
// }];
// const AdvancedSearch: React.FC<IAdvancedSearchProps> = (props) => {
//     const sp: SPFI = getSP(props.context);
//     const [searchText, setSearchText] = useState<string>('');
//     // const [searchFilter, setSearchFilter] = useState<string>('');
//     const [searchFilter, setSearchFilter] = useState<Map<string,string>>(new Map<string,string>());
//     // const [searchPath, setSearchPath] = useState<string>(props.context.pageContext.site.absoluteUrl);
//     const [searchPath, setSearchPath] = useState<string>("site:"+props.context.pageContext.site.absoluteUrl);
//     const [tsearchpaths, settSearchPath] = useState<Map<string,string[]>>();
//     const [searchResult, setSearchResult] = useState<IDocumentDisplayFields[]>([]);
//     const [searchQueryRefiners, setSearchQueryRefiners] = useState<string[]>([]);
//     const [searchfields, setsearchfields] = useState<IField[]>(fieldsinit);
//     const [searchRefiners, setSearchRefiners] = useState<SearchAggregation[]>([]);
//     // const [searchRefinerFilters, setSearchRefinerFilters] = useState<string[]>([]);
//     const [searchRefinerFilters, setSearchRefinerFilters] = useState<Map<string,string[]>>(new Map<string,string[]>());
//     //const [startDate, setStartDate] = useState<string>('');
//     const [startDate, setStartDate] = useState<Map<string,string>>(new Map<string,string>());
//     //const [endDate, setEndDate] = useState<string>('');
//     const [endDate, setEndDate] = useState<Map<string,string>>(new Map<string,string>());
//     const elementRef = useRef<HTMLDivElement>(null);
//     const useHide = useContext(UserContext);

//     //let _searchpaths:string[]=[]; 
//     //let _searchpaths:Map<string,string[]>=new Map<string,string[]>(); 
//     let rootpath=window.location.origin;


//     useEffect(() => {
//         const searchQueryFromUrl = getUrlParameter("searchquery");
//         const searchPathFromUrl = getUrlParameter("searchpath");
//         const _searchQuery = searchQueryFromUrl ? decodeURIComponent(searchQueryFromUrl) : getSearchFilter();
//         const _searchPath = searchPathFromUrl ? decodeURIComponent(searchPathFromUrl) : searchPath;
//         settSearchPath(new Map<string,string[]>()); 
//         setSearchText(_searchQuery);
//         runSearch(_searchQuery,_searchQuery, _searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
//     }, []);

//     const runSearch = async (searchText: string, searchFilters: string, searchPath: string, refiners: string[] = [], refinerfilters: string[] = []) => {
//         let qyerytext = `${searchText} IsDocument:True ${searchFilters} ${searchPath}`;
//         let graphcl = await (props.context as BaseWebPartContext).msGraphClientFactory.getClient("3");
//         let mssearch = new GraphSearchHelper(graphcl);
//         let hitcont = await mssearch.searchAll(qyerytext, 500, refiners, refinerfilters);
//         const searchres = (hitcont.hits) ? hitcont.hits : []
//         const results: Partial<ISearchHitResource>[] = (hitcont.hits) ? hitcont.hits.map((hit) => {
//             const resource: Partial<ISearchHitResource> = hit.resource as ISearchHitResource;
//             return resource;
//         }) : [];
//         let resultsdoc: IDocumentDisplayFields[] = searchres.map(filehit => {
//             let file: Partial<ISearchHitResource> = filehit.resource;
//             let tRes: IDocumentDisplayFields = { Title: file.name, Size: file.size, Extension: file.name.split('.').pop(), Path: file.webUrl, Summary: filehit.summary,Properties:file.listItem.fields }
//             return tRes;
//         });
//         setSearchResult(removeDuplicates(resultsdoc,'Path'));
//         setSearchRefiners(hitcont.aggregations);
//     }

//     const getSearchFilter=(searchFiltertemp?:Map<string,string>):string=>{
   
//         if(!searchFiltertemp) searchFiltertemp=new Map([...searchFilter]);

//         let refinefilter= (Array.from(searchFiltertemp.values()).length>0)?Array.from(searchFiltertemp.values()).join(" "):"";
//         return refinefilter;  
//     }

//     const getSearchRefineFiltersArray=(searchRefinerFilterstemp?:Map<string,string[]>):string[]=>{
   
//         if(!searchRefinerFilterstemp) searchRefinerFilterstemp=new Map([...searchRefinerFilters]);
//         let refinefilterarray=
//         Array.from(searchRefinerFilterstemp.entries()).map(([key, value]) => {
//             return (value.length>1)?`${key}:or(${value.join(',')})`:`${key}:${value}`;
//           });
//         return refinefilterarray;  
//     }

//     const searchClickHandler: React.MouseEventHandler = (ev) => {
//         ev.preventDefault();
//         runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
//     }

//     const searchTextChangeHandler: React.ChangeEventHandler = (ev) => {
//         setSearchText((ev.target as HTMLInputElement).value);
//     }

//     const handleDateTimeFilter = (filter: string) => {

//         // setSearchFilter(searchFilter+" "+filter);
//         // runSearch(searchText, filter, searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
//     }

//     const GetFieldType = (fieldname: string) => {
//         try{
//             console.log('fieldname',fieldname);
//             console.log('searchfields',searchfields);
//             return searchfields.filter(f => f.fieldname.toLowerCase() == fieldname.toLowerCase())[0].fieldtype;
//         }
//         catch(ex)
//         {
            
//             console.log("error",ex);
//         }
//     }
    

//     const handleCheckboxChange = (refinerName: string, value: string, checked: boolean) => {
//         console.log(refinerName);
    
//         setSearchRefinerFilters((prevFilters) => {
//             // Create a new Map to preserve immutability
//             const updatedFilters = new Map([...prevFilters]);
    
//             if (checked) {
//                 // If the checkbox is checked, add the value
//                 if (updatedFilters.has(refinerName)) {
//                     const existingValues = [...(updatedFilters.get(refinerName) || [])];
//                     updatedFilters.set(refinerName, [...existingValues, value]);
//                 } else {
//                     updatedFilters.set(refinerName, [value]);
//                 }
//             } else {
//                 // If the checkbox is unchecked, remove the value
//                 const existingValues = [...(updatedFilters.get(refinerName) || [])];
//                 const filteredValues = existingValues.filter((v) => v !== value);
    
//                 if (filteredValues.length > 0) {
//                     updatedFilters.set(refinerName, filteredValues);
//                 } else {
//                     updatedFilters.delete(refinerName);
//                 }
//             }
    
//             return updatedFilters;
//         });
//     };
    

//     // Generate filter query based on input
//     const generateFilter = (datefield: string,startdt?:string,enddt?:string): string => {
//         let stdate=(startdt)?startdt:startDate.get(datefield);
//         let enddate=(enddt)?enddt:endDate.get(datefield);
//         if (stdate && enddate) {
//             // return `LastModifiedTime:Range(${startDate}..${endDate})`;
//             return `(${datefield}>=${stdate} AND ${datefield}<=${enddate})`;
//         } else if (stdate) {
//             return `${datefield}>=${stdate}`;
//         } else if (enddate) {
//             return `${datefield}<=${enddate}`;
//         }
//         return '';
//     };

//     const handleDateFilterChange = (datefield: string) => {
//         // let filter = '';
//         // if (startDate && endDate) {
//         //     filter = `LastModifiedTime:Range(${startDate}..${endDate})`;
//         // } else if (startDate) {
//         //     filter = `LastModifiedTime>=${startDate}`;
//         // } else if (endDate) {
//         //     filter = `LastModifiedTime<=${endDate}`;
//         // }

//         // onFilterChange(filter); // Pass filter to parent
//         // const filter = generateFilter(datefield);
//         // setSearchFilter(filter);
//         // runSearch(searchText, filter, searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
//     };

//     const handleApply = (datefield: string) => {
//         // const filter = generateFilter(datefield);
//         // let newfilter = searchFilter +" "+filter;
//         // setSearchFilter(newfilter);
//         //runSearch(searchText, filter, searchPath, searchQueryRefiners, searchRefinerFilters);
//     };

//     const handleApplyFilters = () => {
//         // const filter = generateFilter(datefield);
//         // setSearchFilter(filter);
//         // runSearch(searchText, filter, searchPath, searchQueryRefiners, searchRefinerFilters);

//         runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
//     };

//     const GetStat=()=>{
//         return Array.from(searchRefinerFilters.keys()).length;
//     }

//     const ClearFilter=(filterfield:string)=>{

//         const updatedFilters = new Map([...searchFilter]);
//         updatedFilters.delete(filterfield);
//         setSearchFilter(updatedFilters);
//         const stdt=new Map([...startDate]);
//         stdt.delete(filterfield);
//         setStartDate(stdt);
//         const endt=new Map([...endDate]);
//         endt.delete(filterfield);
//         setEndDate(endt);
//         runSearch(searchText, getSearchFilter(updatedFilters), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());

//     }

//     const ClearRefinerFilter=(filterfield:string)=>{

//         const updatedFilters = new Map([...searchRefinerFilters]);
//         updatedFilters.delete(filterfield);
//         setSearchRefinerFilters(updatedFilters);
//         runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray(updatedFilters));

//     }

//     const handleKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
//         if (event.key === "Enter") {
//             event.preventDefault();
//             console.log("Enter key pressed. Current value:", event.currentTarget.value);
//             runSearch(searchText, getSearchFilter(), searchPath, searchQueryRefiners, getSearchRefineFiltersArray());
            
//         }
//     };

//     // const  UpdatedSearchPaths=(allpath:{selepath:string,typeofpath:enumPathType}[])=>
//     // {
//     //     allpath.forEach(a=>{        
//     //         let patht;
//     //         if(a.typeofpath==enumPathType.library) patht=rootpath+a.selepath;
//     //         else patht=a.selepath
//     //         _searchpaths.push(path);            
//     //     });
//     //     _searchpaths=getDeepestPaths(_searchpaths);
//     //     let updatedpath=_searchpaths.map(d=>"path:"+d).join(' ');
//     //     setSearchPath(updatedpath);
//     //     return updatedpath;
//     // }

//     const getUpdatedSearchPath=(_srchpaths?:Map<string,string[]>)=>{

//         let allpaths:string[]=[];
//         let _searchpaths:Map<string,string[]>;
//         if(_srchpaths) _searchpaths=_srchpaths;
//         else _searchpaths=new Map([...tsearchpaths]);
       
//         Array.from(_searchpaths.keys()).forEach(s=>{
//             // if(_searchpaths.get(s).length==0)
//             // {
//             //     // allpaths.push(s);
//             //     allpaths.push("site:"+s);
//             // }
//             // else
//             // {
//             //     // allpaths=allpaths.concat(_searchpaths.get(s));
//             //     allpaths=allpaths.concat(_searchpaths.get(s).map(p=>`listid:${p}`));
//             // }

//             // allpaths.push("site:"+s);
//             allpaths.push("site:"+"\""+s+"\"");
 
//             // allpaths.push("path:"+s);
//             allpaths=allpaths.concat(_searchpaths.get(s).map(p=>`listid:${p}`));

//             // if(_searchpaths.get(s).length!=0)
//             // {
//             //    allpaths=allpaths.concat(_searchpaths.get(s).map(p=>`listid:${p}`));
//             // }
//         })

//         // let updatedpath=(allpaths.length>0)?allpaths.map(d=>"path:"+d).join(' '):("path:"+props.context.pageContext.site.absoluteUrl);
//         let updatedpath= (allpaths.length>0)?allpaths.join(" "):("site:"+props.context.pageContext.site.absoluteUrl);
//         // let updatedpath= (allpaths.length>0)?allpaths.join(" "):("path:"+props.context.pageContext.site.absoluteUrl);
//         setSearchPath(updatedpath);
//         return updatedpath;

//     }

//     const getDeepestPaths=(paths: string[]): string[]=> {
//         const getDepth = (path: string): number => {
//             // Remove the protocol and domain, if present, and count the segments
//             const cleanedPath = path.replace(/^https?:\/\/[^/]+/, "");
//             return cleanedPath.split("/").filter(Boolean).length;
//         };
    
//         const groupedByRoot: Record<string, string[]> = {};
    
//         // Group paths by their root (protocol + domain)
//         paths.forEach((path) => {
//             const root = path.match(/^https?:\/\/[^/]+/)?.[0] || ""; // Extract root
//             if (!groupedByRoot[root]) {
//                 groupedByRoot[root] = [];
//             }
//             groupedByRoot[root].push(path);
//         });
    
//         const deepestPaths: string[] = [];
    
//         // Find the deepest path in each group
//         Object.values(groupedByRoot).forEach((group) => {
//             let maxDepth = 0;
//             let deepestGroup: string[] = [];
    
//             group.forEach((path) => {
//                 const depth = getDepth(path);
//                 if (depth > maxDepth) {
//                     maxDepth = depth;
//                     deepestGroup = [path];
//                 } else if (depth === maxDepth) {
//                     deepestGroup.push(path);
//                 }
//             });
    
//             deepestPaths.push(...deepestGroup);
//         });
    
//         return deepestPaths;
//     }

//     return (
//         <div id="wrapper" ref={elementRef}>
//             <div className="app-menu" id="myHeader">
//                 <VerticalSideBar _context={sp} />
//             </div>

//             <div className="content-page">
//                 <HorizontalNavbar _context={sp} siteUrl={props.siteUrl} context={props.context}  />
//                 <div className="content" style={{ marginLeft: `${!useHide ? '80px' : '230px'}`, marginTop: '1.5rem' }}>
//                     <section className='container-fluid'>
//                         <div className='row'>
//                             <header className="p-3 pb-0 col-12">
//                                 {/* <h1 style={{fontSize:'20px', fontWeight:'600'}} className='text-dark'>Search</h1> */}
//                                 <form>
//                                     <div className="input-group">
//                                         <input style={{ padding: '.75rem .75rem' }} type="text" className="form-control" placeholder="Search Files.." onChange={searchTextChangeHandler} onKeyDown={handleKeyPress} value={searchText}/>
//                                         <button style={{ fontSize: '1rem' }} className="btn btn-primary" type="button" onClick={searchClickHandler}>Advance Search</button>
//                                     </div>
//                                 </form>
//                             </header>

//                             <main className="mt-3 p-0 pb-3 col-12">
//                                 <div className='row mt-0 p-0'>
//                                     <section style={{paddingLeft:'25px'}}className='col-sm-3 nwsera'>
//                                         <div  style={{background:'transparent', border:'0px solid #ccc',width:'95%',}}  className='card'>
//                                         <h5 style={{textAlign:'left', fontSize:'18px',borderBottom:'1px solid #ccc',  margin:'inherit'}} className='text-dark p-0 pt-0 mb-2 fw-bold'>Filter</h5> 
//                                         <DMSSearchFilterCheckedDropDown context={props.context} 
                                        
//                                         onMultiFieldSelect={selfields => {
//                                             let selfld = [...searchQueryRefiners];
//                                             // console.log('fieldtype', fldtype);
//                                             // selfld.push(selfields.map(s=>s.field));
//                                             selfld=selfld.concat(selfields.map(s=>s.field));

//                                             setSearchQueryRefiners(selfld);
//                                             let flds = [...searchfields];
//                                             // flds.push({ fieldname: fld, fieldtype: fldtype });
//                                             flds=flds.concat(selfields.map(s=>({fieldname:s.field,fieldtype:s.fieldtype})));
//                                             setsearchfields(flds);
//                                             runSearch(searchText, getSearchFilter(), searchPath, selfld, getSearchRefineFiltersArray());
//                                         }} 
                                        
//                                         // onMultiFieldSelect={selfields => {
//                                         //     let selfld = [...searchQueryRefiners];
//                                         //     let flds = [...searchfields];
                                        
//                                         //     // Extract fields and field objects from selfields
//                                         //     const selectedFields = selfields.map(s => s.field);
//                                         //     const selectedFieldObjects = selfields.map(s => ({ fieldname: s.field, fieldtype: s.fieldtype }));
                                        
//                                         //     // Add new fields to searchQueryRefiners
//                                         //     selfld = selfld.filter(field => selectedFields.includes(field));
//                                         //     selfld = selfld.concat(selectedFields.filter(field => !selfld.includes(field)));
                                        
//                                         //     // Add new field objects to searchfields
//                                         //     flds = flds.filter(fld => selectedFields.includes(fld.fieldname));
//                                         //     flds = flds.concat(selectedFieldObjects.filter(fld => !flds.some(existingFld => existingFld.fieldname === fld.fieldname)));
                                        
//                                         //     setSearchQueryRefiners(selfld);
//                                         //     setsearchfields(flds);
//                                         //     runSearch(searchText, getSearchFilter(), searchPath, selfld, getSearchRefineFiltersArray());
//                                         // }}

//                                         onSiteSelect={(sites)=>{
//                                           console.log("Selected istes",sites)
//                                           let _searchpaths=new Map([...tsearchpaths]);
                                          
//                                           sites.forEach(s=>{
//                                              if(!_searchpaths.has(s.SiteURL))//if(!allkeys.some(s1=>s.SiteURL.toLowerCase()==s1.toLowerCase()))
//                                              {
//                                                 _searchpaths.set(s.SiteURL,[]);
//                                              }
//                                           })
//                                           let allkeys=Array.from(_searchpaths.keys());
//                                           let siteurls=Array.from(sites.map(s=>s.SiteURL));
//                                           allkeys.forEach(s=>{
//                                             if(!siteurls.some(p=>p==s))//if(!allkeys.some(s1=>s.SiteURL.toLowerCase()==s1.toLowerCase()))
//                                             {
//                                                _searchpaths.delete(s);
//                                             }
//                                           })

//                                           settSearchPath(_searchpaths);

//                                           let newserachpath=getUpdatedSearchPath(_searchpaths);//UpdatedSearchPaths(s.map(s1=>({selepath:s1.SiteURL,typeofpath:enumPathType.site })))
//                                           runSearch(searchText, getSearchFilter(), newserachpath, searchfields.map(sf=>sf.fieldname), getSearchRefineFiltersArray());

//                                         }} 
                                        
//                                         onLibrarySelect={(libs)=>{
//                                             let _searchpaths=new Map([...tsearchpaths]);
//                                             let allkeys=Array.from(_searchpaths.keys());
                                            
//                                             libs.forEach(l=>{
//                                                 let liburl=rootpath+l.FolderPath;
//                                                 //let liburl=l.ListID;
//                                                 let libsite= allkeys.find(a=>liburl.toLowerCase().startsWith(a.toLowerCase()))
//                                                 if(libsite)
//                                                     {
//                                                         let oldlibs= _searchpaths.get(libsite);
//                                                         // if(!oldlibs.some(s=>s==liburl)) oldlibs.push(liburl);
//                                                         if(!oldlibs.some(s=>s==l.ListID)) oldlibs.push(l.ListID);
//                                                         _searchpaths.set(libsite,oldlibs);
//                                                     }
//                                             })

//                                         //    let alllibpaths= libs.map(l=>rootpath+l.FolderPath);
//                                           let alllibpaths= libs.map(l=>l.ListID);
//                                             allkeys.forEach(k=>{
//                                                let tlibs= _searchpaths.get(k);   
//                                                let newtlib=tlibs;                                         
//                                                tlibs.forEach(tlib=>{
//                                                  if(!alllibpaths.some(al=>al==tlib))
//                                                  {
//                                                     newtlib.splice(newtlib.indexOf(tlib));
//                                                  }

//                                                })
//                                                _searchpaths.set(k,newtlib);
//                                             })
                                            
//                                             settSearchPath(_searchpaths);
//                                             let newserachpath=getUpdatedSearchPath(_searchpaths);//UpdatedSearchPaths(s.map(s1=>({selepath:s1.FolderPath,typeofpath:enumPathType.library })))
//                                             runSearch(searchText, getSearchFilter(), newserachpath, searchfields.map(sf=>sf.fieldname), getSearchRefineFiltersArray());

//                                         }}/>
//                                         </div>
//                                         <div style={{width:'95%'}}>
//                                             <h5 style={{textAlign:'left', fontSize:'18px',borderBottom:'1px solid #ccc'}} className='text-dark  mb-2 fw-bold'>Refiners</h5>
//                                             <div className='row p-0'>
//                                                 {searchRefiners?.map(refiner => (
//                                                     <div  style={{border:'0px solid #1fb0e5', borderRadius:'0px', background:'transparent'}} key={refiner.field} className="card col-12 mb-3 pt-2">
//                                                         <h6 style={{textAlign:'left', fontSize:'16px', margin:'inherit'}} className='mt-0 mb-1'>{GetFieldName(refiner.field)}</h6>

//                                                         <div className="form-check">
//                                                             {
//                                                                 (GetFieldType(refiner.field) == enumfieldtype.DateTime) ?
//                                                                     (<div className='col-12'>
//                                                                         <div className='col'><label htmlFor="startDate" className="form-label mb-1 mt-1">Start Date</label></div>
//                                                                         <div className="col">
//                                                                         <input
//                                                                             type="date"
//                                                                             id="startDate"
//                                                                             className="form-control"
//                                                                             value={startDate.get(refiner.field) || ""}
//                                                                             onChange={(e) => {
//                                                                                 // Create a new Map to ensure immutability
//                                                                                 const updatedStartDate = new Map([...startDate]);
//                                                                                 updatedStartDate.set(refiner.field, e.target.value);
//                                                                                 setStartDate(updatedStartDate);
//                                                                                 let dtfilter=generateFilter(refiner.field,e.target.value);

//                                                                                 const updatedSearchFilter = new Map([...searchFilter]);
//                                                                                 updatedSearchFilter.set(refiner.field, dtfilter);
//                                                                                 setSearchFilter(updatedSearchFilter);
//                                                                                 // handleFilterChange();
//                                                                             }}
//                                                                         />
                                                                            
//                                                                         </div>
//                                                                         <div className="col">  <label htmlFor="endDate" className="form-label mb-1 mt-1">End Date</label> </div>
//                                                                         <div className="col">

//                                                                         <input
//                                                                             type="date"
//                                                                             id="endDate"
//                                                                             className="form-control"
//                                                                             value={endDate.get(refiner.field) || ""}
//                                                                             onChange={(e) => {
//                                                                                 // Create a new Map to ensure immutability
//                                                                                 const updatedEndDate = new Map([...endDate]);
//                                                                                 updatedEndDate.set(refiner.field, e.target.value);
//                                                                                 setEndDate(updatedEndDate);
//                                                                                 let dtfilter=generateFilter(refiner.field,null,e.target.value);

//                                                                                 const updatedSearchFilter = new Map([...searchFilter]);
//                                                                                 updatedSearchFilter.set(refiner.field, dtfilter);
//                                                                                 setSearchFilter(updatedSearchFilter);
//                                                                             }}
//                                                                         />
//                                                                         </div>
//                                                                         {/* <div className="col mt-1 mb-3">
//                                                                             <button type='button' className="btn btn-primary float-end mt-2 mb-3" onClick={() => ClearFilter(refiner.field)}>
//                                                                                 Clear
//                                                                             </button>
//                                                                         </div> */}
//                                                                     </div>)
//                                                                     :
//                                                                     <>
//                                                                     <FilterCheckBox refiner={refiner} handleCheckboxChange={handleCheckboxChange} searchRefinerFilters={searchRefinerFilters} />
//                                                                     <div style={{clear:'both', position:'relative'}} className="col mt-1 mb-3">
//                                                                         <button type='button' className="btn btn-secondary newbuttons float-end mt-2 mb-3 me-2" onClick={() => ClearRefinerFilter(refiner.field)}>
//                                                                             Clear
//                                                                         </button>
//                                                                     </div>
//                                                                     </>
                                                                    
                                                                            
//                                                             }
//                                                         </div>
//                                                     </div>
//                                                 ))}
//                                                 {(searchRefiners && searchRefiners.length>0)?<div className="col">
//                                                     <button type='button' className="btn btn-success neewd" onClick={() => handleApplyFilters()}>
//                                                         Apply Filters
//                                                     </button>
//                                                 </div>:<></>}
//                                             </div>
//                                         </div>
//                                     </section>
//                                     <section className="col-sm-9">
//                                         {/* {searchResult.map(res => (
//                                             <div className="col mt-1 mb-1 search-result">
//                                                 <div className="card h-100">
//                                                     <div className="card-body">
//                                                         <h5 className="card-title">{res.Title}</h5>
//                                                         <p className="card-text">{res.Summary}</p>
//                                                     </div>
//                                                     <TagsComponent tags={Object.entries(res.Properties).map(([key, value]) =>`${fieldnamesmapping[key]?fieldnamesmapping[key]:key}: ${value}` )}/>
//                                                 </div>
//                                             </div>
//                                             // <div>{res.Title}</div>
//                                         ))} */}
//                                         {

//                                         (searchResult && searchResult.length>0)?<SearchResultsWithPagination searchResult={searchResult.map(s=>({Title:s.Title,Summary:s.Summary,Properties:s.Properties,Path:s.Path} as SearchResult))} fieldnamesmapping={fieldnamesmapping} fieldtypemappings={searchfields} />:<div>No Results</div>
                                        
//                                         }
//                                     </section>
//                                 </div>
//                             </main>
//                         </div>
//                     </section>
//                 </div>
//             </div>
//         </div>
//     );
// }

// const DMSMain: React.FC<IAdvancedSearchProps> = (props) => {
//     return (
//         <Provider>
//             <AdvancedSearch {...props} />
//         </Provider>
//     );
// }

// export default DMSMain;

export interface IListingState {  
  edItm: any;
  items: any[];  
  showform:boolean;
  process:string;
  siteUrl: string;
  currentPage: number;
  itemsPerPage: number;
  visiblePageStart: number;
  totalItems: number;
  sortColumn: string;
  loading:boolean;
  sortDirection: string;
  searchValues: { [key: string]: string };
}
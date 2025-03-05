export interface IListingState {  
  edItm: any;
  items: any[];  
  showform:boolean;
  process:string;
  siteUrl: string;
  currentPage: number;
  itemsPerPage: number;
  totalItems: number;
  sortColumn: string;
  sortDirection: string;
  searchValues: { [key: string]: string };
}
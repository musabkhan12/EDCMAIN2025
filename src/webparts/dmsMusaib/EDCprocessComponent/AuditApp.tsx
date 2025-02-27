// import * as React from 'react';
// import type { IAuditAppProps } from './IAuditAppProps';
// import { IState } from './IState';
// import "@pnp/sp/site-users/web";
// // import { BrowserRouter as Router, Route, Link, Navigate, Routes, HashRouter } from 'react-router-dom';
// import {Listing} from './ListingComponent/Listing';
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/files";
// import "@pnp/sp/folders";
// import { Stack, StackItem } from 'office-ui-fabric-react';
// import { Navigation } from './NavComponent/Nav';
// import { FormComponent } from './FormComponent/Form';
// import { EditComponent } from './EditComponent/EditComponent';

// export default class AuditApp extends React.Component<IAuditAppProps, IState> {
 
//   constructor (props : IAuditAppProps, state:  IState){
//     super(props);
//     }
   
//   public render(): React.ReactElement<IAuditAppProps> {   
//     <Router>
       
//     <Link to={'/Reports'} className="nav-link">My Route New Details</Link>
//     <Link to={'/New'} className="nav-link">My Route New </Link>
// </Router>
//         return (
//           <HashRouter >
//             <Stack horizontal>
             
             
//             </Stack>
//             <StackItem>
//             <Routes>
//             <Route path='/'  element={(props:any)=><FormComponent  userDisplayName={this.props.userDisplayName}  {...props} userid={this.props.userid} context={this.props.context}/>} />
//             <Route path='/listing'  element={()=><Listing userid={this.props.userid} context={this.props.context}/>} />
//             <Route path='/:type/:postId/:itmId'  element={(props:any)=><FormComponent  userDisplayName={this.props.userDisplayName}  {...props}  userid={this.props.userid}  context={this.props.context}/>} />        
//             <Route path='/:type/:postId'   element={(props:any)=><FormComponent  userDisplayName={this.props.userDisplayName}  {...props}  userid={this.props.userid}  context={this.props.context}/>} />        
//             <Navigate to='/' />
//                   </Routes>
//             </StackItem>
          
//           </HashRouter>
      
        
       
//     );

    
//   }
// }
// // import React from "react";
// // // import { HashRouter, Routes, Route, Navigate, Link } from "react-router-dom";
// // import { Stack, StackItem } from "@fluentui/react";
// // import {FormComponent} from "../EDCprocessComponent/FormComponent/Form"
// // import {Listing}  from "../EDCprocessComponent/ListingComponent/Listing";

// // export default class AuditApp extends React.Component<IAuditAppProps, IState> {
// //   constructor(props: IAuditAppProps) {
// //     super(props);
// //   }

// //   public render(): React.ReactElement<IAuditAppProps> {
// //     return (
// //       <HashRouter>
// //         {/* Navigation Links */}
// //         <nav>
// //           <Link to="/Reports" className="nav-link">My Route New Details</Link>
// //           <Link to="/New" className="nav-link">My Route New</Link>
// //         </nav>

// //         {/* Main Content */}
// //         <Stack horizontal>
// //           <StackItem>
// //             <Routes>
// //               <Route 
// //                 path="/new" 
// //                 element={<FormComponent userDisplayName={this.props.userDisplayName} userid={this.props.userid} context={this.props.context} />} 
// //               />
// //               <Route 
// //                 path="/listing" 
// //                 element={<Listing userid={this.props.userid} context={this.props.context} />} 
// //               />
// //               <Route 
// //                 path="/:type/:postId/:itmId" 
// //                 element={<FormComponent userDisplayName={this.props.userDisplayName} userid={this.props.userid} context={this.props.context} />} 
// //               />
// //               <Route 
// //                 path="/:type/:postId" 
// //                 element={<FormComponent userDisplayName={this.props.userDisplayName} userid={this.props.userid} context={this.props.context} />} 
// //               />

// //               {/* Redirect unknown routes to home */}
// //               <Route path="*" element={<Navigate to="/" />} />
// //             </Routes>
// //           </StackItem>
// //         </Stack>
// //       </HashRouter>
// //     );
// //   }
// // }

import * as React from 'react';
import * as ReactDom from 'react-dom';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// import ArgPocMain2 from './components/DMStest';
import ArgPocMain2 from './components/DMSAdmin';
import EntityMapping from './components/EntityMapping';
import Devision from './components/Division';
import Department from './components/Department';
import DMSAdmin from './components/DMSAdmin';
// import DMSMaster from './components/DMSMaster'
import DMSMain from './components/DMSMain'
import { getSP } from "./loc/pnpjsConfig";
import { IDmsMusaibProps } from './components/IDmsMusaibProps';
import DMSMyApproval from './components/MyApprovals';
import DMSMyApprovalAction from './components/ApprovalAction';
export interface IDmsMusaibWebPartProps {
  description: string;
}
// import BasicForm from './components/CreateEntity'
export default class DmsMusaibWebPart extends BaseClientSideWebPart<IDmsMusaibWebPartProps> {


  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  public render(): void {
    const element: React.ReactElement<IDmsMusaibProps> = React.createElement(
       DMSMain,
       //DMSMyApproval,
      // DMSMyApprovalAction,
      // DMSMaster,
      // ArgPocMain2,
      // EntityMapping,
      // Devision,
      // Department,
      // BasicForm,
       //DMSAdmin,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        currentUserID: this.context.pageContext.legacyPageContext["userId"],
      }
    );

    ReactDom.render(element, this.domElement);
  }
  // public render(): void {
  //   const element: React.ReactElement<IDmsMusaibProps> = React.createElement(
  //     ArgPocMain2,
  //     {
  //       description: this.properties.description,
  //       isDarkTheme: this._isDarkTheme,
  //       environmentMessage: this._environmentMessage,
  //       hasTeamsContext: !!this.context.sdks.microsoftTeams,
  //       userDisplayName: this.context.pageContext.user.displayName,
  //       context: this.context,
  //       siteUrl: this.context.pageContext.web.absoluteUrl,
  //     }
  //   );

  //   ReactDom.render(element, this.domElement);
  // }

  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

 
}

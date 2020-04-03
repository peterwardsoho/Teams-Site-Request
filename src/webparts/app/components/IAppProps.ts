import { CompactPeoplePicker, IPersonaProps, IBasePickerSuggestionsProps } from 'office-ui-fabric-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from "@microsoft/sp-http";


export interface IAppProps {
  description: string;
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  // defaultValues?: IPersonaProps[];
  // multi?: boolean;
}

export interface IAppState {
  addUsers: string[];
  siteTitle:string;
  siteTitle_Error:boolean;
  description:string;
  tenentURL:string;
  template:string;
  sucessMessage:boolean;
  errorMessage:boolean;
}

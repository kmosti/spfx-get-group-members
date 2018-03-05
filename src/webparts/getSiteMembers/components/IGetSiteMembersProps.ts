import { IPropertyPaneDropdownProps } from '@microsoft/sp-webpart-base';

export interface IGetSiteMembersProps {
  description?: string;
  siteGroup?: number;
  groupTitle?: string;
}

export interface IGetSiteMembersState {
  loading?: boolean;
  error?: string;
  groupTitle?: string;
  results?: IGroupMember[];
  showError?: boolean;
}

export interface IGroup {
  Id?: number;
  Title?: string;
}

export interface IGroupMember {
  Id?: number;
  Title?: string;
  email?: string;
}

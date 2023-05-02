export interface IVacationRequestStates {
  IsLoading: boolean;
  CompanyDepartments: any;
  isFutuerData: boolean;
  popoverOpen: boolean;
  isBiggerThenStartData: boolean;
  isPolicyVacationChacke: boolean;
  isManager: string;
  siteGroups: any;
  requestValues: itemObject;
  approvalData: approvalObject;
}
export interface itemObject {
  currentUser: any;
  currentGroup: any;
  Url: any;
  UrlDescription: any;
  RequestDate: any;
  CompanyManagerId: number;
  CompanyDepartmenId: number;
  DepartmentManagerId: number;
  FromDate: Date;
  ToDate: Date;
  numberOfDays: number;
  haveVacationDays: string;
  vacationDuringActiveProject: string;
  vacationDuringActiveProjectEX: string;
  policyVacationChacke: boolean;
}
export interface approvalObject {
  /* Department Manager Approval Status */
  approvalDManagerStatus: string;

  /* Company Manager Approval Status */
  approvalManagerStatus: string;
}

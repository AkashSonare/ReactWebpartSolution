import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
export interface IStateCard{    
    fileschoice: any;
    sheetchoice: any;
    dpselectedItem?: { key: string | string | undefined };    
    dpsheetItem?: { key: string | string | undefined };    
    dpselectedItems: IDropdownOption[];
    dpsheetselectedItems: IDropdownOption[];
    sheetname: string;
    filename: string;
    disabled: boolean;
    checked: boolean;
}

export interface IStateInsert{
    selectedItems: any[];
    name: string;
    description: string;
    address: string;
    permenantaddress: string;
    permenantpincode: string;
    pincode: string;
    dpselectedItem?: { key: string | number | undefined };
    dpselectedcity?: { key: string | number | undefined };
    dpselectedstate?: {key: string | number | undefined};
    dpselectedpermstate?: {key: string | number | undefined};
    dpselectedpermcity?: {key: string | number | undefined};
    disablepermaddress: false;
    dpcity: IDropdownOption[];
    dpStates: IDropdownOption[];
    dppermcity: IDropdownOption[];
    termKey?: string | number;
    dpselectedItems: IDropdownOption[];
    disableAddressToggle: boolean;
    defaultaddresscheck: boolean;
    disableToggle: boolean;
    defaultChecked: boolean;
    pplPickerType:string;
    userIDs: number[];
    userManagerIDs: number[];
    usermanageremailid: string[];
    hideDialog: boolean;
    status: string;
    isChecked: boolean;
    showPanel: boolean;
    required:string;
    onSubmission:boolean;
    termnCond:boolean;
    cityid: string;
    stateid: string;
    statepermid: string;
    citypermid: string;
}
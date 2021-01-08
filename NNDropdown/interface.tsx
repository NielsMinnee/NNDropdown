import { IInputs } from "./generated/ManifestTypes";

export interface Globals {
    context: ComponentFramework.Context<IInputs>;
    enableLogging: boolean;   
}

export interface Setting {
    primaryEntityId: string;
    primaryEntityName: string;
    primaryFieldName: string;
    relationShipName: string;
    relationShipEntityName : string;
    targetEntityName: string;    
    targetEntityFilter: string; 
}

export interface EntityReference {
    entityType: string;
    id: string;
}

export interface DropDownOption {
    key: string;
    text: string;
}

export interface DropDownData{
    allOptions: DropDownOption[];
    selectedOptions: string[];
}
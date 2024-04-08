import * as React from 'react';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import * as operations from './operations';
import { EntityReference, Setting, DropDownData } from "./interface";
import { IInputs } from './generated/ManifestTypes';
import { useState } from 'react';
import { inputProperties } from '@fluentui/react';

const dropdownStyles: Partial<IDropdownStyles> = { title: { border: 'none', backgroundColor: '#F5F5F5' }};

export const NNDropdownControl: React.FC<{ context: ComponentFramework.Context<IInputs>, setting: Setting, dropdowndata: DropDownData }> = ({ context, setting, dropdowndata }) => {

  console.log("In render Control", setting, dropdowndata);

  let targetRef: EntityReference = { entityType: setting.primaryEntityName, id: setting.primaryEntityId }; 

  const [selectedKeys, setSelectedKeys] = useState<string[]>([]);

  const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, index: number): void => {

    if (item) {

      setSelectedKeys(
        item.selected ? [...selectedKeys, item.key as string] : selectedKeys.filter(key => key !== item.key),
      );

      if (item.selected === true) {
        operations._writeLog('Associate this item', item);
        let relatedRef: EntityReference = { entityType: setting.targetEntityName, id: item.key.toString() };
        operations._associateRecord(context, setting, targetRef, relatedRef); 
      }
      else if (item.selected === false) {
        operations._writeLog('Disassocatie this item', item);
        let relatedRef: EntityReference = { entityType: setting.targetEntityName, id: item.key.toString() };
        operations._disAssociateRecord(context, setting, targetRef, relatedRef);
      }
    }
  };

  return (
    <Dropdown
      placeholder="---" //Equal to PowerPlatformStandard  //Select options 
      //label="Multi-select controlled example" // No using a label, Power Platform Provides Label
      //selectedKeys={selectedKeys}     
      //@ts-ignore
      onChange={onChange}
      multiSelect
      defaultSelectedKeys={dropdowndata.selectedOptions} //selectedOptions  
      options={dropdowndata.allOptions} //allOptions 
      styles={dropdownStyles} 
      disabled={context.mode.isControlDisabled}
    />
  );
};
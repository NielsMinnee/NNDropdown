import * as React from 'react';
//import { Dropdown, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import * as operations from './operations';
import { EntityReference, Setting, DropDownData } from "./interface";
import { IInputs } from './generated/ManifestTypes';
import { useState } from 'react';

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { } }; //width: 300

export const NNDropdownControl: React.FC<{ context: ComponentFramework.Context<IInputs>, setting: Setting, dropdowndata: DropDownData }> = ({ context, setting, dropdowndata }) => {

  console.log("In render Control", setting, dropdowndata);

  let targetRef: EntityReference = { entityType: setting.primaryEntityName, id: setting.primaryEntityId }; //operations.setting.primaryEntityName

  const [selectedKeys, setSelectedKeys] = useState<string[]>([]);

  const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, index: number): void => {

    if (item) {

      setSelectedKeys(
        item.selected ? [...selectedKeys, item.key as string] : selectedKeys.filter(key => key !== item.key),
      );

      if (item.selected === true) {
        operations._writeLog('Associate this item', item);
        let relatedRef: EntityReference = { entityType: setting.targetEntityName, id: item.key.toString() };
        operations._associateRecord(context, setting, targetRef, relatedRef); //operations.globals.context
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
      //label="Multi-select controlled example"
      //selectedKeys={selectedKeys}     
      //@ts-ignore
      onChange={onChange}
      multiSelect
      defaultSelectedKeys={dropdowndata.selectedOptions} //selectedOptions  
      options={dropdowndata.allOptions} //allOptions 
      styles={dropdownStyles}
    />
  );
};
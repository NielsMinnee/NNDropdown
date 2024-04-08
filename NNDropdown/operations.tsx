import { IInputs } from "./generated/ManifestTypes";
import * as webAPIHelper from "./webAPIHelper";
import { EntityReference, Setting, DropDownOption, DropDownData, Globals } from "./interface";
import * as ReactDOM from 'react-dom';
import * as React from 'react'
import * as dropdown from './fluentUIDropdown';

export async function _sleep(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export function _writeLog(message: string, data?: any) {
  const condition = true;
  if (condition) { //Needs to be set back to a shared value. Had issue when the control was on the form multiple times. Settings/Globals were mixed.
    console.log(message, data);
  }
}

export async function _getAvailableOptions(context: ComponentFramework.Context<IInputs>, setting: Setting) {
  const baseFetchXml = `<fetch><entity name="${setting.targetEntityName}" /></fetch>`;
  const userFetchXml = setting.targetEntityFilter ? setting.targetEntityFilter : "";
  const fetchXml = userFetchXml != "" ? userFetchXml : baseFetchXml;
  _writeLog("Using this FetchXML for all available options", fetchXml);

  const allOptionsSet = await webAPIHelper.retrieveDataFetchXML(context, setting.targetEntityName, fetchXml);
  _writeLog("Retrieved Data RAW allOptionsSet", allOptionsSet);

  let allOptions: Array<DropDownOption> = new Array;

  allOptionsSet.entities.forEach(entity => {
    let item: DropDownOption = { key: entity[`${setting.targetEntityName}id`], text: entity[setting.primaryFieldName] }
    allOptions.push(item);
  });

  allOptions.sort((a,b) => (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0))

  return allOptions;
}

export async function _currentOptions(context: ComponentFramework.Context<IInputs>, setting: Setting) {
  const fetchXml: string =
    `<fetch>
      <entity name="${setting.relationShipEntityName}" >
        <filter>
          <condition attribute="${setting.primaryEntityName}id" operator="eq" value="${setting.primaryEntityId}" />
        </filter>
      </entity>
    </fetch>`;

  _writeLog("Using this FetchXML for all selected options", fetchXml);

  const currentOptionsSet = await webAPIHelper.retrieveDataFetchXML(context, setting.relationShipEntityName, fetchXml);
  _writeLog("Retrieved Data RAW currentOptionsSet", currentOptionsSet);

  let currentOptions: Array<string> = new Array;

  currentOptionsSet.entities.forEach(entity => {
    currentOptions.push(entity[`${setting.targetEntityName}id`]);
  });

  return currentOptions;
}

export function _proccessSetting(context: ComponentFramework.Context<IInputs>) {
  let setting: Setting = {
    //@ts-ignore 
    primaryEntityId: context.page.entityId ? context.page.entityId : "",
    //@ts-ignore 
    primaryEntityName: context.page.entityTypeName ? context.page.entityTypeName : "",
    primaryFieldName: context.parameters.primaryfieldname.raw ? context.parameters.primaryfieldname.raw : "",
    relationShipName: context.parameters.relationshipname.raw ? context.parameters.relationshipname.raw : "",
    relationShipEntityName: context.parameters.relationshipentityname.raw ? context.parameters.relationshipentityname.raw : "",
    targetEntityName: context.parameters.targetentityname.raw ? context.parameters.targetentityname.raw : "",
    targetEntityFilter: context.parameters.targetentityfilter.raw ? context.parameters.targetentityfilter.raw : "",
  }
  return setting;
}

export function _processGlobals(context: ComponentFramework.Context<IInputs>) {
  let globals = {
    context: context,
    enableLogging: context.parameters.enableLogging.raw?.toLowerCase() === "true" ? true : false
  }
  return globals;
}

export function _associateRecord(context: ComponentFramework.Context<IInputs>, setting: Setting, targetEntityReference: EntityReference, relatedEntityReference: EntityReference) {
  return new Promise(function (resolve, reject) {

    let manyToManyAssociateRequest = {
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,
        operationName: "Associate"
      }),

      relationship: setting.relationShipName,
      target: targetEntityReference,
      relatedEntities: [relatedEntityReference] //Array to be able to associate multiple
    }

    _writeLog('Created Associate Request Object', manyToManyAssociateRequest);

    //@ts-ignore
    context.webAPI.execute(manyToManyAssociateRequest).then(
      (success: any) => {
        _writeLog("Success", success);
        resolve(success);
      },
      (error: any) => {
        _writeLog("Error", error);
        reject(error);
      }
    )
  })
}

export function _disAssociateRecord(context: ComponentFramework.Context<IInputs>, setting: Setting, targetEntityReference: EntityReference, relatedEntityReference: EntityReference) {
  return new Promise(function (resolve, reject) {

    let manyToManyDisassociateRequest = {
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,
        operationName: "Disassociate"
      }),

      relationship: setting.relationShipName,
      target: targetEntityReference,
      relatedEntityId: relatedEntityReference.id
    }

    _writeLog('Created Disassociate Request Object', manyToManyDisassociateRequest);

    //@ts-ignore
    context.webAPI.execute(manyToManyDisassociateRequest).then(
      (success: any) => {
        _writeLog("Success", success);
        resolve(success);
      },
      (error: any) => {
        _writeLog("Error", error);
        reject(error);
      }
    )
  })
}

export async function _execute(context: ComponentFramework.Context<IInputs>, container: HTMLDivElement) {

  const _globals = _processGlobals(context);
  _writeLog("Retrieved and Set Globals", _globals);

  const _setting = _proccessSetting(context);
  _writeLog("Retrieved Settings", _setting);

  if (_setting.primaryEntityId) {

    const dropDownData: DropDownData = {
      allOptions: await _getAvailableOptions(context, _setting),
      selectedOptions: await _currentOptions(context, _setting),
    }

    _writeLog("Retrieved DropdownData", dropDownData);

    ReactDOM.render(React.createElement(dropdown.NNDropdownControl, { context: context, setting: _setting, dropdowndata: dropDownData }), container);
  }
  else {
    //@ts-ignore
    const msg = <div>This record hasn't been created yet. To enable this control, create the record.</div>;
    ReactDOM.render(msg, container);
  }
}
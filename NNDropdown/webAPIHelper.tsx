import { IInputs } from "./generated/ManifestTypes";
import { _writeLog } from "./operations";

export let retrieveDataFetchXML = (context: ComponentFramework.Context<IInputs>, entityName: string, fetchXML: string) => {
    return new Promise<ComponentFramework.WebApi.RetrieveMultipleResponse>(function (resolve, reject) {

        fetchXML = encodeURIComponent(fetchXML); //Filters with a like condition includes %, this causes an error.

        context.webAPI.retrieveMultipleRecords(entityName, `?fetchXml=${fetchXML}`).then(
            function success(result) {                
                _writeLog('Fetch Success', result);
                resolve(result);
            },
            function (error) {
                console.log('Fetch Error', error);
                _writeLog('Fetch Error', error);
                reject(error);
            }
        )
    });
}   
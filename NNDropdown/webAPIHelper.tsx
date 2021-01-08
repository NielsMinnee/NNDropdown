import { IInputs } from "./generated/ManifestTypes";

export let retrieveDataFetchXML = (context: ComponentFramework.Context<IInputs>, entityName: string, fetchXML: string) => {
    return new Promise<ComponentFramework.WebApi.RetrieveMultipleResponse>(function (resolve, reject) {
        context.webAPI.retrieveMultipleRecords(entityName, fetchXML).then(
            function success(result) {
                resolve(result);
            },
            function (error) {
                reject(error);
            }
        )
    });
}   
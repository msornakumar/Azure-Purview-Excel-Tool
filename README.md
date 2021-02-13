# Azure-Purview-Excel-Tool

Azure Purview is cloud-native Fully managed PaaS service to address discovery and compliance needs.Once the discovery is done by setting up scanning , users can search the metadata using Azure Purview Studio by providng search criteria. Right now there is no feature available to export the results to an excel file.This tool writen in Python will export the search results from purview to excel using rest API.

# How to use the tool
## Prerequisites

1) Create an Azure Purview Account. Setup and run the scans to bring in the Asset metadata.

2) Create a Service priciple in Azure Active Directory. Create a client credential on the Service Priciple. Store the Service Principle ID and Secret it in a Azure Key vault. Make sure to give access to the SPN to the Purview account.

    Refer to below link on how to create a service principle account.

    https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-service-principal-portal


    Refer to Prerequisite section of the below link

    https://docs.microsoft.com/en-us/azure/purview/tutorial-using-rest-apis


3) Install python 

    Download and Install Python from below link

    https://www.python.org/downloads/

4) Install below packages

    pip install azure-identity
    pip install azure-keyvault-secrets
    pip install XlsxWriter
    pip install argparse

<br/>

## How to run the tool

## Parameters

### <ins>To be setup in the script file and saved before running</ins>

|Parameter|Type|Description|
|----------|----|-----------|
|KeyVaultURI|Manadatory|Keyvault URI in which Service account SPNId and Password is stored|
|KeyVaultNameForPurviewAppSPNID|Manadatory|Keyvault Secret Name in which Service account SPNId is stored|
|KeyVaultNameForPurviewAppSPNSecret|Manadatory|Keyvault URI in which Service account is stored|
|PurviewTenantID|Manadatory|TenantID in which Purview Account is deployed|
|PurviewColumnAssetType|Optional|The column sheet will be generated for the below Asset Type only. In case if column data is missing for any table Asset type , Try adding the asset type here.|

<br/>


### <ins>To be provided in command line or setup in script and saved before the run</ins>

|Parameter|Type|Default|CommandLine Switch                    |Description|
|----------|----|----|--------------|-----------|
|PurviewAccountName|Manadatory|No Default|-P or --PurviewAccountName|Purview Account Name|
|PurviewSearchKeyWords|Optional|*|-S or --SearchKeyword |Keywords to be searched. Mutiple Keywords can be provided comma seperated. Use * for search all.If left Blank , * will be used|
|PurviewExcelOutputPath|Optional|Blank|-E or --ExcelOutPath|Path in which Excel Output to be return. If left blank , Will be written to current path|
|PurviewRestAPILimit|Mandatory|50|-R or --RestAPILimit|This is API limit to send guids as bulk to get the entity details. Incase if the Entitites API is very slow or failing this can be made as a lower number.|

### <ins> Running from Command Line </ins>

1) Navigate to the path of the script and run command with filename with the required parameter.



### <ins> Running from Visual Code </ins>

1) Follow the below tutorial to setup Code for Python

    https://code.visualstudio.com/docs/python/python-tutorial

2) Install the Azure Account Addin for the Visual Code from below link and login using the account which has permission to Purview

    https://marketplace.visualstudio.com/items?itemName=ms-vscode.azure-account#:~:text=Commands%20%20%20%20Command%20%20%20,your%20Azure%20subscription.%20%205%20more%20rows

3) Run the script after setting up the parameters mentioned above.


Note : While running the script , Ignore the below message as this error will not stop the execution.

EnvironmentCredential.get_token failed: EnvironmentCredential authentication unavailable. Environment variables are not fully configured.
ManagedIdentityCredential.get_token failed: ManagedIdentityCredential authentication unavailable, no managed identity endpoint found.
SharedTokenCacheCredential.get_token failed: SharedTokenCacheCredential authentication unavailable. No accounts were found in the cache.



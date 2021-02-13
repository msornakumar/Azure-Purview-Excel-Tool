import os
import cmd
import sys
from azure.keyvault.secrets import SecretClient
from azure.identity import DefaultAzureCredential
import requests
import xlsxwriter
import json
import argparse


############################### Command Line Parameters or can be set in the file itself #######################################################################################
PurviewAccountName = ""         # Purview Account Name 
PurviewSearchKeyWords = ""      # Keywords to be searched. Mutiple Keywords can be provided comma seperated. Use * for search all.If left Blank , * will be used
PurviewExcelOutputPath = ""     # Path in which Excel Output to be return. If left blank , Will be written to current path
PurviewRestAPILimit = 50        # This is API limit to send guids as bulk to get the entity details. Incase if the Entitites API is very slow or failing this can be 
                                #     made as a lower number.

########################################### Purview Settings to be done in file ################################################################################################
KeyVaultURI = f""                           # Keyvault URI in which Service account SPNId and Password is stored
KeyVaultNameForPurviewAppSPNID = ""         # Keyvault Secret Name in which Service account SPNId is stored
KeyVaultNameForPurviewAppSPNSecret = ""     # Keyvault URI in which Service account is stored
PurviewTenantID = ""                        # TenantID in which Purview Account is deployed


########################################### Optional Setting ################################################################################################

# The column sheet will be generated for these Asset Type only. 
# In case if column data is missing for any table Asset type , Try adding the asset type here.

PurviewColumnAssetType = [                                              
'AZURE_DATA_EXPLORER_TABLE',
'AZURE_DATALAKE_GEN1_PATH',
'AZURE_DATALAKE_GEN1_RESOURCE_SET',
'AZURE_DATALAKE_GEN2_PATH',
'AZURE_DATALAKE_GEN2_RESOURCE_SET',
'AZURE_SQL_DW_TABLE',
'AZURE_SQL_TABLE',
'AZURE_MARIADB_TABLE',
'AZURE_MYSQL_TABLE',
'AZURE_POSTGRESQL_TABLE',
'AZURE_SQL_DW_TABLE',
'AZURE_SQL_MI_TABLE',
'AZURE_SYNAPSE_DEDICATED_SQL_TABLE',
'AZURE_SYNAPSE_SERVERLESS_SQL_TABLE',
'AZURE_TABLE',
'HBASE_TABLE',
'HIVE_TABLE',
'MSSQL_TABLE',
'RDBMS_TABLE',
'SAP_ECC_TABLE',
'SAP_S4HANA_TABLE',
'TERADATA_TABLE',
]

############################### Command Line Parameters or can be set in the file itself ######################################################


 
# Initialize parser
parser = argparse.ArgumentParser()
parser.add_argument("-P", "--PurviewAccountName", help = "PurviewAccountName")
parser.add_argument("-S", "--SearchKeyword", help = "SearchKeyword")
parser.add_argument("-E", "--ExcelOutPath", help = "ExcelOutPath")
parser.add_argument("-R", "--RestAPILimit", help = "RestAPILimit")
args = parser.parse_args()

if args.PurviewAccountName is not None:
        PurviewAccountName = args.PurviewAccountName

if args.SearchKeyword is not None:
        PurviewSearchKeyWords = args.SearchKeyword

if args.ExcelOutPath is not None:
        PurviewExcelOutputPath = args.ExcelOutPath

if args.RestAPILimit is not None:
        PurviewRestAPILimit = args.RestAPILimite


#################################################### Validate Inputs #####################################


if len(KeyVaultURI) == 0 :
    print('Please configure the KeyVaultURI in the script')
    sys.exit()

if len(KeyVaultNameForPurviewAppSPNID) == 0 :
    print('Please configure the KeyVaultNameForPurviewAppSPNID in the script')
    sys.exit()

if len(KeyVaultNameForPurviewAppSPNSecret) == 0 :
    print('Please configure the KeyVaultNameForPurviewAppSPNSecret in the script')
    sys.exit()

if len(PurviewTenantID) == 0 :
    print('Please configure the PurviewTenantID in the script')
    sys.exit()


if len(PurviewAccountName) == 0 :
    print('Please input the Purview Account Name')
    sys.exit()



try:
    #################################################### Get SPN ID & Secret from Key Valut #####################################
    
    try :
        credential = DefaultAzureCredential()
    except Exception as e:
        ResultOutput =  'Error Occured :' + str(e)
        print(ResultOutput)

    
    client = SecretClient(vault_url=KeyVaultURI, credential=credential)
    PurviewAppSPNID = client.get_secret(KeyVaultNameForPurviewAppSPNID).value
    PurviewAppSPNSecret = client.get_secret(KeyVaultNameForPurviewAppSPNSecret).value

    

    #################################################### Get Oath Token #########################################################
    url = "https://login.microsoftonline.com/" + PurviewTenantID + "/oauth2/token"

    payload="grant_type=client_credentials&client_id=" + PurviewAppSPNID + " &client_secret=" + PurviewAppSPNSecret + "&resource=73c2949e-da2d-457a-9607-fcc665198967"
    headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
    }

    OathTokenresponse = requests.request("POST", url, headers=headers, data=payload)
    OathTokenresponseJSON = OathTokenresponse.json()

    OathBearerToken = OathTokenresponseJSON["access_token"]

    print("")
    print("")
    print("")

    #################################################### Create excel and write the header ################################################
    if os.path.exists("PurviewOutputExcel.xlsx"):
        os.remove("PurviewOutputExcel.xlsx")

    if len(PurviewExcelOutputPath) == 0 :
        PurviewExcelOutputPath = os.path.dirname(os.path.realpath(__file__))

    xlPath = PurviewExcelOutputPath + '\\PurviewOutputExcel.xlsx'  

    xlWorkBook = xlsxwriter.Workbook(xlPath)

    xlWorkBookHeaderFormat = xlWorkBook.add_format()
    xlWorkBookHeaderFormat.set_bold()
    xlWorkBookHeaderFormat.set_font_color('white')
    xlWorkBookHeaderFormat.set_bg_color('#6495ed')


    xlAssetWorkSheet = xlWorkBook.add_worksheet(name='PurviewAssets')
    xlAssetRowheader =['ID','QualifiedName','Name','Description','EntityType','AssetType','GlossaryTerm','Contact','Owner','Typename']
    xlAssetRow = 0
    xlAssetCol = 0
    xlAssetWorkSheet.write_row(xlAssetRow,xlAssetCol,tuple(xlAssetRowheader),xlWorkBookHeaderFormat)
    xlAssetWorkSheet.set_column('A:K', 20)

    xlAssetColumnSheet = xlWorkBook.add_worksheet(name='PurviewColumns')
    xlAssetColumnRowheader =['TableGuid','TableName','ColumnGuid','QualifiedName','Name','DataType','Length','Precision','Scale','Description','GlossaryTerm','Classification']
    xlAssetColumnRow = 0
    xlAssetColumnCol = 0
    xlAssetColumnSheet.write_row(xlAssetColumnRow,xlAssetColumnCol,tuple(xlAssetColumnRowheader),xlWorkBookHeaderFormat)
    xlAssetColumnSheet.set_column('A:L', 20)

    #################################################### Loop for offset ################################################

    PurviewSearchCount = 0 
    PurviewRestAPIOffset = 0

    while PurviewRestAPIOffset <= PurviewSearchCount:

        #################################################### Connect to Purview & Fire REst API ################################################

        PurviewURL = "https://" + PurviewAccountName + ".catalog.purview.azure.com"
        PurviewRestURL = PurviewURL +  "/api/atlas/v2/search/advanced"
        payload=("{\r\n    \"keywords\": \"" + PurviewSearchKeyWords + "\",\r\n    \"offset\": \"" + str(PurviewRestAPIOffset) + "\",\r\n    \"limit\": \""+ str(PurviewRestAPILimit) + "\"\r\n}")
        headers ={}
        headers ['Authorization'] = "Bearer " + OathBearerToken
        headers ['Content-Type'] = "application/json"

        response = requests.request("POST", PurviewRestURL, headers=headers, data=payload)
        PurviewSearchresponseJSON = response.json()


        if PurviewSearchCount == 0 :
            PurviewSearchCount  = PurviewSearchresponseJSON["@search.count"]
            if PurviewSearchCount == 0 :
                ResultOutput = 'No results found for the keyword'
                print(ResultOutput)
                sys.exit()
            else :
                ResultOutput =  str(PurviewSearchCount) + ' results found for the keyword'
                print(ResultOutput)


        ResultOutput =  'Processing ' + str(PurviewRestAPIOffset) + '-' + str(PurviewRestAPIOffset + PurviewRestAPILimit) + ' of ' +  str(PurviewSearchCount) 
        print(ResultOutput)


        ######################################################################################################################################################
        # Write Row by Row to the file
        ######################################################################################################################################################

        PurviewAssetGuids = ""
        Name = ""

        for PurviewAssetProperties in PurviewSearchresponseJSON["value"] :
            
            Id = PurviewAssetProperties["id"]

            
            Name = Name + PurviewAssetProperties["name"] + ","

            QualifiedName = PurviewAssetProperties["qualifiedName"]
            # print (Id + " == " + QualifiedName)

            Name = PurviewAssetProperties["name"]
            Description = PurviewAssetProperties["description"]
            Owner = PurviewAssetProperties["owner"]
            EntityType = PurviewAssetProperties["entityType"]

            #print(PurviewColumnAssetType.count(EntityType.upper()))

            #if not EntityType.upper().startswith("ADF_") :
            if PurviewColumnAssetType.count(EntityType.upper()) !=0 :
                PurviewAssetGuids = PurviewAssetGuids + Id + "&guid="

            ArraySeperator =','
            Classification = ArraySeperator.join(PurviewAssetProperties["classification"])
            if len(PurviewAssetProperties["assetType"]) > 0 :
                AssetType = PurviewAssetProperties["assetType"][0]
            else :
                AssetType = ""

            GlossaryTerm = ""    

            for AssetTerm in PurviewAssetProperties["term"]:
                GlossaryTerm = GlossaryTerm + AssetTerm["name"] +  ","
            
            GlossaryTerm = GlossaryTerm[:-1]

            Contact = ""

            # for Contact in PurviewAssetProperties["term"]:
            #     ContactExpert = Contact[""]

            xlAssetDetailsRow =[Id,QualifiedName,Name,Description,EntityType,AssetType,GlossaryTerm,Contact,Owner]
            
            xlAssetRow += 1
            xlAssetWorkSheet.write_row(xlAssetRow,xlAssetColumnCol,tuple(xlAssetDetailsRow))    

        PurviewAssetGuids = PurviewAssetGuids[:-6]

        if len(PurviewAssetGuids) != 0 :

            PurviewRestURL = PurviewURL +  "/api/atlas/v2/entity/bulk?guid=" + PurviewAssetGuids
            headers ={}
            headers ['Authorization'] = "Bearer " + OathBearerToken
            headers ['Content-Type'] = "application/json"
            payload=("{}")

            

            RetryCount = 0
            PurviewrestAPIStatusCode = 200

            while RetryCount <= 3 :
                response = requests.request("GET", PurviewRestURL, headers=headers, data=payload)
                PurviewrestAPIStatusCode = response.status_code
                if response.status_code == 200 :
                    RetryCount = 0
                    break
                else :
                    # print(Name)
                    # print(PurviewRestURL)
                    # print(ResultOutput)
                    RetryCount += 1
                    ResultOutput = "Rest API call returned status code " + str(PurviewrestAPIStatusCode) + '.Running retry attempt ' +  str(RetryCount)  
                    print(ResultOutput)


            if RetryCount >= 3 :
                xlWorkBook.close()
                ResultOutput = "Error : Three retry attempts failed for the API. Please change the filter to retrieve less number of output or modify the PurviewRestAPILimit parameter to reduce the limit"
                print(ResultOutput)
                sys.exit()
            else : 
                PurviewGuidSearchresponseJSON = response.json()
            
            

            for AssetColumnGuid in PurviewGuidSearchresponseJSON["referredEntities"]:
                #print(json.dumps(PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['relationshipAttributes']))
                Typename = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['typeName']

                if "COLUMN" in Typename.upper() :

                    TableGuid = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['relationshipAttributes']['table']['guid']
                    TableName = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['relationshipAttributes']['table']['displayText']
                    ColumnGuid = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['guid']
                    QualifiedName = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["qualifiedName"]
                    Name = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["name"]
                    DataType = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["data_type"]
                    Length = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["length"]
                    Precision = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["precision"]
                    Scale = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["scale"]
                    Description = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['attributes']["description"]
                    
                    Meanings = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['relationshipAttributes']['meanings']
                    GlossaryTerm = ""
                    for MeaningsProperties in Meanings:
                        GlossaryTerm = GlossaryTerm + MeaningsProperties["displayText"] + ","
                    GlossaryTerm = GlossaryTerm[:-1]

                    
                    ClassificationAssociated =""
                    #print(PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid])
                    if 'classifications' in PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid] :
                        Classifications = PurviewGuidSearchresponseJSON["referredEntities"][AssetColumnGuid]['classifications']
                        
                    
                        for Classification in Classifications:
                            ClassificationAssociated = ClassificationAssociated + Classification["typeName"] + ","

                        ClassificationAssociated = ClassificationAssociated[:-1]


                    xlAssetColumnRowDetails =[TableGuid,TableName,ColumnGuid,QualifiedName,Name,DataType,Length,Precision,Scale,Description,GlossaryTerm,ClassificationAssociated]
                    xlAssetColumnRow += 1

                    xlAssetColumnSheet.write_row(xlAssetColumnRow,xlAssetColumnCol,tuple(xlAssetColumnRowDetails))    

        PurviewRestAPIOffset = PurviewRestAPIOffset + PurviewRestAPILimit 

    xlWorkBook.close()

    ResultOutput =  'Processing Completed.'
    print(ResultOutput)

except Exception as e:
    try :
        xlWorkBook.close()
    except :
        test = 1
    ResultOutput =  'Error Occured :' + str(e)
    print(ResultOutput)

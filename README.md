# cucm_soap

## Overview

The aim of this script is to leverage data in Excel tables to make AXL SOAP requests to Cisco Unified Communications Manager and IMP

CUC is not currently supported due to the syntax being a bit different

It is intended to be a command line tool, with parameters passed to it via arguments

## Capabilities

-   Send an AXL SOAP request per row in an Excel table
-   Allow for the capabilities of Excel to be utilized in the background for data formatting, the script utilizes the
actual values in the cells
-   Create Excel tables and populate them with information returned from AXL SOAP requests
-   Preview mode

## Setup

-   Download and install python 3
-   Download and install git
-   Downloading venv and setting up a virtual environment is recommended but not strictly required
-   Clone this repository
```
git clone https://github.com/sz-hunor/cucm_soap
```
-   Install python script dependencies
```
pip install -r requirements.txt
```
-   Download the AXLSoap.xsd and AXLSoap.xsd files from CUCM and place them in the same directory as the script

## Excel Syntax and Structure

The AXL SOAP requests use JSON like structures as the bodies of the requests, these can contain items such as
nested objects, lists and ```None``` values, Excel on the other hand is a flat data structure.
The script uses special syntax to help map the complex JSON like structures to Excel and vice-versa.
  
-   The first row in the Excel sheet maps to the elements in the SOAP request body

-   The values in each row map to the values of these elements

The script will make a separate AXL SOAP call per row in the Excel table


The intention was to keep the values in the data rows free of any special syntactical requirements to keep populating
them with data as straight forward as possible
-   There is one exception to this and that is between empty cells and cells with the ```None``` value:
    - The JSON like structures in the AXL SOAP requests have many elements that are optional, these element can be fully
omitted. An empty cell has this effect, any elements with no value in the Excel table will simply not be included in the 
SOAP request body
    - There is also the case where an element needs to contain the value ```None```, this is to over-write existing data in an
element. As an example, to delete the description of a phone, a description with the value ```None``` can be sent. The syntax
in the Excel table is to actually write the word "none" in the cell

To map nested objects the ":" symbol can be used in the header, to denote the nested structure

-   For example, the column with the header "top_level:nesting_1:nesting_2:nesting_n" with the value "value" would
be interpreted as the JSON structure as:

```json
{
   "top_level":{
      "nesting_1":{
         "nesting_2":{
            "nesting_n":"value"
         }
      }
   }
}
```

Having multiple rows with the same name will get the values of these rows combined into a list

-   For example, a row with header "Header" and value "Value_1" followed by "Header" and "Value_2", "Header" and "Value_3"
would be interpreted as the JSON structure as:

```json
{
   "Header":[
      "Value_1",
      "Value_2",
      "Value_3"
   ]
}
```

The last special syntax is for forcing values to be lists, this is useful for cases where a list with a single item
needs to be created. This is accomplished by surrounding the header with square brackets

-   For example, a row with header "[Header]" and value "Value_1" would be interpreted as the JSON structure as:

```json
{
   "Header":[
      "Value_1"
   ]
}
```

-   The square brackets can be used for headers of rows with identical names as well that would already be merged into
lists, adding square brackets around these headers has no impact. This is the format the export function will write headers
in, if it is a list in the JSON object, it gets square brackets in Excel

For an example Excel file to help visualize some of this simply send a get or list request that will return a fully formatted
Excel file with actual data

Example getUser request: ```python soap_cucm.py --cucm=<value> --user=<value> --pass=<value> --req_json="{'name':'<username>'}" --request=getUser --output=<value>```

Running this will also print out the returned JSON object, that can then be contrasted to the data in the Excel to better
understand the syntax


## CUCM AXL schema

The CUCM AXL schema is contained in the AXLSoap.xsd and AXLSoap.xsd file while enum type element options are contained
in the AXLEnums.xsd file

These files can be opened with tools such as SoapUI, which allow for listing all the AXL SOAP requests available as well
as how the body of these requests needs to be constructed

An CUCM AXL schema can also be found here: https://developer.cisco.com/docs/axl-schema-reference/


Using the methods in the Syntax and Structure section, a valid request body needs to be described. However, most body
elements are optional and can simply be omitted, and it's possible to construct simpler calls, test them and add
complexity

Another option is to use one of the get or list requests and save the return to an Excel file, this will automatically
construct the headers in the correct format and the values can then be simply modified

## AXL Versioning

AXL API requests can specify the AXL schema version in the request header ```SOAPAction: CUCM:DB ver=14.0``` as well as in the 
XML namespace URI for the SOAP Envelope element within body of the request ```xmlns:ns="http://www.cisco.com/AXL/API/14.0"```. 
Both SOAPAction header and the XML namespace URI should indicate the same version. SOAPAction request header is optional 
but if supplied it supersedes the version specified XML namespace URI. AXL schema version in the XML namespace URI is 
always required.

The version specified in the AXL API request indicates which AXL schema version the request payload will follow and the 
response should follow.

CUCM maintains backwards compatibility with the running release minus 2 versions

The versioning for this script is done by simply copying the correct .wsdl file into the same directory as the script
or using the ```--wsdl=<file path and name>``` parameter to specify the correct wsdl file for the version of CUCM targeted

## CUCM AXL Performance

The AXL SOAP Service has dynamic throttling that is always enabled. Upon receiving an AXL write request, the CUCM 
publisher node via Cisco Database Layer Monitor service dynamically evaluates all database change notification queues 
across all nodes in the cluster and if any one node has more then 1500 change notification messages pending in its database 
queues, the AXL write requests will be rejected with a "503 Service Unavailable" response code. Even if the CUCM Cluster 
is keeping up with change notification processing and DB queues are NOT exceeding a depth of 1500, only 1500 AXL Write 
requests per minute are allowed.

The database change notifications queue can be monitored via the following CUCM Performance Counter on each node 
(\\cucm\DB Change Notification Server\QueuedRequestsInDB). This counter can be viewed using the Real Time Monitoring 
Tool (RTMT).

AXL read requests are NOT throttled even while write requests are being throttled.

In addition to AXL requests throttling, the following AXL query limits are always enforced: A single request is limited 
to <8MB of data. Concurrent requests are limited to <16MB of data.

More information can be found here: https://developer.cisco.com/docs/axl/#!axl-developer-guide/data-throttling-and-performance


## Usage and options

The script is meant to be used from the command line with options supplied as arguments

The arguments need to be supplied in the ```-<argument> <value>``` format for short notation and in ```--<argument>=<value>```
for long notation, arguments are case-sensitive

python cucm_soap.py ```-<argument>=<value> -<argument>=<value> ...```

### Options

-   -c, --cucm
    -   The IP or FQDN of the CUCM
    -   The script will automatically encapsulate this in ```"https://"<value>":8443/axl/"```, if this is not the correct AXL
interface, modify the script

-   -u, --user
    -   A user on the CUCM with **Standard AXL API Access** role assigned to it

-   -p, --pass
    -   The password of the user

-   -e, --excel
    -   Path and filename of the input Excel file

-   -v, --verify
    -   Controls weather the server certificate is verified or not
    -   It is an optional argument and if not present server certificate will not be verified
    -   If present it can either take the value "true", this verifies the server certificate against the default trust 
anchors or it can be the path and filename to a .PEM or CRT file, this verifies the server certificate against the trust
anchor specified in the file
    -   The .PEM or .CRT file must be a bundle, meaning it contains a root CA and an intermediate CA and must be in 
X.509 ASCII format

-   -s, --sheet
    -   Name of the sheet to use from the Excel document

-   -r, --request
    -   Name of the AXL SOAP request, this is case-sensitive and needs to be the exact name of the AXL request described
in the AXLAPI.wsdl file

-   -j, --req_json
    -   Instead of creating a request from an Excel file, it is possible to simply manually type out the request body in
JSON format
    -   Single quotes must be used around every key and value in the JSON object
    -   Encapsulate the entire JSON object in double quotes if it contains spaces

-   -o, --output
    -   Path and filename of the output Excel file

-   -w, --wsdl
    -   Path and filename of the AXLAPI.wsdl file, this argument does not need to be provided if the AXLAPI.wsdl is in the
same directory as the script

-   -x, --xsd
    -   Path and filename of the AXLSoap.xsd file, this argument does not need to be provided if the AXLSoap.xsd is in the
same directory as the script

-   -p, --preview
    -   This is a switch and takes no argument

## Modes

The script can be used in two "modes":

-   The mode where an actual AXL SOAP request is sent to a CUCM, for this the mandatory parameters are
    -   ```python soap_cucm.py --cucm=<value> --user=<value> --pass=<value> --excel=<value> --sheet=<value> --request=<value>```
    - or if using the ```--req_json``` argument:
    - -   ```python soap_cucm.py --cucm=<value> --user=<value> --pass=<value> --req_json=<value> --request=<value>```
    -   ```--output=<value>``` is an optional parameter to this syntax
    -   ```--verify=("true"|<value>)``` is an optional parameter to this syntax


-   The preview mode, this is meant to have a way to visually check the JSON like object that would get sent as the request
body based on the info in an Excel file. For this the syntax is:
    -   ```python soap_cucm.py --excel=<value> --sheet=<value> --preview```

## Warning

This script can be used to send any arbitrary data as any arbitrary AXL SOAP request, it is completely possibly to
over-write, delete or otherwise damage existing configuration

Use caution, use test labs, use test data and familiarize yourself with the quirks of SOAP

## Useful examples
Retreiving certificate related data:
```python cucm_soap.py --cucm=<value> --user=<value> --pass=<value> --req_json="{'sql':'select certificate.servername, typecertificateservice.name, certificate.serialnumber, certificate.certificate, certificate.subjectname, certificate.issuername from certificate inner join certificateservicecertificatemap on certificate.pkid = certificateservicecertificatemap.fkcertificate inner join typecertificateservice on certificateservicecertificatemap.tkcertificateservice = typecertificateservice.enum'}" --request=executeSQLQuery --output=<value>```

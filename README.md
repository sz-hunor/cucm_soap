# cucm_soap

## Overview

-   The aim of this script is to leverage data in Excel tables to make AXL SOAP requests to Cisco Unified Communications Manager
-   It is intended to be a command line tool, with parameters passed to it via arguments

## Capabilities

-   Send an AXL SOAP request per row in an Excel table
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

## Excel Syntax

-   The AXL SOAP requests use JSON like structures as the bodies of the requests, these can contain items such assssssssssssssssssssssssssssssssssssssssssssssssssssss
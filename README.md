# HelloID-Conn-Prov-Target-ExchangeOnline

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.

<p align="center">
    <img src="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline/blob/main/Logo.png?raw=true">
</p>

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [HelloID-Conn-Prov-Target-ExchangeOnline](#helloid-conn-prov-target-exchangeonline)
  - [Table of Contents](#table-of-contents)
  - [Supported features:](#supported-features)
  - [Introduction](#introduction)
  - [Getting started](#getting-started)
    - [Requirements](#requirements)
      - [App Registration \& Certificate Setup](#app-registration--certificate-setup)
      - [HelloID-specific configuration](#helloid-specific-configuration)
      - [Convert .pfx to base64 string](#convert-pfx-to-base64-string)
    - [Connection settings](#connection-settings)
    - [Correlation configuration](#correlation-configuration)
    - [Available lifecycle actions](#available-lifecycle-actions)
    - [Field mapping](#field-mapping)
  - [Remarks](#remarks)
  - [Development resources](#development-resources)
    - [API endpoints](#api-endpoints)
    - [API documentation](#api-documentation)
      - [Installing the Microsoft Exchange Online PowerShell module](#installing-the-microsoft-exchange-online-powershell-module)
      - [Creating the Entra ID App Registration and certificate](#creating-the-entra-id-app-registration-and-certificate)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Supported features:
| Feature                             | Supported | Actions                                 | Remarks                                                                                                                                        |
| ----------------------------------- | --------- | --------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| **Account Lifecycle**               | ✅         | Create, Update, Enable, Disable, Delete | Create is for correlating, Enable/Disable for hidefromaddresslist, update for editing mailbox attributes and delete for out of office messages |
| **Permissions**                     | ✅         | Retrieve, Grant, Revoke                 | Groups, shared mailboxes, folder permissions, litigation hold, regional configuration                                                          |
| **Resources**                       | ✅         | -                                       | For groups and shared mailboxes                                                                                                                |
| **Entitlement Import: Accounts**    | ✅         | -                                       |                                                                                                                                                |
| **Entitlement Import: Permissions** | ✅         | -                                       | Only for the shared mailboxes and groups scripts. ⚠️Warning the new shared mailbox scripts are not backwards compatible!⚠️                       |

## Introduction
For this connector we have the option to correlate to and/or update Exchange Online (Office 365) users and/or mailboxes and provision permission(s) to a group and/or shared mailbox.
  >**Only Exchange and Cloud-only groups are supported**

If you want to create Exchange Online (Office 365) users and/or mailboxes, please use the built-in Microsoft (Entra ID) Active Directory target system. Or setup Business Rules to provision an Office 365 license group, Microsoft will automatically provision a mailbox for this user.

## Getting started

### Requirements

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID, an App Registration. During the setup process, you’ll create a new App Registration in the Entra portal, assign the necessary API permissions (such as user and group read/write), and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:
- [App-only authentication with certificate (Exchange Online)](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

1. **HelloID Environment**:
   - Set up your _HelloID_ environment.
   - Install the _HelloID_ Provisioning agent **On-Premises**.
2. **Microsoft Exchange Online PowerShell Module**:
   - [Download link](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
     - When timeouts occur while using the `Get-EXOMailbox` command (e.g., when retrieving permissions), downgrading to `v3.4.0` may be a solution.
     - [Microsoft documentation](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)
3. **Graph API Credentials**:
   - Create an **App Registration** in Microsoft Entra ID.
   - Add API permissions for your app:
     - **Application permissions**:
       - From the **Request API Permissions** screen click `Office 365 Exchange Online`.
          > _The Office 365 Exchange Online might not be a selectable API. In this case, select "APIs my organization uses" and search here for "Office 365 Exchange Online"_
       - `Exchange.ManageAsApp`: Manage Exchange As Application.
   - **Certificate:**
       - Upload the public key file (.cer) in Entra ID.
4. **Assign Entra ID roles to the application**:
   - The **Exchange Administrator** role is required for some operations.  
   - For most mailbox and group management tasks, the **Exchange Recipient Administrator** role is sufficient.
   - Examples:  
     - **Manage shared mailboxes** → Exchange Recipient Administrator  
     - **Manage distribution lists** → Exchange Recipient Administrator  
     - **Manage mail-enabled security groups** → Exchange Administrator (required only if using `BypassSecurityGroupManagerCheck` with `Add-DistributionGroupMember`)  

#### Convert .pfx to base64 string
HelloID requires a base64 string to import the certificate. With the example below, it is possible to create a base64 string

```Powershell
$filePath = 'C:\Cert'
$pfxCertName = 'Cert.pfx'
$pfxPath = "$filePath\$pfxCertName"

$fileContentBytes = [System.IO.File]::ReadAllBytes("$pfxPath")
[System.Convert]::ToBase64String($fileContentBytes) | Set-Content "$filePath\HelloID_Cert_Base64.txt"
```

### Connection settings
The following settings are required to connect.

| Setting                                | Description                                                                                | Mandatory |
| -------------------------------------- | ------------------------------------------------------------------------------------------ | --------- |
| Entra ID Organization                  | The name of the organization to connect to and where the Entra ID App Registration exists. | Yes       |
| Entra ID Tenant ID                     | Id of the Entra ID tenant                                                                  | Yes       |
| Entra ID App Id                        | The Application (client) ID of the Entra ID App Registration with Exchange Permissions     | Yes       |
| Entra ID App Certificate Base64 String | The certificate converted to a base64 string                                               | Yes       |
| Entra ID App Certificate Password      | The certificate password                                                                   | Yes       |


> [!IMPORTANT]
> Please note: You must use the primary .onmicrosoft.com domain of the organization. Using anything else may lead to inconsistent results.

### Correlation configuration

The correlation configuration is used to specify which properties will be used to match an existing account within _Exchange Online_ to a person in _HelloID_.

| Setting                   | Value                                     |
| ------------------------- | ----------------------------------------- |
| Enable correlation        | `True`                                    |
| Person correlation field  | `Accounts.<yourSystem>.userPrinicpalName` |
| Account correlation field | `userPrinicpalName`                       |

> [!TIP]
> _For more information on correlation, please refer to our correlation [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems/correlation.html) pages_.

### Available lifecycle actions

The following lifecycle actions are available:

| Action                                      | Description                                                                                      |
| ------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| correlateOnly - create.ps1                  | Correlate to an account                                                                          |
| create.ps1                                  | Correlate account and sets Hide from address list to mapped value (default true), only if mapped |
| delete.ps1                                  | Set mailbox auto reply configuration (only when none is configured)                              |
| disable.ps1                                 | Sets Hide from address list to mapped value (default true)                                       |
| enable.ps1                                  | Sets Hide from address list to mapped value (default false)                                      |
| update.ps1                                  | Sets custom attribute(s)                                                                         |
| sharedMailboxes - permissions.ps1           | List sharedMailboxes as permissions                                                              |
| sharedMailboxes - grantPermission.ps1       | Grant sharedMailbox membership to an account                                                     |
| sharedMailboxes - revokePermission.ps1      | Revoke sharedMailbox membership from an account                                                  |
| sharedMailboxes - resources.ps1             | Create sharedMailboxes from resources                                                            |
| sharedMailboxes - subPermissions.ps1        | Grant/Revoke sharedMailbox membership from an account                                            |
| groups - permissions.ps1                    | List distribution groups as permissions                                                          |
| groups - grantPermission.ps1                | Grant distribution group membership to an account                                                |
| groups - revokePermission.ps1               | Revoke distribution group membership from an account                                             |
| groups - resources.ps1                      | Create distribution groups from resources                                                        |
| groups - subPermissions.ps1                 | Grant/Revoke distribution group membership from an account                                       |
| folderPermission - permissions.ps1          | Mailbox folder permissions settings                                                              |
| folderPermission - grantPermission.ps1      | Grant folder permissions settings                                                                |
| regionalConfiguration - permissions.ps1     | Mailbox regional configuration settings                                                          |
| regionalConfiguration - grantPermission.ps1 | Grant regional configuration settings                                                            |
| configuration.json                          | Default _configuration.json_                                                                     |
| fieldMapping.json                           | Default _fieldMapping.json_                                                                      |

### Field mapping

The field mapping can be imported by using the _fieldMapping.json_ file.

## Remarks

In some cases, Exchange Online takes more than 70 seconds to return an error. For example when using the `Set-MailboxRegionalConfiguration` with an invalid date format. For this reason, if you get the 30 seconds timeout then we recommend testing locally on the agent server.

Added enhanced functionality to update **EmailAddresses (proxy addresses)**. The script now ensures that existing proxy addresses are preserved, and new ones are added with the correct primary (SMTP:) and secondary (smtp:) casing.

## Development resources

### API endpoints

The following endpoints are used by the connector

| Endpoint                          | Description                                   |
| --------------------------------- | --------------------------------------------- |
| Get-EXOMailbox                    | Get a mailbox                                 |
| Set-Mailbox                       | Set a mailbox                                 |
| Add-MailboxPermission             | Add a mailbox to a permission                 |
| Remove-MailboxPermission          | Remove a mailbox from a permission            |
| Add-RecipientPermission           | Add a mailbox to a permission                 |
| New-Mailbox                       | Creates a mailbox                             |
| Remove-RecipientPermission        | Remove a mailbox from a permission            |
| Get-EXORecipient                  | Get mailboxes for permissions                 |
| Get-DistributionGroup             | Get distribution groups for permissions       |
| Add-DistributionGroupMember       | Add a distribution group to a permission      |
| Remove-DistributionGroupMember    | Remove a distribution group from a permission |
| New-DistributionGroup             | Creates a distribution group                  |
| Set-DistributionGroup             | Set a distribution group                      |
| Get-MailboxFolderStatistics       | Get mailbox statistics                        |
| Set-MailboxFolderPermission       | Set mailbox statistics                        |
| Set-MailboxRegionalConfiguration  | Set mailbox regional configuration            |
| Set-MailboxAutoReplyConfiguration | Set mailbox auto reply configuration          |
| Get-MailboxAutoReplyConfiguration | Get mailbox auto reply configuration          |

### API documentation

#### Installing the Microsoft Exchange Online PowerShell module
Since we use the cmdlets from the Microsoft Exchange Online PowerShell module, it is required this module is installed and available for the service account.
Please follow the [Microsoft documentation on how to install the module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module). 

#### Creating the Entra ID App Registration and certificate
_The steps below are based on the [Microsoft documentation](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) as of the moment of release. The Microsoft documentation should always be leading and is susceptible to change. The steps below might not reflect those changes._
>**Please note that our steps differ from the current documentation as we use Access Token Based Authentication instead of Certificate Based Authentication**

## Getting help

> [!TIP]
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems.html) pages_.

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/

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
    - [HelloID Icon URL](#helloid-icon-url)
    - [Requirements](#requirements)
      - [App Registration \& Certificate Setup](#app-registration--certificate-setup)
      - [HelloID-specific configuration](#helloid-specific-configuration)
      - [Convert .pfx to base64 string](#convert-pfx-to-base64-string)
    - [Connection settings](#connection-settings)
    - [Correlation configuration](#correlation-configuration)
    - [Available lifecycle actions](#available-lifecycle-actions)
    - [Permissions](#permissions)
      - [Static permissions](#static-permissions)
      - [Dynamic permissions (sub permissions) and resources](#dynamic-permissions-sub-permissions-and-resources)
      - [Other permissions](#other-permissions)
    - [Field mapping](#field-mapping)
  - [Upgrade notes for v4.0.0](#upgrade-notes-for-v400)
  - [Remarks](#remarks)
  - [Development resources](#development-resources)
    - [Commands and endpoints](#commands-and-endpoints)
    - [API documentation](#api-documentation)
      - [Installing the Microsoft Exchange Online PowerShell module](#installing-the-microsoft-exchange-online-powershell-module)
      - [Creating the Entra ID App Registration and certificate](#creating-the-entra-id-app-registration-and-certificate)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Supported features:
| Feature                             | Supported | Actions                                 | Remarks                                                                                                                                                                  |
| ----------------------------------- | --------- | --------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Account Lifecycle**               | ✅         | Create, Update, Enable, Disable, Delete | Create is for correlating, Enable/Disable for hide from address list, Update for editing mailbox attributes and Delete for setting the out of office (auto reply) state |
| **Permissions**                     | ✅         | Retrieve, Grant, Revoke                 | Distribution groups and mail-enabled security groups, shared mailboxes, folder permissions, out of office and regional configuration                                      |
| **Dynamic permissions**             | ✅         | Grant, Revoke (sub permissions)         | Groups and shared mailboxes, based on the department of the contracts of the person                                                                                      |
| **Resources**                       | ✅         | -                                       | Creates distribution groups and shared mailboxes from the departments in HelloID                                                                                         |
| **Entitlement Import: Accounts**    | ✅         | -                                       | The correlation attribute must also be set in this script. The default value is `employeeId`                                                                             |
| **Entitlement Import: Permissions** | ✅         | -                                       | Groups and shared mailboxes, static and dynamic. ⚠️The shared mailbox scripts in the folder _new separate entitlements_ are not backwards compatible!⚠️                    |

## Introduction
For this connector we have the option to correlate to and/or update Exchange Online (Office 365) users and/or mailboxes and provision permission(s) to a group and/or shared mailbox.
  >**Only Exchange and Cloud-only groups are supported**

If you want to create Exchange Online (Office 365) users and/or mailboxes, please use the built-in Microsoft (Entra ID) Active Directory target system. Or setup Business Rules to provision an Office 365 license group, Microsoft will automatically provision a mailbox for this user.

## Getting started

### HelloID Icon URL

URL of the icon used for the HelloID Provisioning target system.

```
https://raw.githubusercontent.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline/refs/heads/main/Icon.png
```

### Requirements

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID App Registration. During the setup process you create a new App Registration in the Entra portal, assign the necessary API permissions and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:
- [App-only authentication with certificate (Exchange Online)](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

1. **HelloID Environment**:
   - Set up your _HelloID_ environment.
   - Install the _HelloID_ Provisioning agent **On-Premises**. The scripts only run on the agent: they use the _ExchangeOnlineManagement_ PowerShell module.
2. **Microsoft Exchange Online PowerShell module**:
   - Install the [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) module in **Windows PowerShell 5.1** for **all users** (`Install-Module ExchangeOnlineManagement -Scope AllUsers`), so that the service account of the agent can use it.
   - The connector is tested with version `3.11.1`. Older versions are not tested.
   - [Microsoft documentation](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)
3. **App Registration permissions (Entra ID)**:
   - Add API permissions for your app:
     - **Application permissions**:
       - `Exchange.ManageAsApp`: Manage Exchange As Application.
       - `User.Read.All`: Microsoft Graph, used to look up the Entra ID user. ⚠️ This is not the default `User.Read` type `Delegated` ⚠️
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

| Setting                                         | Description                                                                                       | Mandatory |
| ----------------------------------------------- | ------------------------------------------------------------------------------------------------- | --------- |
| Organization                                    | The primary `.onmicrosoft.com` domain of the organization to connect to                           | Yes       |
| App Registration Directory (tenant) ID          | Id of the Entra ID tenant                                                                         | Yes       |
| App Registration Application (client) ID        | The Application (client) ID of the App Registration with the Exchange and Graph permissions       | Yes       |
| App Registration Certificate Base 64 String     | The certificate (.pfx) converted to a base64 string                                               | Yes       |
| App Registration Certificate Password           | The certificate password                                                                          | Yes       |

> [!IMPORTANT]
> Please note: You must use the primary .onmicrosoft.com domain of the organization. Using anything else may lead to inconsistent results.

### Correlation configuration

The correlation configuration is used to specify which properties will be used to match an existing account within _Exchange Online_ to a person in _HelloID_.

| Setting                   | Value                             |
| ------------------------- | --------------------------------- |
| Enable correlation        | `True`                            |
| Person correlation field  | `PersonContext.Person.ExternalId` |
| Account correlation field | `employeeId`                      |

> [!NOTE]
> HelloID will also queries the Entra ID user. This way its possible to use Entra ID attributes to correlate the accounts. 
> **The import account script also needs to be updated if `employeeId` is not you're correlation field.**

> [!TIP]
> _For more information on correlation, please refer to our correlation [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems/correlation.html) pages_.

### Available lifecycle actions

The following lifecycle actions are available:

| Action                                                  | Description                                                                                                  |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| correlateOnly - create.ps1                              | Correlate to an account                                                                                      |
| correlateOnly - import.ps1                              | Import accounts (correlate only)                                                                             |
| create.ps1                                              | Correlate account and set Hide from address list to the mapped value (default true), only if mapped          |
| import.ps1                                              | Import accounts: the mailboxes, matched to the Entra ID users                                                |
| update.ps1                                              | Set custom attribute(s) and email addresses                                                                  |
| enable.ps1                                              | Set Hide from address list to the mapped value (default false)                                               |
| disable.ps1                                             | Set Hide from address list to the mapped value (default true)                                                |
| delete.ps1                                              | Set the mailbox auto reply (out of office) state                                                             |
| groups - permissions.ps1                                | List distribution groups and mail-enabled security groups as permissions                                     |
| groups - grantPermission.ps1                            | Grant group membership to an account                                                                         |
| groups - revokePermission.ps1                           | Revoke group membership from an account                                                                      |
| groups - importPermission.ps1                           | Import the members of the groups                                                                             |
| groups - subPermissions.ps1                             | Grant/Revoke group membership (dynamic, based on the department of the contracts)                            |
| groups - importSubPermissions.ps1                       | Import the members of the dynamic groups                                                                     |
| groups - resources.ps1 (resources/groups)               | Create distribution groups from resources                                                                    |
| sharedMailboxes - permissions.ps1                       | List shared mailboxes as permissions                                                                         |
| sharedMailboxes - grantPermission.ps1                   | Grant shared mailbox permissions (Full Access, Send As) to an account                                        |
| sharedMailboxes - revokePermission.ps1                  | Revoke shared mailbox permissions from an account                                                            |
| sharedMailboxes - importPermission.ps1                  | Import the Full Access permissions of the shared mailboxes                                                   |
| sharedMailboxes - subPermission.ps1                     | Grant/Revoke shared mailbox permissions (dynamic, based on the department of the contracts)                  |
| sharedMailboxes - importSubPermissions.ps1              | Import Full Access, Send As and Send on Behalf of the dynamic shared mailboxes                               |
| sharedMailboxes - resources.ps1 (resources/sharedMailboxes) | Create shared mailboxes from resources                                                                   |
| sharedMailboxes (new separate entitlements)             | Shared mailbox permissions as separate entitlements for Full Access, Send As and Send on Behalf             |
| folderPermission - permissions.ps1 / grantPermission.ps1 | Mailbox folder permissions                                                                                  |
| outOfOffice - permissions.ps1 / grantPermission.ps1 / revokePermission.ps1 | Toggle the out of office reply (see [permissions/outOfOffice/README.md](permissions/outOfOffice/README.md)) |
| regionalConfiguration - permissions.ps1 / grantPermission.ps1 | Mailbox regional configuration                                                                          |
| configuration.json                                      | Default _configuration.json_                                                                                 |
| fieldMapping.json                                       | Default _fieldMapping.json_                                                                                  |

### Permissions

#### Static permissions

`groups` lists the groups that are not synchronized from an on-premises Active Directory (`IsDirSynced -eq False`): distribution groups and mail-enabled security groups. The `sharedMailboxes` scripts grant Full Access and/or Send As (set `$accessRights` at the top of the grant and revoke scripts; Send on Behalf is also supported).

#### Dynamic permissions (sub permissions) and resources

With the sub permissions a person gets the groups and shared mailboxes of the departments of his or her contracts:

- The `resources` scripts create a distribution group (`dg_<name>`) or shared mailbox (`smb_<name>`) per department. The `ExternalId` of the department is stored in `CustomAttribute1` (the correlation attribute), and the marker `HelloID Dynamic Distribution Group` or `HelloID Dynamic Shared Mailbox` in `CustomAttribute2`.
- The sub permission scripts look up the group or mailbox with `CustomAttribute1` = `Contract.Department.ExternalId` and grant or revoke it, based on the contracts that are in conditions.
- The import sub permission scripts use the marker in `CustomAttribute2` to find the dynamic groups and mailboxes (the filter is configurable at the top of the script, an alternative is the display name prefix).
- A resource that was created halfway (without the custom attributes) is recovered in the next run via its email address.

#### Other permissions

- **Folder permissions**, **out of office** and **regional configuration** are set as permissions on the mailbox. The out of office permission needs an extra configuration key, see [permissions/outOfOffice/README.md](permissions/outOfOffice/README.md).
- The folder _sharedMailboxes (new separate entitlements)_ is not backwards compatible with the shared mailbox permissions: it generates separate entitlements for Full Access, Send As and Send on Behalf. Implement it by adding a new permissions tile and selecting the entitlements in the business rules. See the README in that folder.

### Field mapping

The field mapping can be imported by using the _fieldMapping.json_ file.

## Upgrade notes for v4.0.0

> [!WARNING]
> **Version 4.0.0 contains a breaking change.** Groups (distribution groups and mail-enabled security groups) are now referenced by their `ExchangeObjectId` instead of their `Guid`.
>
> - In practice the `ExchangeObjectId` and `Guid` of a group are the same value, so existing entitlements and sub permissions keep working. Verify this in your own tenant before upgrading: compare both properties with `Get-DistributionGroup | Select-Object Guid, ExchangeObjectId`.
> - `grantPermission.ps1`, `revokePermission.ps1` and the sub permissions accept both values.
> - Update all group scripts together (`permissions/groups`, `resources/groups`) so that the listing, imports and grant/revoke use the same reference.
> - Shared mailbox SendAs permissions are retrieved per shared mailbox on the `Guid`. For tenants with many shared mailboxes the import takes longer than before.
> - The scripts are tested with the _ExchangeOnlineManagement_ module in Windows PowerShell 5.1 (the HelloID agent).

## Remarks

In some cases, Exchange Online takes more than 70 seconds to return an error. For example when using the `Set-MailboxRegionalConfiguration` with an invalid date format. For this reason, if you get the 30 seconds timeout then we recommend testing locally on the agent server.

The update script preserves existing **EmailAddresses (proxy addresses)** and adds new ones with the correct primary (`SMTP:`) and secondary (`smtp:`) casing.

Error messages in the audit log are shown as `message Error code: code`. The detailed error is written to the warnings.

## Development resources

### Commands and endpoints

The connector uses the following commands of the _ExchangeOnlineManagement_ module (Connect-ExchangeOnline with certificate-based authentication) and the Microsoft Graph API:

| Command / endpoint                    | Description                                                    |
| ------------------------------------- | -------------------------------------------------------------- |
| Connect-ExchangeOnline                | Connect with the App Registration and certificate              |
| Get-EXOMailbox                        | Get mailboxes                                                  |
| Set-Mailbox                           | Set a mailbox (attributes, Send on Behalf, hide from address list) |
| New-Mailbox                           | Create a shared mailbox (resources)                            |
| Get-EXORecipient                      | Get shared mailboxes                                           |
| Get-EXOMailboxPermission              | Get Full Access permissions (imports)                          |
| Get-EXORecipientPermission            | Get Send As permissions (imports)                              |
| Add-MailboxPermission                 | Grant Full Access                                              |
| Remove-MailboxPermission              | Revoke Full Access                                             |
| Add-RecipientPermission               | Grant Send As                                                  |
| Remove-RecipientPermission            | Revoke Send As                                                 |
| Get-DistributionGroup                 | Get distribution groups and mail-enabled security groups       |
| Get-DistributionGroupMember           | Get group members (imports)                                    |
| Add-DistributionGroupMember           | Add a member to a group                                        |
| Remove-DistributionGroupMember        | Remove a member from a group                                   |
| New-DistributionGroup                 | Create a distribution group (resources)                        |
| Set-DistributionGroup                 | Set a distribution group (custom attributes)                   |
| Get-MailboxFolderStatistics           | Get the mailbox folders (folder permissions)                   |
| Set-MailboxFolderPermission           | Set mailbox folder permissions                                 |
| Set-MailboxRegionalConfiguration      | Set the mailbox regional configuration                         |
| Get-MailboxAutoReplyConfiguration     | Get the out of office configuration                            |
| Set-MailboxAutoReplyConfiguration     | Set the out of office configuration                            |
| GET https://graph.microsoft.com/v1.0/users | Find the Entra ID user (create and import)                |

### API documentation

#### Installing the Microsoft Exchange Online PowerShell module
Since we use the cmdlets from the Microsoft Exchange Online PowerShell module, it is required that this module is installed for all users and available for the service account of the agent (Windows PowerShell 5.1).
Please follow the [Microsoft documentation on how to install the module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module).

#### Creating the Entra ID App Registration and certificate
_The steps below are based on the [Microsoft documentation](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps). The Microsoft documentation should always be leading and is susceptible to change._
>**The connector uses certificate-based authentication: `Connect-ExchangeOnline` for Exchange Online and a signed token (the same certificate) for the Microsoft Graph API.**

## Getting help

> [!TIP]
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems.html) pages_.

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/

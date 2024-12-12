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
  - [Introduction](#introduction)
  - [Getting started](#getting-started)
    - [Prerequisites](#prerequisites)
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
        - [Application Registration](#application-registration)
        - [Configuring App Permissions](#configuring-app-permissions)
        - [Assign Entra ID roles to the application](#assign-entra-id-roles-to-the-application)
        - [Authentication and Authorization](#authentication-and-authorization)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Introduction
For this connector we have the option to correlate to and/or update Exchange Online (Office 365) users and/or mailboxes and provision permission(s) to a group and/or shared mailbox.
  >**Only Exchange and Cloud-only groups are supported**

If you want to create Exchange Online (Office 365) users and/or mailboxes, please use the built-in Microsoft (Entra ID) Active Directory target system. Or setup Business Rules to provision an Office 365 license group, Microsoft will automatically provision a mailbox for this user.

## Getting started

### Prerequisites

1. **HelloID Environment**:
   - Set up your _HelloID_ environment.
   - Install the _HelloID_ Provisioning agent **On-Premises**.
2. **Microsoft Exchange Online PowerShell Module v3.3.0**:
   - [Download link](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.3.0)
     - Using a higher version than `v3.3.0` of the Exchange Online Management module can result in timeouts when using the `Get-EXOMailbox` command.
   - [Microsoft documentation](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)
3. **Graph API Credentials**:
   - Create an **App Registration** in Microsoft Entra ID.
   - Add API permissions for your app:
     - **Application permissions**:
       - From the **Request API Permissions** screen click `Office 365 Exchange Online`.
          > _The Office 365 Exchange Online might not be a selectable API. In this case, select "APIs my organization uses" and search here for "Office 365 Exchange Online"__
       - `Exchange.ManageAsApp`: Manage Exchange As Application.
   - Create access credentials for your app:
     - Create a **client secret** for your app.
4. **Assign Entra ID roles to the application**:
   - The **Exchange Administrator** role should provide the required permissions 

### Connection settings
The following settings are required to connect.

| Setting               | Description                                                                                | Mandatory |
| --------------------- | ------------------------------------------------------------------------------------------ | --------- |
| Entra ID Organization | The name of the organization to connect to and where the Entra ID App Registration exists. | Yes       |
| Entra ID Tenant ID    | Id of the Entra ID tenant                                                                  | Yes       |
| Entra ID App Id       | The Application (client) ID of the Entra ID App Registration with Exchange Permissions     | Yes       |
| Entra ID App Secret   | Secret of the Entra ID App Registration with Exchange Permissions                          | Yes       |

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

| Action                                      | Description                                                |
| ------------------------------------------- | ---------------------------------------------------------- |
| create.ps1                                  | Correlate to an account                                    |
| delete.ps1                                  | Set mailbox auto reply configuration                       |
| disable.ps1                                 | Sets Hide from address list to true                        |
| enable.ps1                                  | Sets Hide from address list to false                       |
| sharedMailboxes - permissions.ps1           | List sharedMailboxes as permissions                        |
| sharedMailboxes - grantPermission.ps1       | Grant sharedMailbox membership to an account               |
| sharedMailboxes - revokePermission.ps1      | Revoke sharedMailbox membership from an account            |
| sharedMailboxes - resources.ps1             | Create sharedMailboxes from resources                      |
| sharedMailboxes - subPermissions.ps1        | Grant/Revoke sharedMailbox membership from an account      |
| groups - permissions.ps1                    | List distribution groups as permissions                    |
| groups - grantPermission.ps1                | Grant distribution group membership to an account          |
| groups - revokePermission.ps1               | Revoke distribution group membership from an account       |
| groups - resources.ps1                      | Create distribution groups from resources                  |
| groups - subPermissions.ps1                 | Grant/Revoke distribution group membership from an account |
| folderPermission - permissions.ps1          | Mailbox folder permissions settings                        |
| folderPermission - grantPermission.ps1      | Grant folder permissions settings                          |
| regionalConfiguration - permissions.ps1     | Mailbox regional configuration settings                    |
| regionalConfiguration - grantPermission.ps1 | Grant regional configuration settings                      |
| litigationHold - permissions.ps1            | Mailbox litigation hold settings                           |
| litigationHold - grantPermission.ps1        | Grant litigation hold settings                             |
| litigationHold - revokePermission.ps1       | Revoke litigation hold settings                            |
| configuration.json                          | Default _configuration.json_                               |
| fieldMapping.json                           | Default _fieldMapping.json_                                |

### Field mapping

The field mapping can be imported by using the _fieldMapping.json_ file.

## Remarks

In some cases, Exchange Online takes more than 70 seconds to return an error. For example when using the `Set-MailboxRegionalConfiguration` with an invalid date format. For this reason, if you get the 30 seconds timeout then we recommend testing locally on the agent server.

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


### API documentation

#### Installing the Microsoft Exchange Online PowerShell module
Since we use the cmdlets from the Microsoft Exchange Online PowerShell module, it is required this module is installed and available for the service account.
Please follow the [Microsoft documentation on how to install the module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module). 

#### Creating the Entra ID App Registration and certificate
_The steps below are based on the [Microsoft documentation](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) as of the moment of release. The Microsoft documentation should always be leading and is susceptible to change. The steps below might not reflect those changes._
>**Please note that our steps differ from the current documentation as we use Access Token Based Authentication instead of Certificate Based Authentication**

##### Application Registration
The first step is to register a new **Entra ID Application**. The application is used to connect to Exchange and to manage permissions.

* Navigate to **App Registrations** in  Entra ID, and select “New Registration” (**Microsoft Entra admin center > Applications > App registrations > New Application Registration**).
* Next, give the application a name. In this example we are using “**ExO PowerShell CBA**” as application name.
* Specify who can use this application (**Accounts in this organizational directory only**).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “**Register**” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

##### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

* To assign your application the right permissions, navigate to **Microsoft Entra admin center > Applications > App registrations**.
* Select the application we created before, and select “**API Permissions**” or “**View API Permissions**”.
* To assign a new permission to your application, click the “**Add a permission**” button.
* From the “**Request API Permissions**” screen click “**Office 365 Exchange Online**”.
  > _The Office 365 Exchange Online might not be a selectable API. In this case, select "APIs my organization uses" and search here for "Office 365 Exchange Online"__
* For this connector the following permissions are used as **Application permissions**:
  *	Manage Exchange As Application ***Exchange.ManageAsApp***
* To grant admin consent to our application press the “**Grant admin consent for TENANT**” button.

##### Assign Entra ID roles to the application
Entra ID has more than 50 admin roles available. The **Exchange Administrator** role should provide the required permissions for any task in Exchange Online PowerShell. However, some actions may not be allowed, such as managing other admin accounts, for this the Global Administrator would be required. and Exchange Administrator roles. Please note that the required role may vary based on your configuration.
* To assign the role(s) to your application, navigate to **Microsoft Entra admin center > Roles and admins**.
* On the Roles and administrators page that opens, find and select one of the supported roles e.g. “**Exchange Administrator**” by clicking on the name of the role (not the check box) in the results.
* On the Assignments page that opens, click the “**Add assignments**” button.
* In the Add assignments flyout that opens, **find and select the app that we created before**.
* When you're finished, click **Add**.
* Back on the Assignments page, **verify that the app has been assigned to the role**.

For more information about the permissions, please see the Microsoft docs:
* [Permissions in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/permissions-exo).
* [Find the permissions required to run any Exchange cmdlet](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps).
* [View and assign administrator roles in Entra ID](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps).

##### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the **Client ID**, go to the **Microsoft Entra admin center > Applications > App registrations**.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a **Client Secret**.
*	From the Entra ID Portal, go to **Microsoft Entra admin center > Applications > App registrations**.
*	Select the application we have created before, and select "**Certificates and Secrets**". 
*	Under “Client Secrets” click on the “**New Client Secret**” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At last we need to get the **Tenant ID**. This can be found in the Entra ID Portal by going to **Microsoft Entra admin center > Overview**.

## Getting help

> [!TIP]
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems.html) pages_.

> [!TIP]
> _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/provisioning/806-helloid-provisioning-helloid-conn-prov-target-exchangeonline)_

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/

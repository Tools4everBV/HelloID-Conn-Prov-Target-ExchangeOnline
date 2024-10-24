
# HelloID-Conn-Prov-Target-ExchangeOnline

> [!WARNING]
> The readme of this connector has not been converted to the new template. The scripts are ready for a Pull Request

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.

<p align="center">
    <img src="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline/blob/main/Logo.png?raw=true">
</p>

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [HelloID-Conn-Prov-Target-ExchangeOnline](#helloid-conn-prov-target-exchangeonline)
  - [Table of Contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Introduction](#introduction)
  - [Installing the Microsoft Exchange Online PowerShell V3.1 module](#installing-the-microsoft-exchange-online-powershell-v31-module)
  - [Creating the Azure AD App Registration and certificate](#creating-the-azure-ad-app-registration-and-certificate)
    - [Application Registration](#application-registration)
    - [Configuring App Permissions](#configuring-app-permissions)
    - [Assign Azure AD roles to the application](#assign-azure-ad-roles-to-the-application)
    - [Authentication and Authorization](#authentication-and-authorization)
    - [Connection settings](#connection-settings)
    - [Remarks](#remarks)
  - [Getting help](#getting-help)
  - [HelloID Docs](#helloid-docs)

## Requirements
- Installed and available **Microsoft Exchange Online PowerShell V3.1 module**. Please see the [Microsoft documentation](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps) for more information. The download [can be found here](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0).
- Required to run **On-Premises** since it is not allowed to import a module with the Cloud Agent.
- An **App Registration in Azure AD** is required.

## Introduction
For this connector we have the option to correlate to and/or update Exchange Online (Office 365) users and/or mailboxes and provision permission(s) to a group and/or shared mailbox.
  >**Only Exchange and Cloud-only groups are supported**

If you want to create Exchange Online (Office 365) users and/or mailboxes, please use the built-in Microsoft (Azure) Active Directory target system. Or setup Business Rules to provision an Office 365 license group, Microsoft will automatically provision a mailbox for this user.

<!-- GETTING STARTED -->
## Installing the Microsoft Exchange Online PowerShell V3.1 module
Since we use the cmdlets from the Microsoft Exchange Online PowerShell module, it is required this module is installed and available for the service account.
Please follow the [Microsoft documentation on how to install the module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module). 


## Creating the Azure AD App Registration and certificate
> _The steps below are based on the [Microsoft documentation](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) as of the moment of release. The Microsoft documentation should always be leading and is susceptible to change. The steps below might not reflect those changes._
> >**Please note that our steps differ from the current documentation as we use Access Token Based Authentication instead of Certificate Based Authentication**

### Application Registration
The first step is to register a new **Azure Active Directory Application**. The application is used to connect to Exchange and to manage permissions.

* Navigate to **App Registrations** in Azure, and select “New Registration” (**Azure Portal > Azure Active Directory > App Registration > New Application Registration**).
* Next, give the application a name. In this example we are using “**ExO PowerShell CBA**” as application name.
* Specify who can use this application (**Accounts in this organizational directory only**).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “**Register**” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

* To assign your application the right permissions, navigate to **Azure Portal > Azure Active Directory > App Registrations**.
* Select the application we created before, and select “**API Permissions**” or “**View API Permissions**”.
* To assign a new permission to your application, click the “**Add a permission**” button.
* From the “**Request API Permissions**” screen click “**Office 365 Exchange Online**”.
  > _The Office 365 Exchange Online might not be a selectable API. In thise case, select "APIs my organization uses" and search here for "Office 365 Exchange Online"__
* For this connector the following permissions are used as **Application permissions**:
  *	Manage Exchange As Application ***Exchange.ManageAsApp***
* To grant admin consent to our application press the “**Grant admin consent for TENANT**” button.

### Assign Azure AD roles to the application
Azure AD has more than 50 admin roles available. The **Exchange Administrator** role should provide the required permissions for any task in Exchange Online PowerShell. However, some actions may not be allowed, such as managing other admin accounts, for this the Global Administrator would be required. and Exchange Administrator roles. Please note that the required role may vary based on your configuration.
* To assign the role(s) to your application, navigate to **Azure Portal > Azure Active Directory > Roles and administrators**.
* On the Roles and administrators page that opens, find and select one of the supported roles e.g. “**Exchange Administrator**” by clicking on the name of the role (not the check box) in the results.
* On the Assignments page that opens, click the “**Add assignments**” button.
* In the Add assignments flyout that opens, **find and select the app that we created before**.
* When you're finished, click **Add**.
* Back on the Assignments page, **verify that the app has been assigned to the role**.

For more information about the permissions, please see the Microsoft docs:
* [Permissions in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/permissions-exo).
* [Find the permissions required to run any Exchange cmdlet](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps).
* [View and assign administrator roles in Azure Active Directory](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps).

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the **Client ID**, go to the **Azure Portal > Azure Active Directory > App Registrations**.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a **Client Secret**.
*	From the Azure Portal, go to **Azure Active Directory > App Registrations**.
*	Select the application we have created before, and select "**Certificates and Secrets**". 
*	Under “Client Secrets” click on the “**New Client Secret**” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At last we need to get the **Tenant ID**. This can be found in the Azure Portal by going to **Azure Active Directory > Overview**.

### Connection settings
The following settings are required to connect.

| Setting               | Description                                                                                                                                                                                                                             |
| --------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Azure AD Organization | The name of the organization to connect to and where the Azure AD App Registration exists. **Please note: You must use the primary .onmicrosoft.com domain of the organization. Using anything else may lead to inconsistent results.** |
| Azure AD Tenant ID    | Id of the Azure tenant                                                                                                                                                                                                                  |
| Azure AD App Id       | The Application (client) ID of the Azure AD App Registration with Exchange Permissions                                                                                                                                                  |
| Azure AD App Secret   | Secret of the Azure AD App Registration with Exchange Permissions                                                                                                                                                                       |

### Remarks

In some cases, Exchange Online takes more than 70 seconds to return an error. For example when using the 'Set-MailboxRegionalConfiguration' with an invalid date format. For this reason, if you get the 30 seconds timeout then we recommend testing locally on the agent server.

## Getting help
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/hc/en-us/articles/360012518799-How-to-add-a-target-system) pages_

> _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/provisioning/806-helloid-provisioning-helloid-conn-prov-target-exchangeonline)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/

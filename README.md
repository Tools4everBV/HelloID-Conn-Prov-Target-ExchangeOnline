| :warning: Warning |
|:---------------------------|
| This connector is written and tested with the EXO module v3.1. Please make sure you have installed, at least, this version. |

| :information_source: Information |
|:---------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.       |
<br />

<p align="center">
  <img src="https://user-images.githubusercontent.com/69046642/160915847-b8a72368-931c-45d1-8f93-9cc7bb974ca8.png">
</p>

## Versioning
| Version | Description | Date |
| - | - | - |
| 2.0.0   | Use of Access Token to authenticate and no longer use additional PS sessions | 2023/06/09  |
| 1.0.0   | Initial release | 2022/03/30  |

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Versioning](#versioning)
- [Table of Contents](#table-of-contents)
- [Requirements](#requirements)
- [Introduction](#introduction)
- [Installing the Microsoft Exchange Online PowerShell V2 module](#installing-the-microsoft-exchange-online-powershell-v2-module)
- [Creating the Azure AD App Registration and certificate](#creating-the-azure-ad-app-registration-and-certificate)
  - [Application Registration](#application-registration)
  - [Configuring App Permissions](#configuring-app-permissions)
  - [Assign Azure AD roles to the application](#assign-azure-ad-roles-to-the-application)
  - [Authentication and Authorization](#authentication-and-authorization)
  - [Connection settings](#connection-settings)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)

## Requirements
- Installed and available [Microsoft Exchange Online PowerShell V2 module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)
- Required to run **On-Premises** since it is not allowed to import a module with the Cloud Agent.
- An __App Registration in Azure AD__ is required. __Please follow the [Microsoft documentation](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#step-3-generate-a-self-signed-certificate:~:text=Appendix-,Step%201%3A%20Register%20the%20application%20in%20Azure%20AD,-Note) as reference to configure the App Registration correctly__

## Introduction
For this connector we have the option to correlate to existing Exchange Online (Office 365) users and provision groupmemberships and/or permission(s) to a shared mailbox.
  >__Only Exchange and Cloud-only groups are supported__

If you want to create Exchange Online (Office 365) users, please use the built-in Microsoft (Azure) Active Directory target system. If a user exists in Azure AD, Microsoft will automatically sync this to Exchange Online (Office 365).

<!-- GETTING STARTED -->
## Installing the Microsoft Exchange Online PowerShell V2 module
By using this connector you will have the ability to manage groupmemberships and/or permission(s) to a shared mailbox.
Since we use the cmdlets from the Microsoft Exchange Online PowerShell V2 module, it is required this module is installed and available for the service account.
Please follow the [Microsoft documentation on how to install the module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exo-v2-module). 


## Creating the Azure AD App Registration and certificate
> _The steps below are based on the [Microsoft documentation](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) as of the moment of release. The Microsoft documentation should always be leading and susceptible to change. The steps below might not reflect those changes._

### Application Registration
The first step is to register a new <b>Azure Active Directory Application</b>. The application is used to connect to Exchange and to manage permissions.

* Navigate to <b>App Registrations</b> in Azure, and select “New Registration” (<b>Azure Portal > Azure Active Directory > App Registration > New Application Registration</b>).
* Next, give the application a name. In this example we are using “<b>ExO PowerShell CBA</b>” as application name.
* Specify who can use this application (<b>Accounts in this organizational directory only</b>).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “<b>Register</b>” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

To assign your application the right permissions, navigate to <b>Azure Portal > Azure Active Directory > App Registrations</b>.
Select the application we created before, and select “<b>API Permissions</b>” or “<b>View API Permissions</b>”.
To assign a new permission to your application, click the “<b>Add a permission</b>” button.
From the “<b>Request API Permissions</b>” screen click “<b>Office 365 Exchange Online</b>”.
For this connector the following permissions are used as <b>Application permissions</b>:
*	Manage Exchange As Application <b><i>Exchange.ManageAsApp</i></b>
> _The Office 365 Exchange Online might not be a selectable API. In thise case, select "APIs my organization uses" and search here for "Office 365 Exchange Online"__

To grant admin consent to our application press the “<b>Grant admin consent for TENANT</b>” button.

### Assign Azure AD roles to the application
Azure AD has more than 50 admin roles available. The Global Administrator and Exchange Administrator roles provide the required permissions for any task in Exchange Online PowerShell. For general instructions about assigning roles in Azure AD, see [View and assign administrator roles in Azure Active Directory](https://learn.microsoft.com/en-us/azure/active-directory/roles/manage-roles-portal).

To assign the role(s) to your application, navigate to <b>Azure Portal > Azure Active Directory > Roles and administrators</b>.
On the Roles and administrators page that opens, find and select one of the supported roles e.g. “<b>Exchange Administrator</b>” by clicking on the name of the role (not the check box) in the results.
On the Assignments page that opens, click the “<b>Add assignments</b>” button.
In the Add assignments flyout that opens, find and select the app that we created before.
When you're finished, click Add.
Back on the Assignments page, verify that the app has been assigned to the role.

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the <b>Client ID</b>, go to the <b>Azure Portal > Azure Active Directory > App Registrations</b>.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a <b>Client Secret</b>.
*	From the Azure Portal, go to <b>Azure Active Directory > App Registrations</b>.
*	Select the application we have created before, and select "<b>Certificates and Secrets</b>". 
*	Under “Client Secrets” click on the “<b>New Client Secret</b>” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At last we need to get the <b>Tenant ID</b>. This can be found in the Azure Portal by going to <b>Azure Active Directory > Overview</b>.

### Connection settings
The following settings are required to connect.

| Setting     | Description |
| ------------ | ----------- |
| Azure AD Organization | The name of the organization to connect to and where the Azure AD App Registration exists __Please note: This has to be the .onmicrosoft domain name__ |
| Azure AD Tenant ID | Id of the Azure tenant |
| Azure AD App Id | The Application (client) ID of the Azure AD App Registration with Exchange Permissions. __Please follow the [Microsoft documentation](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#step-3-generate-a-self-signed-certificate:~:text=Appendix-,Step%201%3A%20Register%20the%20application%20in%20Azure%20AD,-Note) as reference to configure the App Registration correctly__  |
| Azure AD App Secret | Secret of the Azure app |

## Getting help
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/hc/en-us/articles/360012518799-How-to-add-a-target-system) pages_

> _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/provisioning/806-helloid-provisioning-helloid-conn-prov-target-exchangeonline)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/

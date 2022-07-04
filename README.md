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
| 1.0.1   | Updated to only import the modules we use for performance increase | 2022/07/04  |
| 1.0.0   | Initial release | 2022/03/30  |

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Versioning](#versioning)
- [Table of Contents](#table-of-contents)
- [Requirements](#requirements)
- [Introduction](#introduction)
- [Installing the Microsoft Exchange Online PowerShell V2 module](#installing-the-microsoft-exchange-online-powershell-v2-module)
  - [Connection settings](#connection-settings)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)

## Requirements
- Installed and available [Microsoft Exchange Online PowerShell V2 module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)
- To manage users, mailboxes and groups, the service account has to have the role "**Exchange Administrator**" assigned.
- Required to run **On-Premises** since it is not allowed to import a module with the Cloud Agent.
- **Concurrent sessions** in HelloID set to a **maximum of 1**! If this is any higher than 1, this may cause errors, since Exchange only support a maximum of 3 sessions per minute.
- Since we create a Remote PS Session on the agent server (which will containt the Exchange Session, to avoid the Exchange limit of 3 sessions per minute), the service account has to be a member of the group “**Remote Management Users**”.

## Introduction
For this connector we have the option to correlate to existing Exchange Online (Office 365) users and provision groupmemberships and/or permission(s) to a shared mailbox.
  >__Only Exchange and Cloud-only groups are supported__

If you want to create Exchange Online (Office 365) users, please use the built-in Microsoft (Azure) Active Directory target system. If a user exists in Azure AD, Microsoft will automatically sync this to Exchange Online (Office 365).

<!-- GETTING STARTED -->
## Installing the Microsoft Exchange Online PowerShell V2 module
By using this connector you will have the ability to manage groupmemberships and/or permission(s) to a shared mailbox.
Since we use the cmdlets from the Microsoft Exchange Online PowerShell V2 module, it is required this module is installed and available for the service account.
Please follow the [Microsoft documentation on how to install the module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exo-v2-module). 


### Connection settings
The following settings are required to connect.

| Setting     | Description |
| ------------ | ----------- |
| Username | The username of a Global Admin in Exchange Online (Office 365) |
| Password | The password of the Global Admin in Exchange Online (Office 365) |

## Getting help
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/hc/en-us/articles/360012518799-How-to-add-a-target-system) pages_

> _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/provisioning/806-helloid-provisioning-helloid-conn-prov-target-exchangeonline)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/

# SharePoint - Fast Register to Appointment (Application Customizer)
## Summary
SharePoint Online provides a list of appointments (events) on a site. The "Events" web part can be used to display them on a page. Users can access the appointment details via the web part and view all the details. The detailed view provides a link that allows them to add the appointment to their personal calendar. An appointment entry also allows them to maintain an attendee list. The "Events" list provides the "Attendees" column, where multiple people can be added. However, there is no automatic function for this; the list must be edited manually.
The special app extension adds a registration and deregistration function to the detailed view. Users can register for a specific appointment and deregister later with a simple click. Attendees are automatically managed in the attendee column of the list.

![The extended calendar entry](https://www.dev-sky.net/img/apps/fastregisterappointment.png)

*UI of extended event calendar*

## Video
[![SharePoint Online: Schnelle & einfache Terminanmeldung](https://img.youtube.com/vi/_-aTpJPXRdA/hqdefault.jpg)](https://youtu.be/_-aTpJPXRdA)

## Used SharePoint Framework Version
![version](https://img.shields.io/badge/version-1.18-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> You just need a SharePoint list and a view for field order configuration.

## Solution

| Solution    | Author(s)                                                   |
| ----------- | ----------------------------------------------------------- |
| Repository  | Marc André Schröder-Zhou (https://github.com/maschroeder-z) |

## Version history

| Version | Date             | Comments              |
| ------- | ---------------- | --------------------- |
| 1.1     | 30.11.2024       | Upgrade to SPFx 1.18  |
| 1.0     | 30.09.2023       | Initial Release       |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Check your currenr Node version and installed SPFx-Framework version.

## Features
- A "Register" button is added for users who are not registered.
- Clicking "Register" saves the current user to the Participants column.
- A notification appears for users who are already registered to the event. A "Cancel" button is added.
- Clicking "Cancel" removes the current user from the Participants column.

## Help
Please contact me for further help or information about the sample.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development

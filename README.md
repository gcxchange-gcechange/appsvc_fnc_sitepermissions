##  Site Permissions App Function
This app function checks subsite configurations to make sure certain permissions are set. If it detects any misconfigurations it will reconfigure them + inform site owners. The function app is triggered on a timer set for every Saturday at 12 AM.
## How To Setup
You will need to add a file named **local.settings.json** in the **FunctionApp1** folder.  The function app expects the following values:
- **tenantId** - Your azure subscription
- **clientId** - The app registration client ID
- **clientSecret** - The app registration client secret value
- **hubId** - The site ID for the hub site. All subsites will be scanned.
- **appOnlyId** - The app only ID created in SharePoint
- **appOnlySecret** - The app only secret created in SharePoint
- **excludeSiteIds** - A string of site IDs seperated by commas. These sites will be ignored.
- **emailSenderId** - The object ID of the user that will send emails. Make sure this user has a license to send email.
- **groups** - An array of JSON objects seperated by a comma. 

	- **groupName** (the name of the user/group)

	- **groupId** (the ID of the user/group), 

	- **permissionLevel** (The required permission level. This can be **Read**, **Edit**, **Full Control**, or **Site Collection Administrator**.

			"[{\"groupName\": \"sdg_test\", \"groupId\": \"de7c675e-9066-457a-b1f5-8de21b03341e\", \"permissionLevel\": \"Site Collection Administrator\"}]"

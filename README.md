##  Site Permissions App Function
This app function checks subsite configurations to make sure certain permissions are in place. If it detects any misconfigurations it will reconfigure them + inform site owners. The function app is triggered on a timer set for every Saturday at 12 AM.
##  API Permissions
### Microsoft Graph
- **Group.Read.All** - Read all groups
- **GroupMember.Read.All** - Read all group memberships
- **Mail.Send** - Send mail as any user
- **Sites.Read.All** - Read all site collections
- **User.Read.All** - Read all users' full profiles
### App Only
- **Full Control** - Site collection
## How To Setup
You will need to add a file named **local.settings.json** in the **Permissions** folder.  The function app expects the following values:
- **tenantId** - Your azure subscription
- **hubId** - The site ID for the hub site. All subsites will be scanned.
- **clientId** - The app registration client ID
- **appOnlyId** - The app only ID created in SharePoint
- **keyVaultUrl** - The URL to the key vault containing the client and app only secrets.
- **secretNameClient** - The name of the client secret in your key vault
- **secretNameAppOnly** - The name of the app only secret in your key vault
- **excludeSiteIds** - A string of site IDs seperated by commas. These sites will be ignored.
- **emailSenderId** - The object ID of the user that will send emails. Make sure this user has a license to send email.
- **reportOnly** - "1", "on", or "true". This will run the script in report only mode and won't modify any of the misconfigured sites.
- **groups** - A comma seperated string of group names and permission levels. See the example below for formatting. 

	- **Name** the name of the user/group

	- **Id** the object Id of the user/group

	- **AssignedPermissionLevel** The permission level you want to enforce. Only one permission level can be specified for each group. This can be **Read**, **Contribute**, **Edit**, **Design**, **Full Control**, or **Site Collection Administrator**.

			"name1|Id1|Read, name2|Id2|Full Control, name3|Id3|Site Collection Administrator, etc."

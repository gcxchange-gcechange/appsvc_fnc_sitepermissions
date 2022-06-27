##  Site Permissions App Function
This app function checks subsite configurations to make sure certain permissions are set. If it detects any misconfigurations it will reconfigure them + inform site owners. The function app is triggered on a timer set for every Saturday at 12 AM.
##  API Permissions
Your app registration will need the following API permissions
### Microsoft Graph
- **Directory.Read.All** - Read directory data
- **Group.Read.All** - Read all groups
- **GroupMember.ReadWrite.All** - Read and write all group memberships
- **Mail.Send** - Send mail as any user
- **Sites.FullControl.All** - Have full control of all site collections
- **User.Read** - Sign in and read user profile
- **User.Read.All** - Read all users' full profiles
- **User.ReadWrite.All** - Read and write all users' full profiles
### SharePoint
- **Sites.FullControl.All** - Have full control of all site collections
Your app only permissions required are
### App Only
- Full Control on the tenant and site collection
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
- **groups** - A comma seperated string of group names and permission levels. See the example below for formatting. 

	- **Name** the name of the user/group

	- **Id** the object Id of the user/group

	- **AssignedPermissionLevel** The permission level you want to enforce. Only one permission level can be specified for each group. This can be **Read**, **Edit**, **Full Control**, or **Site Collection Administrator**.

			"group1|Id1|Read, group2|Id2|Full Control, group3|Id3|Site Collection Administrator, etc."

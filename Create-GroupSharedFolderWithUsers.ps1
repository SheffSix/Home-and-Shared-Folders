<#
.SYNOPSIS
	Creates a folder and sub-folders based on an Active Directory group and its members.

.DESCRIPTION
	This script creates a folder based on a given Active Directory group inside a given path. The group is then given read-only permissions to the folder, but not sub-folders or files.
	
	A sub-folder called "_Public" is created unless -NoPublicFolder is specified. All users of the group have read-only permissions to this folder, sub-folders and files.
	
	A sub-folder is created for each member of the group and the user is given modify access to the folder.
		
	The created folders have inheritance enabled, so make sure that parent folders' permissions are set up correctly.
	
.NOTES
	File Name    :  Create-GroupSharedFolderWithUsers.ps1
	Author       :  James Beck
	
.PARAMETER Path
	Specify the path to the folder where the new folder is to be created.
	
.PARAMETER GroupName
	The name of the Active Directory group to be given access to the new folder.
	
.PARAMETER FolderName
	Optional. Specify a name for the new folder. Without this option, the folder will be created with the group name.

.PARAMETER NoPublicFolder
	Optional. Do not create a public folder inside the new folder.

.EXAMPLE
	.\Create-GroupSharedFolderWithUsers.ps1 -Path "E:\Shares\Student Share\Mathematics" -GroupName "TG 10A-M1"
	Creates a folder for the group TG 10A-M1 with the same name, inside the folder E:\Shares\Student Share\Mathematics.
	
.EXAMPLE
	.\Create-GroupSharedFolderWithUsers.ps1 -Path "\\FileServer\SharedArea\Product Design" -GroupName "TG 11B-Pd2" -FolderName "Mr Smith Y11"
	Creates a folder for the group TG 11B-Pd2 called Mr Smith Y11, inside the folder \\FileServer\SharedArea\Product Design.
	
.EXAMPLE
	.\Create-GroupSharedFolderWithUsers.ps1 -Path "R:\PE\Shared" -GroupName "TG 9A-Pe1" -NoPublicFolder
	Creates a folder for the group TG 9A-Pe1 with the same name, inside the folder R:\PE\Shared, without creating a public space.
	
#>

Param (
	[string]$Path,
	[string]$GroupName,
	[string]$FolderName,
	[switch]$NoPublicFolder
)

# Check to make sure both parameters have been specified
If (!($Path -or $GroupName)) { write-host "Incorrect parameters"; exit 3}

If (!($FolderName)) { $FolderName = $GroupName }

# Check if the Specified group exists
If (!(Get-ADGroup -LDAPFilter "(Name=$GroupName)")) { write-host "The group $GroupName does not exist"; exit 1 }

#Check if the parent folder exists
If (!(Test-Path -Path "${Path}" -PathType Container)) { write-host "The path $Path does not exist"; exit 2 }

# Steps to create the group folder. Skip if it already exists.
If (!(Test-Path -Path "${Path}\$FolderName")) {
	# Create the folder
	New-Item -Path "${Path}" -Name "${FolderName}" -ItemType Directory
	
	# Set up folder permissions. Give the group read access on this folder only.
	$acl = Get-Acl -Path "${Path}\${FolderName}"
	$perm = "$GroupName",'Read','Allow'
	$rule = New-Object -TypeName system.security.accesscontrol.filesystemaccessrule -argumentlist $perm
	$acl.setaccessrule($rule)
	$acl | Set-Acl -Path "${Path}\${FolderName}"
}

# Create a Public folder with read-only access for the whole group. Skip if it already exists.
if (!(Test-Path -Path "{Path}\${FolderName}\_Public")  -and !($NoPublicFolder)) {
	# Create the folder
	New-Item -Path "${Path}\${FolderName}\_Public" -ItemType Directory
	
	# Set up folder permissions. Give the group read access to the Public folder.
	$acl = Get-Acl -Path "${Path}\${FolderName}\_Public"
	$perm = "$GroupName",'Read,ExecuteFile,ListDirectory','ContainerInherit,ObjectInherit','None','Allow'
	$rule = New-Object -TypeName system.security.accesscontrol.filesystemaccessrule -argumentlist $perm
	$acl.setaccessrule($rule)
	$acl | Set-Acl -Path "${Path}\${FolderName}\_Public"
}

# Steps to create the users' folders.
# Double-check that the group folder exists.
If (Test-Path -Path "${Path}\${FolderName}" -PathType Container) {
	Get-ADGroupMember "$GroupName" | foreach {
		# Check if the user's folder exists. Skip if it does.
		If (!(Test-Path -Path "${Path}\${FolderName}\$($_.Name)")) {
			# Create the user's folder
			New-Item -Path "${Path}\${FolderName}" -Name $_.Name -ItemType Directory
		}
		
		# Double-check the user's folder exists. Skip if it doesn't.
		If (Test-Path -Path "${Path}\${FolderName}\$($_.Name)" -PathType Container) {
			
			# Set up folder permissions. Give the user modify access to this folder, subfolders and files.
			$acl = Get-Acl -Path "${Path}\${FolderName}\$($_.Name)"
			# 'ContainerInherit,ObjectInherit','None' is what we use to specify subfolder and files.
			$perm = $($_.samaccountname),'Read,Modify','ContainerInherit,ObjectInherit','None','Allow'
			$rule = New-Object -TypeName system.security.accesscontrol.filesystemaccessrule -argumentlist $perm
			$acl.setaccessrule($rule)
			$acl | Set-Acl -Path "${Path}\${FolderName}\$($_.Name)"
		}
	}	
}
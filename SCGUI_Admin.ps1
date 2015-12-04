####---------------------------------------------------------------------------------------------####
# Service Center Exchange Gui																		#
# Written by: Austin Heyne																			#
# 																									#
# --Information--																					#
# Program provides an interface to manage day to day operations in Exchange 2010					#
# This version is the administrative interface that has less error handleing and access to 			#
# restricted functions that would not work in the normal version.									#
####---------------------------------------------------------------------------------------------####

function GenerateForm {

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$tabControl1 = New-Object System.Windows.Forms.TabControl
#Tab1
$tabPageDL = New-Object System.Windows.Forms.TabPage
$groupBox2 = New-Object System.Windows.Forms.GroupBox
$CB_Sponsors = New-Object System.Windows.Forms.CheckBox
$RTB_Members = New-Object System.Windows.Forms.RichTextBox
$label4 = New-Object System.Windows.Forms.Label
$RTB_DLOutput = New-Object System.Windows.Forms.RichTextBox
$Create = New-Object System.Windows.Forms.Button
$ListInGAL = New-Object System.Windows.Forms.CheckBox
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$RB_IO = New-Object System.Windows.Forms.RadioButton
$RB_OO = New-Object System.Windows.Forms.RadioButton
$RB_MO = New-Object System.Windows.Forms.RadioButton
$RB_A = New-Object System.Windows.Forms.RadioButton
$TB_EmailAddress = New-Object System.Windows.Forms.TextBox
$label3 = New-Object System.Windows.Forms.Label
$TB_Sponsors = New-Object System.Windows.Forms.TextBox
$label2 = New-Object System.Windows.Forms.Label
$TB_Alias = New-Object System.Windows.Forms.TextBox
$UseCustomAlias = New-Object System.Windows.Forms.CheckBox
$TB_DGN = New-Object System.Windows.Forms.TextBox
$label1 = New-Object System.Windows.Forms.Label
#Tab2
$tabPageIQ = New-Object System.Windows.Forms.TabPage
$B_SetCustom = New-Object System.Windows.Forms.Button
$label_IQ_12 = New-Object System.Windows.Forms.Label
$TB_ProhibitSend = New-Object System.Windows.Forms.TextBox
$label_IQ_11 = New-Object System.Windows.Forms.Label
$label_IQ_10 = New-Object System.Windows.Forms.Label
$TB_IssueWarning = New-Object System.Windows.Forms.TextBox
$label_IQ_9 = New-Object System.Windows.Forms.Label
$B_IQ_SetDefault = New-Object System.Windows.Forms.Button
$label_IQ_Out = New-Object System.Windows.Forms.Label
$B_IQ_Check = New-Object System.Windows.Forms.Button
$TB_IQName = New-Object System.Windows.Forms.TextBox
$label_IQ_7 = New-Object System.Windows.Forms.Label
#Tab3
$tabPageMBP = New-Object System.Windows.Forms.TabPage
$TB_MBP_Target = New-Object System.Windows.Forms.TextBox
$label_MBP_2 = New-Object System.Windows.Forms.Label
$CB_MBP_FullAccess = New-Object System.Windows.Forms.CheckBox
$CB_MBP_SendAs = New-Object System.Windows.Forms.CheckBox
$TB_MBP_User = New-Object System.Windows.Forms.TextBox
$label_mbp_1 = New-Object System.Windows.Forms.Label
$RTB_MBP_Output = New-Object System.Windows.Forms.RichTextBox
$B_MBP_Set = New-Object System.Windows.Forms.Button
#CMSTab
$CMS_Tab = New-Object System.Windows.Forms.TabPage
$CMS_Progress = New-Object System.Windows.Forms.ProgressBar
$CMS_Status = New-Object System.Windows.Forms.Label
$CMS_Members = New-Object System.Windows.Forms.RichTextBox
$CMS_label3 = New-Object System.Windows.Forms.Label
$CMS_Create = New-Object System.Windows.Forms.Button
$CMS_Owners = New-Object System.Windows.Forms.TextBox
$CMS_label2 = New-Object System.Windows.Forms.Label
$CMS_label1 = New-Object System.Windows.Forms.Label
$CMS_Group = New-Object System.Windows.Forms.TextBox
#Create Room Tab
$CR_Tab = New-Object System.Windows.Forms.TabPage
$CR_Progress = New-Object System.Windows.Forms.ProgressBar
$CR_Output = New-Object System.Windows.Forms.Label
$CR_Create = New-Object System.Windows.Forms.Button
$CR_groupBox1 = New-Object System.Windows.Forms.GroupBox
$CR_RB_AprovalList = New-Object System.Windows.Forms.RadioButton
$CR_RB_Noone = New-Object System.Windows.Forms.RadioButton
$CR_RB_Anyone = New-Object System.Windows.Forms.RadioButton
$CR_label5 = New-Object System.Windows.Forms.Label
$CR_Autobook = New-Object System.Windows.Forms.TextBox
$CR_label4 = New-Object System.Windows.Forms.Label
$CR_Approved = New-Object System.Windows.Forms.TextBox
$CR_Restrict = New-Object System.Windows.Forms.CheckBox
$CR_Sponsors = New-Object System.Windows.Forms.TextBox
$CR_label2 = New-Object System.Windows.Forms.Label
$CR_label1 = New-Object System.Windows.Forms.Label
$CR_Owner = New-Object System.Windows.Forms.TextBox
$CR_Alias = New-Object System.Windows.Forms.TextBox
$CR_DispName = New-Object System.Windows.Forms.TextBox
$CR_label6 = New-Object System.Windows.Forms.Label
$CR_label7 = New-Object System.Windows.Forms.Label
#Tab4
$tabPageConsole = New-Object System.Windows.Forms.TabPage
$RTB_Console = New-Object System.Windows.Forms.RichTextBox
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.

#region Create DL
$Create_OnClick= 
{	
	#Clear any text from DLOutput
	$RTB_DLOutput.Text = ""
	. updateDLOutput
	
	$RTB_DLOutput.AppendText("Creating DL...")
	. updateDLOutput
	#accessLog "`nCreating Distribution List"
	
	#Handleing Generated Alias or Custom Alias
	if($UseCustomAlias.Checked){$alias = $TB_Alias.Text}
	else{$alias = $TB_DGN.Text.Replace(" ","").Replace("(","").Replace(")","").Replace(",","").Replace("-","")}
	$RTB_DLOutput.AppendText("`nDGN: " + $TB_DGN.Text)
	#accessLog ("`nDGN: " + $TB_DGN.Text)
	$RTB_DLOutput.AppendText("`nAlias: " + $alias.toString())
	#accessLog ("`nAlias: " + $alias.toString())
	. updateDLOutput
	
	#Handle Sponsor Check
	if($TB_Sponsors.Text -ne $null){
		$sponsorList = $TB_Sponsors.Text.Split(",")
	}else{
		$sponsorList = "NoData"
	}
	if(!$?){
		$RTB_DLOutput.AppendText("`nError Checking Sponsor List")
		. updateDLOutput 
		#accessLog "`nError Checking Sponsor List"
	}else{
		$RTB_DLOutput.AppendText("`nSponsors Approved")
		. updateDLOutput
		#accessLog "`nSponsors Approved"
	}
	
	#Create DL
	New-DistributionGroup -Alias $alias -Name $TB_DGN.Text -Type Distribution -OrganizationalUnit OU="DLs,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu" -SamAccountName $alias -ManagedBy $sponsorList | Out-String -Stream | ForEach-Object {
		$RTB_Console.AppendText("`n" + $_)
		. updateConsole
		#accessLog ("`n" + $_)
	}
	if(!$?){
		$RTB_DLOutput.AppendText("`nError Creating DL, See Console")
		. updateDLOutput
		#accessLog ("`nError Creating DL, See Console")
	}else{
		$RTB_DLOutput.AppendText("`nDL Created, See Console for more information")
		. updateDLOutput
		#accessLog ("`nDL Created, See Console for more information")
	}
	
	#Modify Email Address Policy
	Set-DistributionGroup $TB_DGN.Text -EmailAddressPolicyEnabled $false | Out-String | ForEach-Object{
		$RTB_Console.AppendText("`n" + $_)
		. updateConsole
		#accessLog ("`n" + $_)
	}
		
	#Add Default Email Address
	$emailAddresses = $(Get-DistributionGroup $TB_DGN.Text).EmailAddresses
	$newAddress = "SMTP:" + $alias + "@sooner.net.ou.edu"
	$emailAddresses += $newAddress
	Set-DistributionGroup $TB_DGN.Text -EmailAddresses $emailAddresses  | Out-String | ForEach-Object{
		$RTB_Console.AppendText("`n" + $_)
		. updateConsole
		#accessLog ("`n" + $_)
	}
	foreach($_ in $emailAddresses){
		$RTB_DLOutput.AppendText("`n" + $_)
		. updateDLOutput
		#accessLog ("`n" + $_)
	}
	
	#Add @ou.edu Email Address
	if($TB_EmailAddress.Text.Replace(" ","") -ne ""){
		$newAddress = "SMTP:" + $TB_EmailAddress.Text + "@ou.edu"
		$emailAddresses += $newAddress
		Set-DistributionGroup $TB_DGN.Text -EmailAddresses $emailAddresses | Out-String |ForEach-Object{
			$RTB_Console.AppendText("`n" + $_)
			. updateConsole
			#accessLog ("`n" + $_)
		}
		if(!$?){
			$RTB_DLOutput.AppendText("`nError adding @ou.edu address, See Console")
			. updateDLOutput
			#accessLog ("`nError adding @ou.edu address, See Console")
		}else{
			$RTB_DLOutput.AppendText("`nAdded " + $newAddress)
			. updateDLOutput
			#accessLog ("`nAdded " + $newAddress)
		}		
	}
	
	#List in GAL?
	if($ListInGAL.Checked){
		Set-DistributionGroup $TB_DGN.Text -HiddenFromAddressListsEnabled $false
		#accessLog ("`nListed in GAL")
	}else{
		Set-DistributionGroup $TB_DGN.Text -HiddenFromAddressListsEnabled $true
		#accessLog ("`nHidden from GAL")
	}
	
	#Handle who can send to list
	if($RB_A.Checked){
		Set-DistributionGroup $TB_DGN.Text -RequireSenderAuthenticationEnabled $false
		#accessLog ("`nRequire Sender Authentication Enabled")
	}
		
	if($RB_MO.Checked){
		Set-DistributionGroup $TB_DGN.Text -AcceptMessagesOnlyFromDLMembers $TB_DGN.Text
		#accessLog ("`nAccept messages only from DL members")
	}
	
	if($RB_OO.Checked){
		Import-Module ActiveDirectory #used for Set-ADGroup to change scope to universal
		#This creates another DL called $alias + "-group" that is used as the group that can send to DL
		New-DistributionGroup -Alias ($alias + "-group") -Name ($alias + "-group") -Type Security -OrganizationalUnit  OU="DLs,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu" -SamAccountName ($alias + "-group") 
		#accessLog ($alias + "-group")
		#Make universal group
		Set-ADGroup -Identity ($alias + "-group") -GroupScope Universal
		#Mail Enable group (May error out and silently continue from already being set)
		Enable-DistributionGroup ($alias + "-group")
		#Wait for new DL to replicate
		$RTB_DLOutput.AppendText("`nPlease wait 60 seconds for Owners Group to replicate")
		. updateDLOutput
		Start-Sleep -Seconds 60
		Set-DistributionGroup ($alias + "-group") -HiddenFromAddressListsEnabled $true
		Set-DistributionGroup ($alias + "-group") -AcceptMessagesOnlyFromDLMembers ($alias + "-group")
		$sponsorList | foreach {Add-DistributionGroupMember -Identity ($alias + "-group") -Member $_}
		Set-DistributionGroup $TB_DGN.Text -AcceptMessagesOnlyFromDLMembers ($alias + "-group")
	}
	
	if($RB_IO.Checked){
		Set-DistributionGroup $TB_DGN.Text -RequireSenderAuthenticationEnabled $true
		#accessLog ("`nRequire sender authentication enabled")
	}
	
	#Handle list membership
	#accessLog ("Sponsors list")
	
	if($CB_Sponsors.Checked){
		$sponsorList | foreach {
			Add-DistributionGroupMember -Identity $TB_DGN.Text -Member $_
			#accessLog ("`n" + $_)
		}
		$RTB_DLOutput.AppendText("`nSponsors added to DL")
		. updateDLOutput
	}
	
	if($RTB_Members.Text -ne ""){
		$memberList = $RTB_Members.Text.Split("`n")
		$memberList | foreach {
			Add-DistributionGroupMember -Identity $TB_DGN.Text -Member $_
			#accessLog ("`n" + $_)
			}
		$RTB_DLOutput.AppendText("`nMembers added to DL")
		. updateDLOutput
	}
	
	#accessLog ("Sponsors added to DL")
	
	$RTB_DLOutput.AppendText("`nDL Created, populate through ADUC")
	. updateDLOutput
	#accessLog ("`nDL Created, populate through ADUC")
	
	#Clear Form
	$TB_Alias.Text = ""
	$TB_DGN.Text = ""
	$TB_EmailAddress.Text = ""
	$TB_Sponsors.Text = ""
	$UseCustomAlias.Checked = $false
	$RB_A.Checked = $true
	$RTB_Members.Text = ""
}
#endregion Create DL

#region Increase Quota
$B_IQ_Check_OnClick= 
{
	$IQ_MB = $TB_IQName.Text 
	$IQ_MB = safeGetMailbox($IQ_MB)
	$IQ_MB_USED = $TB_IQName.Text | Get-MailboxStatistics
	$label_IQ_Out.Text = $IQ_MB.SamAccountName + "'s current quota is " + $IQ_MB.ProhibitSendQuota.Value.ToGB() + "GB, Used: " + $IQ_MB_USED.TotalItemSize.Value.ToGB() + " GB."
}

$B_IQ_SetDefault_OnClick= 
{
	if($TB_IQName.Text.Replace(" ","") -eq ""){
		$label_IQ_Out.Text = "Please Input 4x4 or Email"
	}else{
		$IQ_MB = safeGetMailbox($TB_IQName.Text)
		Set-Mailbox -IssueWarningQuota '5.841 GB (6,271,533,056 bytes)' -ProhibitSendQuota '6 GB (6,442,450,944 bytes)' -Identity $IQ_MB
		$IQ_MB = safeGetMailbox($IQ_MB)
		$label_IQ_Out.Text = $IQ_MB.SamAccountName + "'s current quota is " + $IQ_MB.ProhibitSendQuota
	}
	$TB_IQName.Text = ""
}

$B_SetCustom_OnClick= 
{
	$IQ_MB = safeGetMailbox($TB_IQName.Text)
	$IssueWarningBytes = [int]$TB_IssueWarning.Text * 1073741824
	$ProhibitSendBytes = [int]$TB_ProhibitSend.Text * 1073741824
	Set-Mailbox -IssueWarningQuota $IssueWarningBytes -ProhibitSendQuota $ProhibitSendBytes -Identity $IQ_MB
	$IQ_MB = safeGetMailbox($IQ_MB)
	$label_IQ_Out.Text = $IQ_MB.SamAccountName + "'s current quota is " + $IQ_MB.ProhibitSendQuota
	$TB_IQName.Text = ""
	$TB_IssueWarning.Text = ""
	$TB_ProhibitSend.Text = ""
}
#endregion Increase Quota

#region Set MBPerms
$B_MBP_Set_OnClick= 
{
	$RTB_MBP_Output.AppendText("Starting...")
	. updateMBPOutput 
	#Check for input error
	if(($TB_MBP_User.Text.Replace(" ","") -eq "") -or ($TB_MBP_Target.Text.Replace(" ","") -eq "")){
		$RTB_MBP_Output.AppendText("`nMissing input data")
		. updateMBPOutput 
	}else{
		$MBP_User = safeGetMailbox($TB_MBP_User.Text)
		$RTB_MBP_Output.AppendText("`nUser: " + $MBP_User.SamAccountName)
		. updateMBPOutput
		$MBP_Target = safeGetMailbox($TB_MBP_Target.Text)
		$RTB_MBP_Output.AppendText("`nTarget: " + $MBP_Target.SamAccountName)
		. updateMBPOutput
		if($CB_MBP_SendAs.Checked){
			Add-ADPermission -Identity $MBP_User.SamAccountName -User $MBP_Target.SamAccountName -ExtendedRights 'Send-As' | Out-String | ForEach-Object {
				$RTB_MBP_Output.AppendText("`n" + $_)
				. updateMBPOutput}
		}
		if($CB_MBP_FullAccess.Checked){
			Add-MailboxPermission -Identity $MBP_User.SamAccountName -User $MBP_Target.SamAccountName -AccessRights 'FullAccess' | Out-String | ForEach-Object {
				$RTB_MBP_Output.AppendText("`n" + $_)
				. updateMBPOutput}
		}
	}
	$TB_MBP_User.Text = ""
	$TB_MBP_Target.Text = ""
	$CB_MBP_SendAs.Checked = $false
	$CB_MBP_FullAccess.Checked = $false
}
#endregion Set MBPerms

#region CMS Group
$handler_CMS_Create_Click= 
{
	$CMS_Status.text = ""
	$alias = ($CMS_Group.text).Replace(" ","")
	$owners = ($CMS_Owners.text).Replace(" ","")
	$ownerarray = $owners.split(",")

	$CMS_Progress.Value = 0
	$CMS_Progress.Step = 100/11

	$CMS_Status.text = "Creating Group"
	$CMS_Progress.PerformStep()
	$form1.Update()

	new-DistributionGroup -alias $alias -name $CMS_Group.text -type security -org "OU=CMS,OU=DLs,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu" -SamAccountName $alias

	$CMS_Status.text = "Setting Email Policy"
	$CMS_Progress.PerformStep()
	$form1.Update()
	
	set-distributiongroup $alias -BypassSecurityGroupManagerCheck -EmailAddressPolicyEnabled $false

	$CMS_Status.text = "Setting Owners"
	$CMS_Progress.PerformStep()
	$form1.Update()

	foreach ($person in $ownerarray) {
		$DLmangers = (Get-DistributionGroup $alias).ManagedBy
		$DLmangers += $person
		Set-DistributionGroup $alias -ManagedBy $DLmangers
	}

	$CMS_Status.text = "Waiting for Replication"
	
	$i = 0
	do {
		$CMS_Progress.PerformStep()
		$form1.Update()
		sleep 1
		$i ++
	} while ($i -lt 5)

	$CMS_Status.text = "Setting Address"
	$CMS_Progress.PerformStep()
	$form1.Update()
	
	set-distributiongroup $alias -PrimarySMTPAddress $alias@ou.edu
	
	if($CMS_Members.Text -ne ""){
		$CMS_Status.text = "Adding Members"
		$CMS_Progress.PerformStep()
		$form1.Update()
		$memberList = $CMS_Members.Text.Split("`n")
		$memberList | foreach {Add-DistributionGroupMember -Identity $CMS_Group.Text -Member $_}
		$CMS_Status.text = "Members Added"
		$form1.Update()
	}

	$CMS_Status.text = "Verifying Data"
	$CMS_Progress.PerformStep()
	$form1.Update()
		
	$dlname = (Get-DistributionGroup $alias@ou.edu).name
	$dladdress = (Get-DistributionGroup $alias@ou.edu).PrimarySmtpAddress
	
	if ($dlname -ne "" -and $dladdress -ne "") {
		$CMS_Progress.Value = 100
		$CMS_Status.text = "Done: " + $dlname + " / " + $dladdress
		sleep 2
	}else{
		$CMS_Status.text = "Somethings didn't go right. See console"
	}
	
	$CMS_Group.text = ""
	$CMS_Owners.text = ""
	$CMS_Members.text = ""
	$CMS_Progress.Value = 0
}
#endregion Set MBPerms

#region Create Room
$CR_Create_OnClick= 
{
#TODO: Feedback/status updates, progress bad, error handeling,
	$password = ConvertTo-SecureString "#()%&oisdfdskhf30987)(309h" -asplaintext -force
	$upn = $CR_Alias + '@sooner.net.ou.edu'
	New-Mailbox -Name $CR_Alias -displayname $CR_DispName -Alias $CR_Alias -OrganizationalUnit 'sooner.net.ou.edu/Exchange/Resource Mailboxes' -UserPrincipalName $upn -SamAccountName $CR_Alias -FirstName '' -Initials '' -LastName '' -Password $password -ResetPasswordOnNextLogon $false 
	$CR_Output.Text("Creating Mailbox")
	safeGetMailbox($CR_Alias) | set-mailbox -IssueWarningQuota '5.841 GB (6,271,533,056 bytes)' -ProhibitSendQuota '6 GB (6,442,450,944 bytes)' -RoleAssignmentPolicy "OU Default Role Assignment Policy" -RetainDeletedItemsFor 60.0:0:0 -SingleItemRecoveryEnabled $true
	
#Create Resource Mailbox
	$subject = "Resource Mailbox Creation"
	$body = "This is an internal email sent as part of the resource mailbox creation process.  Please delete this email."
	sendMail $account $subject $body
	
	#Create Access Group
	$CR_Alias = $name + "-group"
	$groupCN = "cn=" + $CR_Alias
	$objOU = [ADSI]"LDAP://OU=Groups,OU=Resource Mailboxes,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu"
	
	$objGroup = $objOU.Create("group", $groupCN)
	$objGroup.Put("sAMAccountName", $CR_Alias)

	#Add sponsors to group
	foreach ($User in $CR_Sponsors.split(",")) {
		$objGroup.member.add($(Get-User $User).DistinguishedName)
	}
	
	$ownerCN = $(get-user $CR_Owner).DistinguishedName
	$objGroup.Put("managedBy", $ownerCN)
	$objGroup.Put("groupType", 0x80000008)
	$objGroup.SetInfo()
	
	if (!$?){
		#outData("Error creating security group, " + $groupName + " please contact Exchange Team")
		#Add error handeling
	}
	#outData("Please wait 30 seconds for data to populate")
	#Start-sleep -s 30
	
	add-adpermission $CR_Alias -user $CR_Owner -accessrights WriteProperty -properties 'Member'
	if (!$?){
		#outData "Unable to add permissions, retrying in 30 seconds."
		Start-sleep -s 30
		add-adpermission $CR_Alias -user $CR_Owner -accessrights WriteProperty -properties 'Member'
		if (!$?){
			#outData "Error applying permissions, please contact MS Team"
		}else{
			#outData "Permissions applied successfully"
		}
	}

	$fullGroupName = "sooner\" + $CR_Alias + "-group"
	$fullOwnerName = "sooner\" + $CR_Owner
	add-mailboxpermission $CR_Alias -owner $fullOwnerName
	add-mailboxpermission $CR_Alias -user $fullGroupName -accessRights fullaccess

	set-mailbox $CR_Alias -grantsendonbehalfto $list.split(",")
	
	add-adpermission $(safeGetMailbox($CR_Alias)).name -extendedrights send-as -user $fullGroupName

#End Resource Mailbox Creation
#Create Calendar Mailbox	
	set-mailbox $CR_Alias -type room
<#	if (!$?){
		outData("Unable to set account to type Room")
		$alreadyRoom = queryYN "Is this account already a room?" " "
		if ($alreadyRoom -eq 1){
			outData("Error setting account to type Room, please contact the Microsoft Team")
			Exit -1
		}
	}#>
	enable-distributiongroup ($CR_Alias + "-group")
	#outData("Please wait 90 seconds for data to populate")
	Start-sleep -s 90
	set-distributiongroup ($CR_Alias + "-group") -HiddenFromAddressListsEnabled $true
	set-distributiongroup ($CR_Alias + "-group") -acceptmessagesonlyfromdlmembers ($account + "-group")

	set-calendarprocessing $CR_Alias -resourcedelegates ($CR_Alias + "-group")
	set-calendarprocessing $CR_Alias -allowconflicts $false -automateprocessing autoaccept -maximumdurationinminutes 1440

	if ($CR_Restrict.Checked){
		#submitWithApproval($account)
		#autoScheduleResources($account)
		if($CR_Autobook -ne ""){
			set-calendarprocessing $CR_Alias -bookinpolicy $CR_Autobook.split(",")
		}
		if($CR_RB_Anyone){
			Set-CalendarProcessing $CR_Alias -AllBookInPolicy $true
		}
		if($CR_RB_AprovalList){
			set-calendarprocessing $CR_Alias -allbookinpolicy $false
			set-calendarprocessing $CR_Alias -requestinpolicy $CR_Approved.split(",")
		}
		if($CR_RB_Noone){}
	}
	#outData("Calendar-Only Mailbox Set")
}
#End Create Calendar Mailbox


#endregion Create Room

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 425
$System_Drawing_Size.Width = 660
$form1.ClientSize = $System_Drawing_Size
$form1.MaximumSize = $System_Drawing_Size
$form1.MinimumSize = $System_Drawing_Size
$form1.MaximizeBox = $false
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Program Files\Common Files\batman.ico')
$form1.Name = "SCGUI"
$form1.Text = "SCGUI"
$form1.StartPosition = 1

$tabControl1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 0
$tabControl1.Location = $System_Drawing_Point
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 388
$System_Drawing_Size.Width = 645
$tabControl1.Size = $System_Drawing_Size
$tabControl1.TabIndex = 0

$form1.Controls.Add($tabControl1)

#region Create DL
$tabPageDL.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabPageDL.Location = $System_Drawing_Point
$tabPageDL.Name = "tabPageDL"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPageDL.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 362
$System_Drawing_Size.Width = 637
$tabPageDL.Size = $System_Drawing_Size
$tabPageDL.TabIndex = 0
$tabPageDL.Text = "Create DL"
$tabPageDL.UseVisualStyleBackColor = $True

$tabControl1.Controls.Add($tabPageDL)

$groupBox2.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 428
$System_Drawing_Point.Y = 7
$groupBox2.Location = $System_Drawing_Point
$groupBox2.Name = "groupBox2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 168
$System_Drawing_Size.Width = 200
$groupBox2.Size = $System_Drawing_Size
$groupBox2.TabIndex = 13
$groupBox2.TabStop = $False
$groupBox2.Text = "Members"

$tabPageDL.Controls.Add($groupBox2)

$CB_Sponsors.Checked = $True
$CB_Sponsors.CheckState = 1
$CB_Sponsors.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 150
$CB_Sponsors.Location = $System_Drawing_Point
$CB_Sponsors.Name = "CB_Sponsors"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 187
$CB_Sponsors.Size = $System_Drawing_Size
$CB_Sponsors.TabIndex = 11
$CB_Sponsors.Text = "Add Sponsors?"
$CB_Sponsors.UseVisualStyleBackColor = $True

$groupBox2.Controls.Add($CB_Sponsors)

$RTB_Members.DataBindings.DefaultDataSourceUpdateMode = 0
$RTB_Members.DetectUrls = $False
$RTB_Members.EnableAutoDragDrop = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 14
$RTB_Members.Location = $System_Drawing_Point
$RTB_Members.Name = "RTB_Members"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 131
$System_Drawing_Size.Width = 187
$RTB_Members.Size = $System_Drawing_Size
$RTB_Members.TabIndex = 10
$RTB_Members.Text = ""

$groupBox2.Controls.Add($RTB_Members)

$label4.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 152
$System_Drawing_Point.Y = 158
$label4.Location = $System_Drawing_Point
$label4.Name = "label4"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 51
$label4.Size = $System_Drawing_Size
#$label4.TabIndex = 12
$label4.Text = "@ou.edu"
$label4.add_Click($handler_label4_Click)

$tabPageDL.Controls.Add($label4)

$RTB_DLOutput.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 181
$RTB_DLOutput.Location = $System_Drawing_Point
$RTB_DLOutput.Name = "RTB_DLOutput"
$RTB_DLOutput.ReadOnly = $true
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 175
$System_Drawing_Size.Width = 625
$RTB_DLOutput.Size = $System_Drawing_Size
#$RTB_DLOutput.TabIndex = 11
$RTB_DLOutput.Text = ""

$tabPageDL.Controls.Add($RTB_DLOutput)


$Create.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 330
$System_Drawing_Point.Y = 152
$Create.Location = $System_Drawing_Point
$Create.Name = "Create"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 91
$Create.Size = $System_Drawing_Size
$Create.TabIndex = 12
$Create.Text = "Create"
$Create.UseVisualStyleBackColor = $True
$Create.add_Click($Create_OnClick)

$tabPageDL.Controls.Add($Create)


$ListInGAL.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 155
$ListInGAL.Location = $System_Drawing_Point
$ListInGAL.Name = "ListInGAL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$ListInGAL.Size = $System_Drawing_Size
$ListInGAL.TabIndex = 5
$ListInGAL.Text = "List in GAL?"
$ListInGAL.Checked = $true
$ListInGAL.UseVisualStyleBackColor = $True

$tabPageDL.Controls.Add($ListInGAL)


$groupBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 213
$System_Drawing_Point.Y = 7
$groupBox1.Location = $System_Drawing_Point
$groupBox1.Name = "groupBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 86
$System_Drawing_Size.Width = 208
$groupBox1.Size = $System_Drawing_Size
$groupBox1.TabIndex = 6
$groupBox1.TabStop = $False
$groupBox1.Text = "Who can send to list?"

$tabPageDL.Controls.Add($groupBox1)

$RB_IO.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 116
$System_Drawing_Point.Y = 50
$RB_IO.Location = $System_Drawing_Point
$RB_IO.Name = "RB_IO"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 85
$RB_IO.Size = $System_Drawing_Size
$RB_IO.TabIndex = 9
$RB_IO.TabStop = $True
$RB_IO.Text = "Internal Only"
$RB_IO.UseVisualStyleBackColor = $True

$groupBox1.Controls.Add($RB_IO)


$RB_OO.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 117
$System_Drawing_Point.Y = 20
$RB_OO.Location = $System_Drawing_Point
$RB_OO.Name = "RB_OO"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 85
$RB_OO.Size = $System_Drawing_Size
$RB_OO.TabIndex = 8
$RB_OO.TabStop = $True
$RB_OO.Text = "Owners Only"
$RB_OO.UseVisualStyleBackColor = $True

$groupBox1.Controls.Add($RB_OO)


$RB_MO.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 49
$RB_MO.Location = $System_Drawing_Point
$RB_MO.Name = "RB_MO"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$RB_MO.Size = $System_Drawing_Size
$RB_MO.TabIndex = 7
$RB_MO.TabStop = $True
$RB_MO.Text = "Members Only"
$RB_MO.UseVisualStyleBackColor = $True

$groupBox1.Controls.Add($RB_MO)


$RB_A.Checked = $True
$RB_A.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 19
$RB_A.Location = $System_Drawing_Point
$RB_A.Name = "RB_A"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$RB_A.Size = $System_Drawing_Size
$RB_A.TabIndex = 6
$RB_A.TabStop = $True
$RB_A.Text = "Anyone"
$RB_A.UseVisualStyleBackColor = $True

$groupBox1.Controls.Add($RB_A)

$TB_EmailAddress.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 155
$TB_EmailAddress.Location = $System_Drawing_Point
$TB_EmailAddress.Name = "TB_EmailAddress"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 145
$TB_EmailAddress.Size = $System_Drawing_Size
$TB_EmailAddress.TabIndex = 4

$tabPageDL.Controls.Add($TB_EmailAddress)

$label3.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 137
$label3.Location = $System_Drawing_Point
$label3.Name = "label3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 195
$label3.Size = $System_Drawing_Size
#$label3.TabIndex = 6
$label3.Text = "@ou.edu Address (Optional)"

$tabPageDL.Controls.Add($label3)

$TB_Sponsors.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 114
$TB_Sponsors.Location = $System_Drawing_Point
$TB_Sponsors.Name = "TB_Sponsors"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 197
$TB_Sponsors.Size = $System_Drawing_Size
$TB_Sponsors.TabIndex = 3

$tabPageDL.Controls.Add($TB_Sponsors)

$label2.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 96
$label2.Location = $System_Drawing_Point
$label2.Name = "label2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 197
$label2.Size = $System_Drawing_Size
#$label2.TabIndex = 4
$label2.Text = "Comma-Separated List of Sponsors"

$tabPageDL.Controls.Add($label2)

$TB_Alias.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 73
$TB_Alias.Location = $System_Drawing_Point
$TB_Alias.Name = "TB_Alias"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 197
$TB_Alias.Size = $System_Drawing_Size
$TB_Alias.TabIndex = 2

$tabPageDL.Controls.Add($TB_Alias)


$UseCustomAlias.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 52
$UseCustomAlias.Location = $System_Drawing_Point
$UseCustomAlias.Name = "UseCustomAlias"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 197
$UseCustomAlias.Size = $System_Drawing_Size
$UseCustomAlias.TabIndex = 1
$UseCustomAlias.Text = "Use Custom Alias"
$UseCustomAlias.UseVisualStyleBackColor = $True

$tabPageDL.Controls.Add($UseCustomAlias)

$TB_DGN.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 21
$TB_DGN.Location = $System_Drawing_Point
$TB_DGN.Name = "TB_DGN"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 197
$TB_DGN.Size = $System_Drawing_Size
$TB_DGN.TabIndex = 0

$tabPageDL.Controls.Add($TB_DGN)

$label1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 3
$label1.Location = $System_Drawing_Point
$label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$label1.Size = $System_Drawing_Size
#$label1.TabIndex = 0
$label1.Text = "Distribution Group Name"

$tabPageDL.Controls.Add($label1)
#endregion Create DL

#region Increase Quota
$tabPageIQ.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabPageIQ.Location = $System_Drawing_Point
$tabPageIQ.Name = "tabPageIQ"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPageIQ.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 364
$System_Drawing_Size.Width = 637
$tabPageIQ.Size = $System_Drawing_Size
$tabPageIQ.TabIndex = 0
$tabPageIQ.Text = "Increase Quota"
$tabPageIQ.UseVisualStyleBackColor = $True

$tabControl1.Controls.Add($tabPageIQ)

$B_SetCustom.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 116
$B_SetCustom.Location = $System_Drawing_Point
$B_SetCustom.Name = "B_SetCustom"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 75
$B_SetCustom.Size = $System_Drawing_Size
$B_SetCustom.TabIndex = 11
$B_SetCustom.Text = "Set Custom"
$B_SetCustom.UseVisualStyleBackColor = $True
$B_SetCustom.add_Click($B_SetCustom_OnClick)

$tabPageIQ.Controls.Add($B_SetCustom)

$label_IQ_12.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 188
$System_Drawing_Point.Y = 119
$label_IQ_12.Location = $System_Drawing_Point
$label_IQ_12.Name = "label_IQ_12"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 23
$label_IQ_12.Size = $System_Drawing_Size
$label_IQ_12.TabIndex = 10
$label_IQ_12.Text = "GB"

$tabPageIQ.Controls.Add($label_IQ_12)

$TB_ProhibitSend.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 111
$System_Drawing_Point.Y = 116
$TB_ProhibitSend.Location = $System_Drawing_Point
$TB_ProhibitSend.Name = "TB_ProhibitSend"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 77
$TB_ProhibitSend.Size = $System_Drawing_Size
$TB_ProhibitSend.TabIndex = 9

$tabPageIQ.Controls.Add($TB_ProhibitSend)

$label_IQ_11.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 97
$label_IQ_11.Location = $System_Drawing_Point
$label_IQ_11.Name = "label_IQ_11"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_IQ_11.Size = $System_Drawing_Size
$label_IQ_11.TabIndex = 8
$label_IQ_11.Text = "Prohibit Send:"

$tabPageIQ.Controls.Add($label_IQ_11)

$label_IQ_10.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 81
$System_Drawing_Point.Y = 119
$label_IQ_10.Location = $System_Drawing_Point
$label_IQ_10.Name = "label_IQ_10"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 23
$label_IQ_10.Size = $System_Drawing_Size
$label_IQ_10.TabIndex = 7
$label_IQ_10.Text = "GB"

$tabPageIQ.Controls.Add($label_IQ_10)

$TB_IssueWarning.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 116
$TB_IssueWarning.Location = $System_Drawing_Point
$TB_IssueWarning.Name = "TB_IssueWarning"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 77
$TB_IssueWarning.Size = $System_Drawing_Size
$TB_IssueWarning.TabIndex = 6

$tabPageIQ.Controls.Add($TB_IssueWarning)

$label_IQ_9.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 97
$label_IQ_9.Location = $System_Drawing_Point
$label_IQ_9.Name = "label_IQ_9"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_IQ_9.Size = $System_Drawing_Size
$label_IQ_9.TabIndex = 5
$label_IQ_9.Text = "Issue Warning:"

$tabPageIQ.Controls.Add($label_IQ_9)


$B_IQ_SetDefault.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 70
$B_IQ_SetDefault.Location = $System_Drawing_Point
$B_IQ_SetDefault.Name = "B_IQ_SetDefault"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$B_IQ_SetDefault.Size = $System_Drawing_Size
$B_IQ_SetDefault.TabIndex = 4
$B_IQ_SetDefault.Text = "Set Default 6GB"
$B_IQ_SetDefault.UseVisualStyleBackColor = $True
$B_IQ_SetDefault.add_Click($B_IQ_SetDefault_OnClick)

$tabPageIQ.Controls.Add($B_IQ_SetDefault)

$label_IQ_Out.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 51
$label_IQ_Out.Location = $System_Drawing_Point
$label_IQ_Out.Name = "label_IQ_Out"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 350
$label_IQ_Out.Size = $System_Drawing_Size
$label_IQ_Out.TabIndex = 3

$tabPageIQ.Controls.Add($label_IQ_Out)


$B_IQ_Check.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 23
$B_IQ_Check.Location = $System_Drawing_Point
$B_IQ_Check.Name = "B_IQ_Check"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 75
$B_IQ_Check.Size = $System_Drawing_Size
$B_IQ_Check.TabIndex = 2
$B_IQ_Check.Text = "Check"
$B_IQ_Check.UseVisualStyleBackColor = $True
$B_IQ_Check.add_Click($B_IQ_Check_OnClick)

$tabPageIQ.Controls.Add($B_IQ_Check)

$TB_IQName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 23
$TB_IQName.Location = $System_Drawing_Point
$TB_IQName.Name = "TB_IQName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$TB_IQName.Size = $System_Drawing_Size
$TB_IQName.TabIndex = 1

$tabPageIQ.Controls.Add($TB_IQName)

$label_IQ_7.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 4
$label_IQ_7.Location = $System_Drawing_Point
$label_IQ_7.Name = "label_IQ_7"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$label_IQ_7.Size = $System_Drawing_Size
$label_IQ_7.TabIndex = 0
$label_IQ_7.Text = "4x4 or Email"

$tabPageIQ.Controls.Add($label_IQ_7)

#endregion Increase Quota

#region Set MBPerms

$tabPageMBP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabPageMBP.Location = $System_Drawing_Point
$tabPageMBP.Name = "tabPageMBP"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPageMBP.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 364
$System_Drawing_Size.Width = 637
$tabPageMBP.Size = $System_Drawing_Size
$tabPageMBP.TabIndex = 0
$tabPageMBP.Text = "Mailbox Permissions"
$tabPageMBP.UseVisualStyleBackColor = $True

$tabControl1.Controls.Add($tabPageMBP)
$RTB_MBP_Output.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 9
$System_Drawing_Point.Y = 171
$RTB_MBP_Output.Location = $System_Drawing_Point
$RTB_MBP_Output.Name = "RTB_MBP_Output"
$RTB_MBP_Output.ReadOnly = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 182
$System_Drawing_Size.Width = 619
$RTB_MBP_Output.Size = $System_Drawing_Size
$RTB_MBP_Output.TabIndex = 7
$RTB_MBP_Output.Text = ""

$tabPageMBP.Controls.Add($RTB_MBP_Output)


$B_MBP_Set.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 131
$System_Drawing_Point.Y = 142
$B_MBP_Set.Location = $System_Drawing_Point
$B_MBP_Set.Name = "B_MBP_Set"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$B_MBP_Set.Size = $System_Drawing_Size
$B_MBP_Set.TabIndex = 6
$B_MBP_Set.Text = "Set"
$B_MBP_Set.UseVisualStyleBackColor = $True
$B_MBP_Set.add_Click($B_MBP_Set_OnClick)

$tabPageMBP.Controls.Add($B_MBP_Set)

$TB_MBP_Target.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 9
$System_Drawing_Point.Y = 115
$TB_MBP_Target.Location = $System_Drawing_Point
$TB_MBP_Target.Name = "TB_MBP_Target"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 198
$TB_MBP_Target.Size = $System_Drawing_Size
$TB_MBP_Target.TabIndex = 5

$tabPageMBP.Controls.Add($TB_MBP_Target)

$label_MBP_2.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 96
$label_MBP_2.Location = $System_Drawing_Point
$label_MBP_2.Name = "label_MBP_2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$label_MBP_2.Size = $System_Drawing_Size
$label_MBP_2.TabIndex = 4
$label_MBP_2.Text = "Target 4x4 or Email"

$tabPageMBP.Controls.Add($label_MBP_2)


$CB_MBP_FullAccess.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 74
$CB_MBP_FullAccess.Location = $System_Drawing_Point
$CB_MBP_FullAccess.Name = "CB_MBP_FullAccess"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 104
$CB_MBP_FullAccess.Size = $System_Drawing_Size
$CB_MBP_FullAccess.TabIndex = 3
$CB_MBP_FullAccess.Text = "Full Access"
$CB_MBP_FullAccess.UseVisualStyleBackColor = $True

$tabPageMBP.Controls.Add($CB_MBP_FullAccess)


$CB_MBP_SendAs.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 53
$CB_MBP_SendAs.Location = $System_Drawing_Point
$CB_MBP_SendAs.Name = "CB_MBP_SendAs"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 104
$CB_MBP_SendAs.Size = $System_Drawing_Size
$CB_MBP_SendAs.TabIndex = 2
$CB_MBP_SendAs.Text = "Send As"
$CB_MBP_SendAs.UseVisualStyleBackColor = $True

$tabPageMBP.Controls.Add($CB_MBP_SendAs)

$TB_MBP_User.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 26
$TB_MBP_User.Location = $System_Drawing_Point
$TB_MBP_User.Name = "TB_MBP_User"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$TB_MBP_User.Size = $System_Drawing_Size
$TB_MBP_User.TabIndex = 1

$tabPageMBP.Controls.Add($TB_MBP_User)

$label_mbp_1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 7
$label_mbp_1.Location = $System_Drawing_Point
$label_mbp_1.Name = "label_mbp_1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_mbp_1.Size = $System_Drawing_Size
$label_mbp_1.TabIndex = 0
$label_mbp_1.Text = "User 4x4 or Email"

$tabPageMBP.Controls.Add($label_mbp_1)

#endregion SetMBPerms

#region CMS Group

$CMS_Tab.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$CMS_Tab.Location = $System_Drawing_Point
$CMS_Tab.Name = "CMS_Tab"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$CMS_Tab.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 364
$System_Drawing_Size.Width = 637
$CMS_Tab.Size = $System_Drawing_Size
$CMS_Tab.TabIndex = 0
$CMS_Tab.Text = "CMS Group"
$CMS_Tab.UseVisualStyleBackColor = $True
#$CMS_Tab.add_Click($handler_TabPage:_Click)

$tabControl1.Controls.Add($CMS_Tab)
$CMS_Status.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 321
$CMS_Progress.Location = $System_Drawing_Point
$CMS_Progress.Name = "CMS_Progress"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 200
$CMS_Progress.Size = $System_Drawing_Size
#$CMS_Progress.TabIndex = 9
$CMS_Progress.add_Click($handler_CMS_Progress_Click)

$CMS_Tab.Controls.Add($CMS_Progress)

$CMS_Status.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 303
$CMS_Status.Location = $System_Drawing_Point
$CMS_Status.Name = "CMS_Status"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 400
$CMS_Status.Size = $System_Drawing_Size
#$CMS_Status.TabIndex = 8
#$CMS_Status.TextAlign = 4

$CMS_Tab.Controls.Add($CMS_Status)

$CMS_Members.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 115
$CMS_Members.Location = $System_Drawing_Point
$CMS_Members.Name = "CMS_Members"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 156
$System_Drawing_Size.Width = 197
$CMS_Members.Size = $System_Drawing_Size
$CMS_Members.TabIndex = 3
$CMS_Members.Text = ""

$CMS_Tab.Controls.Add($CMS_Members)

$CMS_label3.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 96
$CMS_label3.Location = $System_Drawing_Point
$CMS_label3.Name = "CMS_label3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$CMS_label3.Size = $System_Drawing_Size
#$CMS_label3.TabIndex = 6
$CMS_label3.Text = "Members"

$CMS_Tab.Controls.Add($CMS_label3)


$CMS_Create.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 129
$System_Drawing_Point.Y = 277
$CMS_Create.Location = $System_Drawing_Point
$CMS_Create.Name = "CMS_Create"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$CMS_Create.Size = $System_Drawing_Size
$CMS_Create.TabIndex = 4
$CMS_Create.Text = "Create"
$CMS_Create.UseVisualStyleBackColor = $True
$CMS_Create.add_Click($handler_CMS_Create_Click)

$CMS_Tab.Controls.Add($CMS_Create)

$CMS_Owners.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 69
$CMS_Owners.Location = $System_Drawing_Point
$CMS_Owners.Name = "CMS_Owners"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CMS_Owners.Size = $System_Drawing_Size
$CMS_Owners.TabIndex = 2

$CMS_Tab.Controls.Add($CMS_Owners)

$CMS_label2.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 50
$CMS_label2.Location = $System_Drawing_Point
$CMS_label2.Name = "CMS_label2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$CMS_label2.Size = $System_Drawing_Size
#$CMS_label2.TabIndex = 2
$CMS_label2.Text = "Owners"

$CMS_Tab.Controls.Add($CMS_label2)

$CMS_label1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 4
$CMS_label1.Location = $System_Drawing_Point
$CMS_label1.Name = "CMS_label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$CMS_label1.Size = $System_Drawing_Size
#$CMS_label1.TabIndex = 1
$CMS_label1.Text = "CMS Group Name"
#$CMS_label1.add_Click($handler_label1_Click)

$CMS_Tab.Controls.Add($CMS_label1)

$CMS_Group.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 23
$CMS_Group.Location = $System_Drawing_Point
$CMS_Group.Name = "CMS_Group"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CMS_Group.Size = $System_Drawing_Size
$CMS_Group.TabIndex = 0
#$CMS_Group.add_TextChanged($handler_textBox1_TextChanged)

$CMS_Tab.Controls.Add($CMS_Group)
#endregion CMS Group

#region Create Room
#$CR_Tab.Controls.Add($tabControl1)
$CR_Tab.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$CR_Tab.Location = $System_Drawing_Point
$CR_Tab.Name = "CR_Tab"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$CR_Tab.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 364
$System_Drawing_Size.Width = 637
$CR_Tab.Size = $System_Drawing_Size
$CR_Tab.TabIndex = 0
$CR_Tab.Text = "Create Room"
$CR_Tab.UseVisualStyleBackColor = $True

$tabControl1.Controls.Add($CR_Tab)
$CR_Progress.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 9
$System_Drawing_Point.Y = 335
$CR_Progress.Location = $System_Drawing_Point
$CR_Progress.Name = "CR_Progress"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 619
$CR_Progress.Size = $System_Drawing_Size
$CR_Progress.TabIndex = 20

$CR_Tab.Controls.Add($CR_Progress)

$CR_Output.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 9
$System_Drawing_Point.Y = 256
$CR_Output.Location = $System_Drawing_Point
$CR_Output.Name = "CR_Output"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 76
$System_Drawing_Size.Width = 619
$CR_Output.Size = $System_Drawing_Size
#$CR_Output.TabIndex = 19

$CR_Tab.Controls.Add($CR_Output)


$CR_Create.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 9
$System_Drawing_Point.Y = 206
$CR_Create.Location = $System_Drawing_Point
$CR_Create.Name = "CR_Create"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 100
$CR_Create.Size = $System_Drawing_Size
#$CR_Create.TabIndex = 18
$CR_Create.Text = "Create"
$CR_Create.UseVisualStyleBackColor = $True
$CR_Create.add_Click($CR_Create_OnClick)

$CR_Tab.Controls.Add($CR_Create)


$CR_groupBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 227
$System_Drawing_Point.Y = 12
$CR_groupBox1.Location = $System_Drawing_Point
$CR_groupBox1.Name = "CR_groupBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 226
$System_Drawing_Size.Width = 250
$CR_groupBox1.Size = $System_Drawing_Size
#$CR_groupBox1.TabIndex = 17
$CR_groupBox1.TabStop = $False
$CR_groupBox1.Text = "Meeting Requests"

$CR_Tab.Controls.Add($CR_groupBox1)

$CR_RB_AprovalList.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 162
$CR_RB_AprovalList.Location = $System_Drawing_Point
$CR_RB_AprovalList.Name = "CR_RB_AprovalList"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 240
$CR_RB_AprovalList.Size = $System_Drawing_Size
#$CR_RB_AprovalList.TabIndex = 8
$CR_RB_AprovalList.TabStop = $True
$CR_RB_AprovalList.Text = "Comma-Separated List of who can subit requests with approval"
$CR_RB_AprovalList.UseVisualStyleBackColor = $True

$CR_groupBox1.Controls.Add($CR_RB_AprovalList)


$CR_RB_Noone.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 118
$System_Drawing_Point.Y = 134
$CR_RB_Noone.Location = $System_Drawing_Point
$CR_RB_Noone.Name = "CR_RB_Noone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$CR_RB_Noone.Size = $System_Drawing_Size
#$CR_RB_Noone.TabIndex = 7
$CR_RB_Noone.TabStop = $True
$CR_RB_Noone.Text = "No one "
$CR_RB_Noone.UseVisualStyleBackColor = $True
#$CR_RB_Noone.add_CheckedChanged($handler_radioButton3_CheckedChanged)

$CR_groupBox1.Controls.Add($CR_RB_Noone)


$CR_RB_Anyone.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 135
$CR_RB_Anyone.Location = $System_Drawing_Point
$CR_RB_Anyone.Name = "CR_RB_Anyone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$CR_RB_Anyone.Size = $System_Drawing_Size
#$CR_RB_Anyone.TabIndex = 6
$CR_RB_Anyone.TabStop = $True
$CR_RB_Anyone.Text = "Anyone"
$CR_RB_Anyone.UseVisualStyleBackColor = $True

$CR_groupBox1.Controls.Add($CR_RB_Anyone)

$CR_label5.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 117
$CR_label5.Location = $System_Drawing_Point
$CR_label5.Name = "label5"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 240
$CR_label5.Size = $System_Drawing_Size
#$CR_label5.TabIndex = 5
$CR_label5.Text = "Who can submit requests with approval?"
#$CR_label5.add_Click($handler_label5_Click)

$CR_groupBox1.Controls.Add($CR_label5)

$CR_Autobook.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 84
$CR_Autobook.Location = $System_Drawing_Point
$CR_Autobook.Name = "CR_Autobook"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 187
$CR_Autobook.Size = $System_Drawing_Size
#$CR_Autobook.TabIndex = 4

$CR_groupBox1.Controls.Add($CR_Autobook)

$CR_label4.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 50
$CR_label4.Location = $System_Drawing_Point
$CR_label4.Name = "label4"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 240
$CR_label4.Size = $System_Drawing_Size
$CR_label4.TabIndex = 3
$CR_label4.Text = "Comma-Separated List of who can auto-book resouces (including sponsors)"

$CR_groupBox1.Controls.Add($CR_label4)

$CR_Approved.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 198
$CR_Approved.Location = $System_Drawing_Point
$CR_Approved.Name = "CR_Approved"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CR_Approved.Size = $System_Drawing_Size
#$CR_Approved.TabIndex = 1
#$CR_Approved.add_TextChanged($handler_textBox4_TextChanged)

$CR_groupBox1.Controls.Add($CR_Approved)


$CR_Restrict.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 19
$CR_Restrict.Location = $System_Drawing_Point
$CR_Restrict.Name = "CR_Restrict"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 188
$CR_Restrict.Size = $System_Drawing_Size
#$CR_Restrict.TabIndex = 0
$CR_Restrict.Text = "Restrict Meeting Requests"
$CR_Restrict.UseVisualStyleBackColor = $True

$CR_groupBox1.Controls.Add($CR_Restrict)


$CR_Sponsors.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 158
$CR_Sponsors.Location = $System_Drawing_Point
$CR_Sponsors.Name = "CR_Sponsors"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CR_Sponsors.Size = $System_Drawing_Size
#$CR_Sponsors.TabIndex = 16

$CR_Tab.Controls.Add($CR_Sponsors)

$CR_label2.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 139
$CR_label2.Location = $System_Drawing_Point
$CR_label2.Name = "label2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 210
$CR_label2.Size = $System_Drawing_Size
#$CR_label2.TabIndex = 15
$CR_label2.Text = "Comma-Separated List of Sponsors (2)"

$CR_Tab.Controls.Add($CR_label2)

$CR_label1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 94
$CR_label1.Location = $System_Drawing_Point
$CR_label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$CR_label1.Size = $System_Drawing_Size
#$CR_label1.TabIndex = 14
$CR_label1.Text = "Account Owner"

$CR_Tab.Controls.Add($CR_label1)

$CR_Owner.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 112
$CR_Owner.Location = $System_Drawing_Point
$CR_Owner.Name = "CR_Owner"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CR_Owner.Size = $System_Drawing_Size
#$CR_Owner.TabIndex = 13

$CR_Tab.Controls.Add($CR_Owner)

$CR_Alias.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 71
$CR_Alias.Location = $System_Drawing_Point
$CR_Alias.Name = "CR_Alias"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CR_Alias.Size = $System_Drawing_Size
#$CR_Alias.TabIndex = 12

$CR_Tab.Controls.Add($CR_Alias)

$CR_DispName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 30
$CR_DispName.Location = $System_Drawing_Point
$CR_DispName.Name = "CR_DispName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$CR_DispName.Size = $System_Drawing_Size
#$CR_DispName.TabIndex = 1

$CR_Tab.Controls.Add($CR_DispName)

$CR_label6.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 12
$CR_label6.Location = $System_Drawing_Point
$CR_label6.Name = "CR_label6"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$CR_label6.Size = $System_Drawing_Size
#$CR_label6.TabIndex = 0
$CR_label6.Text = "Room Display Name (Numbers OK)"
#$CR_label6.add_Click($handler_label_CR_1_Click)

$CR_Tab.Controls.Add($CR_label6)

$CR_label7.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 54
$CR_label7.Location = $System_Drawing_Point
$CR_label7.Name = "CR_label7"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 200
$CR_label7.Size = $System_Drawing_Size
#$CR_label7.TabIndex = 0
$CR_label7.Text = "Room Alias"
#$CR_label7.add_Click($handler_label_CR_1_Click)

$CR_Tab.Controls.Add($CR_label7)
#endregion Create Room

#region Console Page
$tabPageConsole.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabPageConsole.Location = $System_Drawing_Point
$tabPageConsole.Name = "tabPageConsole"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPageConsole.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 362
$System_Drawing_Size.Width = 637
$tabPageConsole.Size = $System_Drawing_Size
$tabPageConsole.TabIndex = 2
$tabPageConsole.Text = "Console"
$tabPageConsole.UseVisualStyleBackColor = $True

$tabControl1.Controls.Add($tabPageConsole)

$RTB_Console.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 4
$RTB_Console.Location = $System_Drawing_Point
$RTB_Console.Name = "RTB_Console"
$RTB_Console.ReadOnly = $true
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 352
$System_Drawing_Size.Width = 627
$RTB_Console.Size = $System_Drawing_Size
$RTB_Console.TabIndex = 0
$RTB_Console.Text = ""

$tabPageConsole.Controls.Add($RTB_Console)
#endregion Console Page

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

} #End Function

function safeGetMailbox($data){
	if($data -eq $null){
		$data = "NoData"
	}
	return get-mailbox $data
}

function updateConsole {
	$RTB_Console.Select($RTB_DLOutput.Text.Length, 0)
	$RTB_Console.ScrollToCaret()
	$form1.Update()
}

function updateDLOutput {
	$RTB_DLOutput.Select($RTB_DLOutput.Text.Length, 0)
	$RTB_DLOutput.ScrollToCaret()
	$form1.Update()
}

function updateMBPOutput {
	$RTB_MBP_Output.Select($RTB_MBP_Output.Text.Length, 0)
	$RTB_MBP_Output.ScrolltoCaret()
	$form1.Update()
}

<#
function accessLog($log) {
	$logFileDir = "\\norfile.net.ou.edu\ouit\Community Experience\Services\Access Logs"
	$logFileName = ((whoami).split("\"))[1] + "{0:ddMMyy}" -f (Get-Date) + ".log"
	$logPathName = Join-Path $logFileDir -ChildPath $logFileName
	
	if(!(Test-Path $logPathName)) {
		New-Item $logPathName -type file
	}
	
	Add-Content $logPathName $log
}
#>

#Sends an email
function sendMail($account, $subject, $body)
{
	$SmtpClient = new-object system.net.mail.smtpClient 
	$SmtpClient.Host = "smtp.ou.edu"

	$address = safeGetMailbox($account)
	$address = $address.PrimarySmtpAddress.toString()

	$MailMessage = New-Object system.net.mail.mailmessage 
	$mailmessage.from = "postmaster@ou.edu"
	$mailmessage.To.add($address)
	$mailmessage.Subject = $subject
	$MailMessage.IsBodyHtml = $False
	$mailmessage.Body = $body
	$mailmessage.Priority = "Normal"
	$smtpclient.Send($mailmessage) 
}

#Register-EngineEvent Powershell.Exiting -Action {}
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer it-aero.sooner.net.ou.edu
#Call the Function
GenerateForm
exit

####---------------------------------------------------------------------------------------------####
# Service Center Exchange GUI																		#
# Written by: Austin Heyne																			#
# 																									#
# --Information--																					#
# Program provides an interface to manage day to day operations in Exchange 2010. This is the		#
# non-administrative version, it has substantailly more error handling than the Admin version but 	#
# also takes longer to open due to creating all the custom error records. It also has sligtly less	#
# functionality as restricted functions have been removed.											#
####---------------------------------------------------------------------------------------------####

#Generated Form Function
function GenerateForm {

	#region Import the Assemblies
	[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
	[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
	#endregion

	#region Form Objects
	$SCEGUI = New-Object System.Windows.Forms.Form
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	$tabControl1 = New-Object System.Windows.Forms.TabControl
	#DL Tab
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
	#Quota Tab
	$tabPageIQ = New-Object System.Windows.Forms.TabPage
	$B_IQ_SetMax = New-Object System.Windows.Forms.Button
	$B_IQ_SetDefault = New-Object System.Windows.Forms.Button
	$label_IQ_Out = New-Object System.Windows.Forms.Label
	$B_IQ_Check = New-Object System.Windows.Forms.Button
	$TB_IQName = New-Object System.Windows.Forms.TextBox
	$label_IQ_7 = New-Object System.Windows.Forms.Label
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
	#roomview
	$RV_Tab = New-Object System.Windows.Forms.TabPage
	$RV_Status = New-Object System.Windows.Forms.Label
	$RV_button1 = New-Object System.Windows.Forms.Button
	$RV_textBox2 = New-Object System.Windows.Forms.TextBox
	$RV_label2 = New-Object System.Windows.Forms.Label
	$RV_label1 = New-Object System.Windows.Forms.Label
	$RV_textBox1 = New-Object System.Windows.Forms.TextBox
	#Console
	$tabPageConsole = New-Object System.Windows.Forms.TabPage
	$RTB_Console = New-Object System.Windows.Forms.RichTextBox
	#endregion Form Objects

	#region Create DL
	$Create_OnClick= 
	{	
		#Clear any text from DLOutput
		. updateDLOutput ""
		. updateDLOutput "Creating DL..."
		
		try{ 
	#-------Import and Sanatize Data-------#
			Write-debug "1"
			if(-not $TB_DGN.Text -match $global:regex_DGN){
				Write-Debug $TB_DGN.Text
				throw $Global:Error_InvalidDGN
			}
			if($UseCustomAlias.Checked -is $true){
				if(-not $TB_Alias.Text -match $global:regex_Alias){
					throw $Global:Error_InvalidAlias
				}
			}
			if(-not $TB_EmailAddress.Text -match $global:regex_Email){
				throw $Global:Error_InvalidEmailAddress
			}
			if(-not $TB_Sponsors.Text -match $global:regex_SponsorList){
				$msg = "There is an invalid charater in the Sponsors list, please correct."
				popupError($msg,0,"Invalid Imput",0x0)
				throw $Global:Error_BadSponsor
			}
			if(-not $RTB_Members.Text -match $global:regex_MemberList){
				$msg = "There is an invalid character in the Members list, please correct."
				popupError($msg,0,"Invalid Imput",0x0)
				throw $Global:Error_BadMember
			}
			
	#-------Alias - Handle Generated Alias or Custom Alias-------#
			if($UseCustomAlias.Checked){
				$DL_Alias = $TB_Alias.Text
			} else {
				$DL_Alias = $TB_DGN.Replace(" ","").Replace("(","").Replace(")","").Replace(",","").Replace("-","")
			}
			#Make sure the group doens't exist already, this happens more than you would think.
			Get-ADGroup $DL_Alias
			if($?){
				throw $Global:Error_AlreadyExists
			}
			
			. updateDLOutput ("DGN: " + $TB_DGN)
			. updateDLOutput ("Alias: " + $DL_Alias)
		
	#-------Sponsors - Handle Sponsor Check-------#
			if($TB_Sponsors.Text.Replace(" ","") -match $global:regex_SponsorList){
				$sponsorList = $S_Sponsors.Text.Replace(" ","").Split(",")
				$sponsorList = $sponsorList | Select-Object -Unique
				
				if($sponsorList.Length -lt 2){
					. updateDLOutput "Not enough sponsors provided."
					throw $Global:Error_InsufficientSponsors
				} else {
					foreach($_ in $sponsorList){
						#Sanitize each sponsor
						if(-not $_ -match $global:regex_4x4){ 
							popupError(($_ + " is not a valid 4x4."),0,"Invalid Imput",0x0)
							throw $Global:Error_BadSponsor
						}
						#check existance
						Get-ADUser $_
						if(!$?){
							popupError(($_ + " is not a valid 4x4"),0,"Invalid 4x4",0x0)
							throw $Global:Error_BadSponsor
						}
					}
					. updateDLOutput "Sponsors Approved"
				}
			} else {
				throw $Global:Error_InvalidSponsorList
			}
			
	#-------Members - Handle Member Check-------#
			if($RTB_Members.Text.Replace(" ","") -match $global:regex_MemberList){
				$memberList = $RTB_Members.Text.Replace(" ","").Split("`n")
				$memberList = $memberList | Select-Object -Unique
				
				if($memberList.Length -gt 0){
					foreach($_ in $memberList){
						#Sanitize each member
						if(-not $_ -match $global:regex_4x4){
							popupError(($_ + " is not a valid 4x4."),0,"Invalid Imput",0x0)
							throw $Global:Error_BadMember
						}
						#check existance
						Get-ADUser $_
						if(!$?){
							popupError(($_ + " is not a valid 4x4"),0,"Invalid 4x4",0x0)
							throw $Global:Error_BadMember
						}
					}
					. updateDLOutput "Members Approved"	
				}
			} else {
				throw $Global:Error_InvalidMemberList
			}

	#-------Create DL-------#
			New-DistributionGroup -Alias $DL_Alias -Name $TB_DGN.text -Type Distribution -OrganizationalUnit OU="DLs,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu" -SamAccountName $DL_Alias -ManagedBy $sponsorList | Out-String -Stream | ForEach-Object {
				. updateConsole $_
			}
			if(!$?){
				. updateDLOutput "Error Creating DL, See Console"
				throw $Global:Error_Unknown
			} else {
				. updateDLOutput "DL Created."
			}
		
	#-------Modify Email Address Policy-------#
			Set-DistributionGroup $TB_DGN -EmailAddressPolicyEnabled $false | Out-String | ForEach-Object {
				. updateConsole $_
			}
			
	#-------Add Default Email Address-------#
			$emailAddresses = $(Get-DistributionGroup $TB_DGN).EmailAddresses
			$newAddress = "SMTP:" + $DL_Alias + "@sooner.net.ou.edu"
			$emailAddresses += $newAddress
			Set-DistributionGroup $TB_DGN -EmailAddresses $emailAddresses  | Out-String | ForEach-Object {
				. updateConsole $_
			}
			. updateDLOutput "Email Addresses:"
			foreach($_ in $emailAddresses){
				. updateDLOutput $_
			}
		
	#-------Add @ou.edu Email Address-------#
			if($TB_EmailAddress.Replace(" ","") -match $global:regex_Email){
			
				#Check the Email address is not in use by DLs or Mailboxes
				Get-DistributionGroup ($TB_EmailAddress.Text.Replace(" ","") + "@ou.edu")
				if($?){
					throw $Global:Error_EmailAddressInUse
				}
				Get-Mailbox ($TB_EmailAddress.Text.Replace(" ","") + "@ou.edu")
				if($?){
					throw $Global:Error_EmailAddressInUse
				}
				
				#Add email address
				$newAddress = "SMTP:" + $TB_EmailAddress.Text.Replace(" ","") + "@ou.edu"
				$emailAddresses += $newAddress
				Set-DistributionGroup $TB_DGN.Text -EmailAddresses $emailAddresses | Out-String | ForEach-Object {
					. updateConsole $_
				}
				if(!$?){
					. updateDLOutput "Error adding @ou.edu address, See Console"
				} else {
					. updateDLOutput "Added " + $newAddress
				}		
			} else {
				#this is handled in initial sanitization, should never reach here.
				throw $Global:Error_Unknown
			}
		
	#-------List in GAL?-------#
			if($ListInGAL.Checked){
				Set-DistributionGroup $TB_DGN -HiddenFromAddressListsEnabled $false
			} else {
				Set-DistributionGroup $TB_DGN -HiddenFromAddressListsEnabled $true
			}
		
	#-------Handle who can send to list-------#
			if($RB_A.Checked){
			#Anyone
				Set-DistributionGroup $TB_DGN -RequireSenderAuthenticationEnabled $false
			}		
			elseif($RB_MO.Checked){
			#Members Only
				Set-DistributionGroup $TB_DGN -AcceptMessagesOnlyFromDLMembers $TB_DGN
			}	
			elseif($RB_OO.Checked){
			#Owners Only
				#This creates another DL called $DL_Alias + "-group" that is used as the group that can send to DL
				New-DistributionGroup -Alias ($DL_Alias + "-group") -Name ($DL_Alias + "-group") -Type Security -OrganizationalUnit  OU="DLs,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu" -SamAccountName ($DL_Alias + "-group") 
			
				#Make universal group
				Set-ADGroup -Identity ($DL_Alias + "-group") -GroupScope Universal
			
				#Mail Enable group (May error out and silently continue from already being set)
				Enable-DistributionGroup ($DL_Alias + "-group") 2> $null
				
				#Wait for new DL to replicate
				. updateDLOutput "Please wait 60 seconds for Owners Group to replicate"
				. progressBar 60 "Waiting for replication..." "Please wait 60 seconds for Owners Group to replicate"
			
				Set-DistributionGroup ($DL_Alias + "-group") -HiddenFromAddressListsEnabled $true
				Set-DistributionGroup ($DL_Alias + "-group") -AcceptMessagesOnlyFromDLMembers ($DL_Alias + "-group")
				$sponsorList | foreach {
					Add-DistributionGroupMember -Identity ($DL_Alias + "-group") -Member $_
				}
				Set-DistributionGroup $TB_DGN -AcceptMessagesOnlyFromDLMembers ($DL_Alias + "-group")
			}	
			elseif($RB_IO.Checked){
			#Internal Only
				Set-DistributionGroup $TB_DGN -RequireSenderAuthenticationEnabled $true
			}
			else{
				#How'd you get here?
				throw $Global:Error_Unknown
			}
		
	#-------Handle list membership-------#
			if($CB_Sponsors.Checked){
				$memberList = $memberList + $sponsorList | Select-Object -Unique
				$memberList | foreach {
					Add-DistributionGroupMember -Identity $TB_DGN -Member $_
				}
				. updateDLOutput "Sponsors and Members added as members to DL"
			} 
			elseif($memberList.Length -gt 0){
				$memberList | foreach {
					Add-DistributionGroupMember -Identity $TB_DGN -Member $_
				}
				. updateDLOutput "Members added to DL"
			}
			else {
				. updateDLOutput "No Members added to DL"
			}
		
			. updateDLOutput "DL created without errors."
		
			#Clear Form
			$TB_Alias.Text = ""
			$TB_DGN.Text = ""
			$TB_EmailAddress.Text = ""
			$TB_Sponsors.Text = ""
			$UseCustomAlias.Checked = $false
			$RB_A.Checked = $true
			$RTB_Members.Text = ""
		} catch {
	#-------Input errors-------#		
			#Catch invalid dgn
			if($error[0].FullyQualifiedErrorId -match "InvalidDGN"){
				$msg = "No distribution group name or an invalid distribution group name was provided."
				popupError $msg 0 "Invalid Input" 0x0
			}
			if($error[0].FullyQualifiedErrorId -match "InvalidAlias"){
				$msg = "No Alias was provided when 'Use Custom Alias' was checked or a bad Alias was provided."
				popupError $msg 0 "Invalid Input" 0x0
			}
			#Catch invalid email address
			if($error[0].FullyQualifiedErrorId -match "InvalidEmailAddress"){	
				$msg = "No email address or an invalid email address was provided."
				popupError $msg 0 "Invalid Input" 0x0
			}
			#Catch input sanitization errors.
			if($error[0].FullyQualifiedErrorId -match "InputSanitizationError"){
				$msg = "Error in sanitizing input. Please validate data fields and try agian. If this error continues please ensure compliance with documented guidelines or contact the Exchange Team."
				popupError $msg 0 "Invalid Input" 0x0 
			}
			#Catch invalid sponsor list input
			if($error[0].FullyQualifiedErrorId -match "NoSponsorList"){
				$msg = "No sponsor list was provided. Every DL requires a minimum of 2 sponsors. If for some reason only one is available please create a case and assign it to the Exchange Team. Do NOT use yours or another 4x4 to comply."
				popupError($msg,0,"Invalid Imput",0x0)
			}
			#Catch invalid data in sponsor list
			if($error[0].FullyQualifiedErrorId -match "InvalidSponsorList"){
				$msg = "There is an invalid character in the sponsor list."
				popupError($msg,0,"Invalid Imput",0x0)
			}
			#Catch invalid data in member list
			if($error[0].FullyQualifiedErrorId -match "InvalidMemberList"){
				$msg = "There is an invalid character in the member list."
				popupError($msg,0,"Invalid Imput",0x0)
			}
			#Catch insufficient number of sponsors
			if($error[0].FullyQualifiedErrorId -match "InsufficientSponsors"){
				$msg = "Not enough sponsors were provided. Every DL requires a minimum of 2 sponsors. If for some reason only one is available please create a case and assign it to the Exchange Team. Do NOT use yours or another 4x4 to comply."
				popupError($msg,0,"Invalid Imput",0x0)
			}
			#Catch bad sponsor 4x4
			if($error[0].FullyQualifiedErrorId -match "BadSponsor"){
				#error handling before throw so we can include offending 4x4
			}
			#Catch bad member 4x4
			if($error[0].FullyQualifiedErrorId -match "BadMember"){
				#error handling before throw so we can include offending 4x4	
			}
			#Catch object already exists
			if($error[0].FullyQualifiedErrorId -match "AlreadyExists"){
				$msg = "The object you are trying to create already exists."
				popupError($msg,0,"Object already exists",0x0)
			}
			if($error[0].FullyQualifiedErrorId -match "EmailAddressInUse"){
				$msg = "The Email Address provided is already in use."
				popupError($msg,0,"Address in use",0x0)
			}
	#-------Script errors-------#
			#Catch internal coding errors with the sanitiazation function.
			if($error[0].FullyQualifiedErrorId -match "MalformedSanitizationRequest"){
				$msg = "An internal error has occured attempting to sanitize user input. Please contact the Exchange Team with error code: DLSE1"
				popupError($msg, 0, "Invalid Sanitization Request", 0x0)
			}
			#Something bad happened and I don't know what
			if($error[0].FullyQualifiedErrorId -match "UnknownScriptError"){
				$msg = "Something bad happened and I don't know what. Check console for errors. DL may have been created, verify existance before trying agian. If you cannot resolve any issues experienced or created contact Exchange Team."
				popupError($msg,0,"Object already exists",0x0)
			}
		}
	}
	#endregion Create DL

	#region Increase Quota
	$B_IQ_Check_OnClick= 
	{
		Write-Debug "IQ Check OnClick"
		Write-Debug $TB_IQName.Text
		if(-not $TB_IQName.Text.Replace(" ","") -match $global:regex_4x4){
			Write-Debug "Bad Input"
			$label_IQ_Out.Text = "Please input valid 4x4."
			$SCEGUI.Update()
		} else {	
			try{
				$IQ_MB_REF = $TB_IQName.Text.Replace(" ","")
				$IQ_MB_OBJ = get-mailbox $IQ_MB_REF
				if(!$?){
					throw $Global:Error_InvalidRef
				}
				$IQ_MB_USED = Get-MailboxStatistics $IQ_MB_REF
				$label_IQ_Out.Text = $IQ_MB_OBJ.SamAccountName + "'s current quota is " + $IQ_MB_OBJ.ProhibitSendQuota.Value.ToGB() + "GB, Used: " + $IQ_MB_USED.TotalItemSize.Value.ToGB() + " GB."
			}catch{
				#Catch invalid 4x4
				if($error[0].FullyQualifiedErrorId -match "InvalidRef"){
					$msg = "The 4x4 or Email provided is not valid for this request."
					popupError $msg 0 "Invalid Input" 0x0
				}
			}
		}
	}

	$B_IQ_SetDefault_OnClick= 
	{
		if(-not $TB_IQName.Text.Replace(" ","") -match $global:regex_4x4){
			Write-Debug "Bad Input"
			$label_IQ_Out.Text = "Please input valid 4x4."
			$SCEGUI.Update()
		}else{
			try{
				$IQ_MB_REF = $TB_IQName.Text.Replace(" ","")
				$IQ_MB_OBJ = get-mailbox $IQ_MB_REF
				if(!$?){
					throw $Global:Error_InvalidRef
				}
				Set-Mailbox -IssueWarningQuota '5.841 GB (6,271,533,056 bytes)' -ProhibitSendQuota '6 GB (6,442,450,944 bytes)' -Identity $IQ_MB_REF
					#This is safe todo because the Regex check ensures that the input is not a wildcard
				$IQ_MB_OBJ = Get-Mailbox $IQ_MB_REF
				$label_IQ_Out.Text = $IQ_MB_OBJ.SamAccountName + "'s current quota is " + $IQ_MB_OBJ.ProhibitSendQuota
			}catch{
				#Catch invalid 4x4
				if($error[0].FullyQualifiedErrorId -match "InvalidRef"){
					$msg = "The 4x4 or Email provided is not valid for this request."
					popupError $msg 0 "Invalid Input" 0x0
				}
			}
		}
		
		#clear form
		$TB_IQName.Text = ""	
	}

	$B_IQ_SetMax_OnClick= 
	{
		if(-not $TB_IQName.Text.Replace(" ","") -match $global:regex_4x4){
			$label_IQ_Out.Text = "Please input valid 4x4."
			$SCEGUI.Update()
		}else{
			try{
				$IQ_MB_REF = $TB_IQName.Text.Replace(" ","")
				$IQ_MB_OBJ = get-mailbox $IQ_MB_REF
				if(!$?){
					throw $Global:Error_InvalidRef
				}
				Set-Mailbox -IssueWarningQuota 11GB -ProhibitSendQuota 12GB -Identity $IQ_MB_REF
					#This is safe todo because the Regex check ensures that the input is not a wildcard
				$IQ_MB_OBJ = Get-Mailbox $IQ_MB_REF
				$label_IQ_Out.Text = $IQ_MB_OBJ.SamAccountName + "'s current quota is " + $IQ_MB_OBJ.ProhibitSendQuota
			}
			catch{
				#Catch invalid 4x4
				if($error[0].FullyQualifiedErrorId -match "InvalidRef"){
					$msg = "The 4x4 or Email provided is not valid for this request."
					popupError $msg 0 "Invalid Input" 0x0
				}
			}
		}
		
		#clear form
		$TB_IQName.Text = ""	
	}
	#endregion Increase Quota

	#region CMS Group
	$handler_CMS_Create_Click= 
	{
		#CMS groups are basically DLs but in a different OU with a little less config.
		
		#clear form
		$CMS_Status.text = ""
		
		#setup progress bar
		$CMS_Progress.Value = 0
		$CMS_Progress.Step = 100/12
		
		try{
			. cmsUpdateProgress "Validating Input Data"
			#pull and sanitize data from form
			if($CMS_Group.Text.Replace(" ","") -match $global:regex_Alias){
				$alias = $CMS_Group.Text.Replace(" ","")
			} else {
				throw $Global:Error_InvalidDGN
			}

			#Sponsors - Handle Sponsor Check
			if($CMS_Owners.Text.Replace(" ","") -match $global:regex_SponsorList){
				$ownerList = $CMS_Owners.Text.Replace(" ","").TrimEnd(",").Split(",")
				$ownerList = $ownerList | Select-Object -Unique
				
				if($ownerList.Length -lt 1){
					throw $Global:Error_InsufficientOwner
				} else {
					foreach($_ in $ownerList){
						#sanitize each owner
						if(-not $_ -match $global:regex_4x4){ 
							popupError(($_ + " is not a valid 4x4."),0,"Invalid 4x4",0x0)
							throw $Global:Error_BadOwner
						}
						Get-ADUser $_
						if(!$?){
							popupError(($_ + " is not a valid 4x4."),0,"Invalid 4x4",0x0)
							throw $Global:Error_BadOwner
						}
					}
				}
			} else {
				$ownerList = "BadData"
				popupError(("There are invalid characters in the Owners' field."),0,"Invalid Input",0x0)
				throw $Global:Error_BadOwner
			}
			
			#Members - Handle Member Check
			if($CMS_Members.Text.Replace(" ","") -match $global:regex_MemberList){
				$memberList = $CMS_Members.Text.Replace(" "."").TrimEnd(",").Split(",")
				$memberList = $memberList | Select-Object -Unique
				
				foreach($_ in $memberList){
					if(-not $_ -match $global:regex_4x4){
						popupError(($_ + " is not a valid 4x4"),0,"Invalid Imput",0x0)
						throw $Global:Error_BadMember
					}
					Get-ADUser $_
					if(!$?){
						popupError(($_ + " is not a valid 4x4"),0,"Invalid Imput",0x0)
						throw $Global:Error_BadMember
					}
				}
			} else {
				popupError(("There are invalid characters in the Members' field"),0,"Invalid Imput",0x0)
				throw $Global:Error_BadMember
			}
			
			#start progress bar
			. cmsUpdateProgress "Data Validated, Creating Group"

			#create dl
			new-DistributionGroup -alias $alias -name $CMS_Group.text -type security -org "OU=CMS,OU=DLs,OU=Exchange,dc=sooner,dc=net,dc=ou,dc=edu" -SamAccountName $alias

			#update progress
			. cmsUpdateProgress "Setting Email Policy"
		
			#config dl
			set-distributiongroup $alias -BypassSecurityGroupManagerCheck -EmailAddressPolicyEnabled $false

			#update progress
			. cmsUpdateProgress "Setting Owners"

			#set managers
			$DLMangers = (Get-DistributionGroup $alias).ManagedBy
			foreach($_ in $ownerList){
				$DLMangers += $_
			}
			Set-DistributionGroup $alias -ManagedBy $DLMangers
			
			#wait for replication
			. cmsUpdateProgress "Waiting for Replication"
		
			$i = 0
			do{
				sleep 1
				$i++
				. cmsUpdateProgress	"Waiting for Replication"
			} while ($i -lt 4)
		
			#update progress
			. cmsUpdateProgress "Setting Address"		
		
			#set email address
			set-distributiongroup $alias -PrimarySMTPAddress $alias@ou.edu
		
			#add members, not required
			if($memberList.Length -gt 0){
				. cmsUpdateProgress "Adding Members"
				
				foreach($_ in $memberList){
					Add-DistributionGroupMember -Identity $alias -Member $_
				}
				
				$CMS_Status.text = "Members Added"
				$SCEGUI.Update()
			}

			#update progress
			. cmsUpdateProgress "Verifying Data"	
		
			#gather updated data
			$dlname = (Get-DistributionGroup $alias@ou.edu).name
			$dladdress = (Get-DistributionGroup $alias@ou.edu).PrimarySmtpAddress
		
			#check values and update final status
			if ($dlname -ne "" -and $dladdress -ne "") {
				#everything worked
				$CMS_Progress.Value = 100
				$CMS_Status.text = "Done: " + $dlname + " / " + $dladdress
				sleep 2
			} else {
				$CMS_Status.text = "Something didn't work right. See console"
			}
		}
		catch{
			if($error[0].FullyQualifiedErrorId -match "InvalidDGN"){
				$msg = "The provided name does not comply with naming requirements."
				popupError($msg,0,"Invalid Imput",0x0)
			}
			if($error[0].FullyQualifiedErrorId -match "BadOwner"){
				#error handling before throw so we can include offending 4x4
			}
			if($error[0].FullyQualifiedErrorId -match "BadMember"){
				#error handling before throw se we can include the offending 4x4
			}
			if($error[0].FullyQualifiedErrorId -match "InsufficientOwners"){
				$msg = "Not enough owners were provided. Every CMS Group requires atleast 1 owner. If for some reason none are available please create a case and assign it to the Exchange Team. Do NOT use yours or another 4x4 to comply."
				popupError($msg,0,"Invalid Imput",0x0)
			}
		}	
		
		#clear form
		$CMS_Group.text = ""
		$CMS_Owners.text = ""
		$CMS_Members.text = ""
		$CMS_Progress.Value = 0
	}
	#endregion CMS Group

	#region Roomview

	$RV_button1_OnClick= 
	{
	Write-Debug "Button Press"

		try{
			Write-Debug "Starting"
			. updateRVStatus "Starting"
			
	#-------Variables-------#
	
			Write-Debug "1"
			$password = ConvertTo-SecureString ("sooner" + (get-date).toString("yyMMdd")) -AsPlainText -Force
			$name = $RV_textBox1.text 
			$alias = $RV_textBox2.text
			
			if(!($name -match $global:regex_Alias)) {throw $Global:Error_InvalidName}
			if(!($alias -match $global:regex_Alias)) {throw $Global:Error_InvalidAlias}
			
			$groupName = $alias + "-group"
			. updateRVStatus "Input approved."
			sleep 1
			
			Write-Debug "Creating"
			Write-Debug $name
			Write-Debug $alias
			Write-Debug $groupName
			
	#-------Create and Configure-------#

			. updateRVStatus "Creating mailbox"
			Write-Debug "New-Mailbox"
			New-Mailbox -Name $name -DisplayName $name -Alias $alias -UserPrincipalName ($alias + "@ou.edu") -Password $password
			. updateRVStatus "Configuring mailbox"
			Write-Debug "Set-Mailbox"
			Set-Mailbox $alias -EmailAddresses @{add=($alias + "@ou.edu")} 
			Set-Mailbox $alias -Type Room -EmailAddressPolicyEnabled $false -PrimarySmtpAddress ($alias + "@ou.edu")
			. updateRVStatus "Creating access group"
			Write-Debug "New-DistributionGroup"
			New-DistributionGroup -DisplayName $groupName -Name $groupName -OrganizationalUnit "ou=Groups,ou=resource mailboxes,ou=exchange,dc=sooner,dc=net,dc=ou,dc=edu"
			
			. updateRVStatus "Waiting for replication."
			sleep 30
			
			Write-Debug "Set-Group"
			Set-Group $groupName -Universal -ErrorAction SilentlyContinue
			
			Write-Debug "Add-ADGroupMember"
			. updateRVStatus "Add-AdGroupMember"
			Add-DistributionGroupMember $groupName -Member "it_roomview"
			
			Write-Debug "Add-MailboxPermissions"
			sleep 10
			. updateRVStatus "Add-MailboxPermissions"
			Add-MailboxPermission -Identity $name -User $groupName -AccessRights FullAccess
			
			. updateRVStatus "Moving Mailbox" 
			Write-Debug "Move-ADObject"
			get-aduser $alias | Move-ADObject -TargetPath "ou=roomview,ou=resource mailboxes,ou=exchange,dc=sooner,dc=net,dc=ou,dc=edu"
			. updateRVStatus $($name + " created and configured; validating...")
			
	#-------Validate Results-------#

			Write-Debug "Validate Results"
			$mailbox =  Get-Mailbox $name
			$group = Get-Group $groupName
			
			if(!$mailbox){throw $Global:Error_MailboxNotFound}
			if(!($name -eq $mailbox.Name -and $name -eq $mailbox.DisplayName)){throw $Global:Error_NameValidation}
			if(!($alias -eq $mailbox.Alias)){throw $Global:Error_AliasValidation}
			if(!((Get-DistributionGroupMember $groupName).Name -contains 'it_roomview')){throw $Global:Error_MembershipValidation}
			
			. updateRVStatus "Results Validated"
			sleep 1
			
	#-------Send Record to Service Now-------#

			. updateRVStatus "Sending Record"
			$body = $global:CurrentUser + " has created a room mailbox for roomview using the following configuration.<br>"
			$body += "----------------------<br>"
			$body += "Name: " + $name + "<br>"
			$body += "Alias: " + $alias + "<br>"
			$body += "Access group: " + $groupName + "<br>"
			
			#sendMail -address -subject "Roomview Room Creation Script" -body $body -from "roomviewscript@ou.edu" -cc "aheyne@ou.edu"
			
	#-------Cleanup-------#
			
			. updateRVStatus "Done"
			$RV_textBox1.Text = ""
			$RV_textBox2.Text = ""
		} catch {
			#--Catch Custom Errors--#
			
			if($error[0].FullyQualifiedErrorId -match "MailboxNotFound"){
				. updateRVStatus  "Mailbox creation error"
				$msg = "The mailbox failed to create."
				popupError($msg, 0, "Mailbox creation Error", 0x0)
			}
			if($error[0].FullyQualifiedErrorId -match "NameValidation"){
				. updateRVStatus  "Error Validating Room Name"
				$msg = "Error Validating Room Name. Resultant name did not match imput."
				popupError($msg, 0, "Name Validation Error", 0x0)
			}
			if($error[0].FullyQualifiedErrorId -match "AliasValidation"){
				. updateRVStatus  "Error Validating Alias"
				$msg = "Error Validating Alias. Resultant alias did not match imput."
				popupError($msg, 0, "Alias Validation Error", 0x0)
			}
			if($error[0].FullyQualifiedErrorId -match "MembershipValidation"){
				. updateRVStatus  "Error Validating Membership"
				$msg = "Error Validating Membership. it_roomview is not a part of the accessgroup."
				popupError($msg, 0, "Membership Validation Error", 0x0)
			}
			if($error[0].FullyQualifiedErrorId -match "InvalidName"){
				. updateRVStatus  "Invalid Input in Name"
				$msg = "Error Validating Name. Name does not conform to naming conventions."
				popupError $msg 0 "Invalid Input" 0x0
			}
			if($error[0].FullyQualifiedErrorId -match "InvalidAlias"){
				. updateRVStatus  "Invalid Input in Alias"
				$msg = "Error Validating Alias. Alias does not conform to naming conventions."
				popupError $msg 0 "Invalid Input" 0x0
			}
		}
	}

	#endregion Roomview

	$OnLoadForm_StateCorrection = {
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$SCEGUI.WindowState = $InitialFormWindowState
	}

	#region Form Code

	#region Main Window

	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 425
	$System_Drawing_Size.Width = 660
	$SCEGUI.ClientSize = $System_Drawing_Size
	$SCEGUI.MaximumSize = $System_Drawing_Size
	$SCEGUI.MinimumSize = $System_Drawing_Size
	$SCEGUI.MaximizeBox = $false
	$SCEGUI.DataBindings.DefaultDataSourceUpdateMode = 0
	#$SCEGUI.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('batman.ico')
	$SCEGUI.Name = "SCGUI"
	$SCEGUI.Text = "SCGUI"
	$SCEGUI.StartPosition = 1

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

	$SCEGUI.Controls.Add($tabControl1)

	#endregion Main Window

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
	$CB_Sponsors.TabIndex = 13
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
	$RTB_Members.TabIndex = 12
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
	$Create.TabIndex = 14
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

	$B_IQ_SetMax.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 95
	$B_IQ_SetMax.Location = $System_Drawing_Point
	$B_IQ_SetMax.Name = "B_SetCustom"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 100
	$B_IQ_SetMax.Size = $System_Drawing_Size
	$B_IQ_SetMax.TabIndex = 5
	$B_IQ_SetMax.Text = "Set Max 12GB"
	$B_IQ_SetMax.UseVisualStyleBackColor = $True
	$B_IQ_SetMax.add_Click($B_IQ_SetMax_OnClick)

	$tabPageIQ.Controls.Add($B_IQ_SetMax)

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
	$label_IQ_Out.TabIndex = 14

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
	$label_IQ_7.TabIndex = 13
	$label_IQ_7.Text = "Input 4x4"

	$tabPageIQ.Controls.Add($label_IQ_7)

	#endregion Increase Quota

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

	#region Roomview

	$RV_Tab.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 22
	$RV_Tab.Location = $System_Drawing_Point
	$RV_Tab.Name = "RV_Tab"
	$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
	$System_Windows_Forms_Padding.All = 3
	$System_Windows_Forms_Padding.Bottom = 3
	$System_Windows_Forms_Padding.Left = 3
	$System_Windows_Forms_Padding.Right = 3
	$System_Windows_Forms_Padding.Top = 3
	$RV_Tab.Padding = $System_Windows_Forms_Padding
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 364
	$System_Drawing_Size.Width = 637
	$RV_Tab.Size = $System_Drawing_Size
	$RV_Tab.TabIndex = 0
	$RV_Tab.Text = "Roomview"
	$RV_Tab.UseVisualStyleBackColor = $True

	$tabControl1.Controls.Add($RV_Tab)

	$RV_Status.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 9
	$System_Drawing_Point.Y = 123
	$RV_Status.Location = $System_Drawing_Point
	$RV_Status.Name = "RV_Status"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 40
	$System_Drawing_Size.Width = 200
	$RV_Status.Size = $System_Drawing_Size
	$RV_Status.TabIndex = 21

	$RV_Tab.Controls.Add($RV_Status)


	$RV_button1.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 134
	$System_Drawing_Point.Y = 93
	$RV_button1.Location = $System_Drawing_Point
	$RV_button1.Name = "RV_button1"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 75
	$RV_button1.Size = $System_Drawing_Size
	$RV_button1.TabIndex = 2
	$RV_button1.Text = "Create"
	$RV_button1.UseVisualStyleBackColor = $True
	$RV_button1.add_Click($RV_button1_OnClick)

	$RV_Tab.Controls.Add($RV_button1)

	$RV_textBox2.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 8
	$System_Drawing_Point.Y = 67
	$RV_textBox2.Location = $System_Drawing_Point
	$RV_textBox2.Name = "RV_textBox2"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 200
	$RV_textBox2.Size = $System_Drawing_Size
	$RV_textBox2.TabIndex = 1

	$RV_Tab.Controls.Add($RV_textBox2)

	$RV_label2.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 9
	$System_Drawing_Point.Y = 48
	$RV_label2.Location = $System_Drawing_Point
	$RV_label2.Name = "RV_label2"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 15
	$System_Drawing_Size.Width = 200
	$RV_label2.Size = $System_Drawing_Size
	$RV_label2.TabIndex = 20
	$RV_label2.Text = "Alias"
	#$RV_label2.add_Click($handler_label2_Click)

	$RV_Tab.Controls.Add($RV_label2)

	$RV_label1.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 8
	$System_Drawing_Point.Y = 3
	$RV_label1.Location = $System_Drawing_Point
	$RV_label1.Name = "RV_label1"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 15
	$System_Drawing_Size.Width = 200
	$RV_label1.Size = $System_Drawing_Size
	$RV_label1.TabIndex = 10
	$RV_label1.Text = "Display Name"
	#$RV_label1.add_Click($handler_label1_Click)

	$RV_Tab.Controls.Add($RV_label1)

	$RV_textBox1.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 8
	$System_Drawing_Point.Y = 21
	$RV_textBox1.Location = $System_Drawing_Point
	$RV_textBox1.Name = "RV_textBox1"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 200
	$RV_textBox1.Size = $System_Drawing_Size
	$RV_textBox1.TabIndex = 0

	$RV_Tab.Controls.Add($RV_textBox1)


	#endregion Roomview

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

	#endregion Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $SCEGUI.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$SCEGUI.add_Load($OnLoadForm_StateCorrection)
	#Show the Form
	$SCEGUI.ShowDialog() | Out-Null

} #End Generate Form

function GenerateSetupForm {

	#region Import the Assemblies
	[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
	[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
	#endregion

	#region Generated Form Objects
	$SetupStatus = New-Object System.Windows.Forms.Form
	$SetupStatus_label = New-Object System.Windows.Forms.Label
	$SetupStatus_PB = New-Object System.Windows.Forms.ProgressBar
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects

	$OnLoadForm_StateCorrection=
	{#Correct the initial state of the form to prevent the .Net maximized form issue
		$SetupStatus.WindowState = $InitialFormWindowState
		$SetupStatus.Close()
	}
	
	$SetupStatus_FormClosing = [System.Windows.Forms.FormClosingEventHandler]{
		SetupErrors
	}

	#----------------------------------------------
	#region Generated Form Code
	$SetupStatus.AutoSize = $True
	$SetupStatus.AutoSizeMode = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 73
	$System_Drawing_Size.Width = 284
	$SetupStatus.ClientSize = $System_Drawing_Size
	$SetupStatus.ControlBox = $False
	$SetupStatus.DataBindings.DefaultDataSourceUpdateMode = 0
	$SetupStatus.Name = "SetupStatus"
	$SetupStatus.SizeGripStyle = 2
	$SetupStatus.Text = "SCEGUI"
	$SetupStatus.TopMost = $True
	$SetupStatus.add_FormClosing($SetupStatus_FormClosing)

	$SetupStatus_label.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 2
	$System_Drawing_Point.Y = 46
	$SetupStatus_label.Location = $System_Drawing_Point
	$SetupStatus_label.Name = "label_SetupStatus"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 280
	$SetupStatus_label.Size = $System_Drawing_Size
	$SetupStatus_label.TabIndex = 1
	$SetupStatus_label.Text = "Setting Up..."
	$SetupStatus_label.TextAlign = 32

	$SetupStatus.Controls.Add($SetupStatus_label)

	$SetupStatus_PB.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 2
	$System_Drawing_Point.Y = 2
	$SetupStatus_PB.Location = $System_Drawing_Point
	$SetupStatus_PB.Name = "SetupStatus_PB"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 37
	$System_Drawing_Size.Width = 280
	$SetupStatus_PB.Size = $System_Drawing_Size
	$SetupStatus_PB.TabIndex = 0
	$SetupStatus_PB.Value = 0
	$SetupStatus_PB.Step = 100/23
	#$SetupStatus_PB.add_Click($handler_progressBar1_Click)

	$SetupStatus.Controls.Add($SetupStatus_PB)

	#endregion Generated Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $SetupStatus.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$SetupStatus.add_Load($OnLoadForm_StateCorrection)
	#Show the Form
	$SetupStatus.ShowDialog() | Out-Null

}

#region Local Functions

function updateRVStatus([string]$text) {
	$RV_Status.Text = $text
	$form1.Update()
}

function cmsUpdateProgress {
	param(
		[Parameter(Mandatory = $false)]$info = "..."
	)
	$CMS_Status.text = $info
	$CMS_Progress.PerformStep()
	$SCEGUI.Update()
}

function setupStatusProgress() {
	$phrases = @("Reticulating Splines","Performing background check...","Encapsulating Retro Encabulator...", "Identifing primary targets...","Looking for intelegent life...","Biding time...","Predicting Questions...","Searching for the meaning of life...")
	$phrase = $phrases | Get-Random
	
	$SetupStatus_label.Text = $phrase
	$SetupStatus_PB.PerformStep()
	$SetupStatus.Update()
	sleep 1
	
}

function progressBar {
	param(
        [Parameter(Mandatory = $true, Position = 0)][Int32]$time,
        [Parameter(Mandatory = $false, Position = 1)][String]$title = "Waiting...",
        [Parameter(Mandatory = $false, Position = 2)][string]$text = "..."
    )
	
	#this line may not be necessary
	#[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
	Add-Type -assembly System.Windows.Forms
	
	$height=100
	$width=400
	$color = "White"

	#create the form
	$progressBarForm = New-Object System.Windows.Forms.Form
	$progressBarForm.Text = $title
	$progressBarForm.Height = $height
	$progressBarForm.Width = $width
	$progressBarForm.BackColor = $color
	$progressBarForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
	#display center screen
	$progressBarForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
	# create label
	$textProgress = New-Object system.Windows.Forms.Label
	$textProgress.Text = $text
	$textProgress.Left=5
	$textProgress.Top= 10
	$textProgress.Width= $width - 20
	#adjusted height to accommodate progress bar
	$textProgress.Height=15
	$textProgress.Font= "Verdana"
	#optional to show border 
	#$textProgress.BorderStyle=1

	#add the label to the form
	$progressBarForm.controls.add($textProgress)
	$progressBar = New-Object System.Windows.Forms.ProgressBar
	$progressBar.Name = 'progressBar1'
	$progressBar.Value = 0
	$progressBar.Style="Continuous"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Width = $width - 40
	$System_Drawing_Size.Height = 20
	$progressBar.Size = $System_Drawing_Size
	$progressBar.Left = 5
	$progressBar.Top = 40
	$progressBarForm.Controls.Add($progressBar)
	$progressBarForm.Show()| out-null

	#give the form focus
	$progressBarForm.Focus() | out-null

	#update the form
	$progressBarForm.Refresh()

	#Count and update
	for($i=0; $i -le $time; $i++){
		[int]$pct = ($i/$time)*100
		$progressBar.Value = $pct
		$progressBarForm.Refresh()
		Start-Sleep -Seconds 1
	}

	$progressBarForm.Close()
}

function popupError($msg, $timeout, $title, $type) {
	$popWindow = New-Object -ComObject Wscript.Shell
	return $popWindow.Popup($msg, $timeout, $title, $type)
}

function safeGetMailbox($data){
	if($data.Length -lt 5 -or $data -eq $null){
		Write-Debug "safeGetMailbox:BadData: " + $data
		$data = "BadData"
	}
	return get-mailbox $data
}

function updateConsole {
	$RTB_Console.Select($RTB_DLOutput.Text.Length, 0)
	$RTB_Console.ScrollToCaret()
	$SCEGUI.Update()
}

function updateDLOutput([string]$data) {
	#CreateDL is the only function that uses this because information in the DL tab is output to a rich text box which requres some work to update.
	$RTB_DLOutput.AppendText("`n")
	$RTB_DLOutput.AppendText($data)
	$RTB_DLOutput.Select($RTB_DLOutput.Text.Length, 0)
	$RTB_DLOutput.ScrollToCaret()
	$SCEGUI.Update()
}

function sendMail {
	param(
		[Parameter(Mandatory = $true)]$address,
		[Parameter(Mandatory = $true)][System.String]$subject,
		[Parameter(Mandatory = $true)][System.String]$body,
		[Parameter(Mandatory = $false)]$from = "postmaster@ou.edu",
		[Parameter(Mandatory = $false)]$host = "asmtp.ou.edu",
		[Parameter(Mandatory = $false)]$cc
	)
	
	$SmtpClient = new-object system.net.mail.smtpClient 
	$SmtpClient.Host = $host

	$mailMessage = New-Object system.net.mail.mailmessage 
	$mailMessage.from = $from
	$mailMessage.To.add($address)
	if($PSBoundParameters.ContainsKey('cc')) {
		$mailMessage.cc.add($cc)
	}
	$mailMessage.Subject = $subject
	$mailMessage.IsBodyHtml = $true
	$mailMessage.Body = $body
	$mailMessage.Priority = "Normal"
	Write-Debug "Sending email to" $address
	do {
		$smtpclient.Send($mailMessage)
	} until($?) #this resends the message if it fails to send by reading the status of the last command
	Write-Debug "Sent"
}

function New-ErrorRecord {
	[cmdletbinding()]
	#Creates an custom ErrorRecord that can be used to report a terminating or non-terminating error. 
    param(
        [Parameter(Mandatory = $true, Position = 0)][System.String]$Exception,
        [Parameter(Mandatory = $true, Position = 1)][Alias('ID')][System.String]$ErrorId,
        [Parameter(Mandatory = $true, Position = 2)]
        [Alias('Category')][System.Management.Automation.ErrorCategory][ValidateSet('NotSpecified', 'OpenError', 'CloseError', 'DeviceError',
            'DeadlockDetected', 'InvalidArgument', 'InvalidData', 'InvalidOperation',
                'InvalidResult', 'InvalidType', 'MetadataError', 'NotImplemented',
                    'NotInstalled', 'ObjectNotFound', 'OperationStopped', 'OperationTimeout',
                        'SyntaxError', 'ParserError', 'PermissionDenied', 'ResourceBusy',
                            'ResourceExists', 'ResourceUnavailable', 'ReadError', 'WriteError',
                                'FromStdErr', 'SecurityError')]$ErrorCategory,
        [Parameter(Mandatory = $true, Position = 3)][System.Object]$TargetObject,
        [Parameter()][System.String]$Message,
        [Parameter()][System.Exception]$InnerException
    )
    begin {
		Write-Debug "NER Begin"
        $exceptions = Get-AvailableExceptionsList 
        $exceptionsList = $exceptions -join "`r`n"
    }
    process {
		Write-Debug "NER Process"
        # trap for any of the "exceptional" Exception objects that made through the filter
        trap [Microsoft.PowerShell.Commands.NewObjectCommand] {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        # verify input exception is "available". if so...
        if ($exceptions -match "^(System\.)?$Exception$") {
            # ...build and save the new Exception depending on present arguments, if it...
            $_exception = if ($Message -and $InnerException) {
                # ...includes a custom message and an inner exception
                New-Object $Exception $Message, $InnerException
            } elseif ($Message) {
                # ...includes a custom message only
                New-Object $Exception $Message
            } else {
                # ...is just the exception full name
                New-Object $Exception
            }
            # now build and output the new ErrorRecord
            New-Object Management.Automation.ErrorRecord $_exception, $ErrorID, $ErrorCategory, $TargetObject
        } else {
            # Exception argument is not "available";
            # warn the user, provide a list of "available" exceptions and...
            Write-Warning "Available exceptions are:`r`n$exceptionsList" 
            $message2 = "Exception '$Exception' is not available."
            $exception2 = New-Object System.InvalidOperationException $message2
            $errorID2 = 'BadException'
            $errorCategory2 = 'InvalidOperation'
            $targetObject2 = 'Get-AvailableExceptionsList'
            $errorRecord2 = New-Object Management.Automation.ErrorRecord $exception2, $errorID2, $errorCategory2, $targetObject2
            # ...report a terminating error to the user
            $PSCmdlet.ThrowTerminatingError($errorRecord2)
        }
    }
}

function Get-AvailableExceptionsList {
	#Retrieves all available Exceptions to construct ErrorRecord objects.
    [CmdletBinding()]
    param()
    end {
		Write-Debug "NER end"
        $irregulars = 'Dispose|OperationAborted|Unhandled|ThreadAbort|ThreadStart|TypeInitialization'
        [AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object {
            $_.GetExportedTypes() -match 'Exception' -notmatch $irregulars |
            Where-Object {
                $_.GetConstructors() -and $(
                $_exception = New-Object $_.FullName
                New-Object Management.Automation.ErrorRecord $_exception, ErrorID, OpenError, Target
                )
            } | Select-Object -ExpandProperty FullName
        } 2> $null #pipes errors to null since our current erroraction is stop in order to catch our nonterminating custom errors.
    }
}

function setupErrors {
	#$errorList = ("InvalidDGN,InvalidAlias,InvalidEmailAddress,InvalidSponsorList,InvalidMemberList,InputSanitizationError,NoSponsorList,InsufficientSponsors,BadSponsor,BadMember,AlreadyExists,EmailAddressInUse,InvalidRef,BadOwner,InsufficientOwner,MalformedSanitizationRequest,Unknown").Split(",")
	#$errorList | foreach {New-ErrorRecord System.Exception $_ InvalidOperation $_}
	
	$Global:Error_InvalidDGN = New-ErrorRecord System.Exception InvalidDGN InvalidOperation "InvalidDGN"
	setupStatusProgress
	$Global:Error_InvalidAlias = New-ErrorRecord System.Exception InvalidAlias InvalidOperation "InvalidAlias"
	setupStatusProgress
	$Global:Error_InvalidEmailAddress = New-ErrorRecord System.Exception InvalidEmailAddress InvalidOperation "InvalidEmailAddress"
	setupStatusProgress
	$Global:Error_InvalidSponsorList = New-ErrorRecord System.Exception InvalidSponsorList InvalidOperation "InvalidSponsorList"
	setupStatusProgress
	$Global:Error_InvalidMemberList = New-ErrorRecord System.Exception InvalidMemberList InvalidOperation "InvalidMemberList"
	setupStatusProgress
	$Global:Error_InputSanitizationError = New-ErrorRecord System.Exception InputSanitizationError InvalidOperation "InputSanitizationError"
	setupStatusProgress
	$Global:Error_NoSponsorList = New-ErrorRecord System.Exception NoSponsorList InvalidOperation "NoSponsorList" 
	setupStatusProgress
	$Global:Error_InsufficientSponsors = New-ErrorRecord System.Exception InsufficientSponsors InvalidOperation "InsufficientSponsors" 
	setupStatusProgress
	$Global:Error_BadSponsor = New-ErrorRecord System.Exception BadSponsor InvalidOperation "BadSponsor"
	setupStatusProgress
	$Global:Error_BadMember = New-ErrorRecord System.Exception BadMember InvalidOperation "BadMember" 
	setupStatusProgress
	$Global:Error_AlreadyExists = New-ErrorRecord System.Exception AlreadyExists InvalidOperation "AlreadyExists"
	setupStatusProgress
	$Global:Error_EmailAddressInUse = New-ErrorRecord System.Exception EmailAddressInUse InvalidOperation "EmailAddressInUse"
	setupStatusProgress
	$Global:Error_InvalidRef = New-ErrorRecord System.Exception InvalidRef InvalidOperation "InvalidRef"
	setupStatusProgress
	$Global:Error_BadOwner = New-ErrorRecord System.Exception BadOwner InvalidOperation "BadOwner"
	setupStatusProgress
	$Global:Error_InsufficientOwner = New-ErrorRecord System.Exception InsufficientOwners InvalidOperation "InsufficientOwners"
	setupStatusProgress
	$Global:Error_NameValidation = New-ErrorRecord System.Exception NameValidation InvalidOperation "NameValidation"
	setupStatusProgress
	$Global:Error_AliasValidation = New-ErrorRecord System.Exception AliasValidation InvalidOperation "AliasValidation"
	setupStatusProgress
	$Global:Error_MailboxNotFound = New-ErrorRecord System.Exception MailboxNotFound InvalidOperation "MailboxNotFound"
	setupStatusProgress
	$Global:Error_MembershipValidation = New-ErrorRecord System.Exception MembershipValidation InvalidOperation "MembershipValidation"
	setupStatusProgress
	$Global:Error_InvalidName = New-ErrorRecord System.Exception InvalidName InvalidOperation "InvalidName"
	setupStatusProgress
	$Global:Error_InvalidAlias = New-ErrorRecord System.Exception InvalidAlias InvalidOperation "InvalidAlias"
	setupStatusProgress
	#Script Errors
	$Global:Error_MalformedSanitizationRequest = New-ErrorRecord System.Exception MalformedSanitizationRequest InvalidOperation "MalformedSanitizationRequest"
	setupStatusProgress
	$Global:Error_Unknown = New-ErrorRecord System.Exception UnknownScriptError InvalidOperation "UnknownScriptError"
	setupStatusProgress
	
}

function setupConnection {
	#Grab a random Exchange Server
	$ExchangeServer = ("aero","astoria","aurora") | Get-Random

	try{
		#Setting error action preference to stop in order for non-terminating errors to trigger try/catch blocks. Used in input sanitization and internal error handling. 
		$ErrorActionPreference = "Stop"
		
		#Get user credentials
		$cred = Get-Credential
		
		#Get user name
		if($cred.UserName -match "sooner"){
			$global:CurrentUser = $cred.UserName.Split("\")[1]
		} else {
			$global:CurrentUser = $cred.UserName
		}

		#Create and open the powershell session
		$global:PSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://it-" + $ExchangeServer + ".sooner.net.ou.edu/PowerShell") -Authentication Default -Credential $cred
		Import-PSSession $global:PSSession -AllowClobber
		
	}
	catch{
		#Catch bad login
		if($error[0].FullyQualifiedErrorId -match "PSSessionOpenFailed"){
			$msg = "Logon credentials were not accepted."
			popupError($msg,0,"Invalid Logon",0x0)
			setupConnection
		}
	}
	finally{
		Import-Module ActiveDirectory
	}
}

#endregion Local Functions

#REGEX expressions used
$global:regex_4x4 = '[a-zA-Z]{1,4}[0-9]{4}'
$global:regex_DGN = '[a-zA-Z0-9-_. ()]{5,256}'
$global:regex_Alias = '[a-zA-Z0-9-_.]{5,256}'
$global:regex_Email = '[a-zA-Z0-9-_.]{5,254}' #Not an actual email regex expression. This just checks that the account name is valid (ie the 'foo' in foo@ou.edu)
$global:regex_SponsorList = '[a-zA-Z0-9,]'
$global:regex_MemberList = '[a-zA-Z]{1,4}[0-9]{4}\n'

$DebugPreference = "Continue"
setupConnection
GenerateSetupForm

#Call the Function
GenerateForm

Remove-PSSession $PSSession
exit

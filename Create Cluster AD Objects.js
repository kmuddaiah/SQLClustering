var output;
System.log("-------------PC2.1 Create Cluster AD Objects-----------")
System.log("------------- Version 1.0.4 - 27th Jan 2022 -----------")
System.log("------------- Version 1.1.6 - 10th Feb 2022 -----------")
System.log("------------- Version 1.1.9 - 16th Feb 2022 -----------")

var session;

System.debug("REGION_Value :" + REGION_Value)
System.debug("QUEST_SERVER_Value :" + QUEST_SERVER_Value)
System.debug("GAD_SVC_ACCOUNT_Value :" + GAD_SVC_ACCOUNT_Value)
System.debug("GAD_AD_VMOU_Value :" + GAD_AD_VMOU_Value)
System.debug("VM_CLUSTERNAME_Value :" + VM_CLUSTERNAME_Value)
System.debug("VM_SQLListener_Value :" + VM_SQLListener_Value)
System.debug("DNFilter_Value :" + DNFilter_Value)
System.debug("defaultZMB_Value :" + defaultZMB_Value)
System.debug("defaultSecOwner_Value :" + defaultSecOwner_Value)
System.debug("ADLocation_Value :" + ADLocation_Value)
System.debug("AOID_Value :" + AOID_Value)
System.debug("AOEmail_Value :" + AOEmail_Value)
System.debug("PBContactID_Value :" + PBContactID_Value)
System.debug("PBContactEmail_Value :" + PBContactEmail_Value)
System.debug("Host_Environment_Value :" + Host_Environment_Value)
System.debug("contactPath_Value :" + contactPath_Value)


try {
	session = PshellHost.openSession();
	var scriptcmd1 = 	'$REGION= "'+REGION_Value+'" \r \n'+
	'	if ($REGION -eq "null") {$REGION = $null} \r \n'+
	'	$QUEST_SERVER= "'+QUEST_SERVER_Value+'" \r \n'+
	'	if ($QUEST_SERVER -eq "null") {$QUEST_SERVER = $null} \r \n'+
	'	$GAD_SVC_ACCOUNT = "'+GAD_SVC_ACCOUNT_Value+'" \r \n'+
	'	if ($GAD_SVC_ACCOUNT -eq "null") {$GAD_SVC_ACCOUNT = $null} \r \n'+
	'	$GAD_SVC_PWD= "'+domain_password+'" \r \n'+
	'	if ($GAD_SVC_PWD -eq "null") {$GAD_SVC_PWD = $null} \r \n'+
	'	$GAD_AD_VMOU= "'+GAD_AD_VMOU_Value+'" \r \n'+
	'	if ($GAD_AD_VMOU -eq "null") {$GAD_AD_VMOU = $null} \r \n'+
	'	$vmClusterName= "'+VM_CLUSTERNAME_Value+'" \r \n'+
	'	if ($vmClusterName -eq "null") {$vmClusterName = $null} \r \n'+
	'	$vmSQLLName= "'+VM_SQLListener_Value+'" \r \n'+
	'	if ($vmSQLLName -eq "null") {$vmSQLLName = $null} \r \n'+
	'	$DNFilter= "'+DNFilter_Value+'" \r \n'+
	'	if ($DNFilter -eq "null") {$DNFilter = $null} \r \n'+
	'	$defaultZMB= "'+defaultZMB_Value+'" \r \n'+
	'	if ($defaultZMB -eq "null") {$defaultZMB = $null} \r \n'+
	'	$defaultSecOwner= "'+defaultSecOwner_Value+'" \r \n'+
	'	if ($defaultSecOwner -eq "null") {$defaultSecOwner = $null} \r \n'+
	'	$ADLocation= "'+ADLocation_Value+'" \r \n'+
	'	if ($ADLocation -eq "null") {$ADLocation = $null} \r \n'+
	'	$AOID = "'+AOID_Value+'"; \r \n'+
	'	if ($AOID -eq "null") {$AOID = $null} \r \n'+
	'	$AOEmail = "'+AOEmail_Value+'"; \r \n'+
	'	if ($AOEmail -eq "null") {$AOEmail = $null} \r \n'+
	'	$PBContactID = "'+PBContactID_Value+'"; \r \n'+
	'	if ($PBContactID -eq "null") {$PBContactID = $null} \r \n'+
	'	$PBContactEmail = "'+PBContactEmail_Value+'"; \r \n'+
	'	if ($PBContactEmail -eq "null") {$PBContactEmail = $null} \r \n'+
	'	$Host_Environment = "'+Host_Environment_Value+'"; \r \n'+
	'	if ($Host_Environment -eq "null") {$Host_Environment = $null} \r \n'+
	'	$contactPath = "'+contactPath_Value+'"; \r \n'+
	'	if ($contactPath -eq "null") {$contactPath = $null} \r \n'+
	'	if (-not(Get-WindowsFeature RSAT-AD-Powershell).InstallState -eq "Installed") {Add-WindowsFeature RSAT-AD-Powershell;} \r \n'+
	'	if (-not(Get-Module |where {$_.Name -eq "ActiveRolesManagementShell"})) {Import-Module ActiveRolesManagementShell -Verbose -ErrorAction SilentlyContinue;} \r \n'+
	'	Write-Host "Sleeping for 5 seconds..."; \r \n'+
	'	Start-Sleep -Seconds 5; \r \n'+
	'	$sCred = $False; \r \n';
	var scriptcmd2 = 'function preStage($dname, $dfun, $dpath, $dcon){ \r \n'+
	'	Write-Host "First check if AD computer object ($dname) already exist"; \r \n'+
	'	Write-Host "Searching AD computer object in $dpath"; \r \n'+
	'	try { \r \n'+
	'		$compu=(Get-QADComputer -LdapFilter "(cn=$dname)" -connection $dcon).dn; \r \n'+
	'	} catch { \r \n'+
	'		Write-Host "Error searching computer object $dname)"; \r \n'+
	'	} \r \n'+
	'	if ($compu -ne $null){ \r \n'+
	'		Write-Host "Computer object ($dname) found in AD. NO need to Create"; \r \n'+
	'		return 0; \r \n'+
	'	} else { \r \n'+
	'		Write-Host "Computer object ($dname) NOT found in AD. Need to Create"; \r \n'+
	'	} \r \n'+
	'	if (($Host_Environment -eq "UAT") -or ($Host_Environment -eq "DEV")) { \r \n'+
	'		$name1="$dname`_$dfun`POwner"; \r \n'+
	'		$name2="$dname`_$dfun`SOwner"; \r \n'+
	'		$email1=$AOEmail; \r \n'+
	'		$email2=$PBContactEmail; \r \n'+
	'		if($email2){ \r \n'+
	'			Write-Host "PrimaryBusinessContactEmail is not null, use it for contact2"; \r \n'+
	'		} else{ \r \n'+
	'			Write-Host "PrimaryBusinessContactEmail is null, use ApplicationOwnerEmail for contact2"; \r \n'+
	'			$email2=$email1; \r \n'+
	'		} \r \n'+
	'		try { \r \n'+
	'			$check_contact1 = Get-QADObject -LdapFilter "(CN=$name1,$contactPath)" -connection $dcon; \r \n'+
	'			$check_contact2 = Get-QADObject -LdapFilter "(CN=$name2,$contactPath)" -connection $dcon; \r \n'+
	'		} catch { \r \n'+
	'			Write-Host $_; \r \n'+
	'		} \r \n'+
	'		if ($check_contact1 -eq $null){ \r \n'+
	'			$Contact1 = New-QADObject -Type contact -Name $name1 -ObjectAttributes @{"mail"=$email1; "sn"=$name1; "displayName"=$name1} -ParentContainer $contactPath -connection $dcon; \r \n'+
	'			Write-Host "Contact1: $Contact1"; \r \n'+
	'		} else { \r \n'+
	'			Write-Host "Contact1: $check_contact1 already exist"; \r \n'+
	'			$Contact1 = $check_contact1; \r \n'+
	'		} \r \n'+
	'		if ($check_contact2 -eq $null){ \r \n'+
	'			$Contact2 = New-QADObject -Type contact -Name $name2 -ObjectAttributes @{"mail"=$email2; "sn"=$name2; "displayName"=$name2} -ParentContainer $contactPath -connection $dcon; \r \n'+
	'			Write-Host "Contact2: $Contact2"; \r \n'+
	'		} else { \r \n'+
	'			Write-Host "Contact2: $check_contact2 already exist"; \r \n'+
	'			$Contact2 = $check_contact2; \r \n'+
	'		} \r \n'+
	'		Write-Host "Creating $dname object"; \r \n'+
	'		Write-Host "ManagedBy is $Contact1"; \r \n'+
	'		$samName = "$dname"+"$"; \r \n'+
	'		$i=1; \r \n'+
	'		$number=11; \r \n'+
	'		$tnumber=$number-1; \r \n'+
	'		do { \r \n'+
	'			try { \r \n'+
	'				Write-Host "Computer object $dname Create Attempt $i/$tnumber"; \r \n'+
	'				$newReturn=New-QADComputer -Name "$dname" -SamAccountName "$samName" -connection $dcon -ParentContainer "$dpath" -Location "$ADLocation" -ManagedBy "$Contact1" -objectAttributes @{edsvaSecondaryOwners="$Contact2";edsaJoinComputerToDomain="$GAD_SVC_ACCOUNT"} -ErrorAction Ignore; \r \n'+
	'				Write-Host "Computer object $dname created: $newReturn"; \r \n'+
	'				break; \r \n'+
	'			} catch { \r \n'+
	'				Write-Host "Error creating computer object $dname, sleep 300 seconds. Attempt $i/$tnumber"; \r \n'+
	'				start-sleep -s 300 \r \n'+
	'			} \r \n'+
	'			$i++; \r \n'+
	'		} while ($i -le $number) \r \n'+
	'	} else { \r \n'+
	'		Write-Host "Creating computer object ($dname) in $dpath"; \r \n'+
	'		Write-Host "Getting $dname computer object ManagedBy fully qualified name"; \r \n'+
	'		if ($AOID -eq $null){ \r \n'+
	'			$ZurichManagedBy="$defaultZMB"; \r \n'+
	'		} else { \r \n'+
	'			Write-Host "Setting Managed by value to $AOID"; \r \n'+
	'			$ZurichManagedBy=$AOID; \r \n'+
	'		} \r \n'+
	'		Write-Host "Searching user $ZurichManagedBy in AD $DNFilter"; \r \n'+
	'		$ZurichManagedByFull=Get-QADUser $ZurichManagedBy -connection $dcon | Where-Object {$_.DN -like "$DNFilter"} | Select -Expand DN -first 1 \r \n'+
	'		Write-Host "Managed By value for the computer object will be $ZurichManagedByFull"; \r \n'+
	'		Write-Host "Getting $dname object Secondary Owner fully qualified name"; \r \n'+
	'		if ($PBContactID -eq $null){ \r \n'+
	'			$SecondaryOwner="$defaultSecOwner"; \r \n'+
	'		} else { \r \n'+
	'			Write-Host "Setting SecondaryOwner value to $PBContactID"; \r \n'+
	'			$SecondaryOwner=$PBContactID; \r \n'+
	'		} \r \n'+
	'		$SecondaryOwnerFull=Get-QADUser $SecondaryOwner -connection $dcon | Where-Object {$_.DN -like "$DNFilter"} | Select -Expand DN -first 1 \r \n'+
	'		if (!$SecondaryOwnerFull) { \r \n'+
	'			Write-Host "Error in preStage ($dfun) - SecondaryOwnerFull is empty" \r \n'+
	'			Exit 1; \r \n'+
	'		} \r \n'+
	'		Write-Host "Secondary owner value for the computer object ($dname) will be $SecondaryOwnerFull"; \r \n'+
	'		Write-Host "Creating computer object with AD users"; \r \n'+
	'		$samName = "$dname"+"$"; \r \n'+
	'		$i=1; \r \n'+
	'		$number=11; \r \n'+
	'		$tnumber=$number-1; \r \n'+
	'		do { \r \n'+
	'			try { \r \n'+
	'				Write-Host "Computer object $dname create attempt $i/$tnumber"; \r \n'+
	'				$newReturn=New-QADComputer -Name "$dname" -SamAccountName "$samName" -connection $dcon -ParentContainer "$dpath" -Location "$ADLocation" -ManagedBy "$ZurichManagedByFull" -objectAttributes @{edsvaSecondaryOwners="$SecondaryOwnerFull";edsaJoinComputerToDomain="$GAD_SVC_ACCOUNT"} \r \n'+
	'				Write-Host "Computer object created: $newReturn"; \r \n'+
	'				break; \r \n'+
	'			} catch { \r \n'+
	'				Write-Host "Error creating computer object $dname, sleep 300 seconds. Attempt $i/$tnumber"; \r \n'+
	'				start-sleep -s 300 \r \n'+
	'			} \r \n'+
	'			$i++; \r \n'+
	'		} while ($i -le $number) \r \n'+
	'	} \r \n'+
	'} \r \n';
	var scriptcmd4 = '$secPass= $null; \r \n'+
	'$cred= $null; \r \n'+
	'$secPass= ConvertTo-SecureString $GAD_SVC_PWD -AsPlainText -Force; \r \n'+
	'$cred= New-object System.Management.Automation.PSCredential -ArgumentList "$GAD_SVC_ACCOUNT", $secPass; \r \n'+
	'$k=1; \r \n'+
	'$number1=11; \r \n'+
	'$tnumber1=$number1-1; \r \n'+
	'do { \r \n'+
	'	try { \r \n'+
	'		Write-Host "Connect-QADService -service $QUEST_SERVER -Proxy -Credential $cred"; \r \n'+
	'		$myConnection = Connect-QADService -service $QUEST_SERVER -Proxy -Credential $cred; \r \n'+
	'		Write-Host "Quest connection attempt $k/$tnumber1 success with ARS Server $QUEST_SERVER"; \r \n'+
	'		break; \r \n'+
	'	} catch { \r \n'+
	'		Write-Host "Quest connection attempt $k/$tnumber1 error with ARS Server $QUEST_SERVER, sleep for 300 and retry"; \r \n'+
	'		start-sleep -s 300 \r \n'+
	'	} \r \n'+
	'	$k++; \r \n'+
	'} while ($k -le $number1) \r \n'+
	'$funName = "Cluster"; \r \n'+
	'preStage $vmClusterName $funName $GAD_AD_VMOU $myConnection \r \n'+
	'$funName = "SQLL1"; \r \n'+
	'preStage $vmSQLLName $funName $GAD_AD_VMOU $myConnection \r \n'+
	'disconnect-qadservice -Connection $myConnection; \r \n'+
	'Write-Host "Disconnect-Qadservice with $QUEST_SERVER"; \r \n';
	var scriptFinal = scriptcmd1+scriptcmd2+scriptcmd3+scriptcmd4;
	output = System.getModule("com.vmware.library.powershell").invokeScript(PshellHost,scriptFinal,session.getSessionId());
} finally {
	if (session) {
		PshellHost.closeSession(session.getSessionId());
	}
}
System.log("-------------End--PC2.1 Create Cluster AD Objects-----------")

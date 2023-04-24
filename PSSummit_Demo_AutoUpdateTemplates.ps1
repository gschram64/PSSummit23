# VMware Windows Server Template Automation Process
# Geoff Schram
# 10/26/2022
# PowerShell Ver. 4 or newer (M&M Cleanup Requires Ver. 5.1)
# Required Modules: PSWindowsUpdate,VMware.VimAutomation.Core
##########################################################################

function Update-VMwareTemplates {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]$OSver
  )
  
  begin {
    function Invoke-StopWatch {
      param (
      [int]$Time = 300, # default to 5 mins
      [string]$Message = "Something" # default message
      )
    
      $clock = 1
      while ($clock -lt $Time) {
        Write-Progress -Id 0 -Activity "Waiting $($Time/60) Minute(s) for $Message to Complete." -Status "[$($clock)/$Time] Completed" -PercentComplete ($clock/$Time*100)
        Start-Sleep -s 1
        $clock++
      }
      Write-Progress -Id 0 -Activity "Waiting $($Time/60) Minute(s) for $Message to Complete." -Status "[$($clock)/$Time] Completed" -Completed
    }
    
    ### Start: Main Variables ###
    try {
      # Send-MailMessage Default values for Errors
      $MailTo = "You@mail.org"
      $MailFrom = "$env:Computername@mail.org"
      $MailSubject = "VMware Windows Server Template Automation Process Error"
      $AllRIT = "All@mail.org"
      $WinTeam = "Win@mail.org"
      $PSEmailServer = "out@mail.org"
    
      # vCenter & Domain Info
      $vCenter1 = "vCenter1"
      $vCenter2 = "vCenter1"
      $Domain = "test.domain.org"
      $DC1 = "dc1"
    
      # Credential Information
      $CredPath = "cPath1"
      $SvcAcct = "svcAcct"
      $SvcAcctCredFile = "svcAcct.ps1.credential"
      $LocalAdmin = "lcAdmin"
      $LocalAdminCredFile = "lcAdmin.ps1.credential"
      $ADCleanupSvcAcct = "ADCsvcAcct"
      $ADCleanupCredFile = "ADCsvcAcct.ps1.credential"
      $SvcAcct2 = "svcAcct2"
      $SvcAcctCredFile2 = "svcAcct2.ps1.credential"
      $SSAdmin = "ssAdmin"
      $SSAdminCredFile = "ssAdmin.ps1.credential"
      $SSDevAPICredPath = "ssDevAPI.ps1.credential"
      $SSProdAPICredPath = "ssProdAPI.ps1.credential"

      # Verbose output of Main Variables
      Write-Verbose "
      Email Vars..
      Email To: $MailTo
      Email From: $MailFrom
      Email Subject: $MailSubject
      Email RIT: $AllRIT
      Email Windows Team: $WinTeam
      Email Server: $PSEmailServer
      
      Infrastructure Vars..
      Sabin vCenter: $vCenter1
      Liberty vCenter: $vCenter2
      Domain: $Domain
      Research Domain Controller: $DC1
      "
    
      ### Start: Creds ###
      <#
      try {
        if (Get-ChildItem $SvcAcctCredFile -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName $SvcAcct
          $creds | Export-Clixml $SvcAcctCredFile
        }
        $SvcAcctpw = (Import-Clixml $SvcAcctCredFile).Password
        $SvcAcctCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SvcAcct, $SvcAcctpw
    
        if (Get-ChildItem $LocalAdminCredFile -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName $LocalAdmin
          $creds | Export-Clixml $LocalAdminCredFile
        }
        $LocalAdminpw = (Import-Clixml $LocalAdminCredFile).Password
        $LocalAdminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $LocalAdmin, $LocalAdminpw
    
        if (Get-ChildItem $ADCleanupCredFile -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName $ADCleanupSvcAcct
          $creds | Export-Clixml $ADCleanupCredFile
        }
        $ADCleanuppw = (Import-Clixml $ADCleanupCredFile).Password
        $ADCleanupCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADCleanupSvcAcct, $ADCleanuppw
    
        if (Get-ChildItem $SvcAcctCredFile2 -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName $SvcAcct2
          $creds | Export-Clixml $SvcAcctCredFile2
        }
        $SvcAcct2pw = (Import-Clixml $SvcAcctCredFile2).Password
        $SvcAcct2Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SvcAcct2, $SvcAcct2pw
    
        if (Get-ChildItem $SSAdminCredFile -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName $SSAdmin
          $creds | Export-Clixml $SSAdminCredFile
        }
        $SSAdminpw = (Import-Clixml $SSAdminCredFile).Password
        $SSAdminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SSAdmin, $SSAdminpw
    
        if (Get-ChildItem $SSDevAPICredPath -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName "YourAPIUser"
          $creds | Export-Clixml $SSDevAPICredPath
        }
        $SSDevApiKey = (Import-Clixml $SSDevAPICredPath).Password
        $SSDevApiKey = [System.Net.NetworkCredential]::new("", $SSDevApiKey).Password
    
        if (Get-ChildItem $SSProdAPICredPath -ErrorAction Ignore) { # Do nothing
        } else {
          $creds = Get-Credential -UserName "YourAPIUser"
          $creds | Export-Clixml $SSProdAPICredPath
        }
        $SSProdApiKey = (Import-Clixml $SSProdAPICredPath).Password
        $SSProdApiKey = [System.Net.NetworkCredential]::new("", $SSProdApiKey).Password
    
      } # end try
      catch {
    
      } # end catch
      #>
      ### End: Creds ###
    
    
      # removes "20" from "2019" and "2022" for test vm deployment (NetBios naming)
      $OSver = $OSver.toString()
      if ($OSver.Length -eq 4) {
        $OSVerAbbr = $OSver.Substring(2,2)
      } else {
        $OSVerAbbr = $OSver
      }
      Write-Verbose "Abbreviated OS Version: $OSVerAbbr"

      # Date format for Logs
      $DT = Get-Date -UFormat "%Y%m%d"
      # Log file location and filename
      $LogFile = "/Users/schje3/Library/CloudStorage/OneDrive-Personal/PSSummit23/TemplateAutoUpdateLog_Win$($OSver)_$DT.log"
      Write-Verbose "Date Format: $DT"
      Write-Verbose "Log File: $LogFile"
    
      # Import the PowerCLI module
      #Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue
      Write-Verbose "Importing Module VMware.VimAutomation.Core using Import-Module..."
    
      # Connect to vCenter
      try {
        #Set-PowerCLIConfiguration -Scope Session -ParticipateInCeip $false -Confirm:$false -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -WarningAction SilentlyContinue
        Write-Verbose "Setting PowerCLIConfiguration using Set-PowerCLIConfiguration..."
      } 
      catch {
    
      } # end try/catch catch empty on purpose

      #Connect-VIServer $vCenter1
      Write-Verbose "Connecting to $vCenter1..."
    
    } # end try
    catch {
    
    } # end catch
    ### End: Main Variables ###
    
    
    ### Start: Determine Template Names, etc. ### - Testing Passed
    try {
      <#   # Check Current Templates - Researved for iteration changes (same month)
        $dte = Get-Date -Format "yyyyMM"
        if (Get-Template "TPLT-Win$($OSver)_D$($dte)ns") {
          # Updating Templates in same month, so versioning will be needed
          $CurrentDevTemplate = "TPLT-Win$($OSver)_vD$($dte)ns"
          $CurrentDevSysprepTemplate = "TPLT-Win$($OSver)_vD$($dte)s"
          $CurrentProdTemplate = "TPLT-Win$($OSver)_vP$($dte)s"
          
          $i = 0
          do {
            $i++
            $dte = Get-Date
            $oldDTE = $(Get-Date).AddMonths(-$i)
            $PreviousTemplateDate = "$($oldDTE.ToString("yyyy"))$($oldDTE.ToString("MM"))"
            $PreviousReleaseProdTemplate = "TPLT-Win$($OSver)_vP$($PreviousTemplateDate)pr"
          } until ((Get-Template $PreviousReleaseProdTemplate -ErrorAction SilentlyContinue) -or ($i -ge $dte.Month))
      
          $UpdatedDevTemplate = "TPLT-Win$($OSver)_vD$(Get-Date -format "yyyyMM")ns"
          $UpdatedDevSysprepTemplate = "TPLT-Win$($OSver)_vD$(Get-Date -format "yyyyMM")s"
          $PreviousProdTemplate = $PreviousReleaseProdTemplate 
          $UpdatedProdTemplate = "TPLT-Win$($OSver)_vP$(Get-Date -format "yyyyMM")s"
        }
      #>  
    
      # Current Dev Templates
      $i = 0
      do {
        $i++
        $dte = Get-Date
        $oldDTE = $(Get-Date).AddMonths(-$i)
        $PreviousTemplateDate = "$($oldDTE.ToString("yyyy"))$($oldDTE.ToString("MM"))"
        $CurrentDevTemplate = "TPLT-Win$($OSver)_vD$($PreviousTemplateDate)ns"
        $CurrentDevSysprepTemplate = "TPLT-Win$($OSver)_vD$($PreviousTemplateDate)s"
        Write-Verbose "Checking for Dev Template: $CurrentDevTemplate..."
      } until (
        #(Get-Template $CurrentDevTemplate -ErrorAction SilentlyContinue)  -or ($i -ge 12))
        ($i -eq 1))
      Write-Verbose "Current Dev Template: $CurrentDevTemplate"
      Write-Verbose "Current Dev Sysprep Template: $CurrentDevSysprepTemplate"

      # Previous Prod Templates
      $i = 0
      do {
        $i++
        $dte = Get-Date
        $oldDTE = $(Get-Date).AddMonths(-$i)
        $PreviousTemplateDate = "$($oldDTE.ToString("yyyy"))$($oldDTE.ToString("MM"))"
        $CurrentProdTemplate = "TPLT-Win$($OSver)_vP$($PreviousTemplateDate)s"
        Write-Verbose "Checking for Prod Template: $CurrentProdTemplate"
      } until (
        #(Get-Template $CurrentProdTemplate -ErrorAction SilentlyContinue) -or ($i -ge 12))
        ($i -eq 1))
      Write-Verbose "Current Prod Template: $CurrentProdTemplate"

      # Get Prod Previous Release Template
      $i = 0
      do {
        $i++
        $dte = Get-Date
        $oldDTE = $(Get-Date).AddMonths(-$i)
        $PreviousTemplateDate = "$($oldDTE.ToString("yyyy"))$($oldDTE.ToString("MM"))"
        $PreviousReleaseProdTemplate = "TPLT-Win$($OSver)_vP$($PreviousTemplateDate)pr"
        Write-Verbose "Checking for Previous Prod Template: $PreviousReleaseProdTemplate"
      } until (
        #(Get-Template $PreviousReleaseProdTemplate -ErrorAction SilentlyContinue) -or ($i -ge 12))
        ($i -eq 1))
      Write-Verbose "Previous Prod Template: $PreviousReleaseProdTemplate"

      # Template Names
      $UpdatedDevTemplate = "TPLT-Win$($OSver)_vD$(Get-Date -format "yyyyMM")ns"
      $UpdatedDevSysprepTemplate = "TPLT-Win$($OSver)_vD$(Get-Date -format "yyyyMM")s"
      $PreviousProdTemplate = $PreviousReleaseProdTemplate 
      $CurrentProdTemplate = $CurrentProdTemplate
      $UpdatedProdTemplate = "TPLT-Win$($OSver)_vP$(Get-Date -format "yyyyMM")s"
      Write-Verbose "
      Updated Dev Template: $UpdatedDevTemplate
      Updated Dev Sysprep Template: $UpdatedDevSysprepTemplate
      Previous Prod Template: $PreviousProdTemplate
      Current Prod Template: $CurrentProdTemplate
      Updated Prod Template: $UpdatedProdTemplate
      "
    
      # Datastore Information
      #$DevDatastore = (Get-Datastore -Id (Get-Template -Name $CurrentDevTemplate).DatastoreIdList[0])
      #$ProdDatastore = (Get-Datastore -Id (Get-Template -Name $CurrentProdTemplate).DatastoreIdList[0])
      #$DevTemplateDir = (Get-Folder -Id (Get-Template -Name $CurrentDevTemplate).FolderId).Name
      #$ProdTemplateDir = (Get-Folder -Id (Get-Template -Name $CurrentProdTemplate).FolderId).Name
      #$ProdPrevReleaseTemplateDir = (Get-Folder -Id (Get-Template -Name $PreviousProdTemplate).FolderId)
      $DevDatastore = "DevDatastore"
      $ProdDatastore = "ProdDatastore"
      $DevTemplateDir = "DevTemplateDir"
      $ProdTemplateDir = "ProdTemplateDir"
      $ProdPrevReleaseTemplateDir = "ProdPrevReleaseTempDir"
      Write-Verbose "
      Dev Datastore: $DevDatastore
      Prod Datastore: $ProdDatastore
      Dev Template Directory: $DevTemplateDir
      Prod Template Directory: $ProdTemplateDir
      Prod Previous Release Template Directory: $ProdPrevReleaseTemplateDir
      "

      # Host Information
      #$DevWindowsCluster = (Get-VMHost -Id (Get-Template -Name $CurrentDevTemplate).HostId).Name
      #$ProdWindowsCluster = (Get-VMHost -Id (Get-Template -Name $CurrentProdTemplate).HostId).Name
      $DevWindowsCluster = "DevWinCluster"
      $ProdWindowsCluster = "ProdWinCluster"
      Write-Verbose "
      Dev Windows Cluster: $DevWindowsCluster
      Prod Windows Cluster: $ProdWindowsCluster
      "

      # Excluded Windows Updates
      $WinUpdatesExclude = "Windows Malicious Software Removal Tool*"
      Write-Verbose "
      Excluded Windows Updates: $WinUpdatesExclude
      "

    } # end try
    catch {
    
    } # end catch
    ### End: Determine Template Names, etc. ###
    
  } # end begin block
  
  process {
    ### Start: Apply Windows Server Updates To Dev Template ###
    try {
      # Start Logging
      "### VMware Windows Server Template Automation Process - $DT ###" | Out-file -Filepath $LogFile -Append
      "" | Out-file -Filepath $LogFile -Append
      "Working on $CurrentDevTemplate..." | Out-file -Filepath $LogFile -Append

      # Create Backup of Dev Template
      "Creating Backup Template of $CurrentDevTemplate using New-Template..." | Out-file -Filepath $LogFile -Append
      #New-Template -Template $CurrentDevTemplate -Name "$($CurrentDevTemplate)_bkp" -Datastore $DevDatastore -Location $DevTemplateDir -VMHost $DevWindowsCluster -Confirm:$false

      # Convert a template to a VM
      "Converting $CurrentDevTemplate to VM using Set-Template and invoking stopwatch for 60 seconds..." | Out-file -Filepath $LogFile -Append
      #Set-Template -Template $CurrentDevTemplate -ToVM -Confirm:$false -RunAsync

      # Make a 60 seconds delay
      Invoke-StopWatch -Time 2

      # Start the virtual machine
      "Powering on $CurrentDevTemplate using Start-VM, Get-VMQuestions, Set-VMQuestion, and invoking stopwatch for 120 seconds..." | Out-file -Filepath $LogFile -Append
      #Start-VM -VM $CurrentDevTemplate | Get-VMQuestion | Set-VMQuestion -DefaultOption -Confirm:$false
      Invoke-StopWatch -Time 2

      <# # Check for Module PSWindowsUpdate
      "Checking for PSWindowsUpdate Module on $CurrentDevTemplate and installing if necessary..." | Out-file -Filepath $LogFile -Append
      Invoke-VMScript -ScriptType PowerShell -ScriptText 'if (Get-Module PSWindowsUpdate) {} else { Install-Module -Name PSWindowsUpdate -Confirm:$false }' -VM $CurrentDevTemplate -GuestCredential $LocalAdminCreds | Out-file -Filepath $LogFile -Append
      Start-Sleep -s 60
      #>

      # Import PSWindowsUpdate Module
      "Importing PSWindowsUpdate Module on $CurrentDevTemplate using Invoke-VMScript and invoking stopwatch for 10 seconds..." | Out-file -Filepath $LogFile -Append
      #Invoke-VMScript -ScriptType PowerShell -ScriptText "Import-Module -Name PSWindowsUpdate" -VM $CurrentDevTemplate -GuestCredential $LocalAdminCreds | Out-file -Filepath $LogFile -Append
      Invoke-StopWatch -Time 2

      # Run the command to install all available updates in the guest OS using VMWare Tools (the update installation log is saved to a file: C:\temp\Update.log)
      "Executing Windows Updates via PowerShell with Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -NotTitle '$WinUpdatesExclude' -AutoReboot on $CurrentDevTemplate using Invoke-VMScript and invoking stopwatch for 60 seconds..." | Out-file -Filepath $LogFile -Append
      #Invoke-VMScript -ScriptType PowerShell -ScriptText "Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -NotTitle '$WinUpdatesExclude' -AutoReboot" -VM $CurrentDevTemplate -GuestCredential $LocalAdminCreds | Out-file -Filepath $LogFile -Append
      Invoke-StopWatch -Time 2

      # Restart VM
      "Restarting $CurrentDevTemplate using Restart-VMGuest and sleeping for 45 minutes (2700s) while updates complete..." | Out-file -Filepath $LogFile -Append
      #Restart-VMGuest -VM $CurrentDevTemplate -Confirm:$false
      Invoke-StopWatch -Time 2 -Message "Post Reboot OS Updates"

      # Run Windows Update a second time
      "Checking for additional Windows Updates on $CurrentDevTemplate using Invoke-VMScript and invoking stopwatch for 60 seconds..." | Out-file -Filepath $LogFile -Append
      #Invoke-VMScript -ScriptType PowerShell -ScriptText "Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -NotTitle '$WinUpdatesExclude' -AutoReboot" -VM $CurrentDevTemplate -GuestCredential $LocalAdminCreds | Out-file -Filepath $LogFile -Append
      Invoke-StopWatch -Time 2

      <# # Clean up the WinSxS component store and optimize the image with DISM
      Invoke-VMScript -ScriptType PowerShell -ScriptText "Dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase" -VM $CurrentDevTemplate -GuestCredential $LocalAdminCreds
      Start-sleep -s 1800 #>

      # Force restart the VM
      "Restarting $CurrentDevTemplate using Restart-VMGuest and sleeping for 15 minutes (900s) in case there are updates to complete..." | Out-file -Filepath $LogFile -Append
      #Restart-VMGuest -VM $CurrentDevTemplate -Confirm:$false
      Invoke-StopWatch -Time 2 -Message "2nd Round Post Reboot OS Updates"

      # Update VMTools version
      "Updating VMware Tools on $CurrentDevTemplate using Update-Tools..." | Out-file -Filepath $LogFile -Append
      #Update-Tools -VM $CurrentDevTemplate

      # Shut the VM down and convert it back to the template
      "Shutting down $CurrentDevTemplate using Shutdown-VMGuest and invoking stopwatch for 180 seconds..." | Out-file -Filepath $LogFile -Append
      #Shutdown-VMGuest -VM $CurrentDevTemplate -Confirm:$false
      Invoke-StopWatch -Time 2 -Message "Shutting Down $CurrentDevTemplate"
      "Converting $CurrentDevTemplate back to a Template and updating its name using Set-VM..." | Out-file -Filepath $LogFile -Append
      #Set-VM -VM $CurrentDevTemplate -ToTemplate -Name $UpdatedDevTemplate -Confirm:$false

      "Finished Processing $CurrentDevTemplate..." | Out-file -Filepath $LogFile -Append

    } # end try
    catch {

    } # end catch
    ### End: Apply Windows Server Updates To Dev Template ###


    ### Start: Sysprep Updated Dev Template ###
    try {
      # Spacing Log File
      "" | Out-file -Filepath $LogFile -Append
      "" | Out-file -Filepath $LogFile -Append
      "" | Out-file -Filepath $LogFile -Append

      # Clone Template with New Template Name for Current Month
      "Cloning Template $UpdatedDevTemplate for Sysprep process using New-Template..." | Out-file -Filepath $LogFile -Append
      #New-Template -Template $UpdatedDevTemplate -Name $UpdatedDevSysprepTemplate -Datastore $DevDatastore -Location $DevTemplateDir -VMHost $DevWindowsCluster -Confirm:$false

      # Convert Template to VM
      "Converting Template $UpdatedDevSysprepTemplate to VM using Set-Template..." | Out-file -Filepath $LogFile -Append
      #Set-Template -Template $UpdatedDevSysprepTemplate -ToVM -Confirm:$false

      # Power on VM
      "Powering on $UpdatedDevSysprepTemplate using Start-VM and invoking stopwatch for 120 seconds..." | Out-file -Filepath $LogFile -Append
      #Start-VM -VM $UpdatedDevSysprepTemplate | Get-VMQuestion | Set-VMQuestion -DefaultOption -Confirm:$false
      Invoke-StopWatch -Time 2 -Message "$UpdatedDevSysprepTemplate Power On"

      # Execute Custom PowerShell Script to Sysprep VM/Template
      "Executing Sysprep PowerShell Script on $UpdatedDevSysprepTemplate using Invoke-VMScript..." | Out-file -Filepath $LogFile -Append
      #Invoke-VMScript -ScriptType PowerShell -ScriptText "Set-Location C:\Windows\System32\Sysprep; ./sysprep.exe /generalize /oobe /unattend:c:\Trashbox\Win$($OSver)_Sysprep_autounattend.xml" -VM $UpdatedDevSysprepTemplate -GuestCredential $LocalAdminCreds -ErrorAction Ignore | Out-file -Filepath $LogFile -Append

      # Verify VM is Powered Off
      "Verifying VM $UpdatedDevSysprepTemplate is Powered Off..." | Out-file -Filepath $LogFile -Append
      $PwrState = "PoweredOn"
      $LoopCount = 0
      While ($PwrState -eq "PoweredOn") {
        #$ChkPwrState = (Get-VM -Name $UpdatedDevSysprepTemplate).PowerState
        $ChkPwrState = "PoweredOff"
        if ($ChkPwrState = "PoweredOff") {
          $PwrState = "PoweredOff"
        } else {
          if ($LoopCount -ge 10) {
            $SendMailParams = @{
              From = $MailFrom
              To = $MailTo
              Subject = "Windows Server $OSver Template Update FAILED during DEV Sysprep Phase"
              Body = "Sysprep of the Updated Template FAILED for some reason and needs to be inspected."
            }
            #Send-MailMessage @SendMailParams
            Write-Host "Sending email for error $($SendMailParams.Subject)..."

            Write-Error "Sysprep of the Updated Template FAILED!"
            #Exit
          } else {
            Invoke-StopWatch -Time 2 #60
          }
        }
        $LoopCount++
      }
      Write-Verbose "$UpdatedDevSysprepTemplate Power State: $PwrState, invoking stopwatch for 120 seconds..."
      Invoke-StopWatch -Time 2

      # Converting Syspreped VM back to Template
      "Converting Syspreped VM to Template $UpdatedDevSysprepTemplateusing Set-VM..." | Out-file -Filepath $LogFile -Append
      #Set-VM -VM $UpdatedDevSysprepTemplate -ToTemplate -Confirm:$false

      "Finished Processing $UpdatedDevSysprepTemplate for Sysprep..." | Out-file -Filepath $LogFile -Append

    } # end try
    catch {

    } # end catch
    ### End: Sysprep Updated Dev Template ###


    ### Start: Test New Syspreped Dev Template ### - Needs updating once RES DNS Scope is updated
    try {
      # Spacing Log File
      "" | Out-file -Filepath $LogFile -Append
      "" | Out-file -Filepath $LogFile -Append
      "" | Out-file -Filepath $LogFile -Append


      # Deploy Test VM
      $TestVMName = "ritwin$($OSVerAbbr)tstd01"
      "Deploying a Test VM ritwin$($OSVerAbbr)tstd01 using Template $UpdatedDevSysprepTemplate using New-VM..." | Out-file -Filepath $LogFile -Append
      #New-VM $DevWindowsCluster -Template $UpdatedDevSysprepTemplate -Name "ritwin$($OSVerAbbr)tstd01" -Location $DevTemplateDir -Datastore $DevDatastore -ResourcePool WindowsDev -Confirm:$false

      # Power on VM
      "Powering on $TestVMName using Start-VM and invoking stopwatch for 120 seconds..." | Out-file -Filepath $LogFile -Append
      #Start-VM -VM $TestVMName | Get-VMQuestion | Set-VMQuestion -DefaultOption -Confirm:$false
      Invoke-StopWatch -Time 2 -Message "$TestVMName Power on"

      # Waiting for Server Build to Complete
      "Waiting 15 minutes (900s) for $TestVMName to complete building..." | Out-file -Filepath $LogFile -Append
      Invoke-StopWatch -Time 2 -Message "$TestVMName Deployment Build"

      # Testing Server for Completeness
      "Checking Post Deployment Results on $UpdatedDevSysprepTemplate using Invoke-VMScript and storing results in variable..." | Out-file -Filepath $LogFile -Append
      #$PostDeploymentResults = Invoke-VMScript -ScriptType PowerShell -ScriptText 'Get-Content C:\tmp\DeploymentResults.txt' -VM $TestVMName -GuestCredential $SvcAcct2Creds
      #$PostDeploymentResults | Out-file -Filepath $LogFile -Append
      $PostDeploymentResults = "PASSED" # For PSsummit23 demo only
      if (
        #$PostDeploymentResults.ScriptOutput -match "FAILED"
        $PostDeploymentResults -eq "FAILED") {
        $SendMailParams = @{
          From = $MailFrom
          To = $MailTo
          Subject = "Windows Server $($OSver) Deployment Testing FAILED: Test VM $TestVMName"
          #Body = "Below is the testing results for OS Deployment of Test VM $TestVMName.
            #$($PostDeploymentResults.ScriptOutput.toString())"
        }
        #Send-MailMessage @SendMailParams
        Write-Host "Sending email for $($SendMailParams.Subject)..."

        Write-Error "Post Deployment Testing Failed!"
        #Exit
      } else {
          $SendMailParams = @{
            From = $MailFrom
            To = $MailTo
            Subject = "Windows Server $($OSver) Deployment Testing PASSED: Test VM $TestVMName"
            #Body = "Below is the testing results for OS Deployment of Test VM $TestVMName.
              #$($PostDeploymentResults.ScriptOutput.toString())"
          }
        #Send-MailMessage @SendMailParams
        Write-Host "Sending email for $($SendMailParams.Subject)..."
      }

    } # end try
    catch {
      $SendMailParams = @{
        From = $MailFrom
        To = $MailTo
        Subject = "Windows Server $($OSver) Deployment Testing FAILED: Test VM $TestVMName"
        Body = "Deployment of Test VM Failed! $(Get-Error -Newest 1) $($error[0].exception.gettype().fullname)
        Ivestigate and Restart Template Update Process."
      }
      #Send-MailMessage @SendMailParams
      Write-Host "Sending email for $($SendMailParams.Subject)..."

      # Any failure in this section should result in the script exiting...
      #Exit
    } # end catch

    # Cleanup Testing Server Objects - Testing Passed
    try {
      # Remove Test VM
      "Shutting down $TestVMName using Shutdown-VMGuest and invoking stopwatch for 120 seconds..." | Out-file -Filepath $LogFile -Append
      #Shutdown-VMGuest -VM $TestVMName -Confirm:$false
      Invoke-StopWatch -Time 2 -Message "$TestVMName Shutdown"

      "Removing $TestVMName from vCenter permanently using Remove-VM..." | Out-file -Filepath $LogFile -Append
      #Remove-VM -VM $TestVMName -DeletePermanently -Confirm:$false

      # Remove Backup of Dev NonSysprep Template
      "Removing $($CurrentDevTemplate)_bkp from vCenter permanently using Remove-Template..." | Out-file -Filepath $LogFile -Append
      #Remove-Template -Template "$($CurrentDevTemplate)_bkp" -DeletePermanently -Confirm:$false

      # Remove AD Object
      "Removing $TestVMName AD Object from AD using Remove-ADComputer..." | Out-file -Filepath $LogFile -Append
      #Remove-ADComputer -Identity $TestVMName -Confirm:$false -Credential $ADCleanupCreds

    } # end try
    catch {

    } # end catch

    # Clean up DNS - Testing Passed
    <#
    try {
      # mmSoap module requires PS Version 5.1
      "Removing DNS Entry's for $TestVMName..." | Out-file -Filepath $LogFile -Append
      $DNSResults = PowerShell.exe D:\Scripts\PSModules\MenandMice\TemplateAutomationDNSCleanup.ps1 -Usr $SSAdminCreds.username -PW $SSAdminCreds.password -TestVMName $TestVMName

      if ($DNSResults -contains "DNS Removal Successful") {
        "Removal of DNS Entry for $TestVMName Successful..." | Out-file -Filepath $LogFile -Append
      } else {
        "Removal of DNS Entry for $TestVMName FAILED..." | Out-file -Filepath $LogFile -Append
      }
    } # end try
    catch {

    } # end catch
    #>
    ### End: Test New Syspreped Dev Template ###


    ### Start: IF Test Passed, Publish Syspreped Dev Template to Prod ###
    try {
      # Clone Dev Template with New Prod Template Name for Current Month
      "Cloning Template $UpdatedDevSysprepTemplate for publishing process using New-Template..." | Out-file -Filepath $LogFile -Append
      #New-Template -Template $UpdatedDevSysprepTemplate -Name $UpdatedProdTemplate -Datastore $ProdDatastore -Location $ProdTemplateDir -VMHost $ProdWindowsCluster -Confirm:$false

      # Update SS VM Template ID's
      "Gathering $UpdatedProdTemplate Template ID *** STACKSTORM Integration ONLY *** ..." | Out-file -Filepath $LogFile -Append
      #$WinTemplateInfo = Get-Folder "Current Release" | Get-Template $UpdatedProdTemplate | Select-Object Name,Id

      # Process Windows Templates
      "Objectifying Template Data From $UpdatedProdTemplate *** STACKSTORM Integration ONLY *** ..." | Out-file -Filepath $LogFile -Append
      $WinReport = [system.collections.generic.list[pscustomobject]]::new()
      foreach ($Template in $WinTemplateInfo) {
        $WinReport.Add([PSCustomObject][ordered]@{
            TemplateName    = $Template.Name
            TemplateID      = ($Template.Id).Replace("VirtualMachine-","")
        })
      }

      $SabinSSKey = $null
      # Get StackStorm Dev Key for Prod Template
      "Gathering StackStorm Dev Key for $UpdatedProdTemplate *** STACKSTORM Integration ONLY *** ..." | Out-file -Filepath $LogFile -Append
      #$SabinSSKey = Get-SSKey -KeyName "TPLT_Win$($OSver)_sabin_latest" -SSEnvironment Dev -SSApiKey $SSDevApiKey

      # Set new key value if key exists
      $SabinSSKey = "demo" # For PSsummit23 demo only
      if ($SabinSSKey) {
        "Setting new template key value for $UpdatedProdTemplate for StackStorm Dev *** STACKSTORM Integration ONLY *** ..." | Out-file -Filepath $LogFile -Append
        #Set-SSKeyValue -KeyName $SabinSSKey.Name -NewKeyValue $WinReport.TemplateID -SSEnvironment Dev -SSApiKey $SSDevApiKey -Confirm:$false
      } else {
        "Something went wrong with attempting to update StackStorm Dev key for $UpdatedProdTemplate...Sending email and exiting..." | Out-file -Filepath $LogFile -Append
        $SendMailParams = @{
          From = $MailFrom
          To = $MailTo
          Subject = "Windows Server $OSver Template Update FAILED during StackStorm Dev Template ID Update."
          Body = "Unable to get/update StackStorm Dev Template ID!"
        }
        #Send-MailMessage @SendMailParams
        Write-Verbose "Sending email for $($SendMailParams.Subject)..."

        Write-Error "Unable to get/update StackStorm Dev Template ID!"
        Exit
      }

      $SabinSSKey = $null
      # Get StackStorm Prod Key for Prod Template
      "Gathering StackStorm Prod Key for $UpdatedProdTemplate *** STACKSTORM Integration ONLY *** ..." | Out-file -Filepath $LogFile -Append
      #$SabinSSKey = Get-SSKey -KeyName "TPLT_Win$($OSver)_sabin_latest" -SSEnvironment Prod -SSApiKey $SSProdApiKey

      # Set new key value if key exists
      $SabinSSKey = "demo" # For PSsummit23 demo only
      if ($SabinSSKey) {
        "Setting new template key value for $UpdatedProdTemplate for StackStorm Prod *** STACKSTORM Integration ONLY *** ..." | Out-file -Filepath $LogFile -Append
        #Set-SSKeyValue -KeyName $SabinSSKey.Name -NewKeyValue $WinReport.TemplateID -SSEnvironment Prod -SSApiKey $SSProdApiKey -Confirm:$false
      } else {
        "Something went wrong with attempting to update StackStorm Prod key for $UpdatedProdTemplate...Sending email and exiting..." | Out-file -Filepath $LogFile -Append
        $SendMailParams = @{
          From = $MailFrom
          To = $MailTo
          Subject = "Windows Server $OSver Template Update FAILED during StackStorm Prod Template ID Update."
          Body = "Unable to get/update StackStorm Prod Template ID!"
        }
        #Send-MailMessage @SendMailParams
        Write-Verbose "Sending email for $($SendMailParams.Subject)..."

        Write-Error "Unable to get/update StackStorm Prod Template ID!"
        #Exit
      }

      # Move Current Prod Template to Previous Release
      "Moving $CurrentProdTemplate to Previous Release Folder using Move-Template and Set-Template..." | Out-file -Filepath $LogFile -Append
      #Move-Template -Template $CurrentProdTemplate -Destination $ProdPrevReleaseTemplateDir -Confirm:$false
      #Set-Template -Template $CurrentProdTemplate -Name "$($CurrentProdTemplate.Replace('s',''))pr" -Confirm:$false

      # Remove Previous Prod Release/bkp
      "Removing Previous Prod Template from Previous Release Folder using Remove-Template..." | Out-file -Filepath $LogFile -Append
      #Remove-Template -Template $PreviousProdTemplate -Confirm:$false -DeletePermanently

      # Remove Previous Dev Release (sysprep)
      "Removing Previous Dev Sysprep Template from Dev Release Folder using Remove-Template..." | Out-file -Filepath $LogFile -Append
      #Remove-Template -Template $CurrentDevSysprepTemplate -Confirm:$false -DeletePermanently

    } # end try
    catch {

    }
    ### End: IF Test Passed, Publish Syspreped Dev Template to Prod ###

    ## Start: Publish Template to Content Library ### - Testing Failes
    try {

      # Publish Prod Template to Content Library
      "Publishing Prod Template ($UpdatedProdTemplate) to Content Library using New-ContentLibraryItem..." | Out-File -Filepath $LogFile -Append
      #New-ContentLibraryItem -Template $UpdatedProdTemplate -ContentLibrary <Your Content Library Name> -Name $UpdatedProdTemplate -Notes "Release v$(Get-Date -Format 'yyyyMM')"
      
      # Remove Previous Prod Template Release from Content Library
      "Removing Previous Prod Template Release ($CurrentProdTemplate) from Content Library Remove-ContentLibraryItem..." | Out-File -Filepath $LogFile -Append
      #Remove-ContentLibraryItem -ContentLibraryItem $CurrentProdTemplate -Confirm:$false
      
    } # end try
    catch {

    } # end catch
    ### End: Publish Template to Content Library ###

    # Publication Announcement
    "Sending Template Publication Email..." | Out-file -Filepath $LogFile -Append
    try {
        $SendMailParams = @{
          From = $WinTeam
          To = $AllRIT
          Subject = "Windows Server $OSver Template Release vP$(Get-Date -Format "yyyyMM") Published"
          Body = "Windows Server $OSver Template Release v$(Get-Date -Format "yyyyMM") has been publised to production ($UpdatedProdTemplate). Previous Template Release v$($CurrentProdTemplate) has been retired.`n`n- RIT Windows Team"
          Attachment = $LogFile
        }
        #Send-MailMessage @SendMailParams
        Write-Verbose "Sending email for $($SendMailParams.Subject)..."
    }
    catch {

    }

  } # end process block
  
  end {
    
  } # end end block
  
} # end Update-VMwareTemplates
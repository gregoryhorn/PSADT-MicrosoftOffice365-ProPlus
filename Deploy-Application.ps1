[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall')]
    [string]$DeploymentType = 'Install',
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [string]$DeployMode = 'Interactive',
    [Parameter(Mandatory = $false)]
    [switch]$AllowRebootPassThru = $false,
    [Parameter(Mandatory = $false)]
    [switch]$TerminalServerMode = $false,
    [Parameter(Mandatory = $false)]
    [switch]$DisableLogging = $false
)

Try
{
    ## Set the script execution policy for this process
    Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch { }
	
    ##*===============================================
    ##* VARIABLE DECLARATION
    ##*===============================================
    ## Variables: Application
    [string]$appVendor = 'Microsoft'
    [string]$appName = 'Office 365 Pro Plus'
    [string]$appVersion = ''
    [string]$appArch = 'X64'
    [string]$appLang = 'EN-US'
    [string]$appRevision = '01'
    [string]$appScriptVersion = '1.0.0'
    [string]$appScriptDate = '07/21/2019'
    [string]$appScriptAuthor = 'Gregory Horn'
    $OSVersion = Get-WmiObject -Class Win32_OperatingSystem
    $OSVersion = $OSVersion.Version
    $LoggedonUser = [bool] (Get-Process explorer –ea 0)
    ##*===============================================
    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)
    [string]$installName = ''
    [string]$installTitle = ''
	
    ##* Do not modify section below
    #region DoNotModify
	
    ## Variables: Exit Code
    [int32]$mainExitCode = 0
	
    ## Variables: Script
    [string]$deployAppScriptFriendlyName = 'Deploy Application'
    [version]$deployAppScriptVersion = [version]'3.6.9'
    [string]$deployAppScriptDate = '02/12/2017'
    [hashtable]$deployAppScriptParameters = $psBoundParameters
	
    ## Variables: Environment
    If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
    [string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
    ## Dot source the required App Deploy Toolkit Functions
    Try
    {
        [string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
        If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
        If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
    }
    Catch
    {
        If ($mainExitCode -eq 0) { [int32]$mainExitCode = 60008 }
        Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
        ## Exit the script, returning the exit code to SCCM
        If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
    }
	
    #endregion
    ##* Do not modify section above
    ##*===============================================
    ##* END VARIABLE DECLARATION
    ##*===============================================
		
    If ($deploymentType -ine 'Uninstall')
    {
        ##*===============================================
        ##* PRE-INSTALLATION
        ##*===============================================
        [string]$installPhase = 'Pre-Installation'
		
        ## Show Welcome Message, close all applications related to Microsoft Office 20xx
        If ($LoggedonUser -eq $True)
        {
            Show-InstallationWelcome -AllowDefer -DeferTimes 5 -CloseApps "iexplore,communicator,ucmapi,excel,groove,microsoftedge,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio" -CloseAppsCountdown 5400 -CheckDiskSpace -PersistPrompt
            Write-Log "User $env:username is logged on. Running Interactive installation."
        }

        Elseif ($runningTaskSequence -eq $True)
        {
            Write-Log "Installation is running in a Task Sequence. Running Non-Interactive installation."
        }
        Elseif ($LoggedonUser -eq $False)
        {
            Write-Log "No user is logged on to the machine. Running Non-Interactive installation."
        }
        
		
        # Uninstall Office 20xx
        If (Test-Path "$envProgramFilesX86\Microsoft Office\Office12")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2007"
            Write-Log "Microsoft Office 2007 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub07.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }

        If (Test-Path "$envProgramFiles\Microsoft Office\Office12")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2007"
            Write-Log "Microsoft Office 2007 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub07.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }

        If (Test-Path "$envProgramFilesX86\Microsoft Office\Office14")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2010"
            Write-Log "Microsoft Office 2010 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub10.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }

        If (Test-Path "$envProgramFiles\Microsoft Office\Office14")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2010"
            Write-Log "Microsoft Office 2010 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub10.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }

        If (Test-Path "$envProgramFilesX86\Microsoft Office\Office15")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2013"
            Write-Log "Microsoft Office 2013 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub13.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }

        If (Test-Path "$envProgramFiles\Microsoft Office\Office15")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2013"
            Write-Log "Microsoft Office 2013 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub13.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }

        If (Test-Path "$envProgramFilesX86\Microsoft Office\Office16")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2016"
            Write-Log "Microsoft Office 2016 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub16.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }
		
        If (Test-Path "$envProgramFiles\Microsoft Office\Office16")
        { 
            Show-InstallationProgress "Uninstalling Microsoft Office 2016"
            Write-Log "Microsoft Office 2016 was detected. Uninstalling..."
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub16.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
        }
		
        $MSOFFICE = Get-InstalledApplication -Name "Microsoft Office"
        [string] $DSIPLAYNAME = $MSOFFICE.DisplayName
        IF ($DSIPLAYNAME -match 'Microsoft Office 365 ProPlus')
        {
            Show-InstallationProgress -StatusMessage "Performing Pre-Install cleanup. $DSIPLAYNAME Was detected and should be removed before proceeding... This may take some time. Please wait...";
            Write-Log -Message "$DSIPLAYNAME Will be uninstalled" -Source "$DSIPLAYNAME uninstall";
            Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrubc2r.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3,16,42"
        }

        ##*===============================================
        ##* INSTALLATION 
        ##*===============================================
        [string]$installPhase = 'Installation'
		
        ## Handle Zero-Config MSI Installations
        If ($useDefaultMsi)
        {
            [hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
            Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
        }
		
        $LanguageMappingHt = @{
            'en-us' = 'Configuration-EN-US_X64.xml'
            'fr-fr' = 'Configuration-FR-FR_X64.xml'
            'de-de' = 'Configuration-DE-DE_X64.xml'
            'nl-nl' = 'Configuration-NL-NL_X64.xml'
        }

        #Get current Culture
        $Culture = Get-Culture
        Write-Output -InputObject "Current culture is $($Culture.Name) / $($Culture.DisplayName)"

        #Get XML file from hash table based on current culture
        #If culture is empty, default to 'en-us'
        if ($Culture.Name -match '[a-z]')
        {
            $XMLFile = $LanguageMappingHt[$Culture.Name]
            Write-Log "$($Culture.Name) detected"
        }
        else
        {
            $XMLFile = $LanguageMappingHt['en-us']
            Write-Log "$($Culture.Name) is not part of the supported language culture list. Using language culture en-US as fallback."
        }
        
        if ($XMLFile)
        {
            Show-InstallationProgress "Installing Microsoft Office 365 ProPlus - $($Culture.DisplayName)"
            Write-Log "$($Culture.Name) detected, installing Microsoft Office 365 ProPlus - $($Culture.DisplayName)"
            Execute-Process "$dirFiles\Setup.exe" -Parameters "/configure $XMLFile"
        }
        else
        {
            Write-Output -InputObject 'Current culture is not supported'
		}
		# Apply Registry Settinsg Post installation - Enable OfficeMgmtCOM, HideUpdateNotifications and disbale EnableAutomaticUpdates.
		
        Execute-Process -FilePath "cmd.exe" -Arguments "/C reg.exe add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate /v OfficeMgmtCOM /t REG_DWORD /d 1 /f" -WindowStyle Hidden
        Execute-Process -FilePath "cmd.exe" -Arguments "/C reg.exe add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate /v HideUpdateNotifications /t REG_DWORD /d 1 /f" -WindowStyle Hidden
        Execute-Process -FilePath "cmd.exe" -Arguments "/C reg.exe add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate /v EnableAutomaticUpdates /t REG_DWORD /d 0 /f" -WindowStyle Hidden
		
        ##*===============================================
        ##* POST-INSTALLATION
        ##*===============================================
        [string]$installPhase = 'Post-Installation'
		
		
        ## Display a message at the end of the install
        If (-not $useDefaultMsi)
        {
        }
    }
    ElseIf ($deploymentType -ieq 'Uninstall')
    {
        ##*===============================================
        ##* PRE-UNINSTALLATION
        ##*===============================================
        [string]$installPhase = 'Pre-Uninstallation'
		
        ## Close all applications related to Office 365
		
        If (($runningTaskSequence -eq $False) -or ($usersLoggedOn -ne $Null))
        {
            Show-InstallationWelcome -AllowDefer -DeferTimes 5 -CloseApps "iexplore,communicator,ucmapi,excel,groove,microsoftedge,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio" -CheckDiskSpace -PersistPrompt
        }
		
		
        ##*===============================================
        ##* UNINSTALLATION
        ##*===============================================
        [string]$installPhase = 'Uninstallation'
		
        ## Handle Zero-Config MSI Uninstallations
        If ($useDefaultMsi)
        {
            [hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
            Execute-MSI @ExecuteDefaultMSISplat
        }
		
        Show-InstallationProgress "Uninstalling Microsoft Office 365 ProPlus"
        Execute-Process "$dirFiles\Setup.exe" -Parameters "/configure uninstall.xml"
        
        # Scrub Microsoft Office 365 ProPlus

        Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirsupportFiles\OffScrub16.vbs`" ProPlus /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
		
		
        ##*===============================================
        ##* POST-UNINSTALLATION
        ##*===============================================
        [string]$installPhase = 'Post-Uninstallation'
		
		
		
    }
	
    ##*===============================================
    ##* END SCRIPT BODY
    ##*===============================================
	
    ## Call the Exit-Script function to perform final cleanup operations
    Exit-Script -ExitCode $mainExitCode
}
Catch
{
    [int32]$mainExitCode = 60001
    [string]$mainErrorMessage = "$(Resolve-Error)"
    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
    Exit-Script -ExitCode $mainExitCode
}
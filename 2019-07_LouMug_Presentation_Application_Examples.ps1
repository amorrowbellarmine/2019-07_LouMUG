#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '7/16/2019 9:01:24 AM'.

# Uncomment the line below if running in an environment where script signing is 
# required.
#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Site configuration
$SiteCode = "<YOUR_SITE_CODE>" # Site code 
$ProviderMachineName = "<YOUR_FQDN_SERVER>" # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams


#The examples here are simplified down to the most basic information needed. You can easily take them and add more processing to auto collect the data and fill in the properties.
#Useful Links:
## https://prajwaldesai.com/ Great how-to articles about pratically every feature of Configureation Manger. The step-by-step guide is what I used for building our environment. Part 20+ go beyond the basic setup
### https://prajwaldesai.com/sccm-2012-r2-step-by-step-guide/
## https://www.systemcenterdudes.com/category/sccm/ Another great resource of how-t arctiles and settigng up Configuration Manager. Their SQL setup is more detailed.
### https://www.systemcenterdudes.com/step-by-step-sccm-1802-upgrade-guide/
## http://www.scconfigmgr.com/ Very useful articles and tools. 
### http://www.scconfigmgr.com/driver-automation-tool/ Automate the install of drivers into Configuration Manager.
## https://docs.microsoft.com/en-us/sccm/ Microsoft's documentation of Configuration Manager
### https://docs.microsoft.com/en-us/sccm/core/plan-design/get-ready
### https://docs.microsoft.com/en-us/sccm/core/plan-design/changes/whats-new-in-version-1802 
### https://docs.microsoft.com/en-us/powershell/sccm/overview





#Building MSI Application

#Create a new Application
$application = New-CMApplication -Name "Adobe Flash Player 32 NPAPI 32.0.0.223" `
    -Publisher "Adobe Systems Incorporated" `
    -SoftwareVersion "32.0.0.223" `
    -IconLocationFile "\\FILESERVER\software\adobe\Adobe Flash Player\Flash_icon.ico"

#Add a deployment type to the Application
Add-CMMsiDeploymentType -ApplicationName "Adobe Flash Player 32 NPAPI 32.0.0.223" `
    -DeploymentTypeName "Install Flash" `
    -ContentLocation "\\FILESERVER\software\Adobe\Adobe Flash Player\Flash 32.0.0.223\install_flash_player_32_plugin.msi"

#Move the Application to the correct folder in Configuration Manager
Move-CMObject -InputObject $application `
-FolderPath "$($SiteCode):\Application\Adobe Software\Adobe Flash"

#Distribute the Application to the DPs
Start-CMContentDistribution -ApplicationName "Adobe Flash Player 32 NPAPI 32.0.0.223" `
-DistributionPointGroupName "<YOUR_DISTRIBUTION_GROUP>"

#Deploy the Application to a collection
New-CMApplicationDeployment -Name "Adobe Flash Player 32 NPAPI 32.0.0.223" `
    -CollectionName "Public Adobe Flash Update" `
    -DeployAction Install -DeployPurpose Required `
    -DeadlineDateTime "2019/7/27 02:00" `
    -AvailableDateTime "2019/7/27 00:00" `
    -TimeBaseOn LocalTime `
    -UserNotification DisplaySoftwareCenterOnly








#Building a Script Install Application

#Create a new Application
$application = New-CMApplication -Name "Firefox 68.0 ESR" `
    -Publisher "Mozilla" `
    -SoftwareVersion "68.0 ESR" `
    -IconLocationFile "\\FILESERVER\software\Firefox\firefox_icon.ico"

#Create a Detection Clause for the deployment
$clause = New-CMDetectionClauseRegistryKeyValue -Value -Hive "LocalMachine" `
    -KeyName "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox 68.0 ESR (x86 en-US)" `
    -ValueName "DisplayVersion" `
    -PropertyType "String" `
    -ExpressionOperator "IsEquals" `
    -ExpectedValue "68.0" `
    -Is64Bit:$false

#Add a deployment type to the Application
Add-CMScriptDeploymentType -ApplicationName "Firefox 68.0 ESR" `
    -DeploymentTypeName "Install Firefox" -ContentLocation "\\FILESERVER\software\Firefox\Firefox 68.0.0 ESR" `
    -InstallCommand '"Firefox Setup 68.0esr.exe" /silent' `
    -UninstallCommand "`"%PROGRAMFILES%\Mozilla Firefox\uninstall\helper.exe`" /silent" `
    -AddDetectionClause $clause `
    -MaximumRuntimeMins 20 `
    -EstimatedRuntimeMins 10 `
    -InstallationBehaviorType InstallForSystem `
    -LogonRequirementType WhereOrNotUserLoggedOn `
    -Force32Bit:$true

#Move the Application to the correct folder in Configuration Manager
Move-CMObject -InputObject $application `
-FolderPath "$($SiteCode):\Application\Mozilla Firefox"

#Distribute the Application to the DPs
Start-CMContentDistribution -ApplicationName "Firefox 68.0 ESR" `
-DistributionPointGroupName "<YOUR_DISTRIBUTION_GROUP>"

#Deploy the Application to a collection
New-CMApplicationDeployment -Name "Firefox 68.0 ESR" `
    -CollectionName "PUBLIC Firefox Update" `
    -DeployAction Install -DeployPurpose Required `
    -DeadlineDateTime "2019/7/27 02:00" `
    -AvailableDateTime "2019/7/27 00:00" `
    -TimeBaseOn LocalTime `
    -UserNotification DisplaySoftwareCenterOnly






#Added a Requriement, such as only Windows 7 Systems

#Create a new Application
$application = New-CMApplication -Name "Adobe Flash Player 32 ActiveX 32.0.0.223" `
    -Publisher "Adobe Systems Incorporated" `
    -SoftwareVersion "32.0.0.223" `
    -IconLocationFile "\\FILESERVER\software\adobe\Adobe Flash Player\Flash_icon.ico"

#Create the Requirement (PRE 1810 METHOD)
#https://thedesktopteam.com/raphael/sccm-2012-add-cmdeploymenttypeglobalcondition/
#You can get the expression base objects by creating an application the traditional GUI method, export it, and look at the object.xml file located inside the zip file under "SMS_Application\<CI Unique ID>".
    #In the object.xml file find the SDMPackageXML tag and then look at the [CDATA... information. Inside the block of XML there is a <Rule> Tag that has all the information you need.
#$ExpressionBase = new-object "Microsoft.ConfigurationManagement.DesiredConfigurationManagement.CustomCollection``1[[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.RuleExpression]]"
#$ExpressionBase.Add("Windows/All_x64_Windows_7_Client")
#$ExpressionBase.Add("Windows/All_x86_Windows_7_Client")
#
#$operator = "OneOf"
#
#$ExpressionOperator = [Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator]::$operator
#$Annotation = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Annotation
#$Annotation.DisplayName = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.LocalizableString -ArgumentList @("DisplayName", "Hello", $null)
#
#$expression = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.OperatingSystemExpression -ArgumentList @($ExpressionOperator, $ExpressionBase)
#
#$Requirement = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Rule -ArgumentList @(
#    "Rule_$([Guid]::NewGuid().ToString())",
#    [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.NoncomplianceSeverity]::None,
#    $Annotation,
#    $expression
#    )

#Create the Requirement (1810+ METHOD)
$Requirement = Get-CMGlobalCondition -Name "Operating System" | `
    Where-Object {$_.Platformtype -eq 1} | `
    New-CMRequirementRuleOperatingSystemValue -RuleOperator OneOf `
    -PlatformString @(
        "Windows/All_x64_Windows_7_Client",
        "Windows/All_x86_Windows_7_Client")

#Add a deployment type to the Application
Add-CMMsiDeploymentType -ApplicationName "Adobe Flash Player 32 ActiveX 32.0.0.223" `
    -DeploymentTypeName "Install Flash" `
    -ContentLocation "\\FILESERVER\software\Adobe\Adobe Flash Player\Flash 32.0.0.223\install_flash_player_32_active_x.msi" `
    -AddRequirement $Requirement

#Move the Application to the correct folder in Configuration Manager
Move-CMObject -InputObject $application `
-FolderPath "$($SiteCode):\Application\Adobe Software\Adobe Flash"

#Distribute the Application to the DPs
Start-CMContentDistribution -ApplicationName "Adobe Flash Player 32 ActiveX 32.0.0.223" `
-DistributionPointGroupName "<YOUR_DISTRIBUTION_GROUP>"

#Deploy the Application to a collection
New-CMApplicationDeployment -Name "Adobe Flash Player 32 ActiveX 32.0.0.223" `
    -CollectionName "Public Adobe Flash Update" `
    -DeployAction Install -DeployPurpose Required `
    -DeadlineDateTime "2019/7/27 02:00" `
    -AvailableDateTime "2019/7/27 00:00" `
    -TimeBaseOn LocalTime `
    -UserNotification DisplaySoftwareCenterOnly




#Add a Dependency to an Application
#Adobe releases updated to Acrobat Reader at MSP files which are MSI upgrade packages. To deploy these updates we need to make sure the base install of Acrobat Reader is already on the system. This is accomplished by another application package in SCCM. It was built as a standard MSI deployment.
#Adobe Acrobat Reader 2017 downloads: ftp://ftp.adobe.com/pub/adobe/reader/win/Acrobat2017/

#Create a new Application for Acrobat Reader 2017 updates.
$application = New-CMApplication -Name "Adobe Acrobat Reader 2017 MUI 17.011.30080123" `
    -Publisher "Adobe Systems Incorporated" `
    -SoftwareVersion "17.011.30080" `
    -IconLocationFile "\\FILESERVER\software\Adobe\Acrobat Acrobat Reader\Adobe Acrobat Reader 2017\AcroRd32_0000.ico"

#Create a Detection Clause for the deployment. Since this is an MSI upgrade package we can use the ProductCode for Adobe Acrobat Reader 2017. This can be found in the base install application package for Reader.
#FYI this cmdlet didn't work until release 1802.
$clause = New-CMDetectionClauseWindowsInstaller -Value `
    -ProductCode "AC76BA86-7AD7-FFFF-7B44-AE1108756300" `
    -ExpectedValue "17.011.30080" `
    -PropertyType ProductVersion `
    -ExpressionOperator isEquals

#Add a deployment type to the Application
$deployment = Add-CMScriptDeploymentType -ApplicationName "Adobe Acrobat Reader 2017 MUI 17.011.30080123" `
    -DeploymentTypeName "Install Acrobat Reader" `
    -ContentLocation "\\FILESERVER\Software\Adobe\Acrobat Acrobat Reader\Adobe Acrobat Reader 2017\Adobe Acrobat Reader 2017 Update 17.011.30080" `
    -InstallCommand 'msiexec /p "AcroRdr2017Upd1701130080_MUI.msp" /qn /norestart' `
    -UninstallCommand 'msiexec /x {AC76BA86-7AD7-FFFF-7B44-AE1108756300} /qn' `
    -AddDetectionClause $clause `
    -MaximumRuntimeMins 20 `
    -EstimatedRuntimeMins 10 `
    -InstallationBehaviorType InstallForSystem `
    -LogonRequirementType WhereOrNotUserLoggedOn

#Add a Dependency
$deployment | New-CMDeploymentTypeDependencyGroup -GroupName "Acrobat Reader Base" `
    | Add-CMDeploymentTypeDependency -DeploymentTypeDependency `
    (Get-CMDeploymentType -ApplicationName "Adobe Acrobat Reader 2017 MUI (base)" `
        -DeploymentTypeName "Adobe Acrobat Reader 2017 MUI - Windows Installer (*.msi file)") `
        -IsAutoInstall $true
    #This is the name of our base installer and deployment type for our Acrobat Reader 2017 base installer. 

#Move the Application to the correct folder in Configuration Manager
Move-CMObject -InputObject $application `
-FolderPath "$($SiteCode):\Application\Adobe Software\Adobe Acrobat Reader"

#Distribute the Application to the DPs
Start-CMContentDistribution -ApplicationName "Adobe Acrobat Reader 2017 MUI 17.011.30080123" `
-DistributionPointGroupName "<YOUR_DISTRIBUTION_GROUP>"

#Deploy the Application to a collection
New-CMApplicationDeployment -Name "Adobe Acrobat Reader 2017 MUI 17.011.30080123" `
    -CollectionName "PUBLIC Adobe Acrobat Reder Update" `
    -DeployAction Install -DeployPurpose Required `
    -DeadlineDateTime "2019/7/27 02:00" `
    -AvailableDateTime "2019/7/27 00:00" `
    -TimeBaseOn LocalTime `
    -UserNotification DisplaySoftwareCenterOnly

#Update a TaskSequence Step
#You can use PowerShell to create or update task sequence steps. In this case we are udating the step "Install Adobe Acrobat Reader" with the new update for Acrobat Reader in the task sequnce called "Test_PowerShell1"
$application = Get-CMApplication -Name "Adobe Acrobat Reader 2017 MUI 17.011.30080123"
Set-CMTSStepInstallApplication -Application $application -TaskSequenceName "Testing_PowerShell1" -StepName "Install Adobe Acrobat Reader"





#Get MSI Product info using PowerShell
#This doesn't work for Chrome. Google uses a 4-part version number, but their MSI installer has a 3-part version number. Google uses an algorithm generate their MSI version number from the product number, but the process losses information and there is no way to go from MSI to Chrome.
function Get-ProductCode {
#http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
    param(
        $location,
            [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion")]
    [string]$Property
    )
    Process {
        try {
            # Read property from MSI database
            $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
            $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($location, 0))
            $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
            $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
            $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
            $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
 
            # Commit database and close view
            $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
            $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
            $MSIDatabase = $null
            $View = $null
 
            # Return the value
            return $Value
        } 
        catch {
            Write-Warning -Message $_.Exception.Message ; break
        }
    }
    End {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}

$location = "\\FILESERVER\software\Adobe\Adobe Flash Player\Flash 32.0.0.223\install_flash_player_32_active_x.msi" 
#the function returns an array of 4 items for each property. I've found that the last item in the array is the value we're looking for.
$Name = (Get-ProductCode -Property ProductName -location $location)[-1]
$Manufacturer = (Get-ProductCode -Property Manufacturer -location $location)[-1]
$Version = (Get-ProductCode -Property ProductVersion -location $location)[-1]
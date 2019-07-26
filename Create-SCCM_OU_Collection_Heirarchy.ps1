[cmdletbinding()]
param(
    [Parameter(Mandatory=$true)][string]$OUDN,
    [Parameter(Mandatory=$true)][string]$SiteCode,
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$CMfolderpath,
    [Parameter(Mandatory=$false)][string][ValidateRange("Device","User")]$CollectionType="Device",
    [Parameter(Mandatory=$true)][string]$LimitingCollectionID
)

BEGIN{

    $startlocation = Get-Location
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
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Server @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams

    try{
        import-module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Error "ActiveDirectory module no installed. Please install AD RSAT tools first."
        break;
    }

    if ($CollectionType -eq "Device") {
        $rootfolder = $SiteCode+":\DeviceCollection\"
    } else {
        $rootfolder = $SiteCode+":\UserCollection\"
    }


    function Get-CollectionFolder {
        [cmdletbinding()]
        param(
            [Parameter(Mandatory=$true)]$OU,
            [Parameter(Mandatory=$true)]$rootfolder,
            [Parameter(Mandatory=$true)]$CMfolderpath
        )

        BEGIN{}

        PROCESS{
            $OUBreakdown = ($OU.CanonicalName -split '(?![\\])\/')
            if ($OUBreakdown.Count -gt 1) {
                $OUPath = $OUBreakdown[1..($OUBreakdown.Count-1)] -join '\'
                $OUPath = "\"+$OUPath
            } else {
                $OUPath = ""
            }
            #$OUPath
            $OUFolder = $rootfolder+$CMfolderpath+$OUPath
            return ($OUFolder)
        }

        END{}
    }

    function Create-Folder {
        param(
            [Parameter(Mandatory=$true)]$OU,
            [Parameter(Mandatory=$true)]$rootfolder,
            [Parameter(Mandatory=$true)]$CMfolderpath
        )

        $OUFolder = get-CollectionFolder -OU $OU -rootfolder $rootfolder -CMfolderpath $CMfolderpath
        if (Test-Path $OUFolder) {
            Write-Host "Folder already exists" $OUFolder
        } else{ 
            if (Check-Parent $OUFolder){
                write-host "Creating Folder" $OUFolder
                New-Item -Path $OUFolder
            } else {
                Write-Host "Parent folder for $OUFolder doesn't exist. Creating that first"
                create-folder -OU (Get-ADOrganizationalUnit -identity (Get-ADParent -dn $OU) -Properties canonicalName) -rootfolder $rootfolder -CMfolderpath $CMfolderpath
                write-host "Creating Folder" $OUFolder
                New-Item -Path $OUFolder
            }
        }
    }


    function Check-Parent {
        param (
            [Parameter(Mandatory=$true)]$OUFolder
        )

        $ParentBreakdown = $OUFolder -split '(?<![\\])\\'
        $ParentFolder = $ParentBreakdown[0..$($ParentBreakdown.Count-2)] -join '\'
        Test-Path -Path $ParentFolder
    }

    function Get-ADParent {
        param(
            [Parameter(Mandatory=$true)][string]$dn
        )

        #from https://www.uvm.edu/~gcd/2012/07/listing-parent-of-ad-object-in-powershell/
         $parts = $dn -split '(?<![\\]),'
         $parts[1..$($parts.Count-1)] -join ','
    }

    Function Create-Collection {
        param(
            [Parameter(Mandatory=$true)]$OU,
            [Parameter(Mandatory=$true)]$CollectionType
        )

        if (Get-CMDeviceCollection -Name $OU.Name){
            Write-Host $OU.name "collection already exists somewhere in SCCM."
        } else {
        
        $CMSchedule = New-CMSchedule -RecurCount 1 -RecurInterval Days -Start (Get-Date -Hour 00 -Minute 00 -Second 00 -Millisecond 000)
        $collection = New-CMCollection -CollectionType $CollectionType -LimitingCollectionId "$LimitingCollectionID" -RefreshSchedule $CMSchedule -Name $OU.Name -RefreshType Periodic
        
        if ($CollectionType -eq "Device") {
            Add-CMDeviceCollectionQueryMembershipRule -CollectionId $collection.CollectionID -RuleName $OU.Name -QueryExpression $("select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemOUName = "+'"'+$OU.canonicalName+'"')
        } else {
            $query = "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.UserOUName = `"$($OU.CanonicalName)`""
            Add-CMUserCollectionQueryMembershipRule -CollectionId $collection.CollectionID -RuleName $OU.Name -QueryExpression $query
        }
    



        $folder = get-CollectionFolder -OU $OU -rootfolder $rootfolder -CMfolderpath $CMfolderpath
        if ($CollectionType -eq "Device") {
            $coll = Get-CMDeviceCollection -Name $collection.Name
        } else {
            $coll = Get-CMUserCollection -Name $collection.Name
        }
        Move-CMObject -InputObject $coll -FolderPath $folder
    }

    }

}

PROCESS {
    
    $OUs = Get-ADOrganizationalUnit -filter * -SearchBase $OUDN -properties canonicalname

    $count = 0
    $total = $OUS.count
    foreach ($OU in $OUS) {
        $count++
        Write-Progress -Activity "Creating Collections" -PercentComplete (($count/$total)*100) -Status "Working on $($OU.Name). $count of $total complete."
        create-folder -OU $OU -rootfolder $rootfolder -CMfolderpath $CMfolderpath 
        create-collection -OU $OU -CollectionType $CollectionType
    }

    #Get-CMCollection -Name "Public Computers" 
}

END{
    
    Set-Location $startlocation

}
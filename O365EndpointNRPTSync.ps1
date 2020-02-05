<#
.SYNOPSIS
    Allows the automation of comparing and adding O365 Endpoints provided by Microsoft to your NRPT Rules
.DESCRIPTION
    Long description

.PARAMETER GPO
    Name of the GPO object you wish to upate or compare with.
.PARAMETER Category
    The Endpoint Category classification you want to use. This will filter out results that aren't of this category. If no category is provided. All Categories are returned.
    Options: [Allow][Optimize][Default]
.PARAMETER Required
    The Endpoint Required classification you want to use. This will filter out results that aren't classified as required. [Note this an old classification. Suggeted to use Category instead.]
.PARAMETER NotRequired
    The Endpoint Required classification you want to use. This will filter out results that aren't classified as "Not Required". [Note this an old classification. Suggeted to use Category instead.]
.PARAMETER Scope
    Chooses the O365 Endpoint Scope you want to return. Default is worldwide. General recommended to not change this unless you specifically know you need another option.
.PARAMETER ExportCommand
    Exports any results as the Powershell commands required to set all the NRPT rules agains the specified GPO.
.PARAMETER ExportCommandFile
    The file exported that contains the the Powershell commands required to set all the NRPT rules agains the specified GPO.
.PARAMETER DAProxyType
    Configures the DAProxy type of the NRPT rules. Options: "NoProxy", "UseDefault", "UseProxyName"
.PARAMETER NameEncoding
    Configures the NameEncoding settings used for the NRPT rules. Options: "Disable", "Utf8WithMapping", "Utf8WithoutMapping", "Punnycode"
.PARAMETER NRPTRuleComment
    Configures the Comment used when adding NRPT rules. This is also used to compare other rules that have been previously added.
.PARAMETER SYNC
    This switch triggeres teh script to create NRPT rules based off the results of any Category or Required filtering.
.PARAMETER SyncWhitelist
    This switch triggeres the script to create NRPT rules based off the WhiteList Rules.
.PARAMETER SyncCustom
    This switch triggeres the script to create NRPT rules based off any Custom Rules provided.

.EXAMPLE
    O365EndpointNRPTSync -GPO "NRPT Rules" -DAProxyType "NoProxy" -Category Allow,Optimize
.EXAMPLE
    Another example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    General notes
.COMPONENT
    The component this cmdlet belongs to
.ROLE
    The role this cmdlet belongs to
.FUNCTIONALITY
    The functionality that best describes this cmdlet
.NOTES
    Author:  Beau McMahon    
    Date:    10JAN2020
    PSVer:   2.0/3.0/4.0/5.0
    Updated: 10JAN2020
    UpdNote: 
#>


[CmdletBinding()]
param (
    
    #This will default to "$($env:TEMP)O365Endpoints.json" if not specified.
    [Parameter(Mandatory = $FALSE)]
    [ValidatePattern(".json$")]
    [string]
    $CachedEndpoints = "$($env:TEMP)\O365Endpoints.json",

    # Parameter help description
    [Parameter(Mandatory = $FALSE)]
    [string]
    $GPO = "",

    # Parameter help description
    [Parameter(Mandatory = $FALSE)]
    [object[]]
    $WhitelistURLs = ".office.com$|.office365.com$|.live.com$|.microsoft.com$|.microsoftonline.com$|.outlook.com$|.office.net$|.lync.com$|.skype.com$|.onenote.com$|.skypeforbusiness.com$|.windows.net$|.sharepoint.com$|.yammer.com$",

    # Comment to add to rules created by this script. Used to track rules to prevent duplication.
    [Parameter(Mandatory = $FALSE)]
    [string]
    $NRPTRuleComment = "Created using O365Endpoint Script",

    # Use when adding a rule
    [Parameter(Mandatory = $FALSE)]
    [ValidateSet("NoProxy", "UseDefault", "UseProxyName")]
    [string]
    $DAProxyType = "NoProxy",

    # Use when adding a rule. Remove the default setting if this isn't needed.
    [Parameter(Mandatory = $FALSE)]
    [ValidateSet("Disable", "Utf8WithMapping", "Utf8WithoutMapping", "Punnycode")]
    [string]
    $NameEncoding = "Utf8WithoutMapping",

    # Scope of the Endpoints JSON file from Microsoft to compare against.
    [Parameter(Mandatory = $FALSE)]        
    [string]
    $Scope = "worldwide",

    # Name of CSV file to export for rules that already exsits.
    [Parameter(Mandatory = $FALSE)]
    [string]
    $PreExistingRulesCSV, #= "PreExistingRules.csv",

    #Import JSON file if you can't access the internet from the machine.
    [Parameter(Mandatory = $FALSE)]
    [string]
    $ImportJSON,

    # Enables the GPO to be updated using the results outputted by the script.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $Sync,

    # Syncs the NRPT rules with the local machine instead of the GPO.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $Local,

    #Enables the GPO to be updated with the Whitelist URLS.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $SyncWhiteList,

    #Enables the GPO to be updated with any URL entered
    [Parameter(Mandatory = $FALSE)]
    [object[]]
    $SyncCustomURL,

    # Choose what list of rules you want to compaire the O365 Endpoints with
    [Parameter(Mandatory = $FALSE)]
    [ValidateSet('WhiteList', "CurrentRules", "Both")]
    [string]
    $Compare,

    # Filters the results retured based off the Category field
    [Parameter(Mandatory = $FALSE)]
    [ValidateSet("Optimize", "Allow", "Default")]
    [object[]]
    $Category,

    # Adds a column to the returned results and filters them to show what's matching or missing.
    [Parameter(Mandatory = $FALSE)]
    [ValidateSet("Match", "Missing")]
    [string]
    $MatchOrMissing,

    # Filter returned result that are required.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $Required,
        
    # Filter returned result that are not required.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $NotRequired,

    # Filename to save report file as
    [Parameter(Mandatory = $FALSE)]
    [string]
    $ReportFile,
            
    # Switch forces download of Endpoints JSON file.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $NoCache,

    # Switched returns the Powershell Commands what would be run.
    [Parameter(Mandatory = $FALSE)]
    [switch]
    $ExportCommand,

    # Saves the Powershell Commands what would be run to a specific file.
    [Parameter(Mandatory = $FALSE)]
    [string]
    $ExportCommandFile,

    #Log file to record updates to NRPT Rules. This is fairly basic logging atm. Only adds an entry when the script creates an entry itself. (When the -Sync* and no -Export parameter is used)
    [string]
    $Log

)

BEGIN
{
    
    Function ArraytoRegex
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $TRUE)]    
            [AllowNull()]            
            [object[]]
            $array,
            
            # Parameter help description
            [Parameter(Mandatory = $FALSE)]
            [switch]
            $Wildcards
        )

        
        IF ($array)
        {
            IF ($Wildcards)
            {
                return (($array.split(",;") -replace "^\.", "*.") -join "$|").trimend("$") + "$"  
            }
            ELSE { return ($array.split(",;").trimstart("*") -join "$|").trimend("$") + "$" }
        }

        
    }

    Function RegexToArray
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $TRUE)]
            [AllowNull()][AllowEmptyString()][AllowEmptyCollection()]                
            [object[]]                
            $string,

            # Parameter help description
            [Parameter(Mandatory = $FALSE)]
            [switch]
            $Wildcards
        )

        IF ($string)
        {
            IF ($Wildcards)
            {
                return $string.split("|").trimend("$") -replace "^\.", "*."
            }
            ELSE { return $string.split("|").trimend("$") }
        }
    }
    
    IF ($ImportJSON)
    {
        IF (Test-Path $ImportJSON) { $Endpoints = Get-Content $ImportJSON | ConvertFrom-Json }
        ELSE { Throw "Unable to find/access:$ImportJSON" }          
    }
    ELSE
    {
        $clientRequestId = [GUID]::NewGuid().Guid
        
        #Check Endpoints Version. Used to cache Endpoints JSON file.
        IF ($NoCache -eq $FALSE)
        {
            $EndpointsVersion = ((Invoke-RestMethod "https://endpoints.office.com/version?clientrequestid=$($clientRequestId)") | where-object { $_.instance -eq $Scope }).latest            
        }
        IF ((Test-Path "$(($CachedEndpoints).replace(".json","_$EndpointsVersion.json"))") -and ($NoCache -eq $FALSE))
        {
            $Endpoints = Get-Content "$(($CachedEndpoints).replace(".json","_$EndpointsVersion.json"))" | ConvertFrom-Json
        }
        ELSE
        {
            [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            $URI = "https://endpoints.office.com/endpoints/$($scope)?clientrequestid=$($clientRequestId)"

            try { $Endpoints = Invoke-RestMethod $URI -ErrorAction stop }catch { Throw "Unable to access Endpoint URL:$URI" }
            $Endpoints | ConvertTo-Json | Set-Content "$(($CachedEndpoints).replace(".json","_$EndpointsVersion.json"))"
        }
    }

    
    #Calculate OutputDir var
    IF ($PSScriptRoot)
    { 
        #Save to the directory where the script is located
        $OutputDir = "$($PSScriptRoot)" 
    }
    ELSEIF ((Get-Location).Provider.tostring() -eq "Microsoft.PowerShell.Core\FileSystem")
    {
        #IF PSScriptRoot is null, check that the current DIR is actually a filesystem before using it
        $OutputDir = "$(((Get-Location).path))" 
    }
    ELSE
    {
        #Worst case scenario, save to systemdrive (should be C:) root
        $OutputDir = "$($env:systemdrive)"
    }



    #Detect Whitelist format and convert to Regex format
    $WhitelistURLs = ArraytoRegex $WhiteListURLs
    

    #Remove any "*" characters.
    #$WhitelistURLs = $WhitelistURLs.replace("*","") #ArraytoRex does this now


    

    
    
    

}
PROCESS
{

    #########################
    # Cache information and prep arrays        
    #########################
    

    #Prepare Arrays Variables
    $UseLocalInternet = @()
    $PreExistingRules = @()

    #Create NRPT Rule Splatt Template
    $NRPTRuleSplattTemplate = @{
                    
        NameSpace = ""
        Comment   = "$NRPTRuleComment"
        DAEnable  = $TRUE
    }
    IF ($Local -ne $TRUE) { $NRPTRuleSplattTemplate += @{GpoName = "$GPO" } }
    IF ($DAProxyType) { $NRPTRuleSplattTemplate += @{DAProxyType = $DAProxyType } }
    IF ($NameEncoding) { $NRPTRuleSplattTemplate += @{NameEncoding = $NameEncoding } }  


    #Cache the current rules to compare against
    IF ($Compare -match "CurrentRules|Both" -or $Sync -or $SyncWhiteList -or $SyncCustomURL)
    {   
        try
        {
            IF ($Local) { $CurrentRules = Get-DnsClientNrptRule -ErrorAction Stop }
            ELSE { $CurrentRules = Get-DnsClientNrptRule -GpoName $GPO -ErrorAction Stop }
            
            IF ($null -eq $CurrentRules -or $CurrentRules.count -eq 0)
            {
                Write-Warning "No NRPT rules found in GPO `"$GPO`". There are no current rules to filter."
            }
        
        }
        catch [Microsoft.Management.Infrastructure.CimException]
        {
            IF ($_.Exception.ErrorData.ErrorSource -eq "NRPTRule" -and $_.Exception.ErrorData.error_Code -eq 1168)  
            {
                Throw "Unable to find GPO provided. Please check that the GPO name `"$GPO`" is valid."    
            }
            ELSEIF($_.Exception.ErrorId -eq "ParameterArgumentValidationError")
            {
                Throw "GPO Name `"$GPO`" is null or invalid."
            }
        }

        
        #TODO:This can probably be removed.
        #IF ($null -eq $CurrentRules) { Write-Warning "Get-DNSClientNRPTRule command returned no results. Comparing rules to the CurrentRules may fail." }
    }
    
    ########################################
    ########################################
    
    #Check the ExportCommandFile path is valid
    try
    {
        IF (([string]::IsNullOrEmpty($ExportCommandFile) -eq $FALSE))
        {
            IF ((Test-Path $ExportCommandFile) -eq $FALSE) { New-Item $ExportCommandFile -ItemType File -ErrorAction STOP | out-null }
        } 
    }
    catch
    {
        Write-Warning "Error using parameter [-ExportCommandFile] with value of `"$ExportCommandFile`""; Throw $_
    }


    #Check the Log path is valid. Add headers if it doesn't exist.
    try
    {
        IF (([string]::IsNullOrEmpty($Log) -eq $FALSE))
        {
            IF ((Test-Path $Log) -eq $FALSE)
            {
                New-Item $Log -ItemType File -ErrorAction STOP | out-null
                #Add Header to Log file
                Add-Content $Log "Date,Action,URL,Details"
            }
        } 
    }
    catch
    {
        Write-Warning "Error using parameter [-Log] with value of `"$Log`""; Throw $_
    }
    
    
    
    #Process any reporting or comparing parameters.
    IF ($Required -or $NotRequired -or $Category -or $Compare)
    {
        
        IF ($Compare -eq "WhiteList") { $CompareList = $WhitelistURLs }            
        
        ELSEIF ($Compare -eq "CurrentRules") { $CompareList = ArraytoRegex ($CurrentRules.namespace) }
        ELSEIF ($Compare -eq "Both") { $CompareList = ArraytoRegex (($CurrentRules.namespace) + $WhitelistURLs) }
        #ELSEIF($Report -ne "Required" -and $Report -ne "NotRequired"){Throw "Need to provided a valid [-ReportCompare] parameter if you want to check Matching or Missing URL's."}

        IF ((RegexToArray $CompareList).count -eq 0)
        {
            #Add the following entry to give the Array at least one entry. This won't ever match because "<>" are invalid URL characters.
            #Required because if the CompareList Array is empty, it inverts the Match or Missing results.
            $CompareList = "<>"
        }


        #Filter ReportResult to limit specified Category or Requred/NotRequired
        $ReportResult = $Endpoints
        IF ($Category)
        {
            $ReportResult = ForEach ($Cat in $Category) { $ReportResult | where-object { $_.category -eq $Cat } }
        }
        
        IF ($Required) { $ReportResult = $ReportResult | where-object { $_.required -eq $TRUE } }
        IF ($NotRequired) { $ReportResult = $ReportResult | where-object { $_.required -eq $FALSE } }

        
        #Filter results if that Match or Missing from the results you are comparing.
        IF ($MatchOrMissing)            
        {
            IF ($Compare)
            {                
                IF ($MatchOrMissing -eq "Match")
                {                        
                    $ReportResult = $ReportResult | select-object @{N = "$($MatchOrMissing)Entries"; E = { ($_.urls | where-object { $_ -match $CompareList }) -join "," } }, * | where-object { $_.MatchEntries -ne "" }
                }
                IF ($MatchOrMissing -eq "Missing")
                {                     
                    $ReportResult = $ReportResult | select-object @{N = "$($MatchOrMissing)Entries"; E = { ($_.urls | where-object { $_ -notmatch $CompareList }) -join "," } }, * | where-object { $_.MissingEntries -ne "" }
                }
            }
            ELSE { Throw "Need to provided a valid [-ReportCompare] parameter if you want to check Matching or Missing URL's." }
        }
        ELSEIF (-not ($MatchOrMissing) -and $Compare) { Throw "[-Compare] parameter will be ignored because no [-MatchOrMissing] parameter is provided." }
        
        <# - This has now been moved towards the bottom of the script

        #This will export the results to a CSV if the ReportFile parameter is passed. If no sync parameter is passed, it will also return the results of the report to the output.
        #If a sync parameter is passed, the results won't be returned because this interferes with adding the rules.
        IF ($ReportFile) { $ReportResult | select-object @{N = "urls"; E = { $_.urls -join " " } }, * -ExcludeProperty urls | Export-Csv -NoTypeInformation -Path "$($OutputDir)\$($ReportFile)" }
        ELSEIF (($PSBoundParameters.Keys -match "^sync").count -eq 0)
        {
            IF($ExportCommand){ Write-Warning "The [-ExportCommand] parameter will be ignored because no sync parameter is provided." }
            return $ReportResult } #* This is important because if return is used here, the reset of the script is ignored.
        
            #>
        
        
    }
    ELSE { $ReportResult = $Endpoints }#Set the report Results here, this is need in case Require,NotRequire or Category isn't passed.


    #Add Whitelisted urls to the NRPT rules
    IF ($SyncWhiteList)
    {
        
        #$WhitelistURLs.split("|").TrimEnd("$") | ForEach-Object { $UseLocalInternet += $_ } 
        RegexToArray $WhitelistURLs | ForEach-Object { $UseLocalInternet += $_ } 
    }

    #Add all urls returned results to the NRPT rules
    IF ($Sync)
    {
        IF ($Required -or $NotRequired -or $Category)
        {
            ForEach ($URL in $ReportResult)
            {
                IF ($SyncWhiteList) { $URL.urls | where-object { $_ -notmatch $WhitelistURLs } | ForEach-Object { $UseLocalInternet += $_ } }
                ELSE { $URL.urls | ForEach-Object { $UseLocalInternet += $_ } }
            }
        }
        ELSE { Throw "No parameter provided that generates a report. Syncing from report requires [-Required],[-NotRequired], or [-Category] parameter." }
    }
    
    #Add any Custom URL's added into the command line to the NRPT rules. Accepts object[] or string[]
    IF ($SyncCustomURL)
    {
        IF ($SyncWhiteList) { $SyncCustomURL | where-object { $UseLocalInternet -notcontains $_ -and $_ -notmatch $WhitelistURLs } | ForEach-Object { $UseLocalInternet += $_ } }
        ELSE { $SyncCustomURL | where-object { $UseLocalInternet -notcontains $_ } | ForEach-Object { $UseLocalInternet += $_ } }
    }

    
    

    #Clean UseLocalInternet variable. Seems to get an empyt entry that causes issues later
    $UseLocalInternet = $UseLocalInternet | where-object { $_ -ne $null }

    

    
    #Proccess the URLS if a Sync parameter has been used. IF no Sync parameter has been used, it will pass to the ELSEIF and return the report results if any.
    #If the ExportCommand parameter has been used, then the output will be the powershell command required to create that entry for each URL.
    IF ($UseLocalInternet.count -ge 1)
    {
        
        IF ($PreExistingRulesCSV)
        {
            #Convert $UseLocalInternet to useable Regex var
            #$UseLocalInternetRegex = ($UseLocalInternet -join "$|").TrimStart("$|").replace("*", "")
            $UseLocalInternetRegex = ArraytoRegex $UseLocalInternet                

            
            #Check for preexisting rules. Save any that already exist but don't have a matching commment.
            ForEach ($URL in $CurrentRules | where-object { $_.namespace -match $UseLocalInternetRegex -and $_.comment -ne $NRPTRuleComment })
            {
                $PreExistingRules += $URL    
            }

            #Export CSV of PreExistingRules
            #Updates CSV output variable with either the root of the script or the current directory if not running as a script.    
            
            IF ($PreExistingRulesCSV -match ".csv$") { $PreExistingRulesCSV = $PreExistingRulesCSV -replace ".csv$", "_$((get-date -Format yyyyMMddHHmm)).csv" }
            ELSE { $PreExistingRulesCSV = "$($PreExistingRulesCSV)_$((get-date -Format yyyyMMddHHmm)).csv" }
            
            #Generate full path for csv
            $PreExistingRulesCSV = "$($OutputDir)\$($PreExistingRulesCSV)"

            #Export Rules that already exists to CSV
            try
            {
                $PreExistingRules | select-object *, @{N = "Namespace"; E = { $_.namespace -join " " } } -ExcludeProperty Namespace | Export-Csv -NoTypeInformation -Path $PreExistingRulesCSV;
                #Purposly using Write-Host here as I don't want this returned if results are passed to a variable. This is information only.
                Write-Host "Pre Existing Rules CSV saved: $PreExistingRulesCSV"; 
            }
            catch { Write-Output "Failed to export Alread Existing Rules CSV." }
        }
                    
        
        #Add Rule for URL if it's not found in the CurrentRules. Trim off any leading wildcard characters as these aren't used.
        ForEach ($URL in $UseLocalInternet | where-object { $CurrentRules.namespace -notcontains $_ })
        {                
            #Copy the NRPTRuleSplattTemplate. Need to use the .Copy() or else any changes to Params changes the template.
            $Params = $NRPTRuleSplattTemplate.psobject.copy()
            $Params.NameSpace = "$URL"
            
            IF ($ExportCommand -or ([string]::IsNullOrEmpty($ExportCommandFile) -eq $FALSE))
            {
                #This will convert the $Params splat to the equivilant powershell command (e.g. Add-DNSClientNRPTRule -Namespace "office.com" -comment "O365 Whitelist" -DAEnable "TRUE")
                IF ($ExportCommandFile) { Add-Content $ExportCommandFile (@("Add-DNSClientNRPTRule", (($Params.psobject.BaseObject.keys | % { "-$_ `"$($Params.$_)`"" }) -join " ")) -join " ") }
                ELSE { Write-Output (@("Add-DNSClientNRPTRule", (($Params.psobject.BaseObject.keys | % { "-$_ `"$($Params.$_)`"" }) -join " ")) -join " ") }
                
            }
            ELSE
            {
                Write-Host "Adding NRPT rule for $URL" -ForegroundColor Green
                Add-DnsClientNrptRule @Params
                IF ($LOG) { Add-Content $Log ("$((Get-Date -Format u)),UPDATE RULES,$($URL)," + (@("Add-DNSClientNRPTRule", (($Params.psobject.BaseObject.keys | % { "-$_ `"$($Params.$_)`"" }) -join " ")) -join " ")) }
                
            }               


        }
        
        IF (([string]::IsNullOrEmpty($ExportCommandFile) -eq $FALSE)) { Write-Host "Exported Powershell commands to: $((Get-Item $ExportCommandFile).Fullname)" }
        
    }        
    ELSEIF (($ReportResult.count -ge 1) -and (($PSBoundParameters.Keys -match "^sync").count -eq 0)) #If a sync param is provided, then this shouldn't run.
    {
        #This will export the results to a CSV if the ReportFile parameter is passed. If no sync parameter is passed, it will also return the results of the report to the output.            
        IF ($ReportFile) { $ReportResult | select-object @{N = "urls"; E = { $_.urls -join " " } }, * -ExcludeProperty urls | Export-Csv -NoTypeInformation -Path "$($OutputDir)\$($ReportFile)" }
        IF ($ExportCommand)
        {
            ForEach ($URL in $ReportResult.urls | Select-Object -Unique)
            {
                #Copy the NRPTRuleSplattTemplate. Need to use the .Copy() or else any changes to Params changes the template.
                $Params = $NRPTRuleSplattTemplate.psobject.copy()
                $Params.NameSpace = "$URL"
                

                #Write-Output (@("Add-DNSClientNRPTRule", (($Params.psobject.BaseObject.keys | % { "-$_ `"$($Params.$_)`"" }) -join " ")) -join " ")

                IF (([string]::IsNullOrEmpty($ExportCommandFile) -eq $FALSE)) { Add-Content $ExportCommandFile (@("Add-DNSClientNRPTRule", (($Params.psobject.BaseObject.keys | % { "-$_ `"$($Params.$_)`"" }) -join " ")) -join " ") }
                ELSE { Write-Output (@("Add-DNSClientNRPTRule", (($Params.psobject.BaseObject.keys | % { "-$_ `"$($Params.$_)`"" }) -join " ")) -join " ") }
            }

            IF (([string]::IsNullOrEmpty($ExportCommandFile) -eq $FALSE)) { Write-Host "Exported Powershell commands to: $((Get-Item $ExportCommandFile).Fullname)" }


        }
        ELSE
        {
            
            return $ReportResult
        }
    }

}

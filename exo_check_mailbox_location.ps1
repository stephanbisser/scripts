<#  
.SYNOPSIS  
    Check storage location of each ExO mailbox  
.DESCRIPTION  
    This script gathers information about the storage location of all mailboxes or for a specific in Exchange Online, in order to check which mailboxes is stored in which datacenter
.PARAMS
    -username (OPTIONAL): If not already connected to Exchange Online, this paramater must be specified with a valid ExO admin username
    -password (OPTIONAL): If not already connected to Exchange Online, this paramater must be specified with a valid ExO admin password
    -primarySMTPAddress (OPTIONAL): This parameter can be specified in order to check the location of a specific mailbox
.NOTES
    File Name  : exo_check_mailbox_location.ps1  
    Author     : Stephan Bisser - stephan@bisser.at - blog.bisser.at
    Date       : 07.04.2017
    Version    : 1.0
.Reference
    The first part with the datacenter abbreviations were taken from https://gallery.technet.microsoft.com/PowerShell-Script-to-a6bbfc2e (
    Joseph Palarchio) - Thanks for that!
#>

param(
    [Parameter(Mandatory=$false)][string]$username,
    [Parameter(Mandatory=$false)][string]$password,
    [Parameter(Mandatory=$false)][string]$primarySMTPAddress
)

### Login to Exchange Online PowerShell
Function LoginExO
{    
    $pwd = ConvertTo-SecureString $password -AsPlainText -Force
    $credentials = New-Object System.Management.Automation.PSCredential $username, $pwd
    Get-PSSession | Remove-PSSession
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
	Import-PSSession $Session
    . Main
}


Function Main
{
    $Datacenter = @{}
    $Datacenter["CP"]=@("LAM","Brazil")
    $Datacenter["GR"]=@("LAM","Brazil")
    $Datacenter["HK"]=@("APC","Hong Kong")
    $Datacenter["SI"]=@("APC","Singapore")
    $Datacenter["SG"]=@("APC","Singapore")
    $Datacenter["KA"]=@("JPN","Japan")
    $Datacenter["OS"]=@("JPN","Japan")
    $Datacenter["TY"]=@("JPN","Japan")
    $Datacenter["AM"]=@("EUR","Amsterdam, Netherlands")
    $Datacenter["DB"]=@("EUR","Dublin, Ireland")
    $Datacenter["HE"]=@("EUR","Helsinki, Finland")
    $Datacenter["VI"]=@("EUR","Vienna, Austria")
    $Datacenter["BL"]=@("NAM","Virginia, USA")
    $Datacenter["SN"]=@("NAM","San Antonio, Texas, USA")
    $Datacenter["BN"]=@("NAM","Virginia, USA")
    $Datacenter["DM"]=@("NAM","Des Moines, Iowa, USA")
    $Datacenter["BY"]=@("NAM","San Francisco, California, USA")
    $Datacenter["CY"]=@("NAM","Cheyenne, Wyoming, USA")
    $Datacenter["CO"]=@("NAM","Quincy, Washington, USA")
    $Datacenter["MW"]=@("NAM","Quincy, Washington, USA")
    $Datacenter["CH"]=@("NAM","Chicago, Illinois, USA")
    $Datacenter["ME"]=@("APC","Melbourne, Victoria, Australia")
    $Datacenter["SY"]=@("APC","Sydney, New South Wales, Australia")
    $Datacenter["KL"]=@("APC","Kuala Lumpur, Malaysia")
    $Datacenter["PS"]=@("APC","Busan, South Korea")
    $Datacenter["YQ"]=@("CAN","Quebec City, Canada")
    $Datacenter["YT"]=@("CAN","Toronto, Canada")
    $Datacenter["MM"]=@("GBR","Durham, England")
    $Datacenter["LO"]=@("GBR","London, England")
    
    ### Check if a primarySMTPAddress was specified
    if ($primarySMTPAddress.Length -eq 0)
    {
        $mailboxes= Get-Mailbox -ResultSize Unlimited | Where {$_.RecipientTypeDetails -ne "DiscoveryMailbox"} | select Name,PrimarySMTPAddress,ServerName
    } else 
    {
        $mailboxes= Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySMTPAddress -eq $primarySMTPAddress} | select Name,PrimarySMTPAddress,ServerName
    }

    ### Create the output object with all necessary information
    $Output = @()
    foreach ($mailbox in $mailboxes){
        $outputObj = New-Object -TypeName PSObject
        $outputObj | Add-Member -Name 'Name' -MemberType NoteProperty -Value $mailbox.Name
        $outputObj | Add-Member -Name 'PrimarySMTPAddress' -MemberType NoteProperty -Value $mailbox.PrimarySmtpAddress
        $ServerNameShort = $mailbox.ServerName.Substring(0,2).toUpper()
        $outputObj | Add-Member -Name 'ServerName' -MemberType NoteProperty -Value $ServerNameShort
        $Output += $outputObj
    }

    ### Get the necessary information and put it together for displaying it
    $info = @()
    foreach($obj in $Output){
        $center = $obj.ServerName
        $location = $Datacenter[$center][1]
        $region = $Datacenter[$center][0]
        $datacenterInfo = New-Object -TypeName PSObject
        $datacenterInfo | Add-Member -Name 'Name' -MemberType NoteProperty -Value $obj.Name
        $datacenterInfo | Add-Member -Name 'PrimarySMTPAddress' -MemberType NoteProperty -Value $obj.PrimarySMTPAddress
        $datacenterInfo | Add-Member -Name 'Datacenter region' -MemberType NoteProperty -Value $region
        $datacenterInfo | Add-Member -Name 'Datacenter location' -MemberType NoteProperty -Value $location
        $info += $datacenterInfo
    }

    ### Output the necessary information about the Name of the mailbox, the datacenter region and the datacenter location
    $info | Format-Table
}

Function StartScript
{
    if ($username.Length -ne 0 -and $password.Length -ne 0)
    {
        Write-Information "Login to Exchange Online will now be done"
        . LoginExO
    } else 
    {
        Write-Information "No login will be triggered as credentials may not be passed or are invalid"
        . Main
    }
}

. StartScript
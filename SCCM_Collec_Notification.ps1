<#
.DESCRIPTION
    This script send mail to SCCM collection members
.NOTES
    Prerequisites : 
        - Notepad++
        - Microsoft edge
        - SCCM Console
#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$True)][string]$TemplateName,
    [Parameter(Mandatory=$False)][string]$CollectionName,
    [Parameter(Mandatory=$False)][string]$TitleMail,
    [Parameter(Mandatory=$False)][ValidateSet("Information", "Warning", "Alert")][string]$MailType
)

# Variables
## Path 
$CurrentDir = Get-Location
$TemplatePath = "$CurrentDir\Templates"
$HtmlPath = "$CurrentDir\Sources\HTML"

## SCCM 
$SiteCode = "XXX"
$ProviderMachineName = "SCCM_Server.domaine.com"

## Other
$AD_Server = "ServerName.domain.com"
$SmtpServer = "ServerName.domain.com"
$EmailSender = "youremailaddress@domain.com"
$Date = Get-Date -Format "yyyyMMdd HH-mm"

## Array
$ArrayMembers = @()


# Functions
function BuildTemplate {
    # open body.html file
    Write-Host -fore yellow "Please enter the content of your email:"
    start notepad++ "$TemplatePath\$TemplateName\body.html"
    pause

    # Reset Template.html
    Copy-Item -Path "$HtmlPath\Template.html" -Destination "$TemplatePath\$TemplateName" -Recurse -Force

    ## Insert body.html in template file
    $Body = (Get-Content -path "$TemplatePath\$TemplateName\body.html" -Raw)
    (Get-Content -path "$TemplatePath\$TemplateName\Template.html" -Raw) -replace 'XXXBODYXXX',$Body | Set-Content -Path "$TemplatePath\$TemplateName\Template.html" -Encoding UTF8

    ## Insert title in template file
    (Get-Content -path "$TemplatePath\$TemplateName\Template.html" -Raw) -replace 'XXXTITLEXXX',$TitleMail | Set-Content -Path "$TemplatePath\$TemplateName\Template.html" -Encoding UTF8

    ## Insert color in template file
    if($MailType -eq "Information"){
        $Color = "#064fb6"
    }elseif($MailType -eq "Warning"){
        $Color = "#f39c12"
    }elseif($MailType -eq "Alert"){
        $Color = "#b60606"
    }

    (Get-Content -path "$TemplatePath\$TemplateName\Template.html" -Raw) -replace 'UUUCOLORUUU',$Color | Set-Content -Path "$TemplatePath\$TemplateName\Template.html" -Encoding UTF8


    ## check html file
    Write-Verbose "Check HTML file"
    &'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' @("$TemplatePath\$TemplateName\Template.html")
}

function New-SCCMSession {
    $initParams = @{}
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root ProviderMachineName @initParams
    }

    Set-Location "$($SiteCode):"
}

function Send_Email{
    param(
        [Parameter(Mandatory=$true)][string]$destinataire
    )
    $TemplateHtml = "$TemplatePath\$TemplateName\Template.html"
    $TemplateMail = (Get-Content -path $TemplateHtml -Raw) -replace 'XXXUSERNAMEXXX',$UserName -replace 'YYYCOMPUTERNAMEYYY',$ComputerName

    $encodingMail = New-Object System.Text.utf8encoding
    $options = @{
        'SmtpServer' = $SmtpServer 
        'To' = $destinataire
        # 'bcc' = "youremail.domain.com"
        'From' = $EmailSender
        'Subject' = $TitleMail
        'Body' = $TemplateMail
        'Encoding' = $encodingMail
        'Priority' = "High"

    }

    Send-MailMessage @options -BodyAsHtml

}


$TestPath = Test-Path "$TemplatePath\$TemplateName"
if($TestPath -eq $False){
    ## Create folder
    New-item -path $TemplatePath -Name $TemplateName -itemType "directory"

    ## Copy template files
    Copy-Item -Path "$HtmlPath\*" -Destination "$TemplatePath\$TemplateName" -Recurse

    Do{
        BuildTemplate

        $CheckHtml = Read-host "Your template is okay : y/n ?"
    }Until($CheckHtml -eq "y" -or $CheckHtml -eq "Y")



}else{
    $TemplateUpdate = Read-host "Do you want update template file ? y/n"
    if($TemplateUpdate -eq "y" -or $TemplateUpdate -eq "Y"){
        Do{
            BuildTemplate
    
            $CheckHtml = Read-host "Your template is okay : y/n ?"
        }Until($CheckHtml -eq "y" -or $CheckHtml -eq "Y")
    }else{
        Write-Verbose "[ERROR] - $TemplateName already existe !"
        exit
    }
    
}

New-SCCMSession

$CheckCollection = [bool](Get-CMCollection -Name $CollectionName)
if($CheckCollection -eq $True){
    $GetMembers = Get-CMCollectionMember -CollectionName $CollectionName
    cd D:\

    foreach ($Item in $GetMembers) {
        $details = @()
        $SamAccountName = $Item.username
        $ComputerType   = $Item.DeviceOS
        $ComputerName   = $Item.Name
        
        if($ComputerType -like "*Workstation*"){
            if($SamAccountName){
                $CheckUserAD = [Bool](Get-ADUser -Filter {SamAccountName -eq $SamAccountName} -Server $AD_Server -ErrorAction SilentlyContinue)
                if($CheckUserAD -eq $True){
                    $GetAdInfo = Get-ADUser $SamAccountName -Properties *
                    $EmailUser = $GetAdInfo.mail
                    if($EmailUser){
                        $UserName = $GetAdInfo.GivenName
                    
                        $details = [ordered]@{            
                            Username	    = $SamAccountName
                            email	        = $EmailUser
                            Computername	= $ComputerName
                        }     
                        $ArrayMembers += New-Object PSObject -Property $details 
                
                        Send_Email -destinataire $EmailUser
                        
                    }else{
                        Write-Warning -Message "Email not found for $Samaccountname"
                    }
                }else{
                    Write-Verbose "[ERROR] - '$SamAccountName' not found in AD"
                }
            }else{
                Write-Verbose "[ERROR] - User not found for $ComputerName"
            }
        }else{
            Write-Warning -Message "Server found, ignore it."
        }
    }

    ## Export members
    $ArrayMembers | Export-CSV -Path "$TemplatePath\$TemplateName\ExportInfo-$($Date).csv" -NoTypeInformation -delimiter ';' -Encoding 'UTF8'

}else{
    Write-Verbose "[ERROR] - Collection not found in SCCM"
    exit
}

Write-Verbose "### End script ###"

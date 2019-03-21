# Login to Office 365 MSOnline first
$Theme = Get-UDTheme Azure
$Company = 'Template'

$Sku = @{
    'DESKLESSPACK' = 'Office 365 (Plan K1)'
    'DESKLESSWOFFPACK' = 'Office 365 (Plan K2)'
    'FLOW_FREE' = 'Flow Free'
    'LITEPACK' = 'Office 365 (Plan P1)'
    'EXCHANGESTANDARD' = 'Office 365 Exchange Online Only'
    'STANDARDPACK' = 'Enterprise Plan E1'
    'STANDARDWOFFPACK' = 'Office 365 (Plan E2)'
    'ENTERPRISEPACK' = 'Enterprise Plan E3'
    'ENTERPRISEPACKLRG' = 'Enterprise Plan E3'
    'ENTERPRISEWITHSCAL' = 'Enterprise Plan E4'
    'STANDARDPACK_STUDENT' = 'Office 365 (Plan A1) for Students'
    'STANDARDWOFFPACKPACK_STUDENT' = 'Office 365 (Plan A2) for Students'
    'ENTERPRISEPACK_STUDENT' = 'Office 365 (Plan A3) for Students'
    'ENTERPRISEWITHSCAL_STUDENT' = 'Office 365 (Plan A4) for Students'
    'STANDARDPACK_FACULTY' = 'Office 365 (Plan A1) for Faculty'
    'STANDARDWOFFPACKPACK_FACULTY' = 'Office 365 (Plan A2) for Faculty'
    'ENTERPRISEPACK_FACULTY' = 'Office 365 (Plan A3) for Faculty'
    'ENTERPRISEWITHSCAL_FACULTY' = 'Office 365 (Plan A4) for Faculty'
    'ENTERPRISEPACK_B_PILOT' = 'Office 365 (Enterprise Preview)'
    'STANDARD_B_PILOT' = 'Office 365 (Small Business Preview)'
    'VISIOCLIENT' = 'Visio Pro Online'
    'POWER_BI_ADDON' = 'Office 365 Power BI Addon'
    'POWER_BI_INDIVIDUAL_USE' = 'Power BI Individual User'
    'POWER_BI_STANDALONE' = 'Power BI Stand Alone'
    'POWER_BI_STANDARD' = 'Power-BI standard'
    'PROJECTESSENTIALS' = 'Project Lite'
    'PROJECTCLIENT' = 'Project Professional'
    'PROJECTONLINE_PLAN_1' = 'Project Online'
    'PROJECTONLINE_PLAN_2' = 'Project Online and PRO'
    'ECAL_SERVICES' = 'ECAL'
    'EMS' = 'Enterprise Mobility Suite'
    'RIGHTSMANAGEMENT_ADHOC' = 'Windows Azure Rights Management'
    'MCOMEETADV' = 'PSTN conferencing'
    'SHAREPOINTSTORAGE' = 'SharePoint storage'
    'PLANNERSTANDALONE' = 'Planner Standalone'
    'CRMIUR' = 'CMRIUR'
    'BI_AZURE_P1' = 'Power BI Reporting and Analytics'
    'INTUNE_A' = 'Windows Intune Plan A'
}

function ConvertLicenseString {
    $String = $args[0] -Replace 'reseller-account:', ''
    switch ($String) {
        'STANDARDPACK' {'E[1]'}
        'ENTERPRISEPACK' {'E[3]'}
        'PROJECTCLIENT' {'Project'}
        'VISIOCLIENT' {'Visio'}
        'MCOMEETADV' {'Skype PSTN'}
        'FLOW_FREE' {'Flow'}
        default {'None'}
    }
}

$Cache:AllAccounts = Get-MsolUser -All | Select-Object DisplayName, FirstName, IsLicensed, LastName, Office, UserPrincipalName, UserType, Licenses, WhenCreated
$Cache:AllLicenses = Get-MsolAccountSku
$Cache:AllRecipients = Get-Recipient | Select DisplayName, PrimarySmtpAddress, RecipientTypeDetails, WhenCreated
$Cache:AllContacts = Get-MailContact | Select Name, PrimarySmtpAddress, HiddenFromAddressListEnabled
$SortedAccounts = @()

foreach ($User in $Cache:AllAccounts) {
    $Licenses = $User.Licenses
    $LicenseArray = $Licenses | ForEach-Object {ConvertLicenseString($_.AccountSkuID)}
    $LicenseString = $LicenseArray -join ', '
    $LicensedSharedMailboxProperties = [PSCustomObject][Ordered]@{
        DisplayName       = $User.DisplayName
        Licenses          = $LicenseString
        UserPrincipalName = $User.UserPrincipalName
        WhenCreated       = $User.WhenCreated
        Office            = $User.Office
        IsLicensed        = $User.IsLicensed
    }

    $SortedAccounts += $LicensedSharedMailboxProperties
}

$Cache:SortedAccounts = $SortedAccounts

$OfficePage = New-UDPage -Name "Office 365" -Icon home -Content {
    New-UDRow {
        New-UDColumn -Size 5 {
            New-UDLayout -Columns 3 {
                New-UDCounter -Title 'Total Accounts' -AutoRefresh -RefreshInterval 300 -TextSize Large -TextAlignment center -Endpoint {
                    $Cache:AllAccounts.Count | ConvertTo-Json
                }

                New-UDCounter -Title 'Licensed Accounts' -AutoRefresh -RefreshInterval 300 -TextSize Large -TextAlignment center -Endpoint {
                    ($Cache:AllAccounts | Where-Object {$_.Licenses -ne $null}).Count | ConvertTo-Json
                }

                New-UDCounter -Title 'Shared Mailboxes' -AutoRefresh -RefreshInterval 300 -TextSize Large -TextAlignment center -Endpoint {
                    ($Cache:AllRecipients | Where-Object {$_.RecipientTypeDetails -eq 'SharedMailbox'}).Count | ConvertTo-Json
                }
            }
        }
    }

    New-UDRow {
        New-UDButton -Text 'All accounts' -OnClick {
            $Session:GridSelection = 'All'
            Sync-UDElement -Id 'AccountGrid'
        }

        New-UDButton -Text 'Licensed accounts' -OnClick {
            $Session:GridSelection = 'Licensed'
            Sync-UDElement -Id 'AccountGrid'
        }
                
        New-UDButton -Text 'Contacts' -OnClick {
            $Session:GridSelection = 'Contacts'
            Sync-UDElement -Id 'AccountGrid'
        }
    }
    
    New-UDElement -Id 'AccountGrid' -Tag div -Endpoint {
        New-UDRow {
            if ($Session:GridSelection -eq 'All') {
                New-UDGrid -Title "Internal Users" -AutoRefresh -RefreshInterval 300 -Headers @('Name', 'Email Address', 'Licenses') -Properties @('DisplayName', 'UserPrincipalName', 'Licenses') -Endpoint {
                    $Cache:SortedAccounts | Out-UDGridData
                }
            }

            if ($Session:GridSelection -eq 'Licensed') {
                New-UDGrid -Title 'Licensed Users' -Headers @('Name', 'Email Address', 'Licenses') -Properties @('DisplayName', 'UserPrincipalName', 'Licenses') -Endpoint {
                    ($Cache:SortedAccounts | Where-Object {$_.IsLicensed -eq $true}) | Out-UDGridData
                }
            }

            if ($Session:GridSelection -eq 'Contacts') {
                New-UDGrid -Title 'Contacts' -Headers @('Name', 'Email Address') -Properties ('Name', 'PrimarySmtpAddress') -AutoRefresh -RefreshInterval 300 -Endpoint {
                    $Cache:AllContacts | Out-UDGridData
                }
            }
        }
    }      
}
    
Get-UDDashboard | Stop-UDDashboard
$Dashboard = New-UDDashboard -Title "$Company Dashboard" -Theme $Theme -Pages @($OfficePage)
Start-UDDashboard -Port 10000 -Dashboard $Dashboard

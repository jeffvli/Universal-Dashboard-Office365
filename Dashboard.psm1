function Start-UDOfficeDashboard {
    # Login to Office 365 MSOnline first
    $Theme = Get-UDTheme Azure
    $Company = 'Template'

    function Convert-LicenseString {
        # Remove first set of characters matching "$STRING:"
        $String = $args[0] -Replace '^[^:]+:\s*', ''
        # Rename licenses to human-readable format
        switch ($String) {
            'DESKLESSPACK'                      {'K[1]'}
            'DESKLESSWOFFPACK'                  {'K[2]'}
            'FLOW_FREE'                         {'Flow Free'}
            'LITEPACK'                          {'{P[1]'}
            'EXCHANGESTANDARD'                  {'Exchange Online'}
            'O365_BUSINESS_ESSENTIALS'          {'Business Essentials'}
            'STANDARDPACK'                      {'E[1]'}
            'STANDARDWOFFPACK'                  {'E[2]'}
            'ENTERPRISEPACK'                    {'E[3]'}
            'ENTERPRISEPACKLRG'                 {'E[3]'}
            'ENTERPRISEWITHSCAL'                {'E[4]'}
            'STANDARDPACK_STUDENT'              {'A[1] for Students'}
            'STANDARDWOFFPACKPACK_STUDENT'      {'A[2] for Students'}
            'ENTERPRISEPACK_STUDENT'            {'A[3] for Students'}
            'ENTERPRISEWITHSCAL_STUDENT'        {'A[4] for Students'}
            'STANDARDPACK_FACULTY'              {'A[1] for Faculty'}
            'STANDARDWOFFPACKPACK_FACULTY'      {'A[2] for Faculty'}
            'ENTERPRISEPACK_FACULTY'            {'A[3] for Faculty'}
            'ENTERPRISEWITHSCAL_FACULTY'        {'A[4] for Faculty'}
            'ENTERPRISEPACK_B_PILOT'            {'Enterprise Preview'}
            'STANDARD_B_PILOT'                  {'Small Business Preview'}
            'VISIOCLIENT'                       {'Visio'}
            'POWER_BI_ADDON'                    {'Power BI Addon'}
            'POWER_BI_INDIVIDUAL_USE'           {'Power BI Individual'}
            'POWER_BI_STANDALONE'               {'Power BI Stand-Alone'}
            'POWER_BI_STANDARD'                 {'Power BI Standard'}
            'PROJECTESSENTIALS'                 {'Project Lite'}
            'PROJECTCLIENT'                     {'Project Professional'}
            'PROJECTPROFESSIONAL'               {'Project Professional'}
            'PROJECTONLINE_PLAN_1'              {'Project P[1]'}
            'PROJECTONLINE_PLAN_2'              {'Project P[2]'}
            'ECAL_SERVICES'                     {'ECAL'}
            'EMS'                               {'Enterprise Mobility Suite'}
            'RIGHTSMANAGEMENT_ADHOC'            {'Windows Azure Rights Management'}
            'MCOMEETADV'                        {'Skype'}
            'SHAREPOINTSTORAGE'                 {'SharePoint'}
            'PLANNERSTANDALONE'                 {'Planner'}
            'CRMIUR'                            {'CMRIUR'}
            'BI_AZURE_P1'                       {'Power BI Reporting and Analytics'}
            'INTUNE_A'                          {'Windows Intune Plan A'}
            'MICROSOFT_BUSINESS_CENTER'         {'Microsoft Business Center'}
            default                             {'Unknown'}
        }
    }

    $Cache:AllAccounts = Get-MsolUser -All | Select-Object DisplayName, FirstName, IsLicensed, LastName, Office, UserPrincipalName, UserType, Licenses, WhenCreated
    $Cache:AllLicenses = Get-MsolAccountSku | Where-Object {$_.ActiveUnits -lt 10000}
    $Cache:AllMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object Name, Alias, UserPrincipalName, RecipientTypeDetails, Identity, DisplayName, ProhibitSendQuota
    $Cache:AllRecipients = Get-Recipient -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, WhenCreated
    $Cache:AllContacts = Get-MailContact -ResultSize Unlimited | Select-Object Name, PrimarySmtpAddress, HiddenFromAddressListEnabled, DistinguishedName
    $Cache:AllGroups = Get-DistributionGroup | Select-Object Name, DisplayName, PrimarySmtpAddress

    $SortedAccounts = @()
    $SortedLicenses = @()

    foreach ($User in $Cache:AllAccounts) {
        $Licenses = $User.Licenses
        $LicenseArray = $Licenses | ForEach-Object {Convert-LicenseString($_.AccountSkuId)} | Sort-Object
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

    foreach ($License in $Cache:AllLicenses) {
        $LicenseName = $License | ForEach-Object {Convert-LicenseString($_.AccountSkuId)}
        $AllLicenseProperties = [PSCustomObject]@{
            License = $LicenseName
            UsedLicenses = $License.ActiveUnits
            AvailableLicenses = ($License.ActiveUnits - $License.ConsumedUnits)
        }

        $SortedLicenses += $AllLicenseProperties
    }

    $Cache:SortedAccounts = $SortedAccounts
    $Cache:SortedLicenses = $SortedLicenses

    $Pages = Get-ChildItem (Join-Path $PSScriptRoot 'pages') -Recurse -File | ForEach-Object {
        & $_.FullName
    }

    Get-UDDashboard | Stop-UDDashboard
    $Dashboard = New-UDDashboard -Title "$Company Office365 Dashboard" -Theme $Theme -Pages $Pages
    Start-UDDashboard -Port 10000 -Dashboard $Dashboard
}
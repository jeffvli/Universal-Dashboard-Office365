function Start-UDOfficeDashboard {
    # Login to Office 365 MSOnline first
    $Theme = Get-UDTheme Azure
    $Company = 'Template'
    
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

    $Functions = Get-ChildItem (Join-Path $PSScriptRoot 'functions') -Recurse -File | ForEach-Object {
        & $_.FullName
    }

    Get-UDDashboard | Stop-UDDashboard
    $Dashboard = New-UDDashboard -Title "$Company Office365 Dashboard" -Theme $Theme -Pages $Pages
    Start-UDDashboard -Port 10000 -Dashboard $Dashboard
}
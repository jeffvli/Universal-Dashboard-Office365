$AccountPage = New-UDPage -Name "Accounts" -Icon home -Content {
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

        New-UDButton -Text 'Shared Mailboxes' -OnClick {
            $Session:GridSelection = 'Shared'
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
                New-UDGrid -Title 'Licensed Users' -AutoRefresh -RefreshInterval 300 -Headers @('Name', 'Email Address', 'Licenses') -Properties @('DisplayName', 'UserPrincipalName', 'Licenses') -Endpoint {
                    ($Cache:SortedAccounts | Where-Object {$_.IsLicensed -eq $true}) | Out-UDGridData
                }
            }

            if ($Session:GridSelection -eq 'Contacts') {
                New-UDGrid -Title 'Contacts' -AutoRefresh -RefreshInterval 300 -Headers @('Name', 'Email Address') -Properties @('Name', 'PrimarySmtpAddress') -Endpoint {
                    $Cache:AllContacts | Out-UDGridData
                }
            }
            
            if ($Session:GridSelection -eq 'Shared') {
                New-UDGrid -Title 'Shared Mailboxes' -AutoRefresh -RefreshInterval 300 -Headers @('Name', 'Email Address') -Properties @('Name', 'UserPrincipalName') -Endpoint {
                    $Cache:AllMailboxes | Where-Object {$_.RecipientTypeDetails -eq 'SharedMailbox'} | Out-UDGridData
                }
            }
        }
    }
}
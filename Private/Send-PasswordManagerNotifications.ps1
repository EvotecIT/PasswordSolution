function Send-PasswordManagerNofifications {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $ManagerSection,
        [System.Collections.IDictionary] $Summary,
        [System.Collections.IDictionary] $CachedUsers,
        [ScriptBlock] $TemplateManager,
        [string] $TemplateManagerSubject,
        [ScriptBlock] $TemplateManagerNotCompliant,
        [string] $TemplateManagerNotCompliantSubject,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $Logging,
        [System.Collections.IDictionary] $GlobalManagersCache
    )
    if ($ManagerSection.Enable) {
        Write-Color -Text "[i] Sending notifications to managers " -Color White, Yellow, White, Yellow, White, Yellow, White
        $CountManagers = 0
        [Array] $SummaryManagersEmails = foreach ($Manager in $Summary['NotifyManager'].Keys) {
            $CountManagers++
            if ($CachedUsers[$Manager]) {
                # This user is "findable" in AD
                $ManagerUser = $CachedUsers[$Manager]
            } elseif ($GlobalManagersCache[$Manager]) {
                # This user is findable in managers cache
                # This is required when user uses `FilterOrganizationalUnit` feature and manager is not in the same OU
                # This causes Manager Data to be not processed in the same way as User Data so we need to process it separately
                $ManagerUser = $GlobalManagersCache[$Manager]
            } else {
                # This user is provided by user in config file
                $ManagerUser = $Summary['NotifyManager'][$Manager]['Manager']
            }
            [Array] $ManagedUsers = $Summary['NotifyManager'][$Manager]['ManagerDefault'].Values.Output
            [Array] $ManagedUsersManagerNotCompliant = $Summary['NotifyManager'][$Manager]['ManagerNotCompliant'].Values.Output

            $EmailSplat = [ordered] @{}

            if ($Summary['NotifyManager'][$Manager].ManagerDefault.Count -gt 0) {
                if ($TemplateManager) {
                    # User uses global template
                    $EmailSplat.Template = $TemplateManager
                } else {
                    # User uses built-in template
                    $EmailSplat.Template = {

                    }
                }
                if ($TemplateManagerSubject) {
                    $EmailSplat.Subject = $TemplateManagerSubject
                } else {
                    $EmailSplat.Subject = "[Password Expiring] Dear Manager - Your accounts are expiring!"
                }
            } elseif ($Summary['NotifyManager'][$Manager].ManagerNotCompliant.Count -gt 0) {
                if ($TemplateManagerNotCompliant) {
                    # User uses global template
                    $EmailSplat.Template = $TemplateManagerNotCompliant
                } else {
                    # User uses built-in template
                    $EmailSplat.Template = {

                    }
                }
                if ($TemplateManagerNotCompliantSubject) {
                    $EmailSplat.Subject = $TemplateManagerNotCompliantSubject
                } else {
                    $EmailSplat.Subject = "[Password Escalation] Accounts are expiring with non-compliant manager"
                }
            }

            $EmailSplat.User = $ManagerUser
            $EmailSplat.ManagedUsers = $ManagedUsers
            $EmailSplat.ManagedUsersManagerNotCompliant = $ManagedUsersManagerNotCompliant
            $EmailSplat.EmailParameters = $EmailParameters

            $EmailSplat.EmailDateFormat = $Logging.EmailDateFormat
            $EmailSplat.EmailDateFormatUTCConversion = $Logging.EmailDateFormatUTCConversion

            if ($ManagerSection.SendToDefaultEmail -ne $true) {
                $EmailSplat.EmailParameters.To = $ManagerUser.EmailAddress
            } else {
                $EmailSplat.EmailParameters.To = $ManagerSection.DefaultEmail
            }
            if ($Logging.NotifyOnManagerSend) {
                Write-Color -Text "[i] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }
            [Array] $DisabledAccounts = foreach ($ManagedUserDN in $Summary['NotifyManager'][$Manager].ManagerDefault.Keys) {
                $AccountToDisable = [PSCustomObject] @{
                    SamAccountName = $ManagedUser.User.SamAccountName
                    Domain         = $ManagedUser.User.Domain
                    Disabled       = $null
                    Error          = $null
                }
                $ManagedUser = $Summary['NotifyManager'][$Manager].ManagerDefault[$ManagedUserDN]
                if ($ManagedUser.Rule -and $null -ne $ManagedUser.Rule.DisableDays) {
                    $CompareSuccess = $false
                    # We need to check if the user is in the disable days list
                    if ($ManagedUser.Rule.DisableType -eq 'in') {
                        $CompareSuccess = $ManagedUser.User.DaysToExpire -in $ManagedUser.Rule.DisableDays
                    } elseif ($ManagedUser.Rule.DisableType -eq 'lt') {
                        $CompareSuccess = $ManagedUser.User.DaysToExpire -lt $ManagedUser.Rule.DisableDays
                    } elseif ($ManagedUser.Rule.DisableType -eq 'gt') {
                        $CompareSuccess = $ManagedUser.User.DaysToExpire -gt $ManagedUser.Rule.DisableDays
                    } elseif ($ManagedUser.Rule.DisableType -eq 'eq') {
                        $CompareSuccess = $ManagedUser.User.DaysToExpire -eq $ManagedUser.Rule.DisableDays
                    } else {
                        Write-Color -Text "[r] Unknown disable type: ", $ManagedUser.Rule.DisableType, " for user ", $ManagedUser.User.DisplayName, " (", $ManagedUser.User.UserPrincipalName, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                    }
                    if ($CompareSuccess) {
                        # Write-Color -Text "[i] Disabling user ", $ManagedUser.User.DisplayName, " (", $ManagedUser.User.UserPrincipalName, ") (DaysToExpire: ", $ManagedUser.User.DaysToExpire, ")" -Color Yellow, White, Magenta, White, Magenta, White, White, Blue
                        #$ManagedUser.User
                        if ($ManagedUser.Rule.DisableWhatIf) {
                            Write-Color -Text "[i] Disabling user ", $ManagedUser.User.DisplayName, " (", $ManagedUser.User.UserPrincipalName, ")", " would be disabled" -Color Cyan, White, Red, Cyan, Red, Yellow
                            $AccountToDisable.Disabled = $false
                            $AccountToDisable.Error = "WhatIf"
                        } else {
                            # Disable the user
                            Write-Color -Text "[i] Disabling user ", $ManagedUser.User.DisplayName, " (", $ManagedUser.User.UserPrincipalName, ")" -Color Cyan, White, Magenta, White, Magenta, White, White, Blue
                            if ($ManagedUser.User.Enabled) {
                                try {
                                    Disable-ADAccount -Identity $ManagedUser.User.DistinguishedName -Confirm:$false -WhatIf:$ManagedUser.Rule.DisableWhatIf -ErrorAction Stop
                                    $AccountToDisable.Disabled = $true
                                    $AccountToDisable.Error = $null
                                } catch {
                                    $AccountToDisable.Disabled = $false
                                    $AccountToDisable.Error = $_.Exception.Message
                                    Write-Color -Text "[r] Disabling user ", $ManagedUser.User.DisplayName, " (", $ManagedUser.User.UserPrincipalName, ") failed: ", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                                }
                            } else {
                                $AccountToDisable.Disabled = $false
                                $AccountToDisable.Error = "Already disabled"
                                Write-Color -Text "[i] User is already disabled: ", $ManagedUser.User.DisplayName, " (", $ManagedUser.User.UserPrincipalName, ")" -Color Cyan, White, Magenta, White, Magenta, White, White, Blue
                            }
                        }
                        $AccountToDisable
                    }
                }
            }
            $EmailResult = Send-PasswordEmail @EmailSplat
            if ($Logging.NotifyOnManagerSend) {
                if ($EmailResult.Error) {
                    if ($EmailResult.SentTo) {
                        Write-Color -Text "[r] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ", error: ", $EmailResult.Error, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                    } else {
                        Write-Color -Text "[r] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, ", error: ", $EmailResult.Error, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                    }
                } else {
                    if ($EmailResult.SentTo) {
                        Write-Color -Text "[r] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                    } else {
                        Write-Color -Text "[r] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                    }
                }
            }

            [PSCustomObject] @{
                DisplayName              = $ManagerUser.DisplayName
                SamAccountName           = $ManagerUser.SamAccountName
                Domain                   = $ManagerUser.Domain
                Status                   = $EmailResult.Status
                StatusWhen               = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                SentTo                   = $EmailResult.SentTo
                StatusError              = $EmailResult.Error
                Accounts                 = $ManagedUsers.SamAccountName
                AccountsCount            = $ManagedUsers.Count
                DisabledAccounts         = $DisabledAccounts.SamAccountName
                DisabledAccountsCount    = $DisabledAccounts.Count
                DisabledAccountsError    = $DisabledAccounts.Error | Sort-Object -Unique
                Template                 = 'Unknown'
                ManagerNotCompliant      = $ManagedUsersManagerNotCompliant.SamAccountName
                ManagerNotCompliantCount = $ManagedUsersManagerNotCompliant.Count
                #ManagerDisabled          = $ManagedUsersManagerDisabled.SamAccountName
                #ManagerDisabledCount     = $ManagedUsersManagerDisabled.Count
                #ManagerMissing           = $ManagedUsersManagerMissing.SamAccountName
                #ManagerMissingCount      = $ManagedUsersManagerMissing.Count
                #ManagerMissingEmail      = $ManagedUsersManagerMissingEmail.SamAccountName
                #ManagerMissingEmailCount = $ManagedUsersManagerMissingEmail.Count
            }
            if ($ManagerSection.SendCountMaximum -gt 0) {
                if ($ManagerSection.SendCountMaximum -le $CountManagers) {
                    Write-Color -Text "[i]", " Send count maximum reached. There may be more managers that match the rule." -Color Red, DarkRed
                    break
                }
            }
        }
        Write-Color -Text "[i] Sending notifications to managers (sent: ", $SummaryManagersEmails.Count, " out of ", $Summary['NotifyManager'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
        $SummaryManagersEmails
    } else {
        Write-Color -Text "[i] Sending notifications to managers is ", "disabled!" -Color White, Yellow, DarkRed
    }
}
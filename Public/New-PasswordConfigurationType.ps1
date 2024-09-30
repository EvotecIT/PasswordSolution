function New-PasswordConfigurationType {
    <#
    .SYNOPSIS
    Configures behavior of password notification emails for different types of users.

    .DESCRIPTION
    Configures behavior of password notification emails for different types of users.
    It is used to define how the notification emails should be sent to different types of users.
    It supports User, Manager, Security, and Admin types.
    Main function is to prevent sending emails to real users during testing phase.

    .PARAMETER Type
    Type of the configuration. Possible values are User, Manager, Security, and Admin.

    .PARAMETER Enable
    Enable sending emails for the specified type.

    .PARAMETER SendCountMaximum
    Maximum number of emails that can be sent for the specified type. This is used to prevent sending emails to a large number of users during testing phase.

    .PARAMETER DefaultEmail
    Default email address to which the emails should be sent. This is used to prevent sending emails to real users during testing phase.
    All emails will be sent to this email address for the specified type, with the exception of Admin type.
    Admin type will send emails to the specified email address as is.

    .PARAMETER AttachCSV
    Attach a CSV file with the list of users to the email.
    This is used to provide additional information about the users to the recipient.
    This is only used for Security type.

    .PARAMETER DisplayName
    Display name of the Admin user. This is used to provide a custom display name for the Admin user.
    If not specified, the default display name is "Administrators".

    .EXAMPLE
    New-PasswordConfigurationType -Type User -Enable -SendCountMaximum 10 -DefaultEmail 'przemyslaw.klys+testgithub1@test.pl'

    .EXAMPLE
    New-PasswordConfigurationType -Type Manager -Enable -SendCountMaximum 10 -DefaultEmail 'przemyslaw.klys+testgithub2@test.pl'

    .EXAMPLE
    New-PasswordConfigurationType -Type Security -Enable -SendCountMaximum 1 -DefaultEmail 'przemyslaw.klys+testgithub3@test.pl' -AttachCSV

    .EXAMPLE
    New-PasswordConfigurationType -Type Admin -Enable -EmailAddress 'przemyslaw.klys+testgithub3@test.pl' -DisplayName 'Administrators'

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('User', 'Manager', 'Security', 'Admin')][string] $Type,
        [switch] $Enable,
        [int] $SendCountMaximum,
        [Alias('EmailAddress')][string] $DefaultEmail,
        [switch] $AttachCSV,
        [string] $DisplayName
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationType$Type"
        Settings = @{
            Enable                 = $Enable.IsPresent
            SendCountMaximum       = $SendCountMaximum
            SendToDefaultEmail     = if ($DefaultEmail) { $true } else { $false }
            DefaultEmail           = $DefaultEmail
            OverwriteEmailProperty = $OverwriteEmailProperty
            AttachCSV              = $AttachCSV.IsPresent
        }
    }

    if ($Type -eq "Admin") {
        $Output.Settings.Manager = @{
            DisplayName  = if (-not $DisplayName) { "Administrators" } else { $DisplayName }
            EmailAddress = $DefaultEmail
        }
    }
    $Output
}
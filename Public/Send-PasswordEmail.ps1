function Send-PasswordEmail {
    [CmdletBinding()]
    param(
        [scriptblock] $Template,
        [PSCustomObject] $User,
        [Array] $ManagedUsers,
        [Array] $ManagedUsersManagerNotCompliant,
        [Array] $SummaryUsersEmails,
        [Array] $SummaryManagersEmails,
        [Array] $SummaryEscalationEmails,
        [string] $TimeToProcess,

        #[Array] $ManagedUsersManagerDisabled,
        #[Array] $ManagedUsersManagerMissing,
        #[Array] $ManagedUsersManagerMissingEmail,
        [System.Collections.IDictionary] $EmailParameters,
        [string] $Subject
    )

    if ($Template) {
        $SourceParameters = [ordered] @{
            ManagerDisplayName                   = $User.DisplayName
            ManagerUsersTable                    = $ManagedUsers
            ManagerUsersTableManagerNotCompliant = $ManagedUsersManagerNotCompliant
            SummaryEscalationEmails              = $SummaryEscalationEmails
            SummaryManagersEmails                = $SummaryManagersEmails
            SummaryUsersEmails                   = $SummaryUsersEmails
            TimeToProcess                        = $TimeToProcess
            # Only works if User is set
            UserPrincipalName                    = $User.UserPrincipalName     # : adm.pklys@ad.evotec.xyz
            SamAccountName                       = $User.SamAccountName        # : adm.pklys
            Domain                               = $User.Domain                # : ad.evotec.xyz
            Enabled                              = $User.Enabled
            EmailAddress                         = $User.EmailAddress          # :
            DateExpiry                           = $User.DateExpiry            # :
            DaysToExpire                         = $User.DaysToExpire          # :
            PasswordExpired                      = $User.PasswordExpired       # : False
            PasswordLastSet                      = $User.PasswordLastSet       # : 05.09.2020 11:07:29
            PasswordNotRequired                  = $User.PasswordNotRequired   # : False
            PasswordNeverExpires                 = $User.PasswordNeverExpires  # : True
            ManagerSamAccountName                = $User.ManagerSamAccountName # : przemyslaw.klys
            ManagerEmail                         = $User.ManagerEmail          # : przemyslaw.klys@evotec.pl
            ManagerStatus                        = $User.ManagerStatus         # : Enabled
            ManagerLastLogonDays                 = $User.ManagerLastLogonDays  # : 0
            Manager                              = $User.Manager               # : Przemysław Kłys
            DisplayName                          = $User.DisplayName           # : Administrator Przemysław Kłys
            GivenName                            = $User.GivenName             # : Administrator Przemysław
            Surname                              = $User.Surname               # : Kłys
            OrganizationalUnit                   = $User.OrganizationalUnit    # : OU=Special,OU=Accounts,OU=Production,DC=ad,DC=evotec,DC=xyz
            MemberOf                             = $User.MemberOf              # : {CN=GDS-TestGroup4,OU=Security,OU=Groups,OU=Production,DC=ad,DC=evotec,DC=xyz, CN=GDS-TestGroup2,OU=Security,OU=Groups,OU=Production,DC=ad,DC=evotec,DC=xyz, CN=Domain Admins,CN=Users,DC=ad,DC=evotec,DC=xyz}
            DistinguishedName                    = $User.DistinguishedName     # : CN=Administrator Przemysław Kłys,OU=Special,OU=Accounts,OU=Production,DC=ad,DC=evotec,DC=xyz
            ManagerDN                            = $User.ManagerDN             # : CN=Przemysław Kłys,OU=Users,OU=Accounts,OU=Production,DC=ad,DC=evotec,DC=xyz
        }
        $Body = EmailBody -EmailBody $Template -Parameter $SourceParameters

        # Below command would require to define variables as they are used in scriptblock
        #$EmailParameters.Subject = $ExecutionContext.InvokeCommand.ExpandString($Subject)
        # following replacement is a bit more cumbersome the the one above but a bit more secure and doesn't require creating 20+ unused variables
        $EmailParameters.Subject = Add-ParametersToString -String $Subject -Parameter $SourceParameters
        $EmailParameters.Body = $Body

        Send-EmailMessage @EmailParameters
    }
}
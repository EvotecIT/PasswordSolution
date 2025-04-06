function Set-PasswordQualityConfiguration {
    [CmdletBinding()]
    param(
        [scriptblock] $ConfigurationDSL
    )
    $OutputInformation = [ordered] @{
        QualityEmail       = [System.Collections.Generic.List[ordered]]::new()
        EmailConfiguration = $null
    }
    if ($ConfigurationDSL) {
        try {
            $ConfigurationExecuted = & $ConfigurationDSL
            foreach ($Configuration in $ConfigurationExecuted) {
                if ($Configuration.Type -eq 'QualityEmail') {
                    $OutputInformation.QualityEmail.Add($Configuration.Settings)
                } elseif ($Configuration.Type -eq 'PasswordConfigurationEmail') {
                    $OutputInformation.EmailConfiguration = $Configuration.Settings
                }
            }
        } catch {
            Write-Color -Text "[e]", " Processing configuration failed because of error in line ", $_.InvocationInfo.ScriptLineNumber, " in ", $_.InvocationInfo.InvocationName, " with message: ", $_.Exception.Message -Color Yellow, White, Red
            return
        }
    }
    $OutputInformation
}
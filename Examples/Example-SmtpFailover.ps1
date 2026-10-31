Import-Module PasswordSolution

# Apply these settings to an existing Start-PasswordSolution configuration.
$EmailParameters = @{
    From                   = 'notifications@example.com'
    Server                 = 'relay-a.example.com', 'relay-b.example.com'
    Port                   = 25
    SecureSocketOptions    = 'StartTls'
    Strategy               = 'Failover'
    RetryCount             = 3
    RetryDelayMilliseconds = 500
    CircuitBreakMinutes    = 2
}

# The same splat works inside the configuration DSL.
New-PasswordConfigurationEmail @EmailParameters

# Endpoints may override the common port and TLS mode:
# Server = 'smtp+starttls://relay-a.example.com:587', 'smtps://relay-b.example.com:465'
# RoundRobin alternates the starting host; Random chooses a starting host at random.
# Transient failures try another eligible host. Permanent failures do not rotate.
# A lost acknowledgement during sending is DeliveryUnknown and needs review before retry.

# Keep a stable SearchPath in the existing configuration to retain delivery results.
# Find-PasswordNotification -SearchPath '.\Search\notifications.xml' -FailedOnly |
#     Select-Object UserPrincipalName, Rule, SentTo, StatusError, Server, DeliveryUnknown
# Start-PasswordSolution -ConfigurationDSL $Configuration -RetryFailedOnly
# Uncertain deliveries are reported for verification and excluded from recovery by default.
# After confirming that they were not delivered, add -IncludeUnknownDelivery to the recovery run.
# Recovery re-evaluates current rules for failed users; it does not replay account-disable actions
# or send manager, security and administrator notifications.

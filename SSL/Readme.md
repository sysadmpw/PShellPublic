# SSL Certificate Renew
This script will grab a PFX file from a location, then attempt to scan each IP in the serverlist.json file, copy and install the PFX file then bind the PFX file to the default website HTTPS binding. This script at this stage is only compatible with Microsoft IIS.
    
EventID Numbers and explanations - #Error, Information, FailureAudit, SuccessAudit, Warning
    1. Information, things that are happening or about to happen  1000 - Entry Type Information
    2. Errors 1001 - Entry Type Error
    3. Successes 1002 - Entry Type SuccessAudit
    4. Warning 1003 - Entry Type Warning
    
    

# SSL Certificate Renew
This script will grab a PFX file from a location, then attempt to scan each IP in the serverlist.json file, copy and install the PFX file then bind the PFX file to the default website HTTPS binding. This script at this stage is only compatible with Microsoft IIS.
    
EventID Numbers and explanations - #Error, Information, FailureAudit, SuccessAudit, Warning
    1. Information, things that are happening or about to happen  1000 - Entry Type Information
    2. Errors 1001 - Entry Type Error
    3. Successes 1002 - Entry Type SuccessAudit
    4. Warning 1003 - Entry Type Warning
    
    
Check the global variables in the script and adjust as needed. Note, this script utlises Invoke-Command with credentials that are pre-created with encrypted AES and Password files. A more secure and elegant solution will be implemented in future revisions.

The script will attempt to bind the SSL script to an Assigned IP or Unassigned based on what is configured on the target IIS server (unassigned being 0.0.0.0).



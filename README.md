# Search-UnifiedAuditLogUserPrompts
Script to query Unified Audit Log with prompts for users who don't know how to edit a script


This started as the script posted to https://docs.microsoft.com/en-us/microsoft-365/compliance/audit-log-search-script?view=o365-worldwide.

That script was prepended with a series of user prompts.  The idea was that it would allow users who are not comfortable with PowerShell to answer a series of questions, and then query the unified Audit Log using the search-unifiedauditlog cmdlet.  This can get the users around any problems that might appear in the graphical user interface on compliance.microsoft.com's audit log query interface.

The list of questions is not exhaustive.  Many customers might want someone familiar with PowerShell to tweak the prompts for their needs, or to translate them to other languages.
The script is intended as an example, not a final version.

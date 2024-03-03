## PowerShell Racktables integration module

For more information check out [winfred.com](https://winfred.com)

This is a powershell module that connects to the Racktables MySQL DB and performs various tasks.

It is not intended for beginners. It is intended for people who have a good understanding of PowerShell.

This will be an intermediary layer between racktables and outputs provided by other powershell code not included here.

It is possible to fully automate syncs of servers and switches between racktables and other systems as the below code has been used for that purpose.

Output from WMI, SCCM, vCenter, etc. can be used to update racktables with the below functions.

Don't just implement the functions, read the help functions provided first and ask questions if it's not entirely clear.

This has been tested extensively in production, but only in a single environment. It might not work in your environment without modification.

This module directly makes changes to the DB, this is inherently somewhat risky, since on updates of the DB, the tables might change.

Always make sure you have a working backup of the DB before running any of these functions.

set MySQLHost.txt in the script root (that's probably the module folder if used as intended) to the hostname/ip of your MySQL server.

MySql.Data.dll is required for this module to work. It can be downloaded from https://dev.mysql.com/downloads/connector/net
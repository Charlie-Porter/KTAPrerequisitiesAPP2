# KTAPrerequisitiesAPP2

This program conveniently installs all KTA prerequisites with minimal effort.

User Guide Edit section

- Download the KTAPrerequisitiesApp from here.
- Select which KTA install type to be installed on the machine.
- Enter the service accounts username eg domain\account.
- If you would like to grant DBCreator, uncheck the "skip granting SQL DB Creator". If not, go to step 7.
- Enter the hostname or IP address of the SQL server location. If it uses a SQL named instance name enter:  MySQLServer\MyNamedInstance  
  a. Note: You can continue installing even if SQL server is not connected or you wish to grant DBCreastor permissions at a different time.
  By default, the "Use Windows Authentication" is checked.  If you wish to use SQL authentication, uncheck "Use Windows authentication".
  a. Enter the Windows or SQL authentication user you wish to use.
- Test the connection to ensure you have connectivity, if any errors, review the connection details, firewall rules and try again.  If SQL uses a non default port, 1433, you may need to add this to the SQL server instance location eg MySQLServer:xxxx\MyNamedInstance  
- Click install

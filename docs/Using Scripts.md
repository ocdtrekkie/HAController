# Using Scripts

HAController supports running external scripts, as a primary method to provide non-standard integrations or execute actions which require closing the HAController application during execution (such as updating HAController or archiving HAController data files). Scripts can only be run from inside the C:\HAC\scripts folder, and scripts must have alphanumeric file names with no symbols or spaces.

## Batchfile ##

Batchfiles are the sole format of external command script which can be triggered by HAController, however, there is no limitation on what can be initiated within a Batchfile. So while it is not possible to run a PowerShell script from HAController, you could write a Batchfile which runs a PowerShell script. Similarly, if you need to execute a command that lives outside of the HAController scripts folder, you merely need to include a Batchfile in the folder which calls it. This step exists to prevent exploits where HAController could be used to execute processes that were not intended.

Batchfiles can be run by HAController using the `run script` command, followed by the name (without extension) of a .bat file.

## HACScript ##

HACScripts are text files containing a series of commands understood by HAController, one per line. This can be used to run multiple commands at a time, such as on startup, or to automate settings changes.

HACScripts can be run by HAController using the `run hacscript` command, followed by the name (without extension) of a .hacscript file.

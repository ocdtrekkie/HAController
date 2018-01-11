# Running Updates

## update now / run update ##

This is now the preferred way to run updates on HAController, as it will not download anything if it is running the current version already, and the download and decompression is handled by the application and logged. While HAController is not distributed as a ClickOnce application, the code does support updates for ClickOnce distributions via this function.

## run script update ##

This method is extremely reliable, as it does not check versions and does not require HAController be working correctly. This script (C:\HAC\scripts\update.bat) will even run if manually called outside HAController, and regardless of version, will download the latest from the Internet and overwrite the existing application code. Note that the updateupdate script will update the update script, which may be prudent to run first.
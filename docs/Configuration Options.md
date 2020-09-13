# Standardized Commands

This is a list of global configuration modes

## Car mode

This can be set with `enable carmode` and `disable carmode`. Car mode is disabled by default.

Car mode will attempt to keep the application window in focus and with the cursor in the text box. This is intended to allow the app to work without the user having a monitor. It also may enable additional audio readout responses.

## Preserve Caps mode

This can be set with `enable caps` and `disable caps`. Preserve Caps mode is disabled by default.

Preseve Caps mode will ensure text box processing doesn't lowercase input like passwords. However, in this mode, all commands much match the case of the command language exactly.

## SmartCOM mode

This can be set with `enable smartcom` and `disable smartcom`. SmartCOM mode is enabled by default.

SmartCOM mode will intelligently attempt to connect to COM port devices even if they are at a different COM port than originally configured. It does this by storing the friendly name of each COM port device, and then searching for that friendly name if a connection on the original COM port fails.
# File Locations

## HAController Software

Currently, everything is built on the assumption that the HAController software stores data and has full access to *C:\HAC*. Some locations, like the database file location, may be configurable in the app's settings file, however, some places these have to be hardcoded, so the results of changing these may be unpredictable.

## Content Storage

It's strongly recommended to keep large content files on a separate drive from your HAController software so you can use slower, higher capacity storage for that use.

## Music

HAController does not explicitly have a setting for where to look for your music, as it uses the existing Windows Media Player functionality for looking at the music library. If you have a separate storage folder for music from the system default, be sure to include that folder in the system's music library.
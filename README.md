This is a fairly simple home automation controller, which focuses on working independent of web services. Eventually it will support a variety of devices and capabilities.

Huge credit for all of the INSTEON code that I built off of goes to **Jonathan Dale** of http://www.madreporite.com who stated that his code from his site is "free for anyone to use". As he does not have a specific license attached to that statement (I asked), I'm going to just plaster his name up here for credit. Basically I made a little program to send Insteon commands with his code, and that worked well and did what I wanted, but when I tried to make sense of the receiving code, I hit a roadblock. So I basically went for the concept to merge that code in directly, and then slowly integrate it with my app afterwards. I dumped it all in and got to work merging it, removing non-existant functions, and removing errors. This merge is still ongoing, while I also work on other points of interest.

## Modules ##

The following modules are core modules which cannot be disabled:
* Form1 is the UI screen which handles user input collection and some shortcut buttons.
* modComputer can monitor and control the local computer, including audio and video recording.
* modConverse is a cross between a command line interface and a conversation system.
* modCrypto handles basic encryption and decryption functionality.
* modDatabase handles the SQLite database that will store persistent data about the system.
* modGlobal mostly just hangs onto stuff that needs to be accessible to the whole program.
* modPersons contains functions for managing users and contacts.
* modRandom generates random numbers and can optionally query Random.org for true random results.

The following modules function optionally and add features:
* modDreamCheeky supports the Dream Cheeky Big Red Button and WebMail Notifier devices.
* modGPS collects location data from NMEA-compatible GPS receivers and manages navigation.
* modInsteon handles Insteon home automation devices, and may eventually add legacy X10.
* modOpenWeatherMap gathers weather data from the OpenWeatherMap API to present to the user.
* modMail handles checking for commands sent by mail and sending email notifications.
* modMapQuest uses the MapQuest API to get location information and directions.
* modMatrixLCD handles Matrix Orbital-compatible LCDs like the Adafruit LCD Backpack.
* modMusic implements Windows Media Player to play music.
* modPihole checks the status API of a Pi-hole on the network.
* modPing handles checks for the Internet and other devices on the network.
* modSpeech contains speech synthesis code.
* modWolframAlpha can request general knowledge from Wolfram Alpha's API.
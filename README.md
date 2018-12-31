This is a fairly simple automation controller, which focuses on working independent of web services. It is unique in that it is intended to support disparate environments such as both the home and the car, and mimic a true personal assistant that is available everywhere you are. My goal in writing this code was to develop something that operated with as simple and minimal code as possible so it would be easy to understand what the application was doing and how to modify it.

Roadmap: https://oasis.sandstorm.io/shared/SZ8UmwS2gAKhei2ruyofZy-tmf3faAwdCbkjAmSEViA

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
* modLibrary will be able to retrieve digital books from a central repository.
* modOpenWeatherMap gathers weather data from the OpenWeatherMap API to present to the user.
* modMail handles checking for commands sent by mail and sending email notifications.
* modMapQuest uses the MapQuest API to get location information and directions.
* modMatrixLCD handles Matrix Orbital-compatible LCDs like the Adafruit LCD Backpack.
* modMusic implements Windows Media Player to play music.
* modPihole checks the status API of a Pi-hole on the network.
* modPing handles checks for the Internet and other devices on the network.
* modSpeech contains speech synthesis code.
* modWolframAlpha can request general knowledge from Wolfram Alpha's API.
* modZWave will handle Z-Wave home automation devices.

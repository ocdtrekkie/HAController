This is a fairly simple home automation controller, which focuses on working independent of web services. Eventually it will support a variety of devices and capabilities.

Huge credit for all of the INSTEON code that I built off of goes to **Jonathan Dale** of http://www.madreporite.com who stated that his code from his site is "free for anyone to use". As he does not have a specific license attached to that statement (I asked), I'm going to just plaster his name up here for credit. Basically I made a little program to send Insteon commands with his code, and that worked well and did what I wanted, but when I tried to make sense of the receiving code, I hit a roadblock. So I basically went for the concept to merge that code in directly, and then slowly integrate it with my app afterwards. I dumped it all in and got to work merging it, removing non-existant functions, and removing errors. This merge is still ongoing, while I also work on other points of interest.

## Modules ##

* Form1 is basically meant to get factored out eventually in it's entirety.
* modDatabase handles the SQLite database that will store persistent data about the system.
* modGlobal will mostly just hang onto stuff that needs to be accessible to the whole program.
* modScheduler handles code that schedules events. (Currently using Quartz.NET)

The following modules should function optionally:
* modInsteon will handle Insteon and X10 home automation devices.
* modOpenWeatherMap gathers weather data from an API to present to the user.
* modMail will handle email triggers and notifications.
* modPing handles checks for the Internet and other devices on the network.
* modSpeech contains speech synthesis code.
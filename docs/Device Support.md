# Device Support

## Currently Supported Devices ##

These are devices the code has been tested to work with.

### INSTEON via Serial ###

* 2413U - PowerLinc USB Interface
* 2421 - TriggerLinc _(same as 2843-222)_
* 2441TH - Thermostat
* 2450 - I/O Linc **as "BuzzLinc"**
* 2450 - I/O Linc **as Sump Pump Alarm**
* 2450 - I/O Linc **as Garage Door & Status Kit (74551)[1]**
* 2477D - SwitchLinc Dual-Band Dimmer Switch
* 2477S - SwitchLinc Dual-Band On/Off Switch
* 2672-222 - LED Bulb 240V Edison
* 2843-222 - Open/Close Sensor _(same as 2421)_
* 2867-222 - Alert Module
* 2868-222 - Siren
* 2982-222 - Smoke Bridge

[1] You may need to use HouseLinc to set the I/O Linc into Momentary B mode.

### USB Devices ###

* Matrix Orbital LCD or compatible (including Adafruit LCD Backpack)
* Elgato Stream Deck (15-Key Version)
* Dream Cheeky 815 WebMail Notifier

## Regarding "Service Devices" ##

"Service Devices" are home automation devices which aren't intended to be directly accessed by a common local protocol, but require an API call to a proprietary company's server in order to control. **Service devices should be avoided.** Pull requests that contribute support for service devices _will_ be accepted, however platform design will not be compromised to accommodate them. Additionally, service device code must be contained within their own module, and any calls to it elsewhere must be easy to remove.
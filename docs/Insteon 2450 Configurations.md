# Insteon 2450

## "BuzzLinc"

The BuzzLinc configuration of an Insteon 2450 I/O Linc will be supported by HAController, so I wanted to document the creation of it here.

> A jumper across GND and N/O on an IOLinc and a piezoelectric buzzer connected to 5V and COM, voila, an Inseton BuzzLinc.

> The (+) terminal of the piezoelectric buzzer should be at 5V and the (-) terminal at COM.

Source: stusviews at http://forum.smarthome.com/topic.asp?TOPIC_ID=8037

## Sump Pump Alarm

The Sump Pump Alarm begins with the BuzzLinc configuration. A float switch is soldered to a headphone jack and plugged into the Sense port.

The alarm will sound regardless of communication with the controller.

Tested with Level Sense float switch, available here: https://www.amazon.com/dp/B01A4RZULM/ or https://www.level-sense.com/products/level-sense-float-switch

## Garage Door Sensor

The Garage Door Sensor uses a Seco-Larm SM-4201-LQ magnetic contact switch, available here: https://www.amazon.com/dp/B011OXPZ7Y/ and documented here: https://www.seco-larm.com/product/sm-4201-lq/
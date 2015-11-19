This is a really simple home automation controller test, which might inevitably become something larger.

As of this initial commit, almost all of the credit for code that "does things" goes to **Jonathan Dale** of http://www.madreporite.com who stated that his code from his site is "free for anyone to use". As he does not have a specific license attached to that statement (I asked), I'm going to just plaster his name up here for credit.

Basically I made a little program to send Insteon commands with his code, and that worked well and did what I wanted, but when I tried to make sense of the receiving code, I hit a roadblock. So I basically went for the concept to merge that code in directly, and then slowly integrate it with my app afterwards. I dumped it all in and got to work merging it, removing non-existant functions, and removing errors. This first commit is post-merge, but pre-making sense of it.

** Modules **

* Form1 is basically meant to get factored out eventually in it's entirety.
* modGlobal will mostly just hang onto stuff that needs to be accessible to the whole program.
* modInsteon will handle Insteon and X10 home automation devices.
* modOpenWeatherMap gathers weather data from an API to present to the user.

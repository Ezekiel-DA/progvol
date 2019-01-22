# The world's dumbest scheduling letter generator

# Running
* Download the [1.1.0 release](https://github.com/Ezekiel-DA/progvol/releases/tag/v1.1.0) for your OS
* Unzip it somewhere
* Fill out `config_example.json` with your content (see inline explanations in the file) and rename it to `config.json`
* Put your logo (as `logo.jpg`) in the same directory
* Drop `schedule.xlsx` in the same directory
* Run `lettergen-win.exe` to generate volunteer letters
* Run `badgegen-win.exe` to generate volunteer badges

# Input format
* One line per volunteer
* repeated volunteer names on subsequent lines may be omitted; the previous name will be reused until a new name is encountered
* column order __must be__: name, day, start time, end time, class title, role, location
* days and times must be parseable by Moment.js

# Notes
* don't forget to close the Word document between runs since Word keeps the file locked
* There is basically zero error handling at this time
* There are all sorts of hardcodes everywhere for my specific purpose; fitness for any specific purpose not guaranteed
* __The output document for volunteer letters uses A4 as its page size__; please change this to Letter manually in Word

# Hacking
* Clone this repo
* `npm install`
* `node lettergen.js` or `node badgegen.js`
# The world's dumbest scheduling letter generator

The NELA logo is the property of NELA. All rights reserved.

# Running
* Download the [1.0.0 release](https://github.com/Ezekiel-DA/progvol/releases/tag/v1.0.0) for your OS
* Unzip it somewhere
* Drop `schedule.xlsx` in the same directory
* Run `progvol-win.exe` (or the appropriate version for your OS)

# Input format
* One line per volunteer
* repeated volunteer names on subsequent lines may be omitted; the previous name will be reused until a new name is encountered
* column order __must be__: name, day, start time, end time, class title, role, location
* days and times must be parseable by Moment.js

# Notes
* don't forget to close the Word document between runs since Word keeps the file locked
* There is basically zero error handling at this time
* __The output document uses A4 as its page size__; please change this to Letter manually in Word

# Hacking
* Clone this repo
* `npm install`
* `node progvol.js`
# The world's dumbest scheduling letter generator

The NELA logo is the property of NELA. All rights reserved.

# Running
`node progvol.js`

The logo file needs to exist and the input data needs to be in a `schedule.xlsx` file in the same directory.

# Input format
* One line per volunteer
* repeated volunteer names on subsequent lines may be omitted; the previous name will be reused until a new name is encountered
* column order __must be__: name, day, start time, end time, class title, role, location
* days and times must be parseable by Moment.js

# Notes
* don't forget to close the Word document between runs since Word keeps the file locked
* There is basically zero error handling at this time
* __The output document uses A4 as its page size__; please change this to Letter manually in Word
# The world's dumbest scheduling letter generator

The NELA logo is the property of NELA. All rights reserved.

# Running
`node index.js`

The logo file needs to exist and the input data needs to be in a `schedule.xlsx` file in the same directory.

# Input format
* One line per volunteer
* repeated volunteer names on subsequent lines may be omitted and will be repeated automatically until the next name
* column order __must be__: name, day, start time, end time, class title, role, location
* days and times must be parseable by Moment.js

# Notes

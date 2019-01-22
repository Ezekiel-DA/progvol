const xlsx = require('node-xlsx')
const moment = require('moment')

/**
 * Ingest scheduling data from an Excel sheet, producing a map of volunteers
 * @param {String} filename - the name (or path) of the xlsx file to process
 * @returns {Object} a map of volunteer names to Objects with these properties: {name: {String} same as the key, day: {Object[]}}
 * Where the day array is an array of Objects, each representing a shift, of this form:
 * {start: {String} start time, end: {String} end time, role: {String}, classTitle: {String}}
 */
async function parseSchedule (filename) {
  let data = xlsx.parse(filename, { cellDates: true })
  let sheet = data[0].data.slice(1)

  // pre process to add names back in
  sheet.forEach((line, idx) => {
    if (!line[0]) {
      line[0] = sheet[idx - 1][0]
    }
  })
  // don't forget to drop the empty (save for a name) very last line that the dumb algorithm above added >.<
  sheet = sheet.slice(0, sheet.length - 1)

  // convert UTC dates to day of the week and AM/PM hours and minutes
  sheet = sheet.map(line => {
    line[1] = moment(line[1]).format('dddd')
    line[2] = moment(line[2]).format('h:mm a')
    line[3] = moment(line[3]).format('h:mm a')
    return line
  })

  let volunteersMap = sheet.reduce((res, line) => {
    let [name, day, start, end, classTitle, role, location] = line
    name = name.trim()
    res[name] = res[name] || { name }
    res[name][day] = res[name][day] || []
    let shift = { role: role ? role.trim() : '', classTitle: classTitle ? classTitle.trim() : '', start, end, location: String(location).trim() }
    res[name][day].push(shift)
    return res
  }, {})

  // make sure shifts are sorted in chronological order
  Object.values(volunteersMap).forEach(volunteer => {
    ['Friday', 'Saturday', 'Sunday'].forEach(day => {
      if (!volunteer[day]) { return }
      volunteer[day].sort((a, b) => {
        let aMoment = moment(a.start, 'hh:mm a')
        let bMoment = moment(b.start, 'hh:mm a')
        return aMoment.isBefore(bMoment) ? -1 : 1
      })
    })
  })

  return volunteersMap
}

module.exports = { parseSchedule }

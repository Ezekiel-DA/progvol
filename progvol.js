const xlsx = require('node-xlsx')
const process = require('process')
const moment = require('moment')
const docx = require('docx')
const fs = require('fs').promises

// Text strings for various parts of the letters
const headerP = new docx.Paragraph()
headerP.addRun(new docx.TextRun('Fetish Fair Fleamarket #52').bold()).style('basic')

const thankYouP = new docx.Paragraph()
thankYouP.addRun(new docx.TextRun('\tThank you for volunteering with the programming team at FFF#52! It is due to the contributions of volunteers like yourself that NELA is able to put on the Flea every year. Your efforts are greatly appreciated.')).style('basic')

const scheduleBelowP = new docx.Paragraph()
scheduleBelowP.addRun(new docx.TextRun('\tBelow is your Schedule for FFF#52. Please take the time to read through your packet and look over the job descriptions and information provided for you. Thank you again for volunteering with us, and have a great Flea!')).style('basic')

const extraShiftsP = new docx.Paragraph()
extraShiftsP.addRun(new docx.TextRun('If you would like to pick up any extra time and earn further perks, please let us know! We would love to have you, and there is always extra work to be done. Please contact with any questions that you may have, and have a great Flea!')).style('basic')

const signaturesP = new docx.Paragraph()
signaturesP.addRun(new docx.TextRun('Connor and Chinci'))
  .addRun(new docx.TextRun('Programming Volunteer Captains').break())
  .addRun(new docx.TextRun('NELA FFF#52').break()).style('basic')

  /**
   * Add a table for all shifts for a given day and volunteer, sorting shifts by start time
   * @param {String} day - one of Friday, Saturday or Sunday right now
   * @param {docx.Document} document
   * @param {Object} volunteer an object having at least one property that matches the input day and contains an array of shifts in this format {start, end, classTitle, role, location}
   * @returns {docx.Document} the input Document, with the table added in-place
   */
function addTable (day, document, volunteer) {
  let dayText = new docx.TextRun(day).bold()
  let [timeHeader, classHeader, positionHeader, locationHeader] = ['Time', 'Class', 'Position', 'Location'].map(text => {
    let run = new docx.TextRun(text).bold().underline()
    let p = new docx.Paragraph().style('basic')
    p.addRun(run)
    return p.center()
  })

  let dayParagraph = new docx.Paragraph().style('day')
  dayParagraph.addRun(dayText)
  document.addParagraph(dayParagraph)

  let table = document.createTable(volunteer[day].length + 1, 4)
  table.setWidth('pct', '100%')
  table.getCell(0, 0).addContent(timeHeader)
  table.getCell(0, 1).addContent(classHeader)
  table.getCell(0, 2).addContent(positionHeader)
  table.getCell(0, 3).addContent(locationHeader)

  volunteer[day].sort((a, b) => {
    let aMoment = moment(a.start, 'hh:mm a')
    let bMoment = moment(b.start, 'hh:mm a')
    return aMoment.isBefore(bMoment) ? -1 : 1
  })

  volunteer[day].forEach((shift, idx) => {
    table.getCell(idx + 1, 0).addContent((new docx.Paragraph(`${shift.start} - ${shift.end}`)).style('basic').center())
    table.getCell(idx + 1, 1).addContent((new docx.Paragraph(shift.classTitle)).style('basic').center())
    table.getCell(idx + 1, 2).addContent((new docx.Paragraph(shift.role)).style('basic').center())
    table.getCell(idx + 1, 3).addContent((new docx.Paragraph(shift.location)).style('basic').center())
  })

  return document.addParagraph(new docx.Paragraph().style('basic'))
}

/**
 * Add a page for the given volunteer to the document
 * @param {docx.Document} document
 * @param {Object} volunteer - the input volunteer, in this format:
 * @param {String} volunteer.name
 * @param {Object[]} [volunteer.Friday] - an array of Objects representing shifts, in the format outlined in addTable()
 * @param {Object[]} [volunteer.Saturday] - as above
 * @param {Object[]} [volunteer.Sunday] - as above
 */
function addVolunteer (document, volunteer) {
  let greeting = new docx.Paragraph().style('basic')
  greeting.addRun(new docx.TextRun(`Dear ${volunteer.name},`))

  document.addParagraph((new docx.Paragraph().style('basic')).pageBreakBefore())
  .addParagraph(greeting)
  .addParagraph(new docx.Paragraph().style('basic'))
  .addParagraph(thankYouP)
  .addParagraph(scheduleBelowP)
  .addParagraph(new docx.Paragraph().style('basic'))

  if (volunteer.Friday) {
    addTable('Friday', document, volunteer)
  }
  if (volunteer.Saturday) {
    addTable('Saturday', document, volunteer)
  }
  if (volunteer.Sunday) {
    addTable('Sunday', document, volunteer)
  }

  document.addParagraph(extraShiftsP).addParagraph(new docx.Paragraph().style('basic')).addParagraph(signaturesP)
  return document
}

/**
 * Expected input: schedule.xlsx and NELA_logo.jpg must exist in the same folder;
 * Expected output: an output.docx Word document containing one page/letter per volunteer detailing their schedule.
 */
async function main () {
  console.log(process.cwd())
  let data = xlsx.parse('schedule.xlsx', { cellDates: true })
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

  let volunteerMap = sheet.reduce((res, line, idx) => {
    let [name, day, start, end, classTitle, role, location] = line
    name = name.trim()
    res[name] = res[name] || { name }
    res[name][day] = res[name][day] || []
    let shift = { role: role ? role.trim() : '', classTitle: classTitle ? classTitle.trim() : '', start, end, location: String(location).trim() }
    res[name][day].push(shift)
    return res
  }, {})

  let document = new docx.Document({
    creator: 'NELA',
    title: 'volunteer schedules',
    description: 'FF#52 volunteer schedules'
  })

  document.Styles.createParagraphStyle('basic', 'basic').basedOn('Normal').next('basic').quickFormat().font('Arial').size(22)
  document.Styles.createParagraphStyle('day', 'day').basedOn('basic').next('basic').quickFormat().font('Arial').size(28)

  let logo = await fs.readFile('./NELA_logo.jpg')
  document.Header.createImage(logo, 250)
  document.Header.addParagraph(headerP)

  Object.values(volunteerMap).forEach(addVolunteer.bind(null, document))
  console.log(`Processed ${Object.keys(volunteerMap).length} volunteers.`)

  return new docx.Packer().toBuffer(document).then(fs.writeFile.bind(fs, 'output.docx'))
  .catch(err => {
    if (err.code !== 'EBUSY') {
      throw (err)
    } 
    console.log('Could not write output document; remember to close it in Word!')
  })
}

main()

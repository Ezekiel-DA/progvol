const extractor = require('./extractor.js')
const config = require('./config.json')
const docx = require('docx')
const fs = require('fs').promises

// Text strings for various parts of the letters
const headerP = new docx.Paragraph()
headerP.addRun(new docx.TextRun(config.eventTitle).bold()).style('basic')

const thankYouP = new docx.Paragraph()
thankYouP.addRun(new docx.TextRun(`\t${config.thankYou}`)).style('basic')

const scheduleBelowP = new docx.Paragraph()
scheduleBelowP.addRun(new docx.TextRun(`\t${config.scheduleIntro}`)).style('basic')

const extraShiftsP = new docx.Paragraph()
extraShiftsP.addRun(new docx.TextRun(config.scheduleOutro)).style('basic')

const signaturesP = config.signatures.reduce((signatures, signatureLine, idx) => {
  let text = new docx.TextRun(signatureLine)
  return signatures.addRun(idx ? text.break() : text)
}, new docx.Paragraph()).style('basic')

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
 * Expected input: schedule.xlsx and logo.jpg must exist in the same folder;
 * Expected output: an output.docx Word document containing one page/letter per volunteer detailing their schedule.
 */
async function main () {
  let volunteerMap = await extractor.parseSchedule('schedule.xlsx')

  let document = new docx.Document({
    creator: config.docCreator,
    title: config.title,
    description: config.description
  })

  document.Styles.createParagraphStyle('basic', 'basic').basedOn('Normal').next('basic').quickFormat().font('Arial').size(22)
  document.Styles.createParagraphStyle('day', 'day').basedOn('basic').next('basic').quickFormat().font('Arial').size(28)

  let logo = await fs.readFile('./logo.jpg')
  document.Header.createImage(logo, 250)
  document.Header.addParagraph(headerP)

  Object.values(volunteerMap).forEach(addVolunteer.bind(null, document))
  console.log(`Processed ${Object.keys(volunteerMap).length} volunteers.`)

  return new docx.Packer().toBuffer(document).then(fs.writeFile.bind(fs, 'progvol schedule letters.docx'))
    .catch(err => {
      if (err.code !== 'EBUSY') {
        throw (err)
      }
      console.log('Could not write output document; remember to close it in Word!')
    })
}

main()

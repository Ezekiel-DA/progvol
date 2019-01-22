const extractor = require('./extractor.js')
const config = require('./config.json')
const Pdfmake = require('pdfmake')
const fs = require('fs')

const fonts = {
  Roboto: { normal: Buffer.from(require('pdfmake/build/vfs_fonts.js').pdfMake.vfs['Roboto-Regular.ttf'], 'base64') }
}

/**
 * Converts units for a given value from inches to points (which is what pdfmake seems to use internally)
 * @param {Number} inch - an element size in inches
 * @returns {Number}
 */
function inchToPoint(inch) {
  return inch * 72
}

async function main() {
  let volunteerMap = await extractor.parseSchedule('schedule.xlsx')

  const widths = new Array(config.badgesPerRow).fill(inchToPoint(config.badgeWidth))
  const heights = new Array(config.badgesPerColumn).fill(inchToPoint(config.badgeHeight))

  let docDefinition = {
    pageSize: 'LETTER',
    //pageMargins: [inchToPoint(0.75), inchToPoint(1.0)],
    pageMargins: 0,
    content: [
      {
        table: {
          widths,
          heights,
          body: [
            ['Volunteer', 'Volunteer'],
            ['Volunteer', 'Volunteer'],
            ['Volunteer', 'Volunteer'],
            ['Volunteer', 'Volunteer'],
          ]
        }
      }
    ]
  }

  let pdf = (new Pdfmake(fonts)).createPdfKitDocument(docDefinition)
  pdf.pipe(fs.createWriteStream('volunteer badges.pdf'))
  pdf.end()
}

main()

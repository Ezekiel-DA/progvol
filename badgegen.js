const extractor = require('./extractor.js')
const config = require('./config.json')
const Pdfmake = require('pdfmake')
const fs = require('fs')
const R = require('ramda')

const fonts = {
  Roboto: {
    normal: 'fonts/Roboto-Regular.ttf',
    bold: 'fonts/Roboto-Medium.ttf',
    italics: 'fonts/Roboto-Italic.ttf',
    bolditalics: 'fonts/Roboto-MediumItalic.ttf'
  }
}

const styles = {
  badgesTable: { alignment: 'center', margin: 0 },
  mainBadgeText: { fontSize: 14, margin: [0, 15, 0, 0] },
  badgeTextLine: { margin: [0, 4, 0, 4] }
}

const noPadding = () => 0
const inchToPoint = inch => inch * 72

const options = {
  tableLayouts: {
    badgesLayout: {
      hLineWidth: () => 0.1, vLineWidth: () => 0.1,
      paddingLeft: noPadding, paddingRight: noPadding, paddingTop: noPadding, paddingBottom: noPadding,
    }
  }
}

const shortRoles = {
  'Door Monitor': 'Doors mon.',
  'Introducer': 'Intro.',
  'Runner': 'Runner',
  'Serenity Space Monitor': 'Srnty. spc. mon.',
  'Masterclass Helper': 'MClass helper',
  'Ticket Assistant': 'Tkt. asst.',
  'PHQ Setup': 'PHQ setup'
}

let firstTable = true
const pageBreakBefore = node => {
  if (!node.table) { return false }
  if (firstTable) { firstTable = false; return false }
  return true
}

const widths = new Array(config.badges.badgesPerRow).fill(inchToPoint(config.badges.badgeWidth))
const heights = new Array(config.badges.badgesPerColumn).fill(inchToPoint(config.badges.badgeHeight))

async function main() {
  let volunteerMap = await extractor.parseSchedule('schedule.xlsx')

  let badgeFronts = Object.values(volunteerMap).map(volunteer => {
    return [
      {
        columns: [
          { image: './logo.jpg', alignment: 'left', width: inchToPoint(1) },
          [{ text: config.badges.logoTextLine1, alignment: 'right', bold: true }, { text: config.badges.logoTextLine2, alignment: 'right' }]
        ],
        margin: 10, fontSize: 9
      },
      {
        stack: [
          { text: volunteer.name, style: 'badgeTextLine' },
          { text: '<pronouns placeholder>', style: 'badgeTextLine' },
          { text: 'Programming Volunteer', style: 'badgeTextLine', bold: true }
        ],
        style: 'mainBadgeText'
      }
    ]
  })

  let badgeBacks = Object.values(volunteerMap).map(volunteer => {
    let shifts = []
    const days = ['Friday', 'Saturday', 'Sunday']
    days.forEach(day => {
      if (volunteer[day]) {
        shifts = shifts.concat(volunteer[day].map(shift => ({ day, start: shift.start, end: shift.end, role: shift.role, classTitle: shift.classTitle, location: shift.location })))
      }
    })
    const shiftTextSize = numShifts => numShifts < 6 ? 9 : 7
    const topMargin = numShifts => {
      let offset = 10 - numShifts
      return offset > 0 ? offset * (shiftTextSize(numShifts) - 2) : 5
    }

    return [
      { text: `${volunteer.name}'${volunteer.name.match(/(\w)\W?$/)[1].toLowerCase() === 's' ? '' : 's'} schedule:`, style: 'mainBadgeText', fontSize: 10 },
      {
        stack: shifts.map(shift => ({
          text: `${shift.day.slice(0, 3)}. ${shift.start}-${shift.end}: ${shortRoles[shift.role]}@${shift.location}${(shift.classTitle && shift.classTitle !== '---') ? ` (${shift.classTitle})` : ''}`, fontSize: shiftTextSize(shifts.length), alignment: 'left'
        })),
        margin: [5, topMargin(shifts.length), 0, 0]
      }
    ]
  })

  const reshape = R.pipe(R.splitEvery(2), R.splitEvery(4))
  badgeFronts = reshape(badgeFronts)
  badgeBacks = reshape(badgeBacks)

  let docDefinition = {
    pageSize: 'LETTER',
    pageMargins: [inchToPoint(config.badges.horizontalMargin), inchToPoint(config.badges.verticalMargin)],
    content: R.zip(
      badgeFronts.map(eightVolunteers => ({ style: 'badgesTable', layout: 'badgesLayout', table: { widths, heights, body: eightVolunteers } })),
      badgeBacks.map(eightVolunteers => ({ style: 'badgesTable', layout: 'badgesLayout', table: { widths, heights, body: eightVolunteers } }))
    ),
    styles,
    pageBreakBefore
  }

  let pdf = (new Pdfmake(fonts)).createPdfKitDocument(docDefinition, options)
  pdf.pipe(fs.createWriteStream('volunteer badges.pdf'))
  pdf.end()
}

main()




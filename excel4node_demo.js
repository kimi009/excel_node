// Require library
var xl = require('excel4node')

// Create a new instance of a Workbook class
var wb = new xl.Workbook({
	defaultFont: {
		size: 16,
		name: 'Calibri',
		color: 'FFFFFFFF',
	},
	// dateFormat: 'm/d/yy hh:mm:ss',
})

// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1')
ws.column(1).setWidth(30)
// Create a reusable style
var style = wb.createStyle({
	font: {
		color: '#FF0800',
		size: 14,
	},
	numberFormat: '$#,##0.00; ($#,##0.00); -',
})

// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1, 1).number(100).style(style)

// Set value of cell B1 to 200 as a number type styled with paramaters of style
ws.cell(1, 2).number(200).style(style)

// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 3).formula('A1 + B1').style(style)

// Set value of cell A2 to 'string' styled with paramaters of style
ws.cell(2, 1).string('string').style(style)

// Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
ws.cell(3, 1)
	.bool(true)
	.style(style)
	.style({ font: { size: 14 } })


// const date = xl.getExcelTS(newDate)
ws.row(4).setHeight(75)
ws.cell(4, 1)
	.date(new Date(+new Date() + 8 * 3600 * 1000).toISOString())
	.style({
		font: {
			color: '#FF0800',
			size: 18,
		},
		// numberFormat: 'yyyy/m/d hh:mm',
		numberFormat: 'yyyy/mm/dd hh:mm:ss',
	})

ws.cell(5, 1)
	.date(new Date('2023-02-03T10:05:54Z'))
	.style({
		font: {
			color: '#FF0800',
			size: 18,
		},
		numberFormat: 'yyyy/mm/dd hh:mm:ss',
	})

wb.write('excel4node.xlsx')

console.log(new Date().toLocaleString())

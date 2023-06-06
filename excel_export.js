// import XLSX from 'xlsx';
const XLSX = require('xlsx-style')
const dayjs = require('dayjs')

// 这里用的是模拟数据
const headers = [
	{ title: '学员名字', dataIndex: 'name', width: 140, type: 's' },
	{ title: '性别', dataIndex: 'sex', width: 50, type: 's' },
	{ title: '联系方式', dataIndex: 'phone', width: 140, type: 'n' },
	{ title: '状态', dataIndex: 'status', width: 140, type: 's' },
	{ title: '扣除课时', dataIndex: 'deduct', width: 100, type: 's' },
	{ title: '已完成/总课时', dataIndex: 'number', width: 130, type: 's' },
	{ title: '入学日期', dataIndex: 'admissionDate', width: 130, type: 'd' },
]
const options = [
	{ title: '这里是标题哦' },
	// { title: '副标题1' },
	// { title: '副标题2' },
	// { title: '副标题3' },
]

/**
 * 定制化导出excel（定制化：附加标题&&样式）
 * @param { 数据源  } datasource
 * @param { 表格副标题 } options
 * @param { 配置文件类型 } type
 * @param { 导出的文件名 } fileName
 */
function exportExcel(datasource, fileName = '未命名') {
	// 处理列宽
	const cloWidth = headers.map((item) => ({ wpx: item.width || 60 }))

	// 处理附加表头
	const _options = options
		.map((item, i) =>
			Object.assign(
				{},
				{
					title: item.title,
					position: String.fromCharCode(65) + (i + 1),
				}
			)
		)
		.reduce(
			(prev, next) =>
				Object.assign({}, prev, {
					[next.position]: {
						v: next.title,
						s: {
							font: {
								sz: 18,
								bold: true,
								vertAlign: true,
								color: { rgb: 'FF0000' },
							},
							alignment: { vertical: 'center', horizontal: 'center' },
							fill: {
								patternType: 'solid',
								bgColor: { rgb: 'FFF000' },
								fgColor: { rgb: 'FFFF00' },
							},
						},
					},
				}),
			{}
		)

	// 处理表头
	const _headers = headers
		.map((item, i) =>
			Object.assign(
				{},
				{
					key: item.dataIndex,
					title: item.title,
					position: String.fromCharCode(65 + i) + (options.length + 1),
				}
			)
		)
		.reduce(
			(prev, next) =>
				Object.assign({}, prev, {
					[next.position]: {
						v: next.title,
						key: next.key,
						s: {
							font: { sz: 14, bold: true, vertAlign: true },
							alignment: { vertical: 'center', horizontal: 'right' },
							fill: { bgColor: { rgb: 'FFF000' }, fgColor: { rgb: 'FABF8F' } },
						},
					},
				}),
			{}
		)

	// 处理数据源
	let firstStep = datasource.map((item, i) =>
		headers.map((col, j) => {
			let colValue = item[col.dataIndex]
			// if (col.type === 'd') {
			// 	//日期类型
			// 	// colValue = dayjs(colValue).add(1, 'day').format('YYYY/MM/DD HH:mm:ss')
			// 	colValue = dayjs().format('YYYY-MM-DD HH:mm:ss')
			//   console.log(dayjs().format('YYYY-MM-DD HH:mm:ss'))
			// }
			return Object.assign(
				{},
				{
					content: colValue,
					type: col.type,
					position: String.fromCharCode(65 + j) + (options.length + i + 2),
				}
			)
		})
	)
	let secondStep = firstStep.reduce((prev, next) => prev.concat(next))
	// sheetObject = {
	//   A1: {
	//     v: '单元格',
	//     t: 's',
	//     s: {
	//       font: {},
	//       fill: {},
	//       numFmt: {},
	//       alignment: {},
	//       border: {}
	//     }
	//   }
	// }
	// v: 表示单元格的值；

	// t：表示单元格值的类型，b：表示Boolean布尔值，n表示number数组，e表示error错误信息，s表示string字符串，d:表示date日期

	// s: 表示单元格的样式

	const _data = secondStep.reduce(
		(prev, next) =>
			Object.assign({}, prev, {
				[next.position]: Object.assign(
					{},
					{
						v: next.content,
						t: next.type,
					},
					next.type === 'd'
						? {
								t: 'n',
								z: 'yyyy-mm-dd hh:mm:ss',
						  }
						: {}
				),
			}),
		{}
	)

	const output = Object.assign({}, _options, _headers, _data)
	const outputPos = Object.keys(output) // 设置表格渲染区域,如从A1到C8

	// 设置单元格样式！！！！   仅xlsx-style生效，js-xlsx写了也不生效
	// 这里对每个单元格设置样式是写死的，每次改样式改都要改这里有点鸡肋
	// output.A1.s = {
	// 	font: { sz: 18, bold: true, vertAlign: true, color: { rgb: 'FF0000' } },
	// 	alignment: { vertical: 'center', horizontal: 'center' },
	// 	fill: {
	// 		patternType: 'solid',
	// 		bgColor: { rgb: 'FFF000' },
	// 		fgColor: { rgb: 'FFFF00' },
	// 	},
	// 	border: {
	// 		bottom: {
	// 			style: 'thin',
	// 			color: { auto: 1 },
	// 		},
	// 	},
	// }
	// output.A2.s = {
	// 	font: { sz: 14, bold: true, vertAlign: true },
	// 	alignment: { vertical: 'center', horizontal: 'right' },
	//   fill: { bgColor: { rgb: 'FFF000' }, fgColor: { rgb: 'FABF8F' } },
	// }
	output.A3.s = {
		font: { sz: 12, bold: true, vertAlign: true },
		alignment: { vertical: 'center', horizontal: 'bottom' },
	}
	output.A4.s = {
		font: { sz: 12, bold: true, vertAlign: true },
		alignment: { vertical: 'center', horizontal: 'bottom' },
	}
	// output.G3.s = {
	// 	font: { sz: 15, bold: true, vertAlign: true },
	// 	alignment: { vertical: 'center', horizontal: 'bottom' },
	// 	numFmt: 'yyyy-mm-dd hh:mm:ss',
	// }

	// 合并单元格
	const merges = options.map((item, i) => ({
		s: { c: 0, r: i },
		e: { c: headers.length - 1, r: i },
	}))
	// console.log(output)

	const wb = {
		SheetNames: ['mySheet'], // 保存的表标题
		Sheets: {
			mySheet: Object.assign(
				{},
				output, // 导出的内容
				{
					'!ref': `${outputPos[0]}:${outputPos[outputPos.length - 1]}`, // 设置填充区域（表格渲染区域）
					'!cols': [...cloWidth],
					'!merges': [...merges],
				}
			),
		},
	}

	// 这种导出方法只适用于js-xlsx，且设置的单元格样式不生效，
	// 直接打开下面这两行就行了，后面的可以省略
	XLSX.writeFile(wb, `${fileName}.xlsx`)
	// XLSX.writeFile(wb, `${fileName}.xlsx`, { cellDates: true })
}

module.exports = exportExcel
// write_ws_xml_cell

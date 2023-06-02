const XLSX = require('xlsx-style')
const path = require('path')
const exportExcel = require('./excel_export')

const fileName = path.resolve(__dirname, '测试导出数据.xlsx')
const wb = XLSX.readFile(fileName, {
	cellDates: true,
})

const ws = wb.Sheets['mySheet']

const ress = XLSX.utils.sheet_to_json(ws, { header: 1, range: 2 })
// console.log(13, ress)
let oldData = ress.reduce((pre, cur) => {
	const [name, sex, phone, status, deduct, number, admissionDate] = cur
	pre.push({
		name,
		sex,
		phone,
		status,
		deduct,
		number,
		admissionDate: new Date(new Date(admissionDate).getTime() + 8 * 3600 * 1000).toISOString(),
	})
	return pre
}, [])

console.log(oldData)

// 新行数据
const newRowData = {
	name: '我是新来的',
	sex: '女',
	phone: '12345678909',
	status: '已签到',
	deduct: 1,
	number: '3/10',
	admissionDate: new Date(new Date('2023/3/2 22:23:11').getTime() + 8 * 3600 * 1000).toISOString(),
}

exportExcel([...oldData, newRowData], '测试导出数据2')

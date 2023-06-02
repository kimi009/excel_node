const exportExcel = require('./excel_export')

function formatDate(date) {
	let newDate = ''
	if (date) {
		newDate = new Date(date).getTime()
	} else {
		newDate = +new Date()
	}
	return new Date(newDate + 8 * 3600 * 1000).toISOString()
}

const datasource = [
	{
		name: '张三',
		sex: '男',
		phone: '12345678909',
		status: '已签到',
		deduct: 1,
		number: '1/10',
		admissionDate: formatDate(),
	},
	{
		name: '李四',
		sex: '男',
		phone: '12345678909',
		status: '旷课',
		deduct: 1,
		number: '1/10',
		admissionDate: formatDate('2022-2-22'),
	},
	{
		name: '王丽',
		sex: '女',
		phone: '12345678909',
		status: '请假',
		deduct: '-',
		number: '0/10',
		admissionDate: formatDate('2023/09/01'),
	},
	{
		name: '赵钱',
		sex: '男',
		phone: '12345678909',
		status: '已签到',
		deduct: 1,
		number: '1/10',
		admissionDate: formatDate('2023/09/03'),
	},
	{
		name: '孙李',
		sex: '男',
		phone: '12345678909',
		status: '已签到',
		deduct: 1,
		number: '1/10',
		admissionDate: formatDate('2023/8/02'),
	},
	{
		name: '马上飘',
		sex: '男',
		phone: '12345678909',
		status: '已签到',
		deduct: 1,
		number: '1/10',
		admissionDate: formatDate('2023/09/01'),
	},
]

exportExcel(datasource, '测试导出数据')

// console.log(new Date('2023/09/01 12:23:23').toLocaleString())

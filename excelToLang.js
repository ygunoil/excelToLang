//导入
var fs = require('fs')
var path = require('path')
const xlsx = require('xlsx');

let workbook = xlsx.readFile('./aa.xlsx');
let sheetNames = workbook.SheetNames;
var obj = {}
var objEn = {}
sheetNames.shift()
sheetNames.pop()
sheetNames.map((v, i) => {
	var encode = v.split('-')[0]
	var sheet = workbook.Sheets[sheetNames[i]]
	var range = xlsx.utils.decode_range(sheet['!ref']);
	var arr = []
	// //循环获取单元格值
	for(let R = range.s.r; R <= range.e.r; ++R) {
		let row_value = [];
		for(let C = range.s.c; C <= range.e.c; ++C) {
			let cell_address = {
				c: C,
				r: R
			}; //获取单元格地址
			let cell = xlsx.utils.encode_cell(cell_address); //根据单元格地址获取单元格
			if(cell != 'A1' && cell != 'B1' && cell != 'C1') {
				//获取单元格值
				if(sheet[cell]) {
					// 如果出现乱码可以使用iconv-lite进行转码
					// row_value += iconv.decode(sheet1[cell].v, 'gbk') + ", ";
					//    if(sheet1[cell].v != 'encode')
					row_value.push(sheet[cell].v)
				} else {
					//     row_value += ", ";
				}
			}
		}
		if(row_value.length) {
			arr.push(row_value)
		}
	}
	obj[encode] = arr
})
Object.keys(obj).map((v, i) => {
	var o = {}
	var oEn = {}
	obj[v].map(item => {
//		console.log(item)
		if(+item[0]<10){
			item[0] = '0'+ item[0]
		}
		o[+item[0]] = item[1]
		if(item[2]== undefined || item[2]== '' || item[2] > 0){
			item[2]='aaa'
		}
		oEn[+item[0]] = item[2]
	})
	obj[v]=o
	objEn[v]=oEn
})
//console.log(obj)
fs.writeFile(path.resolve(__dirname, './zh-CN.json'),JSON.stringify(obj),(err)=>{
		if (err) {
        	console.error(err);
        }
      console.log('----------新增zh-CN.json成功-------------');
})
fs.writeFile(path.resolve(__dirname, './en.json'),JSON.stringify(objEn),(err)=>{
		if (err) {
        	console.error(err);
        }
      console.log('----------新增en.json成功-------------');
})

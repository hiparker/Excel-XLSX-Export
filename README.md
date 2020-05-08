> 详情地址 ： https://www.arcinbj.com/

### 一、需求描述
在很多OA或者CRM项目中，基本上都会涉及到Excel的导入导出的问题。
首先想到了**POI**和阿里的**EasyExcel**。
如果是小打小闹，导几千数据玩玩，服务器本身基本没什么压力，但是动辄导出上万的数据，那服务器肯定是吃不消的（这里指的是没有对导出Excel服务器做优化或者负载处理）

### 二、设计思路
**传统Java后端导出Excel思路**

![前端export2](https://www.arcinbj.com/upload/2020/05/前端export2-e3d15d86a62c446d912cadbbc5dd2768.jpg)

++1.导出Excel，如果在Java后端的话，且导出的数据量比较大，且又处于高并发的情况，服务器内存会被瞬间占满（如果数据量较大，POI会有内存泄漏的风险），CPU占用率也会持续升高（Excel生成二进制文件，是非常吃CPU性能的）++


**前端JavaScript导出Excel思路**

![前端export3](https://www.arcinbj.com/upload/2020/05/前端export3-df51f8985d104c20b7dea8c2dc0b15f4.jpg)

++2.但是 如果把 生成Excel的工作交给前端浏览器去完成，后端这是做一个数据发包，而浏览器拿到数据后在自己本地客户端执行生成文件，占用的CPU资源也是客户端的，即使再大的数据也对服务端没有太大影响++

### 三、技术框架
SheetJS（又名js-xlsx，npm库名称为xlsx，node库也叫node-xlsx，以下简称JX），免费版不支持样式调整。

（顺便吐槽下这些名字乱的不行。。实际上又是同一个东西= =

> JX官方说明文档：https://github.com/SheetJS/js-xlsx

XLSX-Style（npm库命名为xlsx-style，以下简称XS）基于JX二次开发，使其支持样式调整，但其开发停留在2017年，所基于的JX版本老旧，缺失许多方法。因而诞生了这个项目。

> XS官方说明文档：https://github.com/protobi/js-xlsx

XLSX-Style-Utils：其本体为xlsxStyle.utils.js

> XSU原作者开源地址 https://github.com/Ctrl-Ling/XLSX-Style-Utils 以下简称 XSU

XLSX-Export-Utils：其本体为xlsxExport.utils.js 以下简称 XEU

++本项目开源地址 ☆☆☆☆☆++
> 本项目开源地址 https://github.com/hiparker/Excel-XLSX-Export

### 四、兼容性
![前端export6-兼容](https://www.arcinbj.com/upload/2020/05/前端export6-兼容-f5141375611b4d729696899bca8fd80e.png)

### 五、核心包描述
![前端export1-core](https://www.arcinbj.com/upload/2020/05/前端export1-core-6e24ae32064c49479166c99eb01a3b9a.jpg)
> xlsx.core.min.js JX最新版核心文件，建议在将网页表格导成workbook时使用其方法

> xlsxStyle.core.min.js XS最新版核心文件，因为其原本命名与JX一样，避免冲突改名成xlsxStyle

> xlsxStyle.utils.js 基于XS的方法二次封装，更好的控制导出excel的样式。以下简称XSU

> xlsxExport.utils.js XEU本项目核心文件，基于XS 与 XSU的方法二次封装，更好的控制导出excel的样式。以下简称XEU

### 六、代码解析
> excelExport.html
```
<html lang="zh">
    <head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="author" content="Parker Zhou">
		<title>JavaScript导出Excel</title>
    </head>
 
    <body link="blue" vlink="purple">
		<div style="margin:50 auto;width:90%">
			<table id="print-content" border="1" cellpadding="0" cellspacing="0" style='border-collapse:collapse;table-layout:fixed;'></table>
			
			<br>
			<!-- 导出文件-->
			<input type="button" onclick="excelExport()" value="导出表格" ></input>
		<div>
		
	<!-- 引入文件保存js-->
	<script src="js/sheetjs/xlsx.core.min.js" ></script>
	<script src="js/sheetjs/xlsxStyle.core.min.js" ></script>
	<script src="js/sheetjs/xlsxStyle.utils.js" ></script>
	<script src="js/sheetjs/xlsxExport.utils.js" ></script>
	<script>
		// 数据
		var data = {
			"success":true,
			"errorCode":"-1",
			"msg":"导出成功",
			"body":{
				"title":"个人信息",
				"excelData":[
					["序号","姓名","年龄","性别","手机","邮箱","金额","创建日期"],
					[1,"周一",28,"男","13888888881","1@q.com",4123.3,"2020-05-01"],
					[2,"崔二",25,"女","13888888882","2@q.com",23432,"2020-05-03"],
					[3,"张三",15,"男","13888888883","3@q.com",433.14,"2020-05-02"],
					[4,"李四",27,"男","13888888884","4@q.com",6523,"2020-05-01"],
					[5,"王五",18,"男","13888888885","5@q.com",411.36,"2020-05-04"],
					[6,"赵六",21,"男","13888888886","6@q.com",1234,"2020-05-08"],
					[7,"唐七",22,"女","13888888887","7@q.com",4321.75,"2020-05-05"],
					[8,"范八",19,"男","13888888888","8@q.com",4322,"2020-05-06"],
					[9,"薛九",31,"女","13888888889","9@q.com",56465,"2020-05-01"],
					[10,"闫十",45,"男","13888888810","10@q.com",7864,"2020-05-07"]
				]
			}
		};	
		
		// 导出excel
		function excelExport(){
			if(data.success){
				if(null != data.body && undefined != data.body){
					// 调取封装方法-导出excel
					XSExport.excelExport(
						data.body.excelData,
						data.body.title
					);
				}
			}
		}
		
		
		
		
		// ---------------------------------------------------------------------------------
		// 以下不重要
		
		// 页面测试展示使用 - 创建页面表格
		function createTableElement(data){
			var tableBodyHtml = "<tr><td colspan='"+data.body.excelData[0].length+"' style='text-align: center;font-size:22px;font-family: Arial;font-weight: bold;height: 40px;'>"+data.body.title+"</td></tr>";
			// 生成Element
			data.body.excelData.forEach(function(val,index){
				var trBodyHtml = '<tr height="20" style="text-align: center;font-size:12px">';
				val.forEach(function(value){
					// 第一行小标题
					if(0 === index){
						trBodyHtml += '<td style="font-weight: bold;background-color:#808080;color:#ffffff">';
						trBodyHtml += '<div title="'+value+'" style="width: 125px;height: 16px;text-overflow: ellipsis;white-space: nowrap;overflow: hidden;">'+value+'</div>';
						trBodyHtml += '</td>';
					}else{
						trBodyHtml += '<td style="">';
						trBodyHtml += '<div title="'+value+'" style="width: 125px;height: 16px;text-overflow: ellipsis;white-space: nowrap;overflow: hidden;">'+value+'</div>';
						trBodyHtml += '</td>';
					}
				});
				trBodyHtml += '</tr>';
				tableBodyHtml += trBodyHtml;
			});
			document.querySelector('#print-content').innerHTML  = tableBodyHtml;
		}
		createTableElement(data);
		
	</script>
    </body>
</html>
```


> xlsxExport.utils.js
```
/*
@author Parker
@version 2020-05-08
@aim 对xlsx-style方法进行二次封装 方便调用以导出带样式Excel
@aim 对 XSU 进行封装和调用
@usage XSExport.xxxx()

依赖于
	1. xlsx.core.min.js
	2. xlsxStyle.core.min.js
	3. xlsxStyle.utils.js
	
*/
var XSExport = {};

/**
 * 通用的打开下载对话框方法，没有测试过具体兼容性
 * @param url 下载地址，也可以是一个blob对象，必选
 * @param saveName 保存文件名，可选
 */
XSExport.openDownloadDialog = function(url, saveName){
	if(typeof url == 'object' && url instanceof Blob)
	{
		url = URL.createObjectURL(url); // 创建blob地址
	}
	var aLink = document.createElement('a');
	aLink.href = url;
	aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
	var event;
	if(window.MouseEvent) event = new MouseEvent('click');
	else
	{
		event = document.createEvent('MouseEvents');
		event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
	}
	aLink.dispatchEvent(event);
}

/**
 * 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
 * @param sheet sheet数据
 * @param sheetName excel页内签
 */
XSExport.sheet2blob = function(sheet, sheetName) {
	var that = this;
	sheetName = sheetName || 'sheet1';
	var workbook = {
		SheetNames: [sheetName],
		Sheets: {}
	};
	workbook.Sheets[sheetName] = sheet;
	// 生成excel的配置项
	var wopts = {
		bookType: 'xlsx', // 要生成的文件类型
		bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
		type: 'binary'
	};
	
	// 设置样式
	that.setWorkbookStyle(workbook,workbook.SheetNames[0]);
	
	var wbout = xlsxStyle.write(workbook,wopts);
	var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
	// 字符串转ArrayBuffer
	function s2ab(s) {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	}
	
	return blob;
}

/**
 * 自定义对应表格样式
 * @param wb workbook Excel工作布
 * @param sheetName excel页内签名称
 */
XSExport.setWorkbookStyle = function(wb,sheetName){			
	var cols = XSU.getMaxCol(wb,sheetName);//当前最大列
	
	//wb样式处理，调用xlsxStyle.utils方法

	//------------------通用表格样式----------------------------
	XSU.mergeCells(wb,sheetName,"A1",cols); //合并title单元格
	XSU.setFontTypeAll(wb,sheetName,'Arial');//字体：Arial
	XSU.setFontSizeAll(wb,sheetName,10);//字体大小：10
	XSU.setAlignmentHorizontalAll(wb,sheetName,'center');//垂直居中
	XSU.setAlignmentVerticalAll(wb,sheetName,'center');//水平居中
	XSU.setBorderDefaultAll(wb,sheetName);//设置所有单元格默认边框
	//-------------------------个性化----------------------------
	//列宽设置 1wch为1英文字符宽度 (统一放大一下宽度)
	XSU.setColWidthAll(wb,sheetName,15);
	
	//设置A 行 主标题 默认样式 必须最后设置 否则可能会被其他覆盖
	XSU.setTitleStylesDefault(wb,sheetName);
	//设置B 行 小标题 默认样式 必须最后设置 否则可能会被其他覆盖
	XSU.setSecondRowStylesDefault(wb,sheetName);
}


/**
 * 导出excel
 * @param data 原始数据
 *        数据格式为 第一行是 小标题 后续行则是对应行数据
 *		  例：
 *			data = [
 *				['姓名','年龄','性别'],
 *				['张三','12','男'],
 *				['李四','18','女']
 *			]
 * @param title 标题名称（用于excel内 第一行标题 和 导出文件名）
 */
XSExport.excelExport = function(data,title){	
	var that = this;
	var aoa = data;	
	// 插入头部
	var header = [];
	header.push(title);
	var cols = aoa[0].length;
	for(var i=0;i<cols-1;i++){
		header.push("");	
	}
	aoa.unshift(header);
	
	// 生成sheet
	var sheet = XLSX.utils.aoa_to_sheet(aoa);
	// 二进制文件
	var blob = that.sheet2blob(sheet);
	
	that.openDownloadDialog(blob, title+that.dateToStr('yyyyMMddHHmmss')+'.xlsx');
}
/**
 * 日期对象转换为指定格式的字符串
 * f 日期格式,格式定义如下 yyyy-MM-dd HH:mm:ss
 *  date Date日期对象, 如果缺省，则为当前时间
 *
 * YYYY/yyyy/YY/yy 表示年份
 * MM/M 月份
 * W/w 星期
 * dd/DD/d/D 日期
 * hh/HH/h/H 时间
 * mm/m 分钟
 * ss/SS/s/S 秒
 * string 指定格式的时间字符串
 */
XSExport.dateToStr = function(formatStr, date){
	
	formatStr = arguments[0] || "yyyy-MM-dd HH:mm:ss";
	date = arguments[1] || new Date();
	var str = formatStr;
	var Week = ['日','一','二','三','四','五','六'];
	str=str.replace(/yyyy|YYYY/,date.getFullYear());
	str=str.replace(/yy|YY/,(date.getYear() % 100)>9?(date.getYear() % 100).toString():'0' + (date.getYear() % 100));
	str=str.replace(/MM/,date.getMonth()>=9?(date.getMonth() + 1):'0' + (date.getMonth() + 1));
	str=str.replace(/M/g,date.getMonth());
	str=str.replace(/w|W/g,Week[date.getDay()]);

	str=str.replace(/dd|DD/,date.getDate()>9?date.getDate().toString():'0' + date.getDate());
	str=str.replace(/d|D/g,date.getDate());

	str=str.replace(/hh|HH/,date.getHours()>9?date.getHours().toString():'0' + date.getHours());
	str=str.replace(/h|H/g,date.getHours());
	str=str.replace(/mm/,date.getMinutes()>9?date.getMinutes().toString():'0' + date.getMinutes());
	str=str.replace(/m/g,date.getMinutes());

	str=str.replace(/ss|SS/,date.getSeconds()>9?date.getSeconds().toString():'0' + date.getSeconds());
	str=str.replace(/s|S/g,date.getSeconds());

	return str;
}
```

### 七、效果展示
![前端export4](https://www.arcinbj.com/upload/2020/05/前端export4-5fe19b79cf544acd8cac023ec8ea98fa.jpg)

![前端export5](https://www.arcinbj.com/upload/2020/05/前端export5-83906ecb8d104703871477c28ba54093.jpg)

var XLS = require('xlsjs'),
	 fs = require('fs'),
	 sql = require('mssql'),
	 dir ="C:\\DIR\\rasp\\", 		// папка с файлами для обработки
	 dir2 ="C:\\DIR\\rasp_back\\",	// папка с обработанными файлами
	 dir3 ="C:\\DIR\\Plan\\";	// общая папка для BO

var config = {
    user: '***',
    password: '***',
    server: '10.1.**.**', 
    database: 'REGLAM'
}

/*
	Получаю нижнюю ячейку 
*/	 
var f=function(cell){
	var temp = cell.substr(1) 
	return cell.substr(0,1)+(temp*1+1)
}
/*
	Подготовка для вставки в БД
*/
var encode=function(a){
	var temp="";
	switch(a){
		case "КР":  temp="KR"; break;
		case "ПЛ":  temp="PL"; break;
		case "ПВ":  temp="PV"; break;
		case "ЦС":  temp="CS"; break;
		case "ЦСС": temp="CSS";break;
		case "ЦСТ": temp="CST";break;
		case "РФР": temp="RFR";break;
		case "ПР":  temp="PR"; break;
		case "ТР":  temp="TR"; break;
		case "ЦМВ": temp="CMV";break;
		case "ЗРВ": temp="ZRV";break;
		case "МНВ": temp="MNV";break;
		case "ФТГ": temp="FTG";break;
		default: temp=a;
	}
return temp;	
}
/*
	Запись в БД
*/
var insert=function(t,t1){
	
	t1[0]="'"+t1[0].replace(/\//g,".")+"'"
	console.log(t1.toString())
	var request = new sql.Request();
	request.query("INSERT INTO rvprostoi ("+t.toString()+") VALUES ("+t1.toString()+")",
		function(err, recordset){
			if (err)
				console.dir(err);
		}
	);
}
		
/*
	Разбор файлов *.xls
*/
var read = function(name){
	var workbook = XLS.readFile(dir+name);
	var sheet_name_list = workbook.SheetNames;
	//var Sheet1A1 = workbook.Sheets[sheet_name_list[0]]['C3'].v; 					// имя листа
	
	var all = Object.keys(workbook.Sheets[sheet_name_list[0]]).slice(0,-2) 			// выкидываю 2 полседних служебных элемента
	
	var t=['date'],t1=[workbook.Sheets[sheet_name_list[0]][f(all[0])].w]
	for (var i=1;i<all.length/2;i++){
		t.push(encode(workbook.Sheets[sheet_name_list[0]][all[i]].v.trim()))		// род вагона
		t1.push(workbook.Sheets[sheet_name_list[0]][f(all[i])].v)					// кол-во вагонов
	}
	
	insert(t,t1)
	fs.rename(dir+name, dir2+name, function(){
		console.log(name+" was processed...")
	});
	
}
 
/*
	Подключаюсь к БД, читаю все файлы в папке, разбираю их и записываю в БД
*/
sql.connect(config, function(err) {
	fs.readdirSync(dir).map(function(name){
		read(name)
	})
	
	var stream = fs.createWriteStream(dir3+"rv_prostoi.txt");
	stream.write('date_rvpr pv_rvpr\n');
	var request = new sql.Request();
	request.query("select date, pv from rvprostoi order by date",
		function(err, recordset){
			if (err)
				console.dir(err);
				recordset.map(function(a){
					stream.write(a.date+" "+a.pv+'\n');
				})
		}
	);
	
sql.close()
});






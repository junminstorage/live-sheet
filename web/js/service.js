app.service('ExcelService', function(){
	
	//X: sheet parser, either XLSX or XLS, default to XLSX
	var X = XLSX;
	//array of sheet names
	var names = [];
	//hash of worksheets with key on sheet names
	var workSheets = {};
	//the original workbook
	var wb;
	
	function handleFile(files, cb){
		if(!files)
			return
		//files can be either by user select or drag-n-drop
	    //var files = e.target.files || e.dataTransfer.files; 
	    var f = files[0];
	    var reader = new FileReader();
	    var name = f.name;
	    reader.readAsBinaryString(f);
	    reader.onload = function(e){
	        var data = e.target.result;
	        var xls = [0xd0, 0x3c].indexOf(data.charCodeAt(0)) > -1;
	        X = xls? XLS : XLSX;
	        wb = X.read(data, {type: 'binary'});
	        names = wb.SheetNames;
	        workSheets = to_jsonArray(wb);
	        cb();
	    }
	}
	
	function getWorkbookAsArrayBuffer(){
		var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
		return _s2ab(wbout);
	}
	
	
	function _s2ab(s) {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	}
	
	function to_jsonArray(wb){
		var json = {}
		wb.SheetNames.forEach(function(sheet){		
			json[sheet] =  _getJsonArrayFor(wb.Sheets[sheet]);
		})
		return json;
	}
	
	function _getJsonArrayFor(sheet, opts) {
		var jsonArray = [], txt = "", qreg = /"/g;
		var o = opts == null ? {} : opts;
		if(sheet == null || sheet["!ref"] == null) return "";
		var r = X.utils.safe_decode_range(sheet["!ref"]);
		var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
		var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
		var row = [], rr = "", cols = [];
		var i = 0, cc = 0, val;
		var R = 0, C = 0;
		for(C = r.s.c; C <= r.e.c; ++C) cols[C] = X.utils.encode_col(C);
		for(R = r.s.r; R <= r.e.r; ++R) {
			row = [];
			rr = X.utils.encode_row(R);
			for(C = r.s.c; C <= r.e.c; ++C) {
				val = sheet[cols[C] + rr];
				txt = val !== undefined ? ''+X.utils.format_cell(val) : "";
				//for(i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34) {
				//	txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
				row.push(txt);
			}
			
			jsonArray.push(row);
		}
		return jsonArray;
	}
	
	function _process_wb(wb){
		sheets = wb.SheetNames;
		
	    printSheets(wb);
	    workSheets = to_jsonArray(wb);
	    var sheet = wb.SheetNames[0];
	    createsheet(workSheets[sheet]);
	    
	    //dump excel obj to out div
	    var json = to_json(wb);
	    output = JSON.stringify(to_json(wb), 2, 2);
	    output = to_csv(wb);
	    //output = to_formulae(wb);
	    var out = document.getElementById('out');
	    out.innerText = output;
	    
	    return;
	}
	
	return {
		handleFile: handleFile,
		getSheetNames : function(){return names;},
		getSheets : function(){return workSheets;},
		getWorkbook : function(){return wb;},
		getWorkbookAsArrayBuffer : getWorkbookAsArrayBuffer
		
	}
});
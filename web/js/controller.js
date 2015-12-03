app.controller('ExcelController', function($scope, $location,  ExcelService){
		var self = this;
		$scope.sheets = {};
		$scope.names = [];
		$scope.cacheId = $location.search().id; 
		
		if($scope.cacheId) {
			console.log("load sheet: " + $scope.cacheId);
			ExcelService.cacheGet($scope.cacheId);
			
			$scope.names = ExcelService.getSheetNames();	
			var out = $('#sheets'); out.html("");
			out.append("<p>List of worksheets in the uploaded excel</p>");
			$scope.names.forEach(function(sheet){		
				out.append("<span class='sheet' data-sheet='" + sheet +"'>" + sheet + "</span></br>");
			})
			//default is to display the first sheet
			createsheet(ExcelService.getSheets()[$scope.names[0]]);
		}
		
		//file upload
		$scope.uploadFile = function(that){			
			ExcelService.handleFile(that.files, cb);			
		}
		
		//file drag n drop
		function FileSelectHandler(e) {
			// cancel event and hover styling
			FileDragHover(e);
			ExcelService.handleFile(e.dataTransfer.files, cb);
		}
		
		$scope.displaySheet = function(sheet){
			createsheet(ExcelService.getSheets()[sheet]);	
		}
		
		$("#sheets").on("click", "span.sheet", function(e){
			console.log($(this).attr('data-sheet'));
			var sheet = $(this).attr('data-sheet');
			createsheet(ExcelService.getSheets()[sheet]);	
			
		});
		
		//file drag and drop
		var filedrag = $id("filedrag");

		// file drop
		filedrag.addEventListener("dragover", FileDragHover, false);
		filedrag.addEventListener("dragleave", FileDragHover, false);
		filedrag.addEventListener("drop", FileSelectHandler, false);
		filedrag.style.display = "block";
		
		var cb = function(){
			$scope.names = ExcelService.getSheetNames();	
			var out = $('#sheets'); out.html("");
			out.append("<p>List of worksheets in the uploaded excel</p>");
			$scope.names.forEach(function(sheet){		
				out.append("<span class='sheet' data-sheet='" + sheet +"'>" + sheet + "</span></br>");
			})
			//default is to display the first sheet
			createsheet(ExcelService.getSheets()[$scope.names[0]]);
			$scope.cacheId = makeid();
			ExcelService.cacheSave($scope.cacheId);
			out.append("you can access the sheet by this: <a href=?id=" + $scope.cacheId + "> url </a>");
		}
			
		function $id(id) {
			return document.getElementById(id);
		}

		//file drag hover
		function FileDragHover(e) {
			e.stopPropagation();
			e.preventDefault();
			e.target.className = (e.type == "dragover" ? "hover" : "");
		}
		
		function makeid(){
		    var text = "";
		    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

		    for( var i=0; i < 5; i++ )
		        text += possible.charAt(Math.floor(Math.random() * possible.length));

		    return text;
		}
		
		this.onClick = function(target){
			console.log(target);
			$scope.sheets = ExcelService.getSheets();
						
		}
		
		
		function datenum(v, date1904) {
			if(date1904) v+=1462;
			var epoch = Date.parse(v);
			return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
		}
		 
		function sheet_from_array_of_arrays(data, opts) {
			var ws = {};
			var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
			for(var R = 0; R != data.length; ++R) {
				for(var C = 0; C != data[R].length; ++C) {
					if(range.s.r > R) range.s.r = R;
					if(range.s.c > C) range.s.c = C;
					if(range.e.r < R) range.e.r = R;
					if(range.e.c < C) range.e.c = C;
					var cell = {v: data[R][C] };
					if(cell.v == null) continue;
					var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
					
					if(typeof cell.v === 'number') cell.t = 'n';
					else if(typeof cell.v === 'boolean') cell.t = 'b';
					else if(cell.v instanceof Date) {
						cell.t = 'n'; cell.z = XLSX.SSF._table[14];
						cell.v = datenum(cell.v);
					}
					else cell.t = 's';
					
					ws[cell_ref] = cell;
				}
			}
			if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
			return ws;
		}
		 
				 
		function Workbook() {
			if(!(this instanceof Workbook)) return new Workbook();
			this.SheetNames = [];
			this.Sheets = {};
		}
		 
			
		this.download = function(){
			//$scope.sheets = ExcelService.getSheets();
			//$scope.wb = ExcelService.getWorkbook();
			/* original data */
			//var data = [[1,2,3],[true, false, null, "sheetjs"],["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]]
			/*var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);	 
			var ws_name = "SheetJS";
			wb.SheetNames.push(ws_name);
			wb.Sheets[ws_name] = ws;
			var wbout = XLSX.write($scope.wb, {bookType:'xlsx', bookSST:true, type: 'binary'});

			function s2ab(s) {
				var buf = new ArrayBuffer(s.length);
				var view = new Uint8Array(buf);
				for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
				return buf;
			}*/
			saveAs(new Blob([ExcelService.getWorkbookAsArrayBuffer()],{type:"application/octet-stream"}), $scope.cacheId + ".xlsx")
			
		}
		
		function createsheet(json) {
		      var $container = $("#hot"); 
		      var $parent = $container.parent();
		      var offset = $container.offset();
		      var $window = $(window);
		      availableWidth = Math.max($window.width() - 250,600);
		      availableHeight = Math.min($window.height() - 250, 400);
		  
		      /* add header row for table */
		      if(!json) json = [];
		      /* showtime! */
		      $("#hot").handsontable({
		        data: json,
		        contextMenu: true,
		        formulas: true,
		        comments: true,
		        fixedRowsTop: 1,
		        stretchH: 'all',
		        rowHeaders: true,
		       /* columns: cols.map(function(x) { return {data:x}; }),*/
		        colHeaders: true,
		        cells: function (r,c,p) {
		          if(r === 0) this.renderer = boldRenderer;
		        },
		        width: function () { return availableWidth; },
		        height: function () { return availableHeight; },
		        stretchH: 'all'
		      });
		      
		      $("#hot2").handsontable({
		            data: json,
		            formulas: false,
		            startRows: 5,
		            startCols: 3,
		            fixedRowsTop: 1,
		            stretchH: 'all',
		            rowHeaders: true,
		            /*columns: cols.map(function(x) { return {data:x}; }),
		            colHeaders: cols.map(function(x,i) { return XLS.utils.encode_col(i); }),*/
		            cells: function (r,c,p) {
		              if(r === 0) this.renderer = boldRenderer;
		            },
		            width: function () { return availableWidth; },
		            height: function () { return availableHeight; },
		            stretchH: 'all'
		          });
		}
		
		/** Handsontable magic **/
		var boldRenderer = function (instance, td, row, col, prop, value, cellProperties) {
		  Handsontable.TextCell.renderer.apply(this, arguments);
		  $(td).css({'font-weight': 'bold'});
		};
		
});

app.controller("DnDController", function(ExcelService){
	//file drag and drop
	var filedrag = $id("filedrag");

	// file drop
	filedrag.addEventListener("dragover", FileDragHover, false);
	filedrag.addEventListener("dragleave", FileDragHover, false);
	filedrag.addEventListener("drop", FileSelectHandler, false);
	filedrag.style.display = "block";
	
	function $id(id) {
		return document.getElementById(id);
	}

	//file selection
	function FileSelectHandler(e) {
		// cancel event and hover styling
		FileDragHover(e);
		ExcelService.handleFile(e.dataTransfer.files);
	}

	//file drag hover
	function FileDragHover(e) {
		e.stopPropagation();
		e.preventDefault();
		e.target.className = (e.type == "dragover" ? "hover" : "");
	}
	
});


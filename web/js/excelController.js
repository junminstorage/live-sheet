var app = angular.module('excelApp', []);

app.service('ExcelService', function(){
	
	//X: sheet parser, either XLSX or XLS, default to XLSX
	var X = XLSX;
	//array of sheet names
	var names = [];
	//hash of worksheets with key on sheet names
	var workSheets = {};
	
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
	        var wb = X.read(data, {type: 'binary'});
	        names = wb.SheetNames;
	        workSheets = to_jsonArray(wb);
	        cb();
	    }
	}
	
	function to_jsonArray(wb){
		var json = {}
		wb.SheetNames.forEach(function(sheet){		
			json[sheet] =  getJsonArrayFor(wb.Sheets[sheet]);
		})
		return json;
	}
	
	function getJsonArrayFor(sheet, opts) {
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
		getSheets : function(){return workSheets;}
		
	}
});


app.controller('ExcelController', function($scope, ExcelService){
		var self = this;
		$scope.sheets = {};
		$scope.names = [];
		
		$scope.uploadFile = function(that){			
			ExcelService.handleFile(that.files, cb);			
		}
		
		//file drag and drop
		var filedrag = $id("filedrag");

		// file drop
		filedrag.addEventListener("dragover", FileDragHover, false);
		filedrag.addEventListener("dragleave", FileDragHover, false);
		filedrag.addEventListener("drop", FileSelectHandler, false);
		filedrag.style.display = "block";
		
		var cb = function(){
			$scope.names = ExcelService.getSheetNames();	
			var out = $('#sheets');
			out.append("<p>List of worksheets in the uploaded excel</p>");
			$scope.names.forEach(function(sheet){		
				out.append("<span class='sheet'>" + sheet + "</span></br>");
			})
			//default is to display the first sheet
			createsheet(ExcelService.getSheets()[$scope.names[0]]);
			
		}
		
		function $id(id) {
			return document.getElementById(id);
		}

		//file selection
		function FileSelectHandler(e) {
			// cancel event and hover styling
			FileDragHover(e);
			ExcelService.handleFile(e.dataTransfer.files, cb);
		}

		//file drag hover
		function FileDragHover(e) {
			e.stopPropagation();
			e.preventDefault();
			e.target.className = (e.type == "dragover" ? "hover" : "");
		}
		
		
		this.onClick = function(target){
			console.log(target);
			$scope.sheets = ExcelService.getSheets();
		}
		
		var createsheet = function(json) {
		      $container = $("#hot"); 
		      $parent = $container.parent();
		      var offset = $container.offset();
		      $window = $(window);
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

//I simply log the creation / linking of a DOM node to
// illustrate the way the DOM nodes are created with the
// various tracking approaches.
app.directive(
    "bnLogDomCreation",
    function() {
        // I bind the UI to the $scope.
        function link( $scope, element, attributes ) {
            console.log(
                attributes.bnLogDomCreation,
                $scope.$index
            );
        }
        // Return the directive configuration.
        return({
            link: link
        });
    }
);
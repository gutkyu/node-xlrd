'use strict'
var util = require('util'),
	assert = require('assert'),
	comm = require('./common');

var DEBUG = 0;
var OBJ_MSO_DEBUG = 0;

var _WINDOW2_options = {
    // Attribute names and initial values to use in case
    // a WINDOW2 record is not written.
    "show_formulas": 0,
    "show_grid_lines": 1,
    "show_sheet_headers": 1,
    "panes_are_frozen": 0,
    "show_zero_values": 1,
    "automatic_grid_line_colour": 1,
    "columns_from_right_to_left": 0,
    "show_outline_symbols": 1,
    "remove_splits_if_pane_freeze_is_removed": 0,
    // Multiple sheets can be selected, but only one can be active
    // (hold down Ctrl and click multiple tabs in the file in OOo)
    "sheet_selected": 0,
    // "sheet_visible" should really be called "sheet_active"
    // and is 1 when this sheet is the sheet displayed when the file
    // is open. More than likely only one sheet should ever be set as
    // visible.
    // This would correspond to the Book's sheet_active attribute, but
    // that doesn't exist as WINDOW1 records aren't currently processed.
    // The real thing is the visibility attribute from the BOUNDSHEET record.
    "sheet_visible": 0,
    "showInPageBreakPreview" : 0,
};

////
//  Contains the data for one worksheet. 
//
//  In the cell access functions, "rowx" is a row index, counting from zero, and "colx" is a
// column index, counting from zero.
// Negative values for row/column indexes and slice positions are supported in the expected fashion. 
//
//  For information about cell types and cell values, refer to the documentation of the {@link #Cell} class. 
//
//  WARNING: You don't call this class yourself. You access Sheet objects via the Book object that
// was returned when you called xlrd.open_workbook("myfile.xls"). 


module.exports = function Sheet(book, position, name, index){
	var self = this;
    ////
    // Name of sheet.
    var name = name;

    ////
    // A reference to the Book object to which this sheet belongs.
    // Example usage: some_sheet.book.dateMode
    var book = book;
    
    ////
    // Number of rows in sheet. A row index is in range(thesheet.nrows).
    var _nrows = 0;

    ////
    // Nominal number of columns in sheet. It is 1 + the maximum column index
    // found, ignoring trailing empty cells. See also open_workbook(raggedRows=?)
    // and Sheet.{@link //Sheet.rowLength}(row_index).
    var _ncols = 0;

    ////
    // The map from a column index to a {@link //Colinfo} object. Often there is an entry
    // in COLINFO records for all column indexes in range(257).
    // Note that xlrd ignores the entry for the non-existent
    // 257th column. On the other hand, there may be no entry for unused columns.
    // Populated only if open_workbook(formattingInfo=True).
    var colInfoMap = {};

    ////
    // The map from a row index to a {@link //Rowinfo} object. Note that it is possible
    // to have missing entries -- at least one source of XLS files doesn't
    // bother writing ROW records.
    // Populated only if open_workbook(formattingInfo=True).
    var rowInfoMap = {};

    ////
    // List of address ranges of cells containing column labels.
    // These are set up in Excel by Insert > Name > Labels > Columns.
    //  How to deconstruct the list:
    //  
    // for crange in thesheet.colLabelRanges:
    //     rlo, rhi, clo, chi = crange
    //     for rx in xrange(rlo, rhi):
    //         for cx in xrange(clo, chi):
    //             print "Column label at (rowx=%d, colx=%d) is %r" \
    //                 (rx, cx, thesheet.cellValue(rx, cx))
    //  
    var colLabelRanges = [];

    ////
    // List of address ranges of cells containing row labels.
    // For more details, see  colLabelRanges  above.
    var rowLabelRanges = [];

    ////
    // List of address ranges of cells which have been merged.
    // These are set up in Excel by Format > Cells > Alignment, then ticking
    // the "Merge cells" box.
    // Extracted only if open_workbook(formattingInfo=True).
    //  How to deconstruct the list:
    //  
    // for crange in thesheet.mergedCells:
    //     rlo, rhi, clo, chi = crange
    //     for rowx in xrange(rlo, rhi):
    //         for colx in xrange(clo, chi):
    //             // cell (rlo, clo) (the top left one) will carry the data
    //             // and formatting info; the remainder will be recorded as
    //             // blank cells, but a renderer will apply the formatting info
    //             // for the top left cell (e.g. border, pattern) to all cells in
    //             // the range.
    //  
    var mergedCells = [];
    
    ////
    // Mapping of (rowx, colx) to list of (offset, font_index) tuples. The offset
    // defines where in the string the font begins to be used.
    // Offsets are expected to be in ascending order.
    // If the first offset is not zero, the meaning is that the cell's XF's font should
    // be used from offset 0.
    //    This is a sparse mapping. There is no entry for cells that are not formatted with  
    // rich text.
    //  How to use:
    //  
    // runList = thesheet.richTextRunlistMap.get((rowx, colx))
    // if runList:
    //     for offset, font_index in runList:
    //         // do work here.
    //         pass
    //  
    // Populated only if open_workbook(formattingInfo=True).
   
    var richTextRunlistMap = {};

    ////
    // Default column width from DEFCOLWIDTH record, else null.
    // From the OOo docs: 
    // """Column width in characters, using the width of the zero character
    // from default font (first FONT record in the file). Excel adds some
    // extra space to the default width, depending on the default font and
    // default font size. The algorithm how to exactly calculate the resulting
    // column width is not known. 
    // Example: The default width of 8 set in this record results in a column
    // width of 8.43 using Arial font with a size of 10 points.""" 
    // For the default hierarchy, refer to the {@link //Colinfo} class.
    var defaultColWidth = null;

    ////
    // Default column width from STANDARDWIDTH record, else null.
    // From the OOo docs: 
    // """Default width of the columns in 1/256 of the width of the zero
    // character, using default font (first FONT record in the file).""" 
    // For the default hierarchy, refer to the {@link //Colinfo} class.
    var standardWidth = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var defaultRowHeight = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var defaultRowHeightMismatch = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var defaultRowHidden = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var defaultAdditionalSpaceAbove = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var defaultAdditionalSpaceBelow = null;

    ////
    // Visibility of the sheet. 0 = visible, 1 = hidden (can be unhidden
    // by user -- Format/Sheet/Unhide), 2 = "very hidden" (can be unhidden
    // only by VBA macro).
    var visibility = 0;

    ////
    // A 256-element tuple corresponding to the contents of the GCW record for this sheet.
    // If no such record, treat as all bits zero.
    // Applies to BIFF4-7 only. See docs of the {@link //Colinfo} class for discussion.
    var gcw = new Array(256);

    ////
    //  A list of {@link //Hyperlink} objects corresponding to HLINK records found
    // in the worksheet. 
    var hyperlinkList = [];

    ////
    //  A sparse mapping from (rowx, colx) to an item in {@link //Sheet.hyperlinkList}.
    // Cells not covered by a hyperlink are not mapped.
    // It is possible using the Excel UI to set up a hyperlink that 
    // covers a larger-than-1x1 rectangle of cells.
    // Hyperlink rectangles may overlap (Excel doesn't check).
    // When a multiply-covered cell is clicked on, the hyperlink that is activated
    // (and the one that is mapped here) is the last in hyperlinkList.
    var hyperlinkMap = {};

    ////
    //  A sparse mapping from (rowx, colx) to a {@link //Note} object.
    // Cells not containing a note ("comment") are not mapped.
    var cellNoteMap = {};    
    
    ////
    // Number of columns in left pane (frozen panes; for split panes, see comments below in code)
    var vertSplitPos = 0;

    ////
    // Number of rows in top pane (frozen panes; for split panes, see comments below in code)
    var horzSplitPos = 0;

    ////
    // Index of first visible row in bottom frozen/split pane
    var horzSplitFirstVisible = 0;

    ////
    // Index of first visible column in right frozen/split pane
    var vertSplitFirstVisible = 0;

    ////
    // Frozen panes: ignore it. Split panes: explanation and diagrams in OOo docs.
    var splitActivePane = 0;

    ////
    // Boolean specifying if a PANE record was present, ignore unless you're xlutils.copy
    var hasPaneRecord = 0;

    ////
    // A list of the horizontal page breaks in this sheet.
    // Breaks are tuples in the form (index of row after break, start col index, end col index).
    // Populated only if open_workbook(formattingInfo=True).
    var horizontalPageBreaks = [];

    ////
    // A list of the vertical page breaks in this sheet.
    // Breaks are tuples in the form (index of col after break, start row index, end row index).
    // Populated only if open_workbook(formattingInfo=True).
    var verticalPageBreaks = [];
    
	//self.book = book;
	self.biffVersion = book.biffVersion;
	self._position = position;
	self.logfile = book.logfile;
	self.bt = comm.XL_CELL_EMPTY; //unsigned char 1byte
	self.bf = -1; //signed short 2bytes
	self.name = name;
	self.index = index;
	self.verbosity = book.verbosity;
	self.formattingInfo = book.formattingInfo;
	self.raggedRows = book.raggedRows;
	if (self.raggedRows)
		self.putCell = putCellRagged;
	else
		self.putCell = putCellUnragged;
	self._xf_index_to_xl_type_map = book._xf_index_to_xl_type_map;
	//_nrows = 0; // actual, including possibly empty cells
	//_ncols = 0;
	self._maxDataRowIndex = -1; // highest rowx containing a non-empty cell
	self._maxDataColIndex = -1; // highest colx containing a non-empty cell
	self._dimnRows = 0; // as per DIMENSIONS record
	self._dimnCols = 0;
	var _cellValues = [];
	var _cellTypes = [];
	var _cellXFIndexes = [];
	self.defaultColWidth = null;
	self.standardWidth = null;
	self.defaultRowHeight = null;
	self.defaultRowHeightMismatch = 0;
	self.defaultRowHidden = 0;
	self.defaultAdditionalSpaceAbove = 0;
	self.defaultAdditionalSpaceBelow = 0;
	self.colInfoMap = {};
	self.rowInfoMap = {};
	self.colLabelRanges = [];
	self.rowLabelRanges = [];
	self.mergedCells = [];
	self.richTextRunlistMap = {};
	self.horizontalPageBreaks = [];
	self.verticalPageBreaks = [];
	self._xf_index_stats = [0, 0, 0, 0];
	self.visibility = book._sheetVisibility[index]; // from BOUNDSHEET record
	for(var key in _WINDOW2_options) self[key] =_WINDOW2_options[key];
	self.firstVisibleRowIndex = 0;
	self.firstVisibleColIndex = 0;
	self.gridlineColourIndex = 0x40;
	self.gridlineColourRgb = null; // pre-BIFF8
	self.hyperlinkList = [];
	self.hyperlinkMap = {};
	self.cellNoteMap = {};

	// Values calculated by xlrd to predict the mag factors that
	// will actually be used by Excel to display your worksheet.
	// Pass these values to xlwt when writing XLS files.
	// Warning 1: Behaviour of OOo Calc and Gnumeric has been observed to differ from Excel's.
	// Warning 2: A value of zero means almost exactly what it says. Your sheet will be
	// displayed as a very tiny speck on the screen. xlwt will reject attempts to set
	// a mag_factor that is not (10 <= mag_factor <= 400).
	self.cookedPageBreakPreviewMagFactor = 60;
	self.cookedNormalViewMagFactor = 100;

	// Values (if any) actually stored on the XLS file
	self.cachedPageBreakPreviewMagFactor = null; // from WINDOW2 record
	self.cachedNormalViewMagFactor = null; // from WINDOW2 record
	self.sclMagFactor = null; // from SCL record

	self._ixfe = null; // BIFF2 only
	self._cellAttrToXfx = {}; // BIFF2.0 only

	//////// Don't initialise this here, use class attribute initialisation.
	//////// self.gcw = (0, ) * 256 ////////

	if (self.biffVersion >= 80)
		self.utterMaxRows = 65536;
	else
		self.utterMaxRows = 16384;
	self.utterMaxCols = 256;

	self._firstFullRowIndex = -1;

	// self._put_cell_exceptions = 0
	// self._put_cell_row_widenings = 0
	// self._put_cell_rows_appended = 0
	// self._put_cell_cells_appended = 0


    // === Following methods are used in building the worksheet.
    // === They are not part of the API.

    function tidyDimensions(){
        if (self.verbosity >= 3)
            console.log(util.format("tidyDimensions: nrows=%d ncols=%d \n",_nrows, _ncols));
        if(1 && self.mergedCells && self.mergedCells.length){
            var nr = nc = 0;
            var uMaxRows = self.utterMaxRows;
            var uMaxCols = self.utterMaxCols;
			self.mergedCells.forEach(function(x){
				var rlo = x[0], rhi =x[1], clo = x[2], chi = x[3];
                if (!(0 <= rlo &&  rlo  < rhi && rhi <= uMaxRows) || !(0 <= clo && clo < chi && chi <= uMaxCols))
                    console.log(util.format(
                        "*** WARNING: sheet #%d (%r), MERGEDCELLS bad range %r\n",
                        self.index, self.name, crange));
                if (rhi > nr) nr = rhi;
                if (chi > nc) nc = chi;
			});
            if( nc > _ncols)
                _ncols = nc;
            if( nr > _nrows)
                // we put one empty cell at (nr-1,0) to make sure
                // we have the right number of rows. The ragged rows
                // will sort out the rest if needed.
                self.putCell(nr-1, 0, comm.XL_CELL_EMPTY, '', -1);
		}
        if(self.verbosity >= 1 && (_nrows != self._dimnRows || _ncols != self._dimnCols))
            console.log(util.format("NOTE *** sheet %d (%r): DIMENSIONS R,C = %d,%d should be %d,%d\n",
                self.index,
                self.name,
                self._dimnRows,
                self._dimnCols,
                _nrows,
                _ncols
                ));
        if (!self.raggedRows){
            // fix ragged rows
            var ncols = _ncols,
				cTyps = _cellTypes,
				cVals = _cellValues,
				cXfIdxs = _cellXFIndexes,
				fmtInf = self.formattingInfo;
            // for rowx in xrange(_nrows):
			var ubound = self._firstFullRowIndex;
            if(self._firstFullRowIndex == -2)
                ubound = _nrows;
            for(var rowx =0;rowx < ubound ; rowx++){
                var trow = cTyps[rowx];
                var rlen = trow.length;
                var nextra = ncols - rlen;
                if(nextra > 0){
                    cVals[rowx][ncols-1]= '';
					for(var c =rlen;c < ncols-1; c++) cVals[rowx][c]='';

					trow[ncols-1] = self.bt;
					for(var c =rlen;c < ncols-1; c++) trow[c]=self.bt;
					
                    if (fmtInf) {
						cXfIdxs[rowx][ncols-1] = self.bf;
						for(var c =rlen;c < ncols-1; c++) cXfIdxs[rowx][c]=self.bf;
					}
				}
			}
		}
    }
	
	function putCellRagged(rowx, colx, ctype, value, xfIndex){
        if (!ctype){
            // we have a number, so look up the cell type
            ctype = self._xf_index_to_xl_type_map[xfIndex];
		}
        assert(0 <= colx && colx < self.utterMaxCols,'0 <= colx < self.utterMaxCols');
        assert(0 <= rowx && rowx < self.utterMaxRows,'0 <= rowx < self.utterMaxRows');
        var fmtInf = self.formattingInfo;

		var nr = rowx + 1;
		if (elf.nrows < nr){
			for(var i = elf.nrows ; i < nr ; i++ ){
				_cellTypes[i] = [];
				_cellValues[i] =[];
				if(fmtInf) self._cellXFIndexes[i] =[];
			}
			_nrows = nr;
		}
		
		var typesRow = _cellTypes[rowx];
		var valuesRow = _cellValues[rowx];
		var fmtRow = null;
		if (fmtInf)
			fmtRow = _cellXFIndexes[rowx];
		var ltr = typesRow.length;
		if (colx >= _ncols)
			_ncols = colx + 1;
		var numEmpty = colx - ltr;
		
		typesRow[colx] = ctype;
		valuesRow[colx] = value;
		if (fmtInf)
			fmtRow[colx] = xfIndex;
		if(numEmpty > 0){
			for(i = ltr; i < colx;i++) 
				if(typesRow[i] === undefined) typesRow[i] = self.bt;
			for(i = ltr; i < colx;i++) 
				if(valuesRow[i] === undefined) valuesRow[i] = '';
			if (fmtInf) 
				for(i = ltr; i < colx;i++) 
					if(fmtRow[i] === undefined) fmtRow[i] = self.bf;
		}
        
	}
	function extendCells(rowx, colx){
		// print >> self.logfile, "putCell extending", rowx, colx
		// self.extend_cells(rowx+1, colx+1)
		// self._put_cell_exceptions += 1
		var nr = rowx + 1;
		var nc = colx + 1;
		assert( 1 <= nc && nc <= self.utterMaxCols,'1 <= nc <= self.utterMaxCols');
		assert( 1 <= nr && nr <= self.utterMaxRows,'1 <= nr <= self.utterMaxRows');
		if (nc > _ncols){
			_ncols = nc;
			// The row self._firstFullRowIndex and all subsequent rows
			// are guaranteed to have length == _ncols. Thus the
			// "fix ragged rows" section of the tidyDimensions method
			// doesn't need to examine them.
			if (nr < _nrows){
				// cell data is not in non-descending row order *AND*
				// _ncols has been bumped up.
				// This very rare case ruins this optmisation.
				self._firstFullRowIndex = -2;
			}else if (rowx > self._firstFullRowIndex && self._firstFullRowIndex > -2)
				self._firstFullRowIndex = rowx;
		}
		if( nr <= _nrows){
			// New cell is in an existing row, so extend that row (if necessary).
			// Note that nr < _nrows means that the cell data
			// is not in ascending row order!!
			var trow = _cellTypes[rowx];
			var nextra = _ncols - trow.length;
			if (nextra > 0){
				// self._put_cell_row_widenings += 1
				trow[_ncols]=self.bt;
				trow.pop();
				for(c =trow.length;c < _ncols; c++) trow[c]=self.bt;
				if (self.formattingInfo){
					_cellXFIndexes[rowx][_ncols]=self.bf;
					_cellXFIndexes[rowx].pop();
					for(c =trow.length;c < _ncols; c++) _cellXFIndexes[rowx][c]=self.bf;
				}
				_cellValues[rowx][_ncols]='';
				_cellValues[rowx].pop();
				for(c =trow.length;c < _ncols; c++) _cellValues[rowx][c]='';
			}
		}else{
			var fmtInf = self.formattingInfo;
			var nc = _ncols,
				bt = self.bt,
				bf = self.bf;
			for(var r = _nrows; r < nr; r++){
				// self._put_cell_rows_appended += 1
				var ta = new Array(nc);
				for(var c =0;c < nc; c++) ta[c]=bt;
				_cellTypes.push(ta);
				var va = new Array(nc);
				for(var c =0;c < nc; c++) va[c]='';
				_cellValues.push(va);
				if (fmtInf){
					var fa = new Array(nc);
					for(var c =0;c < nc; c++) fa[c]=bf;
					_cellXFIndexes.push(fa);
				}
			}
			_nrows = nr;
		}
	}
    function putCellUnragged(rowx, colx, ctype, value, xfIndex){
		if(!ctype)
            // we have a number, so look up the cell type
            ctype = self._xf_index_to_xl_type_map[xfIndex];
        // assert 0 <= colx < self.utterMaxCols
        // assert 0 <= rowx < self.utterMaxRows
		if(_cellTypes[rowx] === undefined || _cellTypes[rowx][colx] === undefined ){
			extendCells(rowx, colx);
		}
		_cellTypes[rowx][colx] = ctype;
		_cellValues[rowx][colx] = value;
		if(self.formattingInfo)
			_cellXFIndexes[rowx][colx] = xfIndex;
        
	}
	
    //#
    // Value of the cell in the given row and column.
    self.cell = function( rowx, colx){
        var val = _cellValues[rowx][colx];
		if(val === undefined || val === null) 
			return val;
        var typ = _cellTypes[rowx][colx];
		return typ === comm.XL_CELL_DATE ? toDate(val, book.dateMode):val;
	};
    self.cell.getValue = self.cell;
	/*
	//#
    // {@link #Cell} object in the given row and column.
    self.cellObj = function (rowx, colx){
		var xfx = null;
        if (self.formattingInfo) xfx = self.cellXFIndex(rowx, colx);
        return {"type":_cellTypes[rowx][colx],"value":_cellValues[rowx][colx],"xfIndex":xfx};
	}
    */
	
	self.cell.getRaw = function( rowx, colx){
        return _cellValues[rowx][colx];
	};
    
	/*
    //formated value
    self.cell.getText = function(row, col){
		
    }
	*/
    //#
    // Type of the cell in the given row and column.
    // Refer to the documentation of the {@link #Cell} class.
    self.cell.getType=function ( rowx, colx){
        return _cellTypes[rowx][colx];
	};
		
	//#
    // XF index of the cell in the given row and column.
    // This is an index into Book.{@link #Book.xfList}.
    self.cell.getXFIndex = function(rowx, colx){
        self.reqFmtInfo();
        var xfx = _cellXFIndexes[rowx][colx];
        if (xfx > -1){
            self._xf_index_stats[0] += 1;
            return xfx;
		}
        // Check for a row xfIndex
        if((xfx=self.rowInfoMap[rowx].xfIndex)!== undefined){
            if (xfx > -1){
                self._xf_index_stats[1] += 1;
                return xfx;
			}
		}
        
        // Check for a column xfIndex
        if((xfx = self.colInfoMap[colx].xfIndex)!== undefined){
            if (xfx == -1) xfx = 15;
            self._xf_index_stats[2] += 1;
            return xfx;
		}else{
			// If all else fails, 15 is used as hardwired global default xfIndex.
            self._xf_index_stats[3] += 1;
            return 15;
		}
	};
	
	self.row = {};
    // Returns a sequence of the {@link #Cell} objects in the given row.
    self.row.getValues = function(rowx){
		var ret =[], len = _cellValues[rowx].length;
		for(var colx=0; colx < len; colx++){
			ret.push(self.cell(rowx, colx));
		}
        return ret;
	};

	
    // Returns the effective number of cells in the given row. For use with
    // open_workbook(raggedRows=True) which is likely to produce rows
    // with fewer than {@link #Sheet.ncols} cells.
    self.row.getCount = function(rowx){
        return _cellValues[rowx].length;
	};
	Object.defineProperty(self.row, 'count', {get:function(){return _nrows;}});

    // Returns a slice of the types
    // of the cells in the given row.
    self.row.getTypes=function( rowx, startColx, endColx){
		startColx=startColx||0;
        if(endColx === undefined)
            return _cellTypes[rowx].slice(startColx);
        return _cellTypes[rowx].slice(startColx,endColx);
	};

    // Returns a slice of the values
    // of the cells in the given row.
    self.row.getRaws=function(rowx, startColx, endColx){
		startColx=startColx||0;
		//endColx=endColx||null;
        if(endColx === undefined)
            return _cellValues[rowx].slice(startColx);
        return _cellValues[rowx].slice(startColx,endColx);
	};
	
	self.column = {};
	
	Object.defineProperty(self.column, 'count', {get:function(){return _ncols;}});

	// === Methods after this line neither know nor care about how cells are stored.
	
    self.read = function(bk){
        //global rc_stats
        var DEBUG = 0;
        var blah = DEBUG || self.verbosity >= 2;
        var blahRows = DEBUG || self.verbosity >= 4;
        var blahFormulas = 0 && blah;
        var r1c1 = 0;
        var oldPos = bk._position;
        bk._position = self._position;
        var XL_SHRFMLA_ETC_ETC = [comm.XL_SHRFMLA, comm.XL_ARRAY, comm.XL_TABLEOP, comm.XL_TABLEOP2,
            comm.XL_ARRAY2, comm.XL_TABLEOP_B2];            
        var putCell = self.putCell;
        var getBkRecParts = bk.getRecordParts;
        var bv = self.biffVersion;
        var fmtInf = self.formattingInfo;
        var doSstRichText = fmtInf && bk._richTextRunlistMap;
        var rowInfoSharingDict = {};
        var txos = {};
        var eofFound = 0;
        while(true){
            // if DEBUG: print "SHEET.READ: about to read from position %d" % bk._position
            var record = getBkRecParts();
			var rc = record[0], dataLen = record[1], data = record[2];
            // if rc in rc_stats:
            //     rc_stats[rc] += 1
            // else:
            //     rc_stats[rc] = 1
            // if DEBUG: print "SHEET.READ: op 0x%04x, %d bytes %r" % (rc, dataLen, data)
            if (rc == comm.XL_NUMBER){
                // [:14] in following stmt ignores extraneous rubbish at end of record.
                // Sample file testEON-8.xls supplied by Jan Kraus.
				
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4), 
					d = data.readDoubleLE(6);
                // if xfIndex == 0:
                //     fprintf(self.logfile,
                //         "NUMBER: r=%d c=%d xfx=%d %f\n", rowx, colx, xfIndex, d)
                putCell(rowx, colx, null, d, xfIndex);
            }else if( rc == comm.XL_LABELSST){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4), 
					sstIndex = data.readInt32LE(6);;
                // print "LABELSST", rowx, colx, sstIndex, bk._sharedStrings[sstIndex]
                putCell(rowx, colx, comm.XL_CELL_TEXT, bk._sharedStrings[sstIndex], xfIndex);
                if (doSstRichText){
                    var runList = bk._richTextRunlistMap.get(sstIndex);
                    if (runList){
                        var row = self.richTextRunlistMap[rowx]
						if(!row) row = self.richTextRunlistMap[rowx] = [];
						row[colx] = runList;
					}
				}
            }else if( rc == comm.XL_LABEL){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4),
                    strg = '';
                    
                if ( bv < comm.BIFF_FIRST_UNICODE)
                    strg = comm.unpackString(data, 6, bk.encoding || bk.deriveEncoding(), 2);
                else
                    strg = unpackUnicode(data, 6, 2);
                putCell(rowx, colx, comm.XL_CELL_TEXT, strg, xfIndex);
            }else if( rc == comm.XL_RSTRING){
                var rowx = data.readUInt16LE(0),
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4),
                    runList =[],
                    strg = '';
                if ( bv < comm.BIFF_FIRST_UNICODE){
                    var unpack = comm.unpackStringUpdatePos(data, 6, bk.encoding || bk.deriveEncoding(), 2), 
                        pos = unpack[1];
					strg = unpack[0];
                    var nrt = data[pos];
                    pos += 1;
                    while (pos< nrt) runList.append( data.readUInt8(pos++));
                    assert(pos == data.length,'pos == len(data)');
                }else{
					var unpack= comm.unpack_unicode_update_pos(data, 6, 2), 
						pos = unpack[1];
                    strg = unpack[0];
                    var nrt = data.readUInt16LE(0); 
                    pos += 2;
					while (pos< nrt) {
						runList.append( data.readUInt16LE(pos));
						pos+=2;
					}
                    assert(pos == data.length,'pos == len(data)');
				}
                putCell(rowx, colx, comm.XL_CELL_TEXT, strg, xfIndex);
				var row = self.richTextRunlistMap[rowx]
				if(!row) row = self.richTextRunlistMap[rowx] = [];
				row[colx] = runList;
            }else if( rc == comm.XL_RK){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4);
                var d = unpackRK(data, 6, 10);
                putCell(rowx, colx, null, d, xfIndex);
            }else if( rc == comm.XL_MULRK){
                var mulrkRow = data.readUInt16LE(0), 
					mulrkFirst = data.readUInt16LE(2),
					mulrkLast = data.readUInt16LE(data.length-2);
                var pos = 4;
				for(colx = mulrkFirst; colx < mulrkLast+1; colx++){
					xfIndex = data.readUInt16LE(pos)
                    var d = unpackRK(data,pos+2,pos+6);
                    pos += 6;
                    putCell(mulrkRow, colx, null, d, xfIndex);
				}
            }else if(rc == comm.XL_ROW){
                // Version 0.6.0a3: ROW records are just not worth using (for memory allocation).
                // Version 0.6.1: now used for formatting info.
                if ( !fmtInf) continue;
                var rowx = data.readUInt16LE(0), 
					bits1 = data.readUInt16LE(6), 
					bits2 = data.readInt32LE(12);
                if (!(0 <= rowx < self.utterMaxRows)){
                    console.log(util.format("*** NOTE: ROW record has row index %d; should have 0 <= rowx < %d -- record ignored!" 
                        ,rowx, self.utterMaxRows));
                    continue;
				}
                
                var r = rowInfoSharingDict.get(bits1, bits2);
                if(!r){
                    rowInfoSharingDict[bits1][bits2] = r = Rowinfo();
                    // Using upkbits() is far too slow on a file
                    // with 30 sheets each with 10K rows :-(
                    //    upkbits(r, bits1, (
                    //        ( 0, 0x7FFF, 'height'),
                    //        (15, 0x8000, 'has_default_height'),
                    //        ))
                    //    upkbits(r, bits2, (
                    //        ( 0, 0x00000007, 'outline_level'),
                    //        ( 4, 0x00000010, 'outline_group_starts_ends'),
                    //        ( 5, 0x00000020, 'hidden'),
                    //        ( 6, 0x00000040, 'height_mismatch'),
                    //        ( 7, 0x00000080, 'has_default_xf_index'),
                    //        (16, 0x0FFF0000, 'xfIndex'),
                    //        (28, 0x10000000, 'additional_space_above'),
                    //        (29, 0x20000000, 'additional_space_below'),
                    //        ))
                    // So:
                    r.height = bits1 & 0x7fff;
                    r.has_default_height = (bits1 >> 15) & 1;
                    r.outline_level = bits2 & 7;
                    r.outline_group_starts_ends = (bits2 >> 4) & 1;
                    r.hidden = (bits2 >> 5) & 1;
                    r.height_mismatch = (bits2 >> 6) & 1;
                    r.has_default_xf_index = (bits2 >> 7) & 1;
                    r.xfIndex = (bits2 >> 16) & 0xfff;
                    r.additional_space_above = (bits2 >> 28) & 1;
                    r.additional_space_below = (bits2 >> 29) & 1;
                    if(!r.has_default_xf_index)
                        r.xfIndex = -1;
				}
                self.rowInfoMap[rowx] = r;
                if ( 0 && r.xfIndex > -1){
                    console.log(util.format("**ROW %d %d %d\n",self.index, rowx, r.xfIndex));
				}
                if ( blahRows){
                    console.log('ROW', rowx, bits1, bits2);
                    //r.dump(header="--- sh //%d, rowx=%d ---" % (self.index, rowx));
				}
            }else if( comm.XL_FORMULA_OPCODES[0] == rc || comm.XL_FORMULA_OPCODES[1] == rc || comm.XL_FORMULA_OPCODES[2] == rc ){ // 06, 0206, 0406
			//todo 나중에 구현
			/*
                // DEBUG = 1
                // if DEBUG: print "FORMULA: rc: 0x%04x data: %r" % (rc, data)
                if ( bv >= 50){
                    var rowx = data.readUInt16LE(0), 
						colx = data.readUInt16LE(2), 
						xfIndex = data.readUInt16LE(4), 
						result_str = data.slice(6,14), 
						flags = data.readUInt16LE(14);
                    var lenRecordLength = 2; //todo 사용여부 확인
                    var tkarr_offset = 20; //todo 사용여부 확인
                }else if( bv >= 30){
                    var rowx = data.readUInt16LE(0), 
						colx = data.readUInt16LE(2), 
						xfIndex = data.readUInt16LE(4), 
						result_str = data.slice(6,14), 
						flags = data.readUInt16LE(14);
                    var lenRecordLength = 2;
                    var tkarr_offset = 16;
                }else{ // BIFF2
                    rowx, colx, cell_attr,  result_str, flags = local_unpack('<HH3s8sB', data[0:16]);
					var rowx = data.readUInt16LE(0), 
						colx = data.readUInt16LE(2), 
						cell_attr = data.slice(4,7), 
						result_str = data.slice(7,15), 
						flags = data.readUInt8(15);
                    var xfIndex =  self.fixed_BIFF2_xfindex(cell_attr, rowx, colx);
                    var lenRecordLength = 1;
                    var tkarr_offset = 16;
				}
                if ( blahFormulas){ // testing formula dumper
                    //////// XXXX FIXME
                    console.log(util.format("FORMULA: rowx=%d colx=%d"), rowx, colx);
                    var fmlalen = data.readUInt16LE(20);
                    decompile_formula(bk, data[22:], fmlalen, FMLA_TYPE_CELL,
                        browx=rowx, bcolx=colx, blah=1, r1c1=r1c1);
				}
                if ( result_str[6:8] == b"\xFF\xFF"){
                    first_byte = BYTES_ORD(result_str[0]);
                    if ( first_byte == 0)
                        // need to read next record (STRING)
                        gotstring = 0;
                        // if flags & 8:
                        if ( 1){ // "flags & 8" applies only to SHRFMLA
                            // actually there's an optional SHRFMLA or ARRAY etc record to skip over
                            rc2, data2_len, data2 = bk.getRecordParts();
                            if ( rc2 == comm.XL_STRING || rc2 == comm.XL_STRING_B2){
                                gotstring = 1;
                            }else if( rc2 == comm.XL_ARRAY){
                                row1x, rownx, col1x, colnx, arrayFlags, tokslen = \
                                    local_unpack("<HHBBBxxxxxH", data2[:14]);
                                if ( blahFormulas)
                                    fprintf(self.logfile, "ARRAY: %d %d %d %d %d\n",
                                        row1x, rownx, col1x, colnx, arrayFlags);
                                    // dump_formula(bk, data2[14:], tokslen, bv, reldelta=0, blah=1)
                            }else if( rc2 == comm.XL_SHRFMLA){
                                row1x, rownx, col1x, colnx, nfmlas, tokslen = \
                                    local_unpack("<HHBBxBH", data2[:10]);
                                if ( blahFormulas){
                                    fprintf(self.logfile, "SHRFMLA (sub): %d %d %d %d %d\n",
                                        row1x, rownx, col1x, colnx, nfmlas);
                                    decompile_formula(bk, data2[10:], tokslen, FMLA_TYPE_SHARED,
                                        blah=1, browx=rowx, bcolx=colx, r1c1=r1c1);
								}
                            }else if( rc2 not in comm.XL_SHRFMLA_ETC_ETC)
                                raise XLRDError(
                                    "Expected SHRFMLA, ARRAY, TABLEOP* or STRING record; found 0x%04x" % rc2);
                            // if DEBUG: print "gotstring:", gotstring
						}
                        // now for the STRING record
                        if ( not gotstring){
                            rc2, _unused_len, data2 = bk.getRecordParts();
                            if ( rc2 not in (comm.XL_STRING, comm.XL_STRING_B2))
                                raise XLRDError("Expected STRING record; found 0x%04x" % rc2);
						}
                        // if DEBUG: print "STRING: data=%r BIFF=%d cp=%d" % (data2, self.biffVersion, bk.encoding)
                        strg = self.string_record_contents(data2);
                        self.putCell(rowx, colx, comm.XL_CELL_TEXT, strg, xfIndex);
                        // if DEBUG: print "FORMULA strg %r" % strg
                    }else if( first_byte == 1){
                        // boolean formula result
                        value = BYTES_ORD(result_str[2]);
                        putCell(rowx, colx, comm.XL_CELL_BOOLEAN, value, xfIndex);
                    }else if( first_byte == 2){
                        // Error in cell
                        value = BYTES_ORD(result_str[2]);
                        putCell(rowx, colx, comm.XL_CELL_ERROR, value, xfIndex);
                    }else if( first_byte == 3){
                        // empty ... i.e. empty (zero-length) string, NOT an empty cell.
                        putCell(rowx, colx, comm.XL_CELL_TEXT, "", xfIndex);
                    }else
                        raise XLRDError("unexpected special case (0x%02x) in FORMULA" % first_byte);
                else{
                    // it is a number
                    d = local_unpack('<d', result_str)[0];
                    putCell(rowx, colx, null, d, xfIndex);
				}
			*/
            }else if( rc == comm.XL_BOOLERR){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4), 
					value = data.readUInt8(6), 
					isErr = data.readUInt8(7);
                // Note OOo Calc 2.0 writes 9-byte BOOLERR records.
                // OOo docs say 8. Excel writes 8.
                var cellTyp = [comm.XL_CELL_BOOLEAN, comm.XL_CELL_ERROR][isErr];
                // if DEBUG: print "XL_BOOLERR", rowx, colx, xfIndex, value, isErr
                putCell(rowx, colx, cellTyp, value, xfIndex);
            }else if( rc == comm.XL_COLINFO){
			//todo 나중에
			/*
                if ( !fmtInf) continue;
                var c = Colinfo();
                var first_colx, last_colx, c.width, c.xfIndex, flags \
                    = local_unpack("<HHHHH", data[:10]);
                //////// Colinfo.width is denominated in 256ths of a character,
                //////// *not* in characters.
                if ( not(0 <= first_colx <= last_colx <= 256))
                    // Note: 256 instead of 255 is a common mistake.
                    // We silently ignore the non-existing 257th column in that case.
                    print("*** NOTE: COLINFO record has first col index %d, last %d; " \
                        "should have 0 <= first <= last <= 255 -- record ignored!" \
                        % (first_colx, last_colx), file=self.logfile);
                    del c;
                    continue;
                upkbits(c, flags, (
                    ( 0, 0x0001, 'hidden'),
                    ( 1, 0x0002, 'bit1_flag'),
                    // *ALL* colinfos created by Excel in "default" cases are 0x0002!!
                    // Maybe it's "locked" by analogy with XFProtection data.
                    ( 8, 0x0700, 'outline_level'),
                    (12, 0x1000, 'collapsed'),
                    ));
                for colx in xrange(first_colx, last_colx+1):
                    if ( colx > 255) break; // Excel does 0 to 256 inclusive
                    self.colInfoMap[colx] = c;
                    if ( 0)
                        fprintf(self.logfile,
                            "**COL %d %d %d\n",
                            self.index, colx, c.xfIndex);
                if ( blah)
                    fprintf(
                        self.logfile,
                        "COLINFO sheet //%d cols %d-%d: wid=%d xfIndex=%d flags=0x%04x\n",
                        self.index, first_colx, last_colx, c.width, c.xfIndex, flags,
                        );
                    c.dump(self.logfile, header='===');
			*/
            }else if( rc == comm.XL_DEFCOLWIDTH){
                self.defaultColWidth = data.readUInt16LE(0);
                if (0) console.log('DEFCOLWIDTH', self.defaultColWidth);
            }else if( rc == comm.XL_STANDARDWIDTH){
                if ( dataLen != 2) console.log('*** ERROR *** STANDARDWIDTH', dataLen, repr(data));
                self.standardWidth = data.readUInt16LE(0);
                if ( 0) console.log('STANDARDWIDTH', self.standardWidth);
            }else if( rc == comm.XL_GCW){
			//todo 나중에
			/*
                if ( not fmtInf) continue; // useless w/o COLINFO
                assert dataLen == 34;
                assert data[0:2] == b"\x20\x00";
                iguff = unpack("<8i", data[2:34]);
                gcw = [];
                for bits in iguff{
                    for j in xrange(32){
                        gcw.append(bits & 1);
                        bits >>= 1
					}
				}
                self.gcw = tuple(gcw);
                if ( 0){
                    showgcw = "".join(map(lambda x: "F "[x], gcw)).rstrip().replace(' ', '.');
                    print("GCW:", showgcw, file=self.logfile);
				}
			*/
            }else if( rc == comm.XL_BLANK){
                if( !fmtInf) continue;
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xfIndex = data.readUInt16LE(4);
                // if 0: print >> self.logfile, "BLANK", rowx, colx, xfIndex
                putCell(rowx, colx, comm.XL_CELL_BLANK, '', xfIndex);
            }else if( rc == comm.XL_MULBLANK){ // 00BE
                if ( ! fmtInf) continue;
                var nitems = dataLen >> 1;
                var result = [];
				for(i=0; i< nitems;i+=2) result.push(data.readUInt16LE(i));
                var rowx = result[0], mul_first = result[1];
                var mulLast = result[result.length-1];
                // print >> self.logfile, "MULBLANK", rowx, mul_first, mulLast, dataLen, nitems, mulLast + 4 - mul_first
                assert( nitems == mulLast + 4 - mul_first,'nitems == mulLast + 4 - mul_first');
                var pos = 2;
                for(colx = mul_first; colx < mulLast+1; colx++){
                    putCell(rowx, colx, comm.XL_CELL_BLANK, '', result[pos]);
                    pos += 1;
				}
            }else if( rc == comm.XL_DIMENSION || rc == comm.XL_DIMENSION2){
                if ( dataLen == 0)
                    // Four zero bytes after some other record. See github issue 64.
                    continue;
                // if dataLen == 10:
                // Was crashing on BIFF 4.0 file w/o the two trailing unused bytes.
                // Reported by Ralph Heimburger.
				var dimTuple =[];
                if ( bv < 80){
					dimTuple =[data.readUInt16LE(2), data.readUInt16LE(6)];
                }else{
					dimTuple =[data.readInt32LE(4), data.readUInt16LE(10)];
				}
                _nrows = _ncols = 0;
                self._dimnRows = self._dimnCols = dimTuple;
                if ( (bv == 21 || bv ==  30 || bv ==  40) && book.xfList && ! book._xfEpilogueDone)
                    book.xfEpilogue();
                if ( blah)
                    console.log(util.format("sheet %d(%r) DIMENSIONS: ncols=%d nrows=%d",
                        self.index, self.name, self._dimnCols, self._dimnRows
                        ));
            }else if( rc == comm.XL_HLINK){
                self.handle_hlink(data);
            }else if( rc == comm.XL_QUICKTIP){
                self.handle_quicktip(data);
            }else if( rc == comm.XL_EOF){
                DEBUG = 0;
                if (DEBUG) print("SHEET.READ: EOF", file=self.logfile);
                eofFound = 1;
                break;
            }else if( rc == comm.XL_OBJ){
                // handle SHEET-level objects; note there's a separate Book.handle_obj
				/*
                var saved_obj = self.handle_obj(data);
                if (saved_obj) saved_obj_id = saved_obj.id;
                else saved_obj_id = null;
				*/
            }else if( rc == comm.XL_MSO_DRAWING){
                //self.handle_msodrawingetc(rc, dataLen, data);
            }else if( rc == comm.XL_TXO){
                var txo = self.handle_txo(data);
                if (txo && saved_obj_id){
                    txos[saved_obj_id] = txo;
                    saved_obj_id = null;
				}
            }else if( rc == comm.XL_NOTE){
                self.handle_note(data, txos);
            }else if( rc == comm.XL_FEAT11){
                self.handle_feat11(data);
            }else if( comm.bofCodes.indexOf(rc) >= 0){ ////////// EMBEDDED BOF //////////
                var version = data.readUInt16LE(0), 
					boftype = data.readUInt16LE(2);
                if ( boftype != 0x20) // embedded chart
                    console.log(util.format("*** Unexpected embedded BOF (0x%04x) at offset %d: version=0x%04x type=0x%04x" 
                        ,rc, bk._position - dataLen - 4, version, boftype));
                while (1){
					var record = bk.getRecordParts();
                    var code = record[0];
					dataLen = record[1];
					data  = record[2];
                    if(code == comm.XL_EOF)
                        break;
				}
                if ( DEBUG) console.log("---> found EOF");
            }else if( rc == comm.XL_COUNTRY){
                bk.handleCountry(data);
            }else if( rc == comm.XL_LABELRANGES){
                var pos = 0;
                pos = comm.unpack_cell_range_address_list_update_pos(
                        self.rowLabelRanges, data, pos, bv, 8
                        );
                pos = comm.unpack_cell_range_address_list_update_pos(
                        self.colLabelRanges, data, pos, bv, 8
                        );
                assert (pos == dataLen,'pos == dataLen');
            }else if( rc == comm.XL_ARRAY){
                var row1x = data.readUInt16LE(0), 
					rownx = data.readUInt16LE(2), 
					col1x = data.readUInt8(3), 
					colnx = data.readUInt8(4), 
					arrayFlags = data.readUInt8(5), 
					tokslen =  data.readUInt16LE(12);
                    
                if ( blahFormulas)
                    console.log("ARRAY:", row1x, rownx, col1x, colnx, arrayFlags);
                    // dump_formula(bk, data[14:], tokslen, bv, reldelta=0, blah=1)
            }else if( rc == comm.XL_SHRFMLA){
			//todo 나중에
			/*
                row1x, rownx, col1x, colnx, nfmlas, tokslen = \
                    local_unpack("<HHBBxBH", data[:10]);
                if ( blahFormulas){
                    print("SHRFMLA (main):", row1x, rownx, col1x, colnx, nfmlas, file=self.logfile);
                    decompile_formula(bk, data[10:], tokslen, FMLA_TYPE_SHARED,
                        blah=1, browx=rowx, bcolx=colx, r1c1=r1c1);
				}
			*/
            }else if( rc == comm.XL_CONDFMT){
                if ( !fmtInf) continue;
                assert(bv >= 80,'bv >= 80');
				//todo 나중에
				/*
                num_CFs, needs_recalc, browx1, browx2, bcolx1, bcolx2 = \
                    unpack("<6H", data[0:12]);
                if ( self.verbosity >= 1){
                    fprintf(self.logfile,
                        "\n*** WARNING: Ignoring CONDFMT (conditional formatting) record\n" \
                        "*** in Sheet %d (%r).\n" \
                        "*** %d CF record(s); needs_recalc_or_redraw = %d\n" \
                        "*** Bounding box is %s\n",
                        self.index, self.name, num_CFs, needs_recalc,
                        rangename2d(browx1, browx2+1, bcolx1, bcolx2+1),
                        );
				}
                olist = []; // updated by the function
                pos = comm.unpack_cell_range_address_list_update_pos(
                    olist, data, 12, bv, addr_size=8);
                // print >> self.logfile, repr(result), len(result)
                if ( self.verbosity >= 1)
                    fprintf(self.logfile,
                        "*** %d individual range(s):\n" \
                        "*** %s\n",
                        len(olist),
                        ", ".join([rangename2d(*coords) for coords in olist]),
                        );
				*/
            }else if( rc == comm.XL_CF){
                if ( ! fmtInf) continue;
				//todo 나중에
				/*
                cf_type, cmp_op, sz1, sz2, flags = unpack("<BBHHi", data[0:10]);
                font_block = (flags >> 26) & 1;
                bord_block = (flags >> 28) & 1;
                patt_block = (flags >> 29) & 1;
                if ( self.verbosity >= 1)
                    fprintf(self.logfile,
                        "\n*** WARNING: Ignoring CF (conditional formatting) sub-record.\n" \
                        "*** cf_type=%d, cmp_op=%d, sz1=%d, sz2=%d, flags=0x%08x\n" \
                        "*** optional data blocks: font=%d, border=%d, pattern=%d\n",
                        cf_type, cmp_op, sz1, sz2, flags,
                        font_block, bord_block, patt_block,
                        );
                // hex_char_dump(data, 0, dataLen, fout=self.logfile)
                pos = 12;
                if ( font_block)
                    (font_height, font_options, weight, escapement, underline,
                    font_colour_index, two_bits, font_esc, font_underl) = \
                    unpack("<64x i i H H B 3x i 4x i i i 18x", data[pos:pos+118]);
                    font_style = (two_bits > 1) & 1;
                    posture = (font_options > 1) & 1;
                    font_canc = (two_bits > 7) & 1;
                    cancellation = (font_options > 7) & 1;
                    if (self.verbosity >= 1)
                        fprintf(self.logfile,
                            "*** Font info: height=%d, weight=%d, escapement=%d,\n" \
                            "*** underline=%d, colour_index=%d, esc=%d, underl=%d,\n" \
                            "*** style=%d, posture=%d, canc=%d, cancellation=%d\n",
                            font_height, weight, escapement, underline,
                            font_colour_index, font_esc, font_underl,
                            font_style, posture, font_canc, cancellation,
                            );
                    pos += 118;
                if ( bord_block)
                    pos += 8;
                if ( patt_block)
                    pos += 4;
                fmla1 = data[pos:pos+sz1];
                pos += sz1;
                if ( blah and sz1)
                    fprintf(self.logfile,
                        "*** formula 1:\n",
                        )
                    dump_formula(bk, fmla1, sz1, bv, reldelta=0, blah=1);
                fmla2 = data[pos:pos+sz2];
                pos += sz2;
                assert pos == dataLen;
                if ( blah and sz2){
                    fprintf(self.logfile,
                        "*** formula 2:\n",
                        );
                    dump_formula(bk, fmla2, sz2, bv, reldelta=0, blah=1);
				}
				*/
            }else if( rc == comm.XL_DEFAULTROWHEIGHT){
				//todo 나중에
				/*
                if ( dataLen == 4){
                    bits, self.defaultRowHeight = unpack("<HH", data[:4]);
                }else if( dataLen == 2){
                    self.defaultRowHeight, = unpack("<H", data);
                    bits = 0;
                    fprintf(self.logfile,
                        "*** WARNING: DEFAULTROWHEIGHT record len is 2, " \
                        "should be 4; assuming BIFF2 format\n");
                }else{
                    bits = 0;
                    fprintf(self.logfile,
                        "*** WARNING: DEFAULTROWHEIGHT record len is %d, " \
                        "should be 4; ignoring this record\n",
                        dataLen);
                self.defaultRowHeightMismatch = bits & 1;
                self.defaultRowHidden = (bits >> 1) & 1;
                self.defaultAdditionalSpaceAbove = (bits >> 2) & 1;
                self.defaultAdditionalSpaceBelow = (bits >> 3) & 1;
				*/
            }else if( rc == comm.XL_MERGEDCELLS){
                if ( !fmtInf) continue;
                pos = comm.unpack_cell_range_address_list_update_pos(
                    self.mergedCells, data, 0, bv, addr_size=8);
                if ( blah)
                    console.log(util.format(
                        "MERGEDCELLS: %d ranges\n", Math.floor((pos - 2) / 8)));
                assert(pos == dataLen, util.format("MERGEDCELLS: pos=%d dataLen=%d",pos, dataLen));
            }else if( rc == comm.XL_WINDOW2){
                if ( bv >= 80 && dataLen >= 14){
                    var options = data.readUInt16LE(0);
                    self.firstVisibleRowIndex = data.readUInt16LE(2);
					self.firstVisibleColIndex = data.readUInt16LE(4);
                    self.gridlineColourIndex = data.readUInt16LE(6);
                    self.cachedPageBreakPreviewMagFactor = data.readUInt16LE(10);
                    self.cachedNormalViewMagFactor = data.readUInt16LE(12);
                }else{
                    assert( bv >= 30, 'bv >= 30'); // BIFF3-7
                    var options = data.readUInt16LE(0);
                    self.firstVisibleRowIndex = data.readUInt16LE(2);
					self.firstVisibleColIndex = data.readUInt16LE(4);
                    self.gridlineColourRgb = [data.readUInt8(6), data.readUInt8(7),data.readUInt8(8)];
                    self.gridlineColourIndex = nearest_colour_index(
                        book.colourMap, self.gridlineColourRgb, 0);
                    self.cachedPageBreakPreviewMagFactor = 0; // default (60%)
                    self.cachedNormalViewMagFactor = 0; // default (100%)
				}
                // options -- Bit, Mask, Contents:
                // 0 0001H 0 = Show formula results 1 = Show formulas
                // 1 0002H 0 = Do not show grid lines 1 = Show grid lines
                // 2 0004H 0 = Do not show sheet headers 1 = Show sheet headers
                // 3 0008H 0 = Panes are not frozen 1 = Panes are frozen (freeze)
                // 4 0010H 0 = Show zero values as empty cells 1 = Show zero values
                // 5 0020H 0 = Manual grid line colour 1 = Automatic grid line colour
                // 6 0040H 0 = Columns from left to right 1 = Columns from right to left
                // 7 0080H 0 = Do not show outline symbols 1 = Show outline symbols
                // 8 0100H 0 = Keep splits if pane freeze is removed 1 = Remove splits if pane freeze is removed
                // 9 0200H 0 = Sheet not selected 1 = Sheet selected (BIFF5-BIFF8)
                // 10 0400H 0 = Sheet not visible 1 = Sheet visible (BIFF5-BIFF8)
                // 11 0800H 0 = Show in normal view 1 = Show in page break preview (BIFF8)
                // The freeze flag specifies, if a following PANE record (6.71) describes unfrozen or frozen panes.
                for(var key in _WINDOW2_options){
					self[key] =_WINDOW2_options[key] & 1;
					options >>= 1;
				}
            }else if( rc == comm.XL_SCL){
				//todo 나중에
				/*
                num, den = unpack("<HH", data);
                result = 0;
                if ( den)
                    result = (num * 100) // den;
                if ( not(10 <= result <= 400)){
                    if ( DEBUG || self.verbosity >= 0){
                        print((
                            "WARNING *** SCL rcd sheet %d: should have 0.1 <= num/den <= 4; got %d/%d"
                            % (self.index, num, den)
                            ), file=self.logfile);
					}
                    result = 100;
				}
                self.sclMagFactor = result;
				*/
            }else if( rc == comm.XL_PANE){
				//todo 나중에
				/*
                (
                self.vertSplitPos,
                self.horzSplitPos,
                self.horzSplitFirstVisible,
                self.vertSplitFirstVisible,
                self.splitActivePane,
                ) = unpack("<HHHHB", data[:9]);
                self.hasPaneRecord = 1;
				*/
            }else if( rc == comm.XL_HORIZONTALPAGEBREAKS){
                if ( !fmtInf) continue;
				//나중에
				/*
                num_breaks, = local_unpack("<H", data[:2]);
                assert num_breaks * (2 + 4 * (bv >= 80)) + 2 == dataLen;
                pos = 2;
                if ( bv < 80)
                    while pos < dataLen)
                        self.horizontalPageBreaks.append((local_unpack("<H", data[pos:pos+2])[0], 0, 255));
                        pos += 2;
                else{
                    while pos < dataLen){
                        self.horizontalPageBreaks.append(local_unpack("<HHH", data[pos:pos+6]));
                        pos += 6;
					}
				}
				*/
            }else if( rc == comm.XL_VERTICALPAGEBREAKS){
                if ( !fmtInf) continue;
				//todo 나중에
				/*
                num_breaks, = local_unpack("<H", data[:2]);
                assert num_breaks * (2 + 4 * (bv >= 80)) + 2 == dataLen;
                pos = 2;
                if ( bv < 80)
                    while pos < dataLen){
                        self.verticalPageBreaks.append((local_unpack("<H", data[pos:pos+2])[0], 0, 65535));
                        pos += 2;
					}
                else
                    while pos < dataLen){
                        self.verticalPageBreaks.append(local_unpack("<HHH", data[pos:pos+6]));
                        pos += 6;
					}
				*/
            //////// all of the following are for BIFF <= 4W
            }else if( bv <= 45){
				//todo 나중에
				/*
                if ( rc == comm.XL_FORMAT || rc == comm.XL_FORMAT2){
                    bk.handle_format(data, rc);
                }else if( rc == comm.XL_FONT || rc == comm.XL_FONT_B3B4){
                    bk.handle_font(data);
                }else if( rc == comm.XL_STYLE){
                    if ( not book._xfEpilogueDone) book.xfEpilogue();
                    bk.handle_style(data);
                }else if( rc == comm.XL_PALETTE){
                    bk.handle_palette(data);
                }else if( rc == comm.XL_BUILTINFMTCOUNT){
                    bk.handle_builtinfmtcount(data);
                }else if( rc == comm.XL_XF4 || rc == comm.XL_XF3 || rc == comm.XL_XF2){ //////// N.B. not XL_XF
                    bk.handleXf(data);
                }else if( rc == comm.XL_DATEMODE){
                    bk.handle_datemode(data);
                }else if( rc == comm.XL_CODEPAGE){
                    bk.handleCodePage(data);
                }else if( rc == comm.XL_FILEPASS){
                    bk.handle_filepass(data);
                }else if( rc == comm.XL_WRITEACCESS){
                    bk.handle_writeaccess(data);
                }else if( rc == comm.XL_IXFE){
                    self._ixfe = local_unpack('<H', data)[0];
                }else if( rc == comm.XL_NUMBER_B2){
                    rowx, colx, cell_attr, d = local_unpack('<HH3sd', data);
                    putCell(rowx, colx, null, d, self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_INTEGER){
                    rowx, colx, cell_attr, d = local_unpack('<HH3sH', data);
                    putCell(rowx, colx, null, float(d), self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_LABEL_B2){
                    rowx, colx, cell_attr = local_unpack('<HH3s', data[0:7]);
                    strg = unpackString(data, 7, bk.encoding || bk.deriveEncoding(), lenRecordLength=1);
                    putCell(rowx, colx, comm.XL_CELL_TEXT, strg, self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_BOOLERR_B2){
                    rowx, colx, cell_attr, value, isErr = local_unpack('<HH3sBB', data);
                    cellTyp = (comm.XL_CELL_BOOLEAN, comm.XL_CELL_ERROR)[isErr];
                    // if DEBUG: print "XL_BOOLERR_B2", rowx, colx, cell_attr, value, isErr
                    putCell(rowx, colx, cellTyp, value, self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_BLANK_B2){
                    if ( not fmtInf) continue;
                    rowx, colx, cell_attr = local_unpack('<HH3s', data[:7]);
                    putCell(rowx, colx, comm.XL_CELL_BLANK, '', self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_EFONT){
                    bk.handle_efont(data);
                }else if( rc == comm.XL_ROW_B2){
                    if ( not fmtInf) continue;
                    rowx, bits1, bits2 = local_unpack('<H4xH2xB', data[0:11]);
                    if ( not(0 <= rowx < self.utterMaxRows)){
                        print("*** NOTE: ROW_B2 record has row index %d; " \
                            "should have 0 <= rowx < %d -- record ignored!" \
                            % (rowx, self.utterMaxRows), file=self.logfile);
                        continue;
					}
                    if ( not (bits2 & 1)){  // has_default_xf_index is false
                        xfIndex = -1;
                    }else if( dataLen == 18){
                        // Seems the XF index in the cell_attr is dodgy
                         xfx = local_unpack('<H', data[16:18])[0];
                         xfIndex = self.fixed_BIFF2_xfindex(cell_attr=null, rowx=rowx, colx=-1, true_xfx=xfx);
                    }else{
                        cell_attr = data[13:16];
                        xfIndex = self.fixed_BIFF2_xfindex(cell_attr, rowx, colx=-1);
					}
                    key = (bits1, bits2, xfIndex);
                    r = rowInfoSharingDict.get(key);
                    if ( r is null){
                        rowInfoSharingDict[key] = r = Rowinfo();
                        r.height = bits1 & 0x7fff;
                        r.has_default_height = (bits1 >> 15) & 1;
                        r.has_default_xf_index = bits2 & 1;
                        r.xfIndex = xfIndex;
                        // r.outline_level = 0             // set in __init__
                        // r.outline_group_starts_ends = 0 // set in __init__
                        // r.hidden = 0                    // set in __init__
                        // r.height_mismatch = 0           // set in __init__
                        // r.additional_space_above = 0    // set in __init__
                        // r.additional_space_below = 0    // set in __init__
					}
                    self.rowInfoMap[rowx] = r;
                    if ( 0 and r.xfIndex > -1){
                        fprintf(self.logfile,
                            "**ROW %d %d %d\n",
                            self.index, rowx, r.xfIndex);
					}
                    if ( blahRows){
                        print('ROW_B2', rowx, bits1, has_defaults, file=self.logfile);
                        r.dump(self.logfile,
                            header="--- sh //%d, rowx=%d ---" % (self.index, rowx));
					}
                }else if( rc == comm.XL_COLWIDTH){ // BIFF2 only
                    if ( not fmtInf) continue;
                    first_colx, last_colx, width\
                        = local_unpack("<BBH", data[:4]);
                    if ( not(first_colx <= last_colx)){
                        print("*** NOTE: COLWIDTH record has first col index %d, last %d; " \
                            "should have first <= last -- record ignored!" \
                            % (first_colx, last_colx), file=self.logfile);
                        continue;
					}
                    for colx in xrange(first_colx, last_colx+1){
                        if ( colx in self.colInfoMap){
                            c = self.colInfoMap[colx];
                        }else{
                            c = Colinfo();
                            self.colInfoMap[colx] = c;
						}
                        c.width = width;
					}
                    if ( blah)
                        fprintf(
                            self.logfile,
                            "COLWIDTH sheet //%d cols %d-%d: wid=%d\n",
                            self.index, first_colx, last_colx, width
                            );
                }else if( rc == comm.XL_COLUMNDEFAULT) // BIFF2 only
                    if ( not fmtInf) continue;
                    first_colx, last_colx = local_unpack("<HH", data[:4]);
                    //////// Warning OOo docs wrong; first_colx <= colx < last_colx
                    if ( blah)
                        fprintf(
                            self.logfile,
                            "COLUMNDEFAULT sheet //%d cols in range(%d, %d)\n",
                            self.index, first_colx, last_colx
                            );
                    if ( not(0 <= first_colx < last_colx <= 256)){
                        print("*** NOTE: COLUMNDEFAULT record has first col index %d, last %d; " \
                            "should have 0 <= first < last <= 256" \
                            % (first_colx, last_colx), file=self.logfile);
                        last_colx = min(last_colx, 256);
					}
                    for colx in xrange(first_colx, last_colx)){
                        offset = 4 + 3 * (colx - first_colx);
                        cell_attr = data[offset:offset+3];
                        xfIndex = self.fixed_BIFF2_xfindex(cell_attr, rowx=-1, colx=colx);
                        if ( colx in self.colInfoMap){
                            c = self.colInfoMap[colx];
                        }else{
                            c = Colinfo();
                            self.colInfoMap[colx] = c;
						}
                        c.xfIndex = xfIndex;
                }else if( rc == comm.XL_WINDOW2_B2){ // BIFF 2 only
                    attr_names = ("show_formulas", "show_grid_lines", "show_sheet_headers",
                        "panes_are_frozen", "show_zero_values");
                    for attr, char in zip(attr_names, data[0:5]))
                        setattr(self, attr, int(char != b'\0'));
                    (self.firstVisibleRowIndex, self.firstVisibleColIndex,
                    self.automatic_grid_line_colour,
                    ) = unpack("<HHB", data[5:10]);
                    self.gridlineColourRgb = unpack("<BBB", data[10:13]);
                    self.gridlineColourIndex = nearest_colour_index(
                        book.colourMap, self.gridlineColourRgb, debug=0);
                    self.cachedPageBreakPreviewMagFactor = 0 ;// default (60%)
                    self.cachedNormalViewMagFactor = 0; // default (100%)
				}
				*/
            }else
                // if DEBUG: print "SHEET.READ: Unhandled record type %02x %d bytes %r" % (rc, dataLen, data)
                ;
		}
        if (!eofFound)
            throw new Error(util.format("Sheet %d (%r) missing EOF record",self.index, self.name));
        tidyDimensions();
        updateCookedMagFactors();
        bk._position = oldPos;
        return 1;
    }
	
	
    function updateCookedMagFactors(){
        // Cached values are used ONLY for the non-active view mode.
        // When the user switches to the non-active view mode,
        // if the cached value for that mode is not valid,
        // Excel pops up a window which says:
        // "The number must be between 10 and 400. Try again by entering a number in this range."
        // When the user hits OK, it drops into the non-active view mode
        // but uses the magn from the active mode.
        // NOTE: definition of "valid" depends on mode ... see below
        var blah = DEBUG || self.verbosity > 0;
        if (self.showInPageBreakPreview){
            if (!self.sclMagFactor) // no SCL record
                self.cookedPageBreakPreviewMagFactor = 100; // Yes, 100, not 60, NOT a typo
            else
                self.cookedPageBreakPreviewMagFactor = self.sclMagFactor;
            var zoom = self.cachedNormalViewMagFactor
            if (!(10 <= zoom && zoom <=400))
                if (blah)
                    console.log(util.format(
                        "WARNING *** WINDOW2 rcd sheet %d: Bad cachedNormalViewMagFactor: %d",
                        self.index, self.cachedNormalViewMagFactor
                        ));
                zoom = self.cookedPageBreakPreviewMagFactor;
            self.cookedNormalViewMagFactor = zoom;
        }else{
            // normal view mode
            if(!self.sclMagFactor) // no SCL record
                self.cookedNormalViewMagFactor = 100;
            else
                self.cookedNormalViewMagFactor = self.sclMagFactor;
            var zoom = self.cachedPageBreakPreviewMagFactor;
            if (zoom == 0)
                // VALID, defaults to 60
                zoom = 60;
            else if (!(10 <= zoom && zoom <= 400)){
                if (blah)
                    console.log(util.format(
                        "WARNING *** WINDOW2 rcd sheet %r: Bad cachedPageBreakPreviewMagFactor: %r"
                        ,self.index, self.cachedPageBreakPreviewMagFactor)
                        );
                zoom = self.cookedNormalViewMagFactor;
			}
            self.cookedPageBreakPreviewMagFactor = zoom;
		}
	}
}


function unpackRK(data, start, end){
    var flags = data[start];
    if (flags & 2){
        // There's a SIGNED 30-bit integer in there!
        var i = data.readInt32LE(start);
        i >>= 2;// div by 4 to drop the 2 flag bits
        if (flags & 1)
            return i / 100.0;
        return i*1.0;
    }else{
        // It's the most significant 30 bits of an IEEE 754 64-bit FP number
		var buf = new Buffer([0,0,0,0,flags & 252,data[start+1],data[start+2],data[start+3]] );
        var d = buf.readDoubleLE(0);
        if (flags & 1)
            return d / 100.0;
        return d;
	}
}

var _JDN_delta = [2415080 - 61, 2416482 - 1];
var _XLDAYS_TOO_LARGE = [2958466, 2958466 - 1462] // This is equivalent to 10000-01-01
function toDate(xldate, dateMode){
    if(dateMode != 0 &&  dateMode != 1 )
        throw new Error('bad date mode : ' + dateMode);
    if (xldate == 0.0)
        return new Date();
    if (xldate < 0.0)
		throw new Error('date is negative : ' + xldate);
    var xldays = Math.floor(xldate);
    var remain = xldate - xldays;
    var seconds = Math.round(remain * 86400.0);
    assert(0 <= seconds && seconds <= 86400, util.format('0 <= %d <= 86400',seconds));
	var hour = 0, minute = 0, second = 0;
    if (seconds == 86400){
        hour = minute = second = 0;
        xldays += 1;
    }else{
        second = seconds % 60; 
		var minutes = Math.floor(seconds/60);
        minute = minutes % 60; 
		hour = Math.floor(minutes/60);
	}
    if(xldays >= _XLDAYS_TOO_LARGE[dateMode])
        throw new Error('date is too large : ' + xldate); 

    if (xldays == 0)
        return new Date(0, 0, 0, hour, minute, second);

    if (xldays < 61 && dateMode == 0)
        throw new Error('date is ambiguous : ' +xldate);

    var jdn = xldays + _JDN_delta[dateMode];
	
    var yreg = (Math.floor((Math.floor(jdn * 4 + 274277) / 146097) * 3 / 4) + jdn + 1363) * 4 + 3;
    var mp = Math.floor((yreg % 1461) / 4) * 535 + 333
    var d = Math.floor((mp % 16384) / 535) + 1
    
    mp >>= 14;
    if(mp >= 10)
        return new Date (Math.floor(yreg / 1461) - 4715, mp - 9, d, hour, minute, second);
    else
        return new Date (Math.floor(yreg / 1461) - 4716, mp + 3, d, hour, minute, second);
}


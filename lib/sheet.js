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
    "show_in_page_break_preview" : 0,
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


module.exports = function Sheet(book, position, name, number){
	var self = this;
    ////
    // Name of sheet.
    var name = name;

    ////
    // A reference to the Book object to which this sheet belongs.
    // Example usage: some_sheet.book.datemode
    var book = book;
    
    ////
    // Number of rows in sheet. A row index is in range(thesheet.nrows).
    var nrows = 0;

    ////
    // Nominal number of columns in sheet. It is 1 + the maximum column index
    // found, ignoring trailing empty cells. See also open_workbook(ragged_rows=?)
    // and Sheet.{@link //Sheet.rowLength}(row_index).
    var ncols = 0;

    ////
    // The map from a column index to a {@link //Colinfo} object. Often there is an entry
    // in COLINFO records for all column indexes in range(257).
    // Note that xlrd ignores the entry for the non-existent
    // 257th column. On the other hand, there may be no entry for unused columns.
    // Populated only if open_workbook(formatting_info=True).
    var colinfo_map = {};

    ////
    // The map from a row index to a {@link //Rowinfo} object. Note that it is possible
    // to have missing entries -- at least one source of XLS files doesn't
    // bother writing ROW records.
    // Populated only if open_workbook(formatting_info=True).
    var rowinfo_map = {};

    ////
    // List of address ranges of cells containing column labels.
    // These are set up in Excel by Insert > Name > Labels > Columns.
    //  How to deconstruct the list:
    //  
    // for crange in thesheet.col_label_ranges:
    //     rlo, rhi, clo, chi = crange
    //     for rx in xrange(rlo, rhi):
    //         for cx in xrange(clo, chi):
    //             print "Column label at (rowx=%d, colx=%d) is %r" \
    //                 (rx, cx, thesheet.cellValue(rx, cx))
    //  
    var col_label_ranges = [];

    ////
    // List of address ranges of cells containing row labels.
    // For more details, see  col_label_ranges  above.
    var row_label_ranges = [];

    ////
    // List of address ranges of cells which have been merged.
    // These are set up in Excel by Format > Cells > Alignment, then ticking
    // the "Merge cells" box.
    // Extracted only if open_workbook(formatting_info=True).
    //  How to deconstruct the list:
    //  
    // for crange in thesheet.merged_cells:
    //     rlo, rhi, clo, chi = crange
    //     for rowx in xrange(rlo, rhi):
    //         for colx in xrange(clo, chi):
    //             // cell (rlo, clo) (the top left one) will carry the data
    //             // and formatting info; the remainder will be recorded as
    //             // blank cells, but a renderer will apply the formatting info
    //             // for the top left cell (e.g. border, pattern) to all cells in
    //             // the range.
    //  
    var merged_cells = [];
    
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
    // runlist = thesheet.rich_text_runlist_map.get((rowx, colx))
    // if runlist:
    //     for offset, font_index in runlist:
    //         // do work here.
    //         pass
    //  
    // Populated only if open_workbook(formatting_info=True).
   
    var rich_text_runlist_map = {};

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
    var defcolwidth = null;

    ////
    // Default column width from STANDARDWIDTH record, else null.
    // From the OOo docs: 
    // """Default width of the columns in 1/256 of the width of the zero
    // character, using default font (first FONT record in the file).""" 
    // For the default hierarchy, refer to the {@link //Colinfo} class.
    var standardwidth = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var default_row_height = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var default_row_height_mismatch = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var default_row_hidden = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var default_additional_space_above = null;

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the  optional  DEFAULTROWHEIGHT record.
    var default_additional_space_below = null;

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
    var hyperlink_list = [];

    ////
    //  A sparse mapping from (rowx, colx) to an item in {@link //Sheet.hyperlink_list}.
    // Cells not covered by a hyperlink are not mapped.
    // It is possible using the Excel UI to set up a hyperlink that 
    // covers a larger-than-1x1 rectangle of cells.
    // Hyperlink rectangles may overlap (Excel doesn't check).
    // When a multiply-covered cell is clicked on, the hyperlink that is activated
    // (and the one that is mapped here) is the last in hyperlink_list.
    var hyperlink_map = {};

    ////
    //  A sparse mapping from (rowx, colx) to a {@link //Note} object.
    // Cells not containing a note ("comment") are not mapped.
    var cell_note_map = {};    
    
    ////
    // Number of columns in left pane (frozen panes; for split panes, see comments below in code)
    var vert_split_pos = 0;

    ////
    // Number of rows in top pane (frozen panes; for split panes, see comments below in code)
    var horz_split_pos = 0;

    ////
    // Index of first visible row in bottom frozen/split pane
    var horz_split_first_visible = 0;

    ////
    // Index of first visible column in right frozen/split pane
    var vert_split_first_visible = 0;

    ////
    // Frozen panes: ignore it. Split panes: explanation and diagrams in OOo docs.
    var split_active_pane = 0;

    ////
    // Boolean specifying if a PANE record was present, ignore unless you're xlutils.copy
    has_pane_record = 0;

    ////
    // A list of the horizontal page breaks in this sheet.
    // Breaks are tuples in the form (index of row after break, start col index, end col index).
    // Populated only if open_workbook(formatting_info=True).
    horizontal_page_breaks = [];

    ////
    // A list of the vertical page breaks in this sheet.
    // Breaks are tuples in the form (index of col after break, start row index, end row index).
    // Populated only if open_workbook(formatting_info=True).
    vertical_page_breaks = [];
    
	//self.book = book;
	self.biffVersion = book.biffVersion;
	self._position = position;
	self.logfile = book.logfile;
	self.bt = comm.XL_CELL_EMPTY; //unsigned char 1byte
	self.bf = -1; //signed short 2bytes
	self.name = name;
	self.number = number;
	self.verbosity = book.verbosity;
	self.formatting_info = book.formatting_info;
	self.ragged_rows = book.ragged_rows;
	if (self.ragged_rows)
		self.put_cell = put_cell_ragged;
	else
		self.put_cell = put_cell_unragged;
	self._xf_index_to_xl_type_map = book._xf_index_to_xl_type_map;
	self.nrows = 0; // actual, including possibly empty cells
	self.ncols = 0;
	self._maxdatarowx = -1; // highest rowx containing a non-empty cell
	self._maxdatacolx = -1; // highest colx containing a non-empty cell
	self._dimnrows = 0; // as per DIMENSIONS record
	self._dimncols = 0;
	var _cellValues = [];
	var _cellTypes = [];
	var _cellXFIndexes = [];
	self.defcolwidth = null;
	self.standardwidth = null;
	self.default_row_height = null;
	self.default_row_height_mismatch = 0;
	self.default_row_hidden = 0;
	self.default_additional_space_above = 0;
	self.default_additional_space_below = 0;
	self.colinfo_map = {};
	self.rowinfo_map = {};
	self.col_label_ranges = [];
	self.row_label_ranges = [];
	self.merged_cells = [];
	self.rich_text_runlist_map = {};
	self.horizontal_page_breaks = [];
	self.vertical_page_breaks = [];
	self._xf_index_stats = [0, 0, 0, 0];
	self.visibility = book._sheet_visibility[number]; // from BOUNDSHEET record
	for(var key in _WINDOW2_options) self[key] =_WINDOW2_options[key];
	self.first_visible_rowx = 0;
	self.first_visible_colx = 0;
	self.gridline_colour_index = 0x40;
	self.gridline_colour_rgb = null; // pre-BIFF8
	self.hyperlink_list = [];
	self.hyperlink_map = {};
	self.cell_note_map = {};

	// Values calculated by xlrd to predict the mag factors that
	// will actually be used by Excel to display your worksheet.
	// Pass these values to xlwt when writing XLS files.
	// Warning 1: Behaviour of OOo Calc and Gnumeric has been observed to differ from Excel's.
	// Warning 2: A value of zero means almost exactly what it says. Your sheet will be
	// displayed as a very tiny speck on the screen. xlwt will reject attempts to set
	// a mag_factor that is not (10 <= mag_factor <= 400).
	self.cooked_page_break_preview_mag_factor = 60;
	self.cooked_normal_view_mag_factor = 100;

	// Values (if any) actually stored on the XLS file
	self.cached_page_break_preview_mag_factor = null; // from WINDOW2 record
	self.cached_normal_view_mag_factor = null; // from WINDOW2 record
	self.scl_mag_factor = null; // from SCL record

	self._ixfe = null; // BIFF2 only
	self._cell_attr_to_xfx = {}; // BIFF2.0 only

	//////// Don't initialise this here, use class attribute initialisation.
	//////// self.gcw = (0, ) * 256 ////////

	if (self.biffVersion >= 80)
		self.utter_max_rows = 65536;
	else
		self.utter_max_rows = 16384;
	self.utter_max_cols = 256;

	self._first_full_rowx = -1;

	// self._put_cell_exceptions = 0
	// self._put_cell_row_widenings = 0
	// self._put_cell_rows_appended = 0
	// self._put_cell_cells_appended = 0


    // === Following methods are used in building the worksheet.
    // === They are not part of the API.

    function tidy_dimensions(){
        if (self.verbosity >= 3)
            console.log(util.format("tidy_dimensions: nrows=%d ncols=%d \n",self.nrows, self.ncols));
        if(1 && self.merged_cells && self.merged_cells.length){
            var nr = nc = 0;
            var umaxrows = self.utter_max_rows;
            var umaxcols = self.utter_max_cols;
			self.merged_cells.forEach(function(x){
				var rlo = x[0], rhi =x[1], clo = x[2], chi = x[3];
                if (!(0 <= rlo &&  rlo  < rhi && rhi <= umaxrows) || !(0 <= clo && clo < chi && chi <= umaxcols))
                    console.log(util.format(
                        "*** WARNING: sheet #%d (%r), MERGEDCELLS bad range %r\n",
                        self.number, self.name, crange));
                if (rhi > nr) nr = rhi;
                if (chi > nc) nc = chi;
			});
            if( nc > self.ncols)
                self.ncols = nc;
            if( nr > self.nrows)
                // we put one empty cell at (nr-1,0) to make sure
                // we have the right number of rows. The ragged rows
                // will sort out the rest if needed.
                self.put_cell(nr-1, 0, comm.XL_CELL_EMPTY, '', -1);
		}
        if(self.verbosity >= 1 && (self.nrows != self._dimnrows || self.ncols != self._dimncols))
            console.log(util.format("NOTE *** sheet %d (%r): DIMENSIONS R,C = %d,%d should be %d,%d\n",
                self.number,
                self.name,
                self._dimnrows,
                self._dimncols,
                self.nrows,
                self.ncols
                ));
        if (!self.ragged_rows){
            // fix ragged rows
            var ncols = self.ncols,
				s_cell_types = _cellTypes,
				s_cell_values = _cellValues,
				s_cell_xf_indexes = _cellXFIndexes,
				s_fmt_info = self.formatting_info;
            // for rowx in xrange(self.nrows):
            if(self._first_full_rowx == -2)
                ubound = self.nrows;
            else
                ubound = self._first_full_rowx;
            for(rowx =0;rowx < ubound ; rowx++){
                var trow = s_cell_types[rowx];
                var rlen = trow.length;
                var nextra = ncols - rlen;
                if(nextra > 0){
                    s_cell_values[rowx][ncols-1]= '';
					for(c =rlen;c < ncols-1; c++) s_cell_values[rowx][c]='';

					trow[ncols-1] = self.bt;
					for(c =rlen;c < ncols-1; c++) trow[c]=self.bt;
					
                    if (s_fmt_info) {
						s_cell_xf_indexes[rowx][ncols-1] = self.bf;
						for(c =rlen;c < ncols-1; c++) s_cell_xf_indexes[rowx][c]=self.bf;
					}
				}
			}
		}
    }
	
	function put_cell_ragged(rowx, colx, ctype, value, xf_index){
debugger;
        if (!ctype){
            // we have a number, so look up the cell type
            ctype = self._xf_index_to_xl_type_map[xf_index];
		}
        assert(0 <= colx && colx < self.utter_max_cols,'0 <= colx < self.utter_max_cols');
        assert(0 <= rowx && rowx < self.utter_max_rows,'0 <= rowx < self.utter_max_rows');
        var fmt_info = self.formatting_info;

		var nr = rowx + 1;
		if (elf.nrows < nr){
			_cellTypes[rowx] = [];
			_cellValues[rowx] =[];
			if (fmt_info) self.formatting_info[rowx] =[];
			self.nrows = nr;
		}else{
			if(_cellTypes[rowx] == undefined) _cellTypes = [];
			if(_cellValues[rowx] == undefined) _cellValues = [];
			if(fmt_info && self.formatting_info[rowx] == undefined) self.formatting_info = [];
		}
		var types_row = _cellTypes[rowx];
		var values_row = _cellValues[rowx];
		var fmt_row = null;
		if (fmt_info)
			fmt_row = _cellXFIndexes[rowx];
		var ltr = types_row.length;
		if (colx >= self.ncols)
			self.ncols = colx + 1;
		var num_empty = colx - ltr;
		
		types_row[colx] = ctype;
		values_row[colx] = value;
		if (fmt_info)
			fmt_row[colx] = xf_index;
		if(num_empty > 0){
			for(i = ltr; i < colx;i++) 
				if(type_row[i] == undefined) type_row[i] = self.bt;
			for(i = ltr; i < colx;i++) 
				if(values_row[i] == undefined) values_row[i] = '';
			if (fmt_info) 
				for(i = ltr; i < colx;i++) 
					if(fmt_row[i] == undefined) fmt_row[i] = self.bf;
		}
        
	}
	function extendCells(rowx, colx, ctype, value, xf_index){
		// print >> self.logfile, "put_cell extending", rowx, colx
		// self.extend_cells(rowx+1, colx+1)
		// self._put_cell_exceptions += 1
		var nr = rowx + 1;
		var nc = colx + 1;
		assert( 1 <= nc && nc <= self.utter_max_cols,'1 <= nc <= self.utter_max_cols');
		assert( 1 <= nr && nr <= self.utter_max_rows,'1 <= nr <= self.utter_max_rows');
		if (nc > self.ncols){
			self.ncols = nc;
			// The row self._first_full_rowx and all subsequent rows
			// are guaranteed to have length == self.ncols. Thus the
			// "fix ragged rows" section of the tidy_dimensions method
			// doesn't need to examine them.
			if (nr < self.nrows){
				// cell data is not in non-descending row order *AND*
				// self.ncols has been bumped up.
				// This very rare case ruins this optmisation.
				self._first_full_rowx = -2;
			}else if (rowx > self._first_full_rowx && self._first_full_rowx > -2)
				self._first_full_rowx = rowx;
		}
		if( nr <= self.nrows){
			// New cell is in an existing row, so extend that row (if necessary).
			// Note that nr < self.nrows means that the cell data
			// is not in ascending row order!!
			var trow = _cellTypes[rowx];
			var nextra = self.ncols - trow.length;
			if (nextra > 0){
				// self._put_cell_row_widenings += 1
				trow[self.ncols]=self.bt;
				trow.pop();
				for(c =trow.length;c < self.ncols; c++) trow[c]=self.bt;
				if (self.formatting_info){
					_cellXFIndexes[rowx][self.ncols]=self.bf;
					_cellXFIndexes[rowx].pop();
					for(c =trow.length;c < self.ncols; c++) _cellXFIndexes[rowx][c]=self.bf;
				}
				_cellValues[rowx][self.ncols]='';
				_cellValues[rowx].pop();
				for(c =trow.length;c < self.ncols; c++) _cellValues[rowx][c]='';
			}
		}else{
			var fmt_info = self.formatting_info;
			var nc = self.ncols,
				bt = self.bt,
				bf = self.bf;
			for(r = self.nrows; r < nr; r++){
				// self._put_cell_rows_appended += 1
				var ta = new Array(nc);
				for(c =0;c < nc; c++) ta[c]=bt;
				_cellTypes.push(ta);
				var va = new Array(nc);
				for(c =0;c < nc; c++) va[c]='';
				_cellValues.push(va);
				if (fmt_info){
					var fa = new Array(nc);
					for(c =0;c < nc; c++) fa[c]=bf;
					_cellXFIndexes.push(fa);
				}
			}
			self.nrows = nr;
		}
	}
    function put_cell_unragged(rowx, colx, ctype, value, xf_index){
		if(!ctype)
            // we have a number, so look up the cell type
            ctype = self._xf_index_to_xl_type_map[xf_index];
        // assert 0 <= colx < self.utter_max_cols
        // assert 0 <= rowx < self.utter_max_rows
		if(_cellTypes[rowx] == undefined || _cellTypes[rowx][colx] == undefined ){
			extendCells(rowx, colx, ctype, value, xf_index);
		}
		_cellTypes[rowx][colx] = ctype;
		_cellValues[rowx][colx] = value;
		if(self.formatting_info)
			_cellXFIndexes[rowx][colx] = xf_index;
        
	}
	 
	//#
    // {@link #Cell} object in the given row and column.
    self.cellObj = function (rowx, colx){
		var xfx = null;
        if (self.formatting_info) xfx = self.cellXFIndex(rowx, colx);
        return {"type":_cellTypes[rowx][colx],"value":_cellValues[rowx][colx],"xfIndex":xfx};
	}
    
    //#
    // Value of the cell in the given row and column.
    self.cell=function( rowx, colx){
        var val = _cellValues[rowx][colx];
        var typ = _cellTypes[rowx][colx];
        return typ == comm.XL_CELL_DATE?toDate(val, book.datemode):val;
	}
    
    self.cellValue=function( rowx, colx){
        return _cellValues[rowx][colx];
	}
    
    //formated value
    self.cellText = function(row, col){
    
    }
    //#
    // Type of the cell in the given row and column.
    // Refer to the documentation of the {@link #Cell} class.
    self.cellType=function ( rowx, colx){
        return _cellTypes[rowx][colx];
	}
		
	//#
    // XF index of the cell in the given row and column.
    // This is an index into Book.{@link #Book.xf_list}.
    self.cellXFIndex = function(rowx, colx){
        self.req_fmt_info();
        var xfx = _cellXFIndexes[rowx][colx];
        if (xfx > -1){
            self._xf_index_stats[0] += 1;
            return xfx;
		}
        // Check for a row xf_index
        if((xfx=self.rowinfo_map[rowx].xf_index)!= undefined){
            if (xfx > -1){
                self._xf_index_stats[1] += 1;
                return xfx;
			}
		}
        
        // Check for a column xf_index
        if((xfx = self.colinfo_map[colx].xf_index)!= undefined){
            if (xfx == -1) xfx = 15;
            self._xf_index_stats[2] += 1;
            return xfx;
		}else{
			// If all else fails, 15 is used as hardwired global default xf_index.
            self._xf_index_stats[3] += 1;
            return 15;
		}
	}
	
    // Returns the effective number of cells in the given row. For use with
    // open_workbook(ragged_rows=True) which is likely to produce rows
    // with fewer than {@link #Sheet.ncols} cells.
    self.rowLength=function(rowx){
        return _cellValues[rowx].length;
	}

    // Returns a sequence of the {@link #Cell} objects in the given row.
    self.row=function(rowx){
		var ret =[];
		for(colx=0;colx<_cellValues[rowx].length;colx++)
			ret.push(self.cellObj(rowx, colx));
        return ret;
	}

    // Returns a slice of the types
    // of the cells in the given row.
    self.rowTypes=function( rowx, start_colx, end_colx){
		start_colx=start_colx||0;

        if(end_colx == undefined)
            return _cellTypes[rowx].slice(start_colx);
		
        return _cellTypes[rowx].slice(start_colx,end_colx);
	}

    // Returns a slice of the values
    // of the cells in the given row.
    self.rowValues=function(rowx, start_colx, end_colx){
		start_colx=start_colx||0;
		end_colx=end_colx||null;
        if(end_colx == undefined)
            return _cellValues[rowx].slice(start_colx);
        return _cellValues[rowx].slice(start_colx,end_colx);
	}
		
	// === Methods after this line neither know nor care about how cells are stored.
	
    self.read = function(bk){
        //global rc_stats
        var DEBUG = 0;
        var blah = DEBUG || self.verbosity >= 2;
        var blah_rows = DEBUG || self.verbosity >= 4;
        var blah_formulas = 0 && blah;
        var r1c1 = 0;
        var oldpos = bk._position;
        bk._position = self._position;
        XL_SHRFMLA_ETC_ETC = [comm.XL_SHRFMLA, comm.XL_ARRAY, comm.XL_TABLEOP, comm.XL_TABLEOP2,
            comm.XL_ARRAY2, comm.XL_TABLEOP_B2];            
        var self_put_cell = self.put_cell;
        var bk_get_record_parts = bk.get_record_parts;
        var bv = self.biffVersion;
        var fmt_info = self.formatting_info;
        var do_sst_rich_text = fmt_info && bk._rich_text_runlist_map;
        var rowinfo_sharing_dict = {};
        var txos = {};
        var eof_found = 0;
        while(true){
            // if DEBUG: print "SHEET.READ: about to read from position %d" % bk._position
            var record = bk_get_record_parts();
			var rc = record[0], data_len = record[1], data = record[2];
            // if rc in rc_stats:
            //     rc_stats[rc] += 1
            // else:
            //     rc_stats[rc] = 1
            // if DEBUG: print "SHEET.READ: op 0x%04x, %d bytes %r" % (rc, data_len, data)
            if (rc == comm.XL_NUMBER){
                // [:14] in following stmt ignores extraneous rubbish at end of record.
                // Sample file testEON-8.xls supplied by Jan Kraus.
				
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4), 
					d = data.readDoubleLE(6);
                // if xf_index == 0:
                //     fprintf(self.logfile,
                //         "NUMBER: r=%d c=%d xfx=%d %f\n", rowx, colx, xf_index, d)
                self_put_cell(rowx, colx, null, d, xf_index);
            }else if( rc == comm.XL_LABELSST){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4), 
					sstindex = data.readInt32LE(6);;
                // print "LABELSST", rowx, colx, sstindex, bk._sharedstrings[sstindex]
                self_put_cell(rowx, colx, comm.XL_CELL_TEXT, bk._sharedstrings[sstindex], xf_index);
                if (do_sst_rich_text){
                    var runlist = bk._rich_text_runlist_map.get(sstindex);
                    if (runlist){
                        var row = self.rich_text_runlist_map[rowx]
						if(!row) row = self.rich_text_runlist_map[rowx] = [];
						row[colx] = runlist;
					}
				}
            }else if( rc == comm.XL_LABEL){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4);
                if ( bv < comm.BIFF_FIRST_UNICODE)
                    strg = comm.unpack_string(data, 6, bk.encoding || bk.derive_encoding(), 2);
                else
                    strg = unpack_unicode(data, 6, 2);
                self_put_cell(rowx, colx, comm.XL_CELL_TEXT, strg, xf_index);
            }else if( rc == comm.XL_RSTRING){
                var rowx = data.readUInt16LE(0),
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4);
                if ( bv < comm.BIFF_FIRST_UNICODE){
                    var unpack = unpack_string_update_pos(data, 6, bk.encoding || bk.derive_encoding(), 2);
					var strg = unpack[0], pos = unpack[1];
                    var nrt = data[pos];
                    pos += 1;
                    var runlist = [];
                    while (pos< nrt) runlist.append( data.readUInt8(pos++));
                    assert(pos == data.length,'pos == len(data)');
                }else{
					var unpack= unpack_unicode_update_pos(data, 6, 2);
                    var strg = unpack[0], 
						pos = unpack[1];
                    var nrt = data.readUInt16LE(0); 
                    pos += 2;
                    var runlist = [];
					while (pos< nrt) {
						runlist.append( data.readUInt16LE(pos));
						pos+=2;
					}
                    assert(pos == data.length,'pos == len(data)');
				}
                self_put_cell(rowx, colx, comm.XL_CELL_TEXT, strg, xf_index);
				var row = self.rich_text_runlist_map[rowx]
				if(!row) row = self.rich_text_runlist_map[rowx] = [];
				row[colx] = runlist;
            }else if( rc == comm.XL_RK){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4);
                var d = unpack_RK(data, 6, 10);
                self_put_cell(rowx, colx, null, d, xf_index);
            }else if( rc == comm.XL_MULRK){
                var mulrk_row = data.readUInt16LE(0), 
					mulrk_first = data.readUInt16LE(2)
					mulrk_last = data.readUInt16LE(data.length-2);
                var pos = 4;
				for(colx = mulrk_first; colx < mulrk_last+1; colx++){
					xf_index = data.readUInt16LE(pos)
                    var d = unpack_RK(data,pos+2,pos+6);
                    pos += 6;
                    self_put_cell(mulrk_row, colx, null, d, xf_index);
				}
            }else if(rc == comm.XL_ROW){
                // Version 0.6.0a3: ROW records are just not worth using (for memory allocation).
                // Version 0.6.1: now used for formatting info.
                if ( !fmt_info) continue;
                var rowx = data.readUInt16LE(0), 
					bits1 = data.readUInt16LE(6), 
					bits2 = data.readInt32LE(12);
                if (!(0 <= rowx < self.utter_max_rows)){
                    console.log(util.format("*** NOTE: ROW record has row index %d; should have 0 <= rowx < %d -- record ignored!" 
                        ,rowx, self.utter_max_rows));
                    continue;
				}
                
                var r = rowinfo_sharing_dict.get(bits1, bits2);
                if(!r){
                    rowinfo_sharing_dict[bits1][bits2] = r = Rowinfo();
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
                    //        (16, 0x0FFF0000, 'xf_index'),
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
                    r.xf_index = (bits2 >> 16) & 0xfff;
                    r.additional_space_above = (bits2 >> 28) & 1;
                    r.additional_space_below = (bits2 >> 29) & 1;
                    if(!r.has_default_xf_index)
                        r.xf_index = -1;
				}
                self.rowinfo_map[rowx] = r;
                if ( 0 && r.xf_index > -1){
                    console.log(util.format("**ROW %d %d %d\n",self.number, rowx, r.xf_index));
				}
                if ( blah_rows){
                    console.log('ROW', rowx, bits1, bits2);
                    //r.dump(header="--- sh //%d, rowx=%d ---" % (self.number, rowx));
				}
            }else if( comm.XL_FORMULA_OPCODES[0] == rc || comm.XL_FORMULA_OPCODES[1] == rc || comm.XL_FORMULA_OPCODES[2] == rc ){ // 06, 0206, 0406
			//todo 나중에 구현
			/*
                // DEBUG = 1
                // if DEBUG: print "FORMULA: rc: 0x%04x data: %r" % (rc, data)
                if ( bv >= 50){
                    var rowx = data.readUInt16LE(0), 
						colx = data.readUInt16LE(2), 
						xf_index = data.readUInt16LE(4), 
						result_str = data.slice(6,14), 
						flags = data.readUInt16LE(14);
                    var lenlen = 2; //todo 사용여부 확인
                    var tkarr_offset = 20; //todo 사용여부 확인
                }else if( bv >= 30){
                    var rowx = data.readUInt16LE(0), 
						colx = data.readUInt16LE(2), 
						xf_index = data.readUInt16LE(4), 
						result_str = data.slice(6,14), 
						flags = data.readUInt16LE(14);
                    var lenlen = 2;
                    var tkarr_offset = 16;
                }else{ // BIFF2
                    rowx, colx, cell_attr,  result_str, flags = local_unpack('<HH3s8sB', data[0:16]);
					var rowx = data.readUInt16LE(0), 
						colx = data.readUInt16LE(2), 
						cell_attr = data.slice(4,7), 
						result_str = data.slice(7,15), 
						flags = data.readUInt8(15);
                    var xf_index =  self.fixed_BIFF2_xfindex(cell_attr, rowx, colx);
                    var lenlen = 1;
                    var tkarr_offset = 16;
				}
                if ( blah_formulas){ // testing formula dumper
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
                            rc2, data2_len, data2 = bk.get_record_parts();
                            if ( rc2 == comm.XL_STRING || rc2 == comm.XL_STRING_B2){
                                gotstring = 1;
                            }else if( rc2 == comm.XL_ARRAY){
                                row1x, rownx, col1x, colnx, array_flags, tokslen = \
                                    local_unpack("<HHBBBxxxxxH", data2[:14]);
                                if ( blah_formulas)
                                    fprintf(self.logfile, "ARRAY: %d %d %d %d %d\n",
                                        row1x, rownx, col1x, colnx, array_flags);
                                    // dump_formula(bk, data2[14:], tokslen, bv, reldelta=0, blah=1)
                            }else if( rc2 == comm.XL_SHRFMLA){
                                row1x, rownx, col1x, colnx, nfmlas, tokslen = \
                                    local_unpack("<HHBBxBH", data2[:10]);
                                if ( blah_formulas){
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
                            rc2, _unused_len, data2 = bk.get_record_parts();
                            if ( rc2 not in (comm.XL_STRING, comm.XL_STRING_B2))
                                raise XLRDError("Expected STRING record; found 0x%04x" % rc2);
						}
                        // if DEBUG: print "STRING: data=%r BIFF=%d cp=%d" % (data2, self.biffVersion, bk.encoding)
                        strg = self.string_record_contents(data2);
                        self.put_cell(rowx, colx, comm.XL_CELL_TEXT, strg, xf_index);
                        // if DEBUG: print "FORMULA strg %r" % strg
                    }else if( first_byte == 1){
                        // boolean formula result
                        value = BYTES_ORD(result_str[2]);
                        self_put_cell(rowx, colx, comm.XL_CELL_BOOLEAN, value, xf_index);
                    }else if( first_byte == 2){
                        // Error in cell
                        value = BYTES_ORD(result_str[2]);
                        self_put_cell(rowx, colx, comm.XL_CELL_ERROR, value, xf_index);
                    }else if( first_byte == 3){
                        // empty ... i.e. empty (zero-length) string, NOT an empty cell.
                        self_put_cell(rowx, colx, comm.XL_CELL_TEXT, "", xf_index);
                    }else
                        raise XLRDError("unexpected special case (0x%02x) in FORMULA" % first_byte);
                else{
                    // it is a number
                    d = local_unpack('<d', result_str)[0];
                    self_put_cell(rowx, colx, null, d, xf_index);
				}
			*/
            }else if( rc == comm.XL_BOOLERR){
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4), 
					value = data.readUInt8(6), 
					is_err = data.readUInt8(7);
                // Note OOo Calc 2.0 writes 9-byte BOOLERR records.
                // OOo docs say 8. Excel writes 8.
                var cellty = [comm.XL_CELL_BOOLEAN, comm.XL_CELL_ERROR][is_err];
                // if DEBUG: print "XL_BOOLERR", rowx, colx, xf_index, value, is_err
                self_put_cell(rowx, colx, cellty, value, xf_index);
            }else if( rc == comm.XL_COLINFO){
			//todo 나중에
			/*
                if ( !fmt_info) continue;
                var c = Colinfo();
                var first_colx, last_colx, c.width, c.xf_index, flags \
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
                    self.colinfo_map[colx] = c;
                    if ( 0)
                        fprintf(self.logfile,
                            "**COL %d %d %d\n",
                            self.number, colx, c.xf_index);
                if ( blah)
                    fprintf(
                        self.logfile,
                        "COLINFO sheet //%d cols %d-%d: wid=%d xf_index=%d flags=0x%04x\n",
                        self.number, first_colx, last_colx, c.width, c.xf_index, flags,
                        );
                    c.dump(self.logfile, header='===');
			*/
            }else if( rc == comm.XL_DEFCOLWIDTH){
                self.defcolwidth = data.readUInt16LE(0);
                if (0) console.log('DEFCOLWIDTH', self.defcolwidth);
            }else if( rc == comm.XL_STANDARDWIDTH){
                if ( data_len != 2) console.log('*** ERROR *** STANDARDWIDTH', data_len, repr(data));
                self.standardwidth = data.readUInt16LE(0);
                if ( 0) console.log('STANDARDWIDTH', self.standardwidth);
            }else if( rc == comm.XL_GCW){
			//todo 나중에
			/*
                if ( not fmt_info) continue; // useless w/o COLINFO
                assert data_len == 34;
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
                if( !fmt_info) continue;
                var rowx = data.readUInt16LE(0), 
					colx = data.readUInt16LE(2), 
					xf_index = data.readUInt16LE(4);
                // if 0: print >> self.logfile, "BLANK", rowx, colx, xf_index
                self_put_cell(rowx, colx, comm.XL_CELL_BLANK, '', xf_index);
            }else if( rc == comm.XL_MULBLANK){ // 00BE
                if ( ! fmt_info) continue;
                var nitems = data_len >> 1;
                var result = [];
				for(i=0; i< nitems;i+=2) result.push(data.readUInt16LE(i));
                var rowx = result[0], mul_first = result[1];
                mul_last = result[result.length-1];
                // print >> self.logfile, "MULBLANK", rowx, mul_first, mul_last, data_len, nitems, mul_last + 4 - mul_first
                assert( nitems == mul_last + 4 - mul_first,'nitems == mul_last + 4 - mul_first');
                var pos = 2;
                for(colx = mul_first; colx < mul_last+1; colx++){
                    self_put_cell(rowx, colx, comm.XL_CELL_BLANK, '', result[pos]);
                    pos += 1;
				}
            }else if( rc == comm.XL_DIMENSION || rc == comm.XL_DIMENSION2){
                if ( data_len == 0)
                    // Four zero bytes after some other record. See github issue 64.
                    continue;
                // if data_len == 10:
                // Was crashing on BIFF 4.0 file w/o the two trailing unused bytes.
                // Reported by Ralph Heimburger.
				var dim_tuple =[];
                if ( bv < 80){
					dim_tuple =[data.readUInt16LE(2), data.readUInt16LE(6)];
                }else{
					dim_tuple =[data.readInt32LE(4), data.readUInt16LE(10)];
				}
                self.nrows = self.ncols = 0;
                self._dimnrows = self._dimncols = dim_tuple;
                if ( (bv == 21 || bv ==  30 || bv ==  40) && book.xf_list && ! book._xf_epilogue_done)
                    book.xf_epilogue();
                if ( blah)
                    console.log(util.format("sheet %d(%r) DIMENSIONS: ncols=%d nrows=%d",
                        self.number, self.name, self._dimncols, self._dimnrows
                        ));
            }else if( rc == comm.XL_HLINK){
                self.handle_hlink(data);
            }else if( rc == comm.XL_QUICKTIP){
                self.handle_quicktip(data);
            }else if( rc == comm.XL_EOF){
                DEBUG = 0;
                if (DEBUG) print("SHEET.READ: EOF", file=self.logfile);
                eof_found = 1;
                break;
            }else if( rc == comm.XL_OBJ){
                // handle SHEET-level objects; note there's a separate Book.handle_obj
				/*
                var saved_obj = self.handle_obj(data);
                if (saved_obj) saved_obj_id = saved_obj.id;
                else saved_obj_id = null;
				*/
            }else if( rc == comm.XL_MSO_DRAWING){
                //self.handle_msodrawingetc(rc, data_len, data);
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
            }else if( comm.bofcodes.indexOf(rc) >= 0){ ////////// EMBEDDED BOF //////////
                var version = data.readUInt16LE(0), 
					boftype = data.readUInt16LE(2);
                if ( boftype != 0x20) // embedded chart
                    console.log(util.format("*** Unexpected embedded BOF (0x%04x) at offset %d: version=0x%04x type=0x%04x" 
                        ,rc, bk._position - data_len - 4, version, boftype));
                while (1){
					var record = bk.get_record_parts();
                    var code = record[0];
					data_len = record[1];
					data  = record[2];
                    if(code == comm.XL_EOF)
                        break;
				}
                if ( DEBUG) console.log("---> found EOF");
            }else if( rc == comm.XL_COUNTRY){
                bk.handle_country(data);
            }else if( rc == comm.XL_LABELRANGES){
                var pos = 0;
                pos = unpack_cell_range_address_list_update_pos(
                        self.row_label_ranges, data, pos, bv, 8
                        );
                pos = unpack_cell_range_address_list_update_pos(
                        self.col_label_ranges, data, pos, bv, 8
                        );
                assert (pos == data_len,'pos == data_len');
            }else if( rc == comm.XL_ARRAY){
                var row1x = data.readUInt16LE(0), 
					rownx = data.readUInt16LE(2), 
					col1x = data.readUInt8(3), 
					colnx = data.readUInt8(4), 
					array_flags = data.readUInt8(5), 
					tokslen =  data.readUInt16LE(12);
                    
                if ( blah_formulas)
                    console.log("ARRAY:", row1x, rownx, col1x, colnx, array_flags);
                    // dump_formula(bk, data[14:], tokslen, bv, reldelta=0, blah=1)
            }else if( rc == comm.XL_SHRFMLA){
			//todo 나중에
			/*
                row1x, rownx, col1x, colnx, nfmlas, tokslen = \
                    local_unpack("<HHBBxBH", data[:10]);
                if ( blah_formulas){
                    print("SHRFMLA (main):", row1x, rownx, col1x, colnx, nfmlas, file=self.logfile);
                    decompile_formula(bk, data[10:], tokslen, FMLA_TYPE_SHARED,
                        blah=1, browx=rowx, bcolx=colx, r1c1=r1c1);
				}
			*/
            }else if( rc == comm.XL_CONDFMT){
                if ( !fmt_info) continue;
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
                        self.number, self.name, num_CFs, needs_recalc,
                        rangename2d(browx1, browx2+1, bcolx1, bcolx2+1),
                        );
				}
                olist = []; // updated by the function
                pos = unpack_cell_range_address_list_update_pos(
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
                if ( ! fmt_info) continue;
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
                // hex_char_dump(data, 0, data_len, fout=self.logfile)
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
                assert pos == data_len;
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
                if ( data_len == 4){
                    bits, self.default_row_height = unpack("<HH", data[:4]);
                }else if( data_len == 2){
                    self.default_row_height, = unpack("<H", data);
                    bits = 0;
                    fprintf(self.logfile,
                        "*** WARNING: DEFAULTROWHEIGHT record len is 2, " \
                        "should be 4; assuming BIFF2 format\n");
                }else{
                    bits = 0;
                    fprintf(self.logfile,
                        "*** WARNING: DEFAULTROWHEIGHT record len is %d, " \
                        "should be 4; ignoring this record\n",
                        data_len);
                self.default_row_height_mismatch = bits & 1;
                self.default_row_hidden = (bits >> 1) & 1;
                self.default_additional_space_above = (bits >> 2) & 1;
                self.default_additional_space_below = (bits >> 3) & 1;
				*/
            }else if( rc == comm.XL_MERGEDCELLS){
                if ( !fmt_info) continue;
                pos = unpack_cell_range_address_list_update_pos(
                    self.merged_cells, data, 0, bv, addr_size=8);
                if ( blah)
                    console.log(util.format(
                        "MERGEDCELLS: %d ranges\n", Math.floor((pos - 2) / 8)));
                assert(pos == data_len, util.format("MERGEDCELLS: pos=%d data_len=%d",pos, data_len));
            }else if( rc == comm.XL_WINDOW2){
                if ( bv >= 80 && data_len >= 14){
                    var options = data.readUInt16LE(0);
                    self.first_visible_rowx = data.readUInt16LE(2);
					self.first_visible_colx = data.readUInt16LE(4);
                    self.gridline_colour_index = data.readUInt16LE(6);
                    self.cached_page_break_preview_mag_factor = data.readUInt16LE(10);
                    self.cached_normal_view_mag_factor = data.readUInt16LE(12);
                }else{
                    assert( bv >= 30, 'bv >= 30'); // BIFF3-7
                    var options = data.readUInt16LE(0);
                    self.first_visible_rowx = data.readUInt16LE(2);
					self.first_visible_colx = data.readUInt16LE(4);
                    self.gridline_colour_rgb = [data.readUInt8(6), data.readUInt8(7),data.readUInt8(8)];
                    self.gridline_colour_index = nearest_colour_index(
                        book.colour_map, self.gridline_colour_rgb, 0);
                    self.cached_page_break_preview_mag_factor = 0; // default (60%)
                    self.cached_normal_view_mag_factor = 0; // default (100%)
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
                            % (self.number, num, den)
                            ), file=self.logfile);
					}
                    result = 100;
				}
                self.scl_mag_factor = result;
				*/
            }else if( rc == comm.XL_PANE){
				//todo 나중에
				/*
                (
                self.vert_split_pos,
                self.horz_split_pos,
                self.horz_split_first_visible,
                self.vert_split_first_visible,
                self.split_active_pane,
                ) = unpack("<HHHHB", data[:9]);
                self.has_pane_record = 1;
				*/
            }else if( rc == comm.XL_HORIZONTALPAGEBREAKS){
                if ( !fmt_info) continue;
				//나중에
				/*
                num_breaks, = local_unpack("<H", data[:2]);
                assert num_breaks * (2 + 4 * (bv >= 80)) + 2 == data_len;
                pos = 2;
                if ( bv < 80)
                    while pos < data_len)
                        self.horizontal_page_breaks.append((local_unpack("<H", data[pos:pos+2])[0], 0, 255));
                        pos += 2;
                else{
                    while pos < data_len){
                        self.horizontal_page_breaks.append(local_unpack("<HHH", data[pos:pos+6]));
                        pos += 6;
					}
				}
				*/
            }else if( rc == comm.XL_VERTICALPAGEBREAKS){
                if ( !fmt_info) continue;
				//todo 나중에
				/*
                num_breaks, = local_unpack("<H", data[:2]);
                assert num_breaks * (2 + 4 * (bv >= 80)) + 2 == data_len;
                pos = 2;
                if ( bv < 80)
                    while pos < data_len){
                        self.vertical_page_breaks.append((local_unpack("<H", data[pos:pos+2])[0], 0, 65535));
                        pos += 2;
					}
                else
                    while pos < data_len){
                        self.vertical_page_breaks.append(local_unpack("<HHH", data[pos:pos+6]));
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
                    if ( not book._xf_epilogue_done) book.xf_epilogue();
                    bk.handle_style(data);
                }else if( rc == comm.XL_PALETTE){
                    bk.handle_palette(data);
                }else if( rc == comm.XL_BUILTINFMTCOUNT){
                    bk.handle_builtinfmtcount(data);
                }else if( rc == comm.XL_XF4 || rc == comm.XL_XF3 || rc == comm.XL_XF2){ //////// N.B. not XL_XF
                    bk.handle_xf(data);
                }else if( rc == comm.XL_DATEMODE){
                    bk.handle_datemode(data);
                }else if( rc == comm.XL_CODEPAGE){
                    bk.handle_codepage(data);
                }else if( rc == comm.XL_FILEPASS){
                    bk.handle_filepass(data);
                }else if( rc == comm.XL_WRITEACCESS){
                    bk.handle_writeaccess(data);
                }else if( rc == comm.XL_IXFE){
                    self._ixfe = local_unpack('<H', data)[0];
                }else if( rc == comm.XL_NUMBER_B2){
                    rowx, colx, cell_attr, d = local_unpack('<HH3sd', data);
                    self_put_cell(rowx, colx, null, d, self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_INTEGER){
                    rowx, colx, cell_attr, d = local_unpack('<HH3sH', data);
                    self_put_cell(rowx, colx, null, float(d), self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_LABEL_B2){
                    rowx, colx, cell_attr = local_unpack('<HH3s', data[0:7]);
                    strg = unpack_string(data, 7, bk.encoding || bk.derive_encoding(), lenlen=1);
                    self_put_cell(rowx, colx, comm.XL_CELL_TEXT, strg, self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_BOOLERR_B2){
                    rowx, colx, cell_attr, value, is_err = local_unpack('<HH3sBB', data);
                    cellty = (comm.XL_CELL_BOOLEAN, comm.XL_CELL_ERROR)[is_err];
                    // if DEBUG: print "XL_BOOLERR_B2", rowx, colx, cell_attr, value, is_err
                    self_put_cell(rowx, colx, cellty, value, self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_BLANK_B2){
                    if ( not fmt_info) continue;
                    rowx, colx, cell_attr = local_unpack('<HH3s', data[:7]);
                    self_put_cell(rowx, colx, comm.XL_CELL_BLANK, '', self.fixed_BIFF2_xfindex(cell_attr, rowx, colx));
                }else if( rc == comm.XL_EFONT){
                    bk.handle_efont(data);
                }else if( rc == comm.XL_ROW_B2){
                    if ( not fmt_info) continue;
                    rowx, bits1, bits2 = local_unpack('<H4xH2xB', data[0:11]);
                    if ( not(0 <= rowx < self.utter_max_rows)){
                        print("*** NOTE: ROW_B2 record has row index %d; " \
                            "should have 0 <= rowx < %d -- record ignored!" \
                            % (rowx, self.utter_max_rows), file=self.logfile);
                        continue;
					}
                    if ( not (bits2 & 1)){  // has_default_xf_index is false
                        xf_index = -1;
                    }else if( data_len == 18){
                        // Seems the XF index in the cell_attr is dodgy
                         xfx = local_unpack('<H', data[16:18])[0];
                         xf_index = self.fixed_BIFF2_xfindex(cell_attr=null, rowx=rowx, colx=-1, true_xfx=xfx);
                    }else{
                        cell_attr = data[13:16];
                        xf_index = self.fixed_BIFF2_xfindex(cell_attr, rowx, colx=-1);
					}
                    key = (bits1, bits2, xf_index);
                    r = rowinfo_sharing_dict.get(key);
                    if ( r is null){
                        rowinfo_sharing_dict[key] = r = Rowinfo();
                        r.height = bits1 & 0x7fff;
                        r.has_default_height = (bits1 >> 15) & 1;
                        r.has_default_xf_index = bits2 & 1;
                        r.xf_index = xf_index;
                        // r.outline_level = 0             // set in __init__
                        // r.outline_group_starts_ends = 0 // set in __init__
                        // r.hidden = 0                    // set in __init__
                        // r.height_mismatch = 0           // set in __init__
                        // r.additional_space_above = 0    // set in __init__
                        // r.additional_space_below = 0    // set in __init__
					}
                    self.rowinfo_map[rowx] = r;
                    if ( 0 and r.xf_index > -1){
                        fprintf(self.logfile,
                            "**ROW %d %d %d\n",
                            self.number, rowx, r.xf_index);
					}
                    if ( blah_rows){
                        print('ROW_B2', rowx, bits1, has_defaults, file=self.logfile);
                        r.dump(self.logfile,
                            header="--- sh //%d, rowx=%d ---" % (self.number, rowx));
					}
                }else if( rc == comm.XL_COLWIDTH){ // BIFF2 only
                    if ( not fmt_info) continue;
                    first_colx, last_colx, width\
                        = local_unpack("<BBH", data[:4]);
                    if ( not(first_colx <= last_colx)){
                        print("*** NOTE: COLWIDTH record has first col index %d, last %d; " \
                            "should have first <= last -- record ignored!" \
                            % (first_colx, last_colx), file=self.logfile);
                        continue;
					}
                    for colx in xrange(first_colx, last_colx+1){
                        if ( colx in self.colinfo_map){
                            c = self.colinfo_map[colx];
                        }else{
                            c = Colinfo();
                            self.colinfo_map[colx] = c;
						}
                        c.width = width;
					}
                    if ( blah)
                        fprintf(
                            self.logfile,
                            "COLWIDTH sheet //%d cols %d-%d: wid=%d\n",
                            self.number, first_colx, last_colx, width
                            );
                }else if( rc == comm.XL_COLUMNDEFAULT) // BIFF2 only
                    if ( not fmt_info) continue;
                    first_colx, last_colx = local_unpack("<HH", data[:4]);
                    //////// Warning OOo docs wrong; first_colx <= colx < last_colx
                    if ( blah)
                        fprintf(
                            self.logfile,
                            "COLUMNDEFAULT sheet //%d cols in range(%d, %d)\n",
                            self.number, first_colx, last_colx
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
                        xf_index = self.fixed_BIFF2_xfindex(cell_attr, rowx=-1, colx=colx);
                        if ( colx in self.colinfo_map){
                            c = self.colinfo_map[colx];
                        }else{
                            c = Colinfo();
                            self.colinfo_map[colx] = c;
						}
                        c.xf_index = xf_index;
                }else if( rc == comm.XL_WINDOW2_B2){ // BIFF 2 only
                    attr_names = ("show_formulas", "show_grid_lines", "show_sheet_headers",
                        "panes_are_frozen", "show_zero_values");
                    for attr, char in zip(attr_names, data[0:5]))
                        setattr(self, attr, int(char != b'\0'));
                    (self.first_visible_rowx, self.first_visible_colx,
                    self.automatic_grid_line_colour,
                    ) = unpack("<HHB", data[5:10]);
                    self.gridline_colour_rgb = unpack("<BBB", data[10:13]);
                    self.gridline_colour_index = nearest_colour_index(
                        book.colour_map, self.gridline_colour_rgb, debug=0);
                    self.cached_page_break_preview_mag_factor = 0 ;// default (60%)
                    self.cached_normal_view_mag_factor = 0; // default (100%)
				}
				*/
            }else
                // if DEBUG: print "SHEET.READ: Unhandled record type %02x %d bytes %r" % (rc, data_len, data)
                ;
		}
        if (!eof_found)
            throw new Error(util.format("Sheet %d (%r) missing EOF record",self.number, self.name));
        tidy_dimensions();
        update_cooked_mag_factors();
        bk._position = oldpos;
        return 1;
    }
	
	
    function update_cooked_mag_factors(){
        // Cached values are used ONLY for the non-active view mode.
        // When the user switches to the non-active view mode,
        // if the cached value for that mode is not valid,
        // Excel pops up a window which says:
        // "The number must be between 10 and 400. Try again by entering a number in this range."
        // When the user hits OK, it drops into the non-active view mode
        // but uses the magn from the active mode.
        // NOTE: definition of "valid" depends on mode ... see below
        var blah = DEBUG || self.verbosity > 0;
        if (self.show_in_page_break_preview){
            if (!self.scl_mag_factor) // no SCL record
                self.cooked_page_break_preview_mag_factor = 100; // Yes, 100, not 60, NOT a typo
            else
                self.cooked_page_break_preview_mag_factor = self.scl_mag_factor;
            var zoom = self.cached_normal_view_mag_factor
            if (!(10 <= zoom && zoom <=400))
                if (blah)
                    console.log(util.format(
                        "WARNING *** WINDOW2 rcd sheet %d: Bad cached_normal_view_mag_factor: %d",
                        self.number, self.cached_normal_view_mag_factor
                        ));
                zoom = self.cooked_page_break_preview_mag_factor;
            self.cooked_normal_view_mag_factor = zoom;
        }else{
            // normal view mode
            if(!self.scl_mag_factor) // no SCL record
                self.cooked_normal_view_mag_factor = 100;
            else
                self.cooked_normal_view_mag_factor = self.scl_mag_factor;
            var zoom = self.cached_page_break_preview_mag_factor;
            if (zoom == 0)
                // VALID, defaults to 60
                zoom = 60;
            else if (!(10 <= zoom && zoom <= 400)){
                if (blah)
                    console.log(util.format(
                        "WARNING *** WINDOW2 rcd sheet %r: Bad cached_page_break_preview_mag_factor: %r"
                        ,self.number, self.cached_page_break_preview_mag_factor)
                        );
                zoom = self.cooked_normal_view_mag_factor;
			}
            self.cooked_page_break_preview_mag_factor = zoom;
		}
	}
}


function unpack_RK(data, start, end){
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
function toDate(xldate, datemode){
    if(datemode != 0 &&  datemode != 1 )
        throw new Error('bad date mode : ' + datemode);
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
		minutes = Math.floor(seconds/60);
        minute = minutes % 60; 
		hour = Math.floor(minutes/60);
	}
    if(xldays >= _XLDAYS_TOO_LARGE[datemode])
        throw new Error('date is too large : ' + xldate); 

    if (xldays == 0)
        return new Date(0, 0, 0, hour, minute, second);

    if (xldays < 61 && datemode == 0)
        throw new Error('date is ambiguous : ' +xldate);

    var jdn = xldays + _JDN_delta[datemode];
	
    var yreg = (Math.floor((Math.floor(jdn * 4 + 274277) / 146097) * 3 / 4) + jdn + 1363) * 4 + 3;
    var mp = Math.floor((yreg % 1461) / 4) * 535 + 333
    var d = Math.floor((mp % 16384) / 535) + 1
    
    mp >>= 14;
    if(mp >= 10)
        return new Date (Math.floor(yreg / 1461) - 4715, mp - 9, d, hour, minute, second);
    else
        return new Date (Math.floor(yreg / 1461) - 4716, mp + 3, d, hour, minute, second);
}


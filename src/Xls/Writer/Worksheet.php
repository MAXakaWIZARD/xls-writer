<?php
/*
*  Module written/ported by Xavier Noguer <xnoguer@rezebra.com>
*
*  The majority of this is _NOT_ my code.  I simply ported it from the
*  PERL Spreadsheet::WriteExcel module.
*
*  The author of the Spreadsheet::WriteExcel module is John McNamara
*  <jmcnamara@cpan.org>
*
*  I _DO_ maintain this code, and John McNamara has nothing to do with the
*  porting of this code to PHP.  Any questions directly related to this
*  class library should be directed to me.
*
*  License Information:
*
*    Spreadsheet_Excel_Writer:  A library for generating Excel Spreadsheets
*    Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
*
*    This library is free software; you can redistribute it and/or
*    modify it under the terms of the GNU Lesser General Public
*    License as published by the Free Software Foundation; either
*    version 2.1 of the License, or (at your option) any later version.
*
*    This library is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*    Lesser General Public License for more details.
*
*    You should have received a copy of the GNU Lesser General Public
*    License along with this library; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

namespace Xls\Writer;

/**
 * Class for generating Excel Spreadsheets
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 * @category FileFormats
 * @package  Spreadsheet_Excel_Writer
 */

class Worksheet extends BIFFwriter
{
    /**
     * Name of the Worksheet
     * @var string
     */
    public $name;

    /**
     * Index for the Worksheet
     * @var integer
     */
    public $index;

    /**
     * Reference to the (default) Format object for URLs
     * @var object Format
     */
    public $url_format;

    /**
     * Reference to the parser used for parsing formulas
     * @var object Format
     */
    public $parser;

    /**
     * Filehandle to the temporary file for storing data
     * @var resource
     */
    public $filehandle;

    /**
     * Boolean indicating if we are using a temporary file for storing data
     * @var bool
     */
    public $using_tmpfile;

    /**
     * Maximum number of rows for an Excel spreadsheet (BIFF5)
     * @var integer
     */
    public $xls_rowmax;

    /**
     * Maximum number of columns for an Excel spreadsheet (BIFF5)
     * @var integer
     */
    public $xls_colmax;

    /**
     * Maximum number of characters for a string (LABEL record in BIFF5)
     * @var integer
     */
    public $xls_strmax;

    /**
     * First row for the DIMENSIONS record
     * @var integer
     * @see _storeDimensions()
     */
    public $dim_rowmin;

    /**
     * Last row for the DIMENSIONS record
     * @var integer
     * @see _storeDimensions()
     */
    public $dim_rowmax;

    /**
     * First column for the DIMENSIONS record
     * @var integer
     * @see _storeDimensions()
     */
    public $dim_colmin;

    /**
     * Last column for the DIMENSIONS record
     * @var integer
     * @see _storeDimensions()
     */
    public $dim_colmax;

    /**
     * Array containing format information for columns
     * @var array
     */
    public $colinfo;

    /**
     * Array containing the selected area for the worksheet
     * @var array
     */
    public $selection;

    /**
     * Array containing the panes for the worksheet
     * @var array
     */
    public $panes;

    /**
     * The active pane for the worksheet
     * @var integer
     */
    public $active_pane;

    /**
     * Bit specifying if panes are frozen
     * @var integer
     */
    public $frozen;

    /**
     * Bit specifying if the worksheet is selected
     * @var integer
     */
    public $selected;

    /**
     * The paper size (for printing) (DOCUMENT!!!)
     * @var integer
     */
    public $paper_size;

    /**
     * Bit specifying paper orientation (for printing). 0 => landscape, 1 => portrait
     * @var integer
     */
    public $orientation;

    /**
     * The page header caption
     * @var string
     */
    public $header;

    /**
     * The page footer caption
     * @var string
     */
    public $footer;

    /**
     * The horizontal centering value for the page
     * @var integer
     */
    public $hcenter;

    /**
     * The vertical centering value for the page
     * @var integer
     */
    public $vcenter;

    /**
     * The margin for the header
     * @var float
     */
    public $margin_head;

    /**
     * The margin for the footer
     * @var float
     */
    public $margin_foot;

    /**
     * The left margin for the worksheet in inches
     * @var float
     */
    public $margin_left;

    /**
     * The right margin for the worksheet in inches
     * @var float
     */
    public $margin_right;

    /**
     * The top margin for the worksheet in inches
     * @var float
     */
    public $margin_top;

    /**
     * The bottom margin for the worksheet in inches
     * @var float
     */
    public $margin_bottom;

    /**
     * First row to reapeat on each printed page
     * @var integer
     */
    public $title_rowmin;

    /**
     * Last row to reapeat on each printed page
     * @var integer
     */
    var $title_rowmax;

    /**
     * First column to reapeat on each printed page
     * @var integer
     */
    public $title_colmin;

    /**
     * First row of the area to print
     * @var integer
     */
    public $print_rowmin;

    /**
     * Last row to of the area to print
     * @var integer
     */
    public $print_rowmax;

    /**
     * First column of the area to print
     * @var integer
     */
    public $print_colmin;

    /**
     * Last column of the area to print
     * @var integer
     */
    public $print_colmax;

    /**
     * Whether to display RightToLeft.
     * @var integer
     */
    public $Arabic;

    /**
     * Whether to use outline.
     * @var integer
     */
    public $outline_on;

    /**
     * Auto outline styles.
     * @var bool
     */
    public $outline_style;

    /**
     * Whether to have outline summary below.
     * @var bool
     */
    public $outline_below;

    /**
     * Whether to have outline summary at the right.
     * @var bool
     */
    public $outline_right;

    /**
     * Outline row level.
     * @var integer
     */
    public $outline_row_level;

    /**
     * Whether to fit to page when printing or not.
     * @var bool
     */
    public $fit_page;

    /**
     * Number of pages to fit wide
     * @var integer
     */
    public $fit_width;

    /**
     * Number of pages to fit high
     * @var integer
     */
    public $fit_height;

    /**
     * Reference to the total number of strings in the workbook
     * @var integer
     */
    public $str_total;

    /**
     * Reference to the number of unique strings in the workbook
     * @var integer
     */
    public $str_unique;

    /**
     * Reference to the array containing all the unique strings in the workbook
     * @var array
     */
    public $str_table;

    /**
     * Number of merged cell ranges in actual record
     * @var int $merged_cells_counter
     */
    public $merged_cells_counter = 0;

    /**
     * Number of actual mergedcells record
     * @var int $merged_cells_record
     */
    public $merged_cells_record = 0;

    /**
     * Merged cell ranges
     * @var array
     */
    public $merged_ranges;

    /**
     * Charset encoding currently used when calling writeString()
     * @var string
     */
    public $input_encoding;

    /**
     * Constructor
     *
     * @param string $name         The name of the new worksheet
     * @param integer $index        The index of the new worksheet
     * @param mixed &$activesheet The current activesheet of the workbook we belong to
     * @param mixed &$firstsheet  The first worksheet in the workbook we belong to
     * @param mixed &$url_format  The default format for hyperlinks
     * @param mixed &$parser      The formula parser created for the Workbook
     * @param string $tmp_dir      The path to the directory for temporary files
     */
    public function __construct(
        $biffVersion,
        $name,
        $index,
        &$activesheet,
        &$firstsheet,
        &$str_total,
        &$str_unique,
        &$str_table,
        &$url_format,
        &$parser,
        $tmp_dir
    ) {
        parent::__construct();
        $this->BIFF_version = $biffVersion;
        $rowmax = 65536; // 16384 in Excel 5
        $colmax = 256;

        $this->name = $name;
        $this->index = $index;
        $this->activesheet = & $activesheet;
        $this->firstsheet = & $firstsheet;
        $this->str_total = & $str_total;
        $this->str_unique = & $str_unique;
        $this->str_table = & $str_table;
        $this->url_format = & $url_format;
        $this->parser = & $parser;

        //$this->ext_sheets      = array();
        $this->filehandle = '';
        $this->using_tmpfile = true;
        //$this->fileclosed      = 0;
        //$this->offset          = 0;
        $this->xls_rowmax = $rowmax;
        $this->xls_colmax = $colmax;
        $this->xls_strmax = 255;
        $this->dim_rowmin = $rowmax + 1;
        $this->dim_rowmax = 0;
        $this->dim_colmin = $colmax + 1;
        $this->dim_colmax = 0;
        $this->colinfo = array();
        $this->selection = array(0, 0, 0, 0);
        $this->panes = array();
        $this->active_pane = 3;
        $this->frozen = 0;
        $this->selected = 0;

        $this->paper_size = 0x0;
        $this->orientation = 0x1;
        $this->header = '';
        $this->footer = '';
        $this->hcenter = 0;
        $this->vcenter = 0;
        $this->margin_head = 0.50;
        $this->margin_foot = 0.50;
        $this->margin_left = 0.75;
        $this->margin_right = 0.75;
        $this->margin_top = 1.00;
        $this->margin_bottom = 1.00;

        $this->title_rowmin = null;
        $this->title_rowmax = null;
        $this->title_colmin = null;
        $this->title_colmax = null;
        $this->print_rowmin = null;
        $this->print_rowmax = null;
        $this->print_colmin = null;
        $this->print_colmax = null;

        $this->print_gridlines = 1;
        $this->screen_gridlines = 1;
        $this->print_headers = 0;

        $this->fit_page = 0;
        $this->fit_width = 0;
        $this->fit_height = 0;

        $this->hbreaks = array();
        $this->vbreaks = array();

        $this->protect = 0;
        $this->password = null;

        $this->col_sizes = array();
        $this->row_sizes = array();

        $this->zoom = 100;
        $this->print_scale = 100;

        $this->outline_row_level = 0;
        $this->outline_style = 0;
        $this->outline_below = 1;
        $this->outline_right = 1;
        $this->outline_on = 1;
        $this->Arabic = 0;

        $this->merged_ranges = array();

        $this->input_encoding = '';

        $this->dv = array();

        $this->tmpDir = $tmp_dir;
        $this->tmpFile = '';

        $this->initialize();
    }

    /**
     * Open a tmp file to store the majority of the Worksheet data. If this fails,
     * for example due to write permissions, store the data in memory. This can be
     * slow for large files.
     *
     * @access private
     */
    public function initialize()
    {
        if ($this->using_tmpfile == false) {
            return;
        }

        if ($this->tmpDir === '' && ini_get('open_basedir') === true) {
            // open_basedir restriction in effect - store data in memory
            // ToDo: Let the error actually have an effect somewhere
            $this->using_tmpfile = false;
            throw new \Exception('Temp file could not be opened since open_basedir restriction in effect - please use setTmpDir() - using memory storage instead');
        }

        // Open tmp file for storing Worksheet data
        if ($this->tmpDir === '') {
            $fh = tmpfile();
        } else {
            // For people with open base dir restriction
            $this->tmpFile = tempnam($this->tmpDir, "Spreadsheet_Excel_Writer");
            $fh = @fopen($this->tmpFile, "w+b");
        }

        if ($fh === false) {
            // If tmpfile() fails store data in memory
            $this->using_tmpfile = false;
        } else {
            // Store filehandle
            $this->filehandle = $fh;
        }
    }

    /**
     * Add data to the beginning of the workbook (note the reverse order)
     * and to the end of the workbook.
     *
     * @access public
     * @see Workbook::storeWorkbook()
     * @param array $sheetnames The array of sheetnames from the Workbook this
     *                          worksheet belongs to
     */
    public function close($sheetnames)
    {
        $num_sheets = count($sheetnames);

        /***********************************************
         * Prepend in reverse order!!
         */

        // Prepend the sheet dimensions
        $this->storeDimensions();

        // Prepend the sheet password
        $this->storePassword();

        // Prepend the sheet protection
        $this->storeProtect();

        // Prepend the page setup
        $this->storeSetup();

        /* FIXME: margins are actually appended */
        // Prepend the bottom margin
        $this->storeMarginBottom();

        // Prepend the top margin
        $this->storeMarginTop();

        // Prepend the right margin
        $this->storeMarginRight();

        // Prepend the left margin
        $this->storeMarginLeft();

        // Prepend the page vertical centering
        $this->storeVcenter();

        // Prepend the page horizontal centering
        $this->storeHcenter();

        // Prepend the page footer
        $this->storeFooter();

        // Prepend the page header
        $this->storeHeader();

        // Prepend the vertical page breaks
        $this->storeVbreak();

        // Prepend the horizontal page breaks
        $this->storeHbreak();

        // Prepend WSBOOL
        $this->storeWsbool();

        // Prepend GRIDSET
        $this->storeGridset();

        //  Prepend GUTS
        if ($this->BIFF_version == 0x0500) {
            $this->storeGuts();
        }

        // Prepend PRINTGRIDLINES
        $this->storePrintGridlines();

        // Prepend PRINTHEADERS
        $this->storePrintHeaders();

        // Prepend EXTERNSHEET references
        if ($this->BIFF_version == 0x0500) {
            for ($i = $num_sheets; $i > 0; $i--) {
                $sheetname = $sheetnames[$i - 1];
                $this->storeExternsheet($sheetname);
            }
        }

        // Prepend the EXTERNCOUNT of external references.
        if ($this->BIFF_version == 0x0500) {
            $this->storeExterncount($num_sheets);
        }

        // Prepend the COLINFO records if they exist
        if (!empty($this->colinfo)) {
            $colcount = count($this->colinfo);
            for ($i = 0; $i < $colcount; $i++) {
                $this->storeColinfo($this->colinfo[$i]);
            }
            $this->storeDefcol();
        }

        // Prepend the BOF record
        $this->storeBof(0x0010);

        /*
        * End of prepend. Read upwards from here.
        ***********************************************/

        // Append
        $this->storeWindow2();
        $this->storeZoom();
        if (!empty($this->panes)) {
            $this->storePanes($this->panes);
        }
        $this->storeSelection($this->selection);
        $this->storeMergedCells();
        /* TODO: add data validity */
        /*if ($this->BIFF_version == 0x0600) {
            $this->storeDataValidity();
        }*/
        $this->storeEof();

        if ($this->tmpFile != '') {
            if ($this->filehandle) {
                //fclose($this->filehandle);
                //$this->filehandle = '';
            }
            @unlink($this->tmpFile);
            $this->tmpFile = '';
            $this->using_tmpfile = true;
        }
    }

    /**
     * Retrieve the worksheet name.
     * This is usefull when creating worksheets without a name.
     *
     * @access public
     * @return string The worksheet's name
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Retrieves data from memory in one chunk, or from disk in $buffer
     * sized chunks.
     *
     * @return string The data
     */
    public function getData()
    {
        $buffer = 4096;

        // Return data stored in memory
        if (isset($this->data)) {
            $tmp = $this->data;
            unset($this->data);
            $fh = $this->filehandle;
            if ($this->using_tmpfile) {
                fseek($fh, 0);
            }
            return $tmp;
        }
        // Return data stored on disk
        if ($this->using_tmpfile) {
            if ($tmp = fread($this->filehandle, $buffer)) {
                return $tmp;
            }
        }

        // No data to return
        return '';
    }

    /**
     * Sets a merged cell range
     *
     * @access public
     * @param integer $first_row First row of the area to merge
     * @param integer $first_col First column of the area to merge
     * @param integer $last_row  Last row of the area to merge
     * @param integer $last_col  Last column of the area to merge
     */
    public function setMerge($first_row, $first_col, $last_row, $last_col)
    {
        if (($last_row < $first_row) || ($last_col < $first_col)) {
            return;
        }

        $max_record_ranges = floor(($this->limit - 6) / 8);
        if ($this->merged_cells_counter >= $max_record_ranges) {
            $this->merged_cells_record++;
            $this->merged_cells_counter = 0;
        }

        // don't check rowmin, rowmax, etc... because we don't know when this
        // is going to be called
        $this->merged_ranges[$this->merged_cells_record][] = array($first_row, $first_col, $last_row, $last_col);
        $this->merged_cells_counter++;
    }

    /**
     * Set this worksheet as a selected worksheet,
     * i.e. the worksheet has its tab highlighted.
     *
     * @access public
     */
    public function select()
    {
        $this->selected = 1;
    }

    /**
     * Set this worksheet as the active worksheet,
     * i.e. the worksheet that is displayed when the workbook is opened.
     * Also set it as selected.
     *
     * @access public
     */
    public function activate()
    {
        $this->selected = 1;
        $this->activesheet = $this->index;
    }

    /**
     * Set this worksheet as the first visible sheet.
     * This is necessary when there are a large number of worksheets and the
     * activated worksheet is not visible on the screen.
     *
     * @access public
     */
    public function setFirstSheet()
    {
        $this->firstsheet = $this->index;
    }

    /**
     * Set the worksheet protection flag
     * to prevent accidental modification and to
     * hide formulas if the locked and hidden format properties have been set.
     *
     * @access public
     * @param string $password The password to use for protecting the sheet.
     */
    public function protect($password)
    {
        $this->protect = 1;
        $this->password = $this->encodePassword($password);
    }

    /**
     * Set the width of a single column or a range of columns.
     *
     * @access public
     * @param integer $firstcol first column on the range
     * @param integer $lastcol  last column on the range
     * @param integer $width    width to set
     * @param mixed $format   The optional XF format to apply to the columns
     * @param integer $hidden   The optional hidden atribute
     * @param integer $level    The optional outline level
     */
    public function setColumn($firstcol, $lastcol, $width, $format = null, $hidden = 0, $level = 0)
    { // added by Dan Lynn <dan@spiderweblabs.com) on 2006-12-06
        // look for any ranges this might overlap and remove, size or split where necessary
        foreach ($this->colinfo as $key => $colinfo) {
            $existing_start = $colinfo[0];
            $existing_end = $colinfo[1];
            // if the new range starts within another range
            if ($firstcol > $existing_start
                && $firstcol < $existing_end
            ) { // trim the existing range to the beginning of the new range
                $this->colinfo[$key][1] = $firstcol - 1;
                // if the new range lies WITHIN the existing range
                if ($lastcol < $existing_end) { // split the existing range by adding a range after our new range
                    $this->colinfo[] = array(
                        $lastcol + 1,
                        $existing_end,
                        $colinfo[2],
                        &$colinfo[3],
                        $colinfo[4],
                        $colinfo[5]
                    );
                }
            } // if the new range ends inside an existing range
            elseif ($lastcol > $existing_start
                && $lastcol < $existing_end
            ) { // trim the existing range to the end of the new range
                $this->colinfo[$key][0] = $lastcol + 1;
            } // if the new range completely overlaps the existing range
            elseif ($firstcol <= $existing_start && $lastcol >= $existing_end) {
                unset($this->colinfo[$key]);
            }
        } // added by Dan Lynn <dan@spiderweblabs.com) on 2006-12-06
        // regenerate keys
        $this->colinfo = array_values($this->colinfo);
        $this->colinfo[] = array($firstcol, $lastcol, $width, &$format, $hidden, $level);
        // Set width to zero if column is hidden
        $width = ($hidden) ? 0 : $width;
        for ($col = $firstcol; $col <= $lastcol; $col++) {
            $this->col_sizes[$col] = $width;
        }
    }

    /**
     * Set which cell or cells are selected in a worksheet
     *
     * @access public
     * @param integer $first_row    first row in the selected quadrant
     * @param integer $first_column first column in the selected quadrant
     * @param integer $last_row     last row in the selected quadrant
     * @param integer $last_column  last column in the selected quadrant
     */
    public function setSelection($first_row, $first_column, $last_row, $last_column)
    {
        $this->selection = array($first_row, $first_column, $last_row, $last_column);
    }

    /**
     * Set panes and mark them as frozen.
     *
     * @access public
     * @param array $panes This is the only parameter received and is composed of the following:
     *                     0 => Vertical split position,
     *                     1 => Horizontal split position
     *                     2 => Top row visible
     *                     3 => Leftmost column visible
     *                     4 => Active pane
     */
    public function freezePanes($panes)
    {
        $this->frozen = 1;
        $this->panes = $panes;
    }

    /**
     * Set panes and mark them as unfrozen.
     *
     * @access public
     * @param array $panes This is the only parameter received and is composed of the following:
     *                     0 => Vertical split position,
     *                     1 => Horizontal split position
     *                     2 => Top row visible
     *                     3 => Leftmost column visible
     *                     4 => Active pane
     */
    public function thawPanes($panes)
    {
        $this->frozen = 0;
        $this->panes = $panes;
    }

    /**
     * Set the page orientation as portrait.
     *
     * @access public
     */
    public function setPortrait()
    {
        $this->orientation = 1;
    }

    /**
     * Set the page orientation as landscape.
     *
     * @access public
     */
    public function setLandscape()
    {
        $this->orientation = 0;
    }

    /**
     * Set the paper type. Ex. 1 = US Letter, 9 = A4
     *
     * @access public
     * @param integer $size The type of paper size to use
     */
    public function setPaper($size = 0)
    {
        $this->paper_size = $size;
    }


    /**
     * Set the page header caption and optional margin.
     *
     * @access public
     * @param string $string The header text
     * @param float $margin optional head margin in inches.
     */
    public function setHeader($string, $margin = 0.50)
    {
        if (strlen($string) >= 255) {
            //carp 'Header string must be less than 255 characters';
            return;
        }
        $this->header = $string;
        $this->margin_head = $margin;
    }

    /**
     * Set the page footer caption and optional margin.
     *
     * @access public
     * @param string $string The footer text
     * @param float $margin optional foot margin in inches.
     */
    public function setFooter($string, $margin = 0.50)
    {
        if (strlen($string) >= 255) {
            //carp 'Footer string must be less than 255 characters';
            return;
        }
        $this->footer = $string;
        $this->margin_foot = $margin;
    }

    /**
     * Center the page horinzontally.
     *
     * @access public
     * @param integer $center the optional value for centering. Defaults to 1 (center).
     */
    public function centerHorizontally($center = 1)
    {
        $this->hcenter = $center;
    }

    /**
     * Center the page vertically.
     *
     * @access public
     * @param integer $center the optional value for centering. Defaults to 1 (center).
     */
    public function centerVertically($center = 1)
    {
        $this->vcenter = $center;
    }

    /**
     * Set all the page margins to the same value in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMargins($margin)
    {
        $this->setMarginLeft($margin);
        $this->setMarginRight($margin);
        $this->setMarginTop($margin);
        $this->setMarginBottom($margin);
    }

    /**
     * Set the left and right margins to the same value in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMargins_LR($margin)
    {
        $this->setMarginLeft($margin);
        $this->setMarginRight($margin);
    }

    /**
     * Set the top and bottom margins to the same value in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMargins_TB($margin)
    {
        $this->setMarginTop($margin);
        $this->setMarginBottom($margin);
    }

    /**
     * Set the left margin in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMarginLeft($margin = 0.75)
    {
        $this->margin_left = $margin;
    }

    /**
     * Set the right margin in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMarginRight($margin = 0.75)
    {
        $this->margin_right = $margin;
    }

    /**
     * Set the top margin in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMarginTop($margin = 1.00)
    {
        $this->margin_top = $margin;
    }

    /**
     * Set the bottom margin in inches.
     *
     * @access public
     * @param float $margin The margin to set in inches
     */
    public function setMarginBottom($margin = 1.00)
    {
        $this->margin_bottom = $margin;
    }

    /**
     * Set the rows to repeat at the top of each printed page.
     *
     * @access public
     * @param integer $first_row First row to repeat
     * @param integer $last_row  Last row to repeat. Optional.
     */
    public function repeatRows($first_row, $last_row = null)
    {
        $this->title_rowmin = $first_row;
        if (isset($last_row)) { //Second row is optional
            $this->title_rowmax = $last_row;
        } else {
            $this->title_rowmax = $first_row;
        }
    }

    /**
     * Set the columns to repeat at the left hand side of each printed page.
     *
     * @access public
     * @param integer $first_col First column to repeat
     * @param integer $last_col  Last column to repeat. Optional.
     */
    public function repeatColumns($first_col, $last_col = null)
    {
        $this->title_colmin = $first_col;
        if (isset($last_col)) { // Second col is optional
            $this->title_colmax = $last_col;
        } else {
            $this->title_colmax = $first_col;
        }
    }

    /**
     * Set the area of each worksheet that will be printed.
     *
     * @access public
     * @param integer $first_row First row of the area to print
     * @param integer $first_col First column of the area to print
     * @param integer $last_row  Last row of the area to print
     * @param integer $last_col  Last column of the area to print
     */
    public function printArea($first_row, $first_col, $last_row, $last_col)
    {
        $this->print_rowmin = $first_row;
        $this->print_colmin = $first_col;
        $this->print_rowmax = $last_row;
        $this->print_colmax = $last_col;
    }


    /**
     * Set the option to hide gridlines on the printed page.
     *
     * @access public
     */
    public function hideGridlines()
    {
        $this->print_gridlines = 0;
    }

    /**
     * Set the option to hide gridlines on the worksheet (as seen on the screen).
     *
     * @access public
     */
    public function hideScreenGridlines()
    {
        $this->screen_gridlines = 0;
    }

    /**
     * Set the option to print the row and column headers on the printed page.
     *
     * @access public
     * @param integer $print Whether to print the headers or not. Defaults to 1 (print).
     */
    public function printRowColHeaders($print = 1)
    {
        $this->print_headers = $print;
    }

    /**
     * Set the vertical and horizontal number of pages that will define the maximum area printed.
     * It doesn't seem to work with OpenOffice.
     *
     * @access public
     * @param  integer $width  Maximun width of printed area in pages
     * @param  integer $height Maximun heigth of printed area in pages
     * @see setPrintScale()
     */
    public function fitToPages($width, $height)
    {
        $this->fit_page = 1;
        $this->fit_width = $width;
        $this->fit_height = $height;
    }

    /**
     * Store the horizontal page breaks on a worksheet (for printing).
     * The breaks represent the row after which the break is inserted.
     *
     * @access public
     * @param array $breaks Array containing the horizontal page breaks
     */
    public function setHPagebreaks($breaks)
    {
        foreach ($breaks as $break) {
            array_push($this->hbreaks, $break);
        }
    }

    /**
     * Store the vertical page breaks on a worksheet (for printing).
     * The breaks represent the column after which the break is inserted.
     *
     * @access public
     * @param array $breaks Array containing the vertical page breaks
     */
    public function setVPagebreaks($breaks)
    {
        foreach ($breaks as $break) {
            array_push($this->vbreaks, $break);
        }
    }


    /**
     * Set the worksheet zoom factor.
     *
     * @access public
     * @param integer $scale The zoom factor
     * @throws \Exception
     */
    public function setZoom($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Zoom factor $scale outside range: 10 <= zoom <= 400");
        }

        $this->zoom = floor($scale);
    }

    /**
     * Set the scale factor for the printed page.
     * It turns off the "fit to page" option
     *
     * @access public
     * @param integer $scale The optional scale factor. Defaults to 100
     */
    public function setPrintScale($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Print scale $scale outside range: 10 <= zoom <= 400");
        }

        // Turn off "fit to page" option
        $this->fit_page = 0;

        $this->print_scale = floor($scale);
    }

    /**
     * Map to the appropriate write method acording to the token recieved.
     *
     * @access public
     * @param integer $row    The row of the cell we are writing to
     * @param integer $col    The column of the cell we are writing to
     * @param mixed $token  What we are writing
     * @param mixed $format The optional format to apply to the cell
     */
    public function write($row, $col, $token, $format = null)
    {
        // Check for a cell reference in A1 notation and substitute row and column
        /*if ($_[0] =~ /^\D/) {
            @_ = $this->substituteCellref(@_);
    }*/

        if (preg_match("/^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/", $token)) {
            // Match number
            return $this->writeNumber($row, $col, $token, $format);
        } elseif (preg_match("/^[fh]tt?p:\/\//", $token)) {
            // Match http or ftp URL
            return $this->writeUrl($row, $col, $token, '', $format);
        } elseif (preg_match("/^mailto:/", $token)) {
            // Match mailto:
            return $this->writeUrl($row, $col, $token, '', $format);
        } elseif (preg_match("/^(?:in|ex)ternal:/", $token)) {
            // Match internal or external sheet link
            return $this->writeUrl($row, $col, $token, '', $format);
        } elseif (preg_match("/^=/", $token)) {
            // Match formula
            return $this->writeFormula($row, $col, $token, $format);
        } elseif ($token == '') {
            // Match blank
            return $this->writeBlank($row, $col, $format);
        } else {
            // Default: match string
            return $this->writeString($row, $col, $token, $format);
        }
    }

    /**
     * Write an array of values as a row
     *
     * @access public
     * @param integer $row    The row we are writing to
     * @param integer $col    The first col (leftmost col) we are writing to
     * @param array $val    The array of values to write
     * @param mixed $format The optional format to apply to the cell
     * @return mixed
     */
    public function writeRow($row, $col, $val, $format = null)
    {
        $retval = '';
        if (is_array($val)) {
            foreach ($val as $v) {
                if (is_array($v)) {
                    $this->writeCol($row, $col, $v, $format);
                } else {
                    $this->write($row, $col, $v, $format);
                }
                $col++;
            }
        } else {
            throw new \Exception('$val needs to be an array');
        }
        return ($retval);
    }

    /**
     * Write an array of values as a column
     *
     * @access public
     * @param integer $row    The first row (uppermost row) we are writing to
     * @param integer $col    The col we are writing to
     * @param array $val    The array of values to write
     * @param mixed $format The optional format to apply to the cell
     * @return mixed
     */
    public function writeCol($row, $col, $val, $format = null)
    {
        $retval = '';
        if (is_array($val)) {
            foreach ($val as $v) {
                $this->write($row, $col, $v, $format);
                $row++;
            }
        } else {
            throw new \Exception('$val needs to be an array');
        }
        return ($retval);
    }

    /**
     * Returns an index to the XF record in the workbook
     *
     * @access private
     * @param mixed &$format The optional XF format
     * @return integer The XF record index
     */
    public function XF(&$format)
    {
        if ($format) {
            return ($format->getXfIndex());
        } else {
            return (0x0F);
        }
    }


    /******************************************************************************
     *******************************************************************************
     *
     * Internal methods
     */


    /**
     * Store Worksheet data in memory using the parent's class append() or to a
     * temporary file, the default.
     *
     * @access private
     * @param string $data The binary data to append
     */
    protected function append($data)
    {
        if ($this->using_tmpfile) {
            // Add CONTINUE records if necessary
            if (strlen($data) > $this->limit) {
                $data = $this->addContinue($data);
            }
            fwrite($this->filehandle, $data);
            $this->datasize += strlen($data);
        } else {
            parent::_append($data);
        }
    }

    /**
     * Substitute an Excel cell reference in A1 notation for  zero based row and
     * column values in an argument list.
     *
     * Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
     *
     * @access private
     * @param string $cell The cell reference. Or range of cells.
     * @throws \Exception
     * @return array
     */
    protected function substituteCellref($cell)
    {
        $cell = strtoupper($cell);

        // Convert a column range: 'A:A' or 'B:G'
        if (preg_match("/([A-I]?[A-Z]):([A-I]?[A-Z])/", $cell, $match)) {
            list($no_use, $col1) = $this->cellToRowcol($match[1] . '1'); // Add a dummy row
            list($no_use, $col2) = $this->cellToRowcol($match[2] . '1'); // Add a dummy row
            return (array($col1, $col2));
        }

        // Convert a cell range: 'A1:B7'
        if (preg_match("/\$?([A-I]?[A-Z]\$?\d+):\$?([A-I]?[A-Z]\$?\d+)/", $cell, $match)) {
            list($row1, $col1) = $this->cellToRowcol($match[1]);
            list($row2, $col2) = $this->cellToRowcol($match[2]);
            return (array($row1, $col1, $row2, $col2));
        }

        // Convert a cell reference: 'A1' or 'AD2000'
        if (preg_match("/\$?([A-I]?[A-Z]\$?\d+)/", $cell)) {
            list($row1, $col1) = $this->cellToRowcol($match[1]);
            return (array($row1, $col1));
        }

        throw new \Exception("Unknown cell reference $cell", 0);
    }

    /**
     * Convert an Excel cell reference in A1 notation to a zero based row and column
     * reference; converts C1 to (0, 2).
     *
     * @access private
     * @param string $cell The cell reference.
     * @return array containing (row, column)
     */
    protected function cellToRowcol($cell)
    {
        preg_match("/\$?([A-I]?[A-Z])\$?(\d+)/", $cell, $match);
        $col = $match[1];
        $row = $match[2];

        // Convert base26 column string to number
        $chars = explode('', $col);
        $expn = 0;
        $col = 0;

        while ($chars) {
            $char = array_pop($chars); // LS char first
            $col += (ord($char) - ord('A') + 1) * pow(26, $expn);
            $expn++;
        }

        // Convert 1-index to zero-index
        $row--;
        $col--;

        return (array($row, $col));
    }

    /**
     * Based on the algorithm provided by Daniel Rentz of OpenOffice.
     *
     * @access private
     * @param string $plaintext The password to be encoded in plaintext.
     * @return string The encoded password
     */
    protected function encodePassword($plaintext)
    {
        $password = 0x0000;
        $i = 1; // char position

        // split the plain text password in its component characters
        $chars = preg_split('//', $plaintext, -1, PREG_SPLIT_NO_EMPTY);
        foreach ($chars as $char) {
            $value = ord($char) << $i; // shifted ASCII value
            $rotated_bits = $value >> 15; // rotated bits beyond bit 15
            $value &= 0x7fff; // first 15 bits
            $password ^= ($value | $rotated_bits);
            $i++;
        }

        $password ^= strlen($plaintext);
        $password ^= 0xCE4B;

        return ($password);
    }

    /**
     * This method sets the properties for outlining and grouping. The defaults
     * correspond to Excel's defaults.
     *
     * @param bool $visible
     * @param bool $symbols_below
     * @param bool $symbols_right
     * @param bool $auto_style
     */
    public function setOutline($visible = true, $symbols_below = true, $symbols_right = true, $auto_style = false)
    {
        $this->outline_on = $visible;
        $this->outline_below = $symbols_below;
        $this->outline_right = $symbols_right;
        $this->outline_style = $auto_style;

        // Ensure this is a boolean vale for Window2
        if ($this->outline_on) {
            $this->outline_on = 1;
        }
    }

    /**
     * This method sets the worksheet direction to right-to-left (RTL)
     *
     * @param bool $rtl
     */
    public function setRTL($rtl = true)
    {
        $this->Arabic = ($rtl ? 1 : 0);
    }

    /******************************************************************************
     *******************************************************************************
     *
     * BIFF RECORDS
     */


    /**
     * Write a double to the specified row and column (zero indexed).
     * An integer can be written as a double. Excel will display an
     * integer. $format is optional.
     *
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *
     * @access public
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param float $num    The number to write
     * @param mixed $format The optional XF format
     * @return integer
     */
    public function writeNumber($row, $col, $num, $format = null)
    {
        $record = 0x0203; // Record identifier
        $length = 0x000E; // Number of bytes to follow

        $xf = $this->XF($format); // The cell format

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xls_rowmax) {
            return (-2);
        }
        if ($col >= $this->xls_colmax) {
            return (-2);
        }
        if ($row < $this->dim_rowmin) {
            $this->dim_rowmin = $row;
        }
        if ($row > $this->dim_rowmax) {
            $this->dim_rowmax = $row;
        }
        if ($col < $this->dim_colmin) {
            $this->dim_colmin = $col;
        }
        if ($col > $this->dim_colmax) {
            $this->dim_colmax = $col;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $xf);
        $xl_double = pack("d", $num);
        if ($this->byte_order) { // if it's Big Endian
            $xl_double = strrev($xl_double);
        }

        $this->append($header . $data . $xl_double);

        return 0;
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * NOTE: there is an Excel 5 defined limit of 255 characters.
     * $format is optional.
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *         -3 : long string truncated to 255 chars
     *
     * @access public
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     * @return integer
     */
    public function writeString($row, $col, $str, $format = null)
    {
        if ($this->BIFF_version == 0x0600) {
            return $this->writeStringBIFF8($row, $col, $str, $format);
        }
        $strlen = strlen($str);
        $record = 0x0204; // Record identifier
        $length = 0x0008 + $strlen; // Bytes to follow
        $xf = $this->XF($format); // The cell format

        $str_error = 0;

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xls_rowmax) {
            return (-2);
        }
        if ($col >= $this->xls_colmax) {
            return (-2);
        }
        if ($row < $this->dim_rowmin) {
            $this->dim_rowmin = $row;
        }
        if ($row > $this->dim_rowmax) {
            $this->dim_rowmax = $row;
        }
        if ($col < $this->dim_colmin) {
            $this->dim_colmin = $col;
        }
        if ($col > $this->dim_colmax) {
            $this->dim_colmax = $col;
        }

        if ($strlen > $this->xls_strmax) { // LABEL must be < 255 chars
            $str = substr($str, 0, $this->xls_strmax);
            $length = 0x0008 + $this->xls_strmax;
            $strlen = $this->xls_strmax;
            $str_error = -3;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row, $col, $xf, $strlen);
        $this->append($header . $data . $str);

        return ($str_error);
    }

    /**
     * Sets Input Encoding for writing strings
     *
     * @access public
     * @param string $encoding The encoding. Ex: 'UTF-16LE', 'utf-8', 'ISO-859-7'
     * @throws \Exception
     */
    public function setInputEncoding($encoding)
    {
        if ($encoding != 'UTF-16LE' && !function_exists('iconv')) {
            throw new \Exception("Using an input encoding other than UTF-16LE requires PHP support for iconv");
        }
        $this->input_encoding = $encoding;
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * This is the BIFF8 version (no 255 chars limit).
     * $format is optional.
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *         -3 : long string truncated to 255 chars
     *
     * @access public
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     * @return integer
     */
    public function writeStringBIFF8($row, $col, $str, $format = null)
    {
        if ($this->input_encoding == 'UTF-16LE') {
            $strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
            $encoding = 0x1;
        } elseif ($this->input_encoding != '') {
            $str = iconv($this->input_encoding, 'UTF-16LE', $str);
            $strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
            $encoding = 0x1;
        } else {
            $strlen = strlen($str);
            $encoding = 0x0;
        }
        $record = 0x00FD; // Record identifier
        $length = 0x000A; // Bytes to follow
        $xf = $this->XF($format); // The cell format

        $str_error = 0;

        // Check that row and col are valid and store max and min values
        if ($this->checkRowCol($row, $col) == false) {
            return -2;
        }

        $str = pack('vC', $strlen, $encoding) . $str;

        /* check if string is already present */
        if (!isset($this->str_table[$str])) {
            $this->str_table[$str] = $this->str_unique++;
        }
        $this->str_total++;

        $header = pack('vv', $record, $length);
        $data = pack('vvvV', $row, $col, $xf, $this->str_table[$str]);
        $this->append($header . $data);

        return $str_error;
    }

    /**
     * Check row and col before writing to a cell, and update the sheet's
     * dimensions accordingly
     *
     * @access private
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @return boolean true for success, false if row and/or col are grester
     *                 then maximums allowed.
     */
    protected function checkRowCol($row, $col)
    {
        if ($row >= $this->xls_rowmax) {
            return false;
        }
        if ($col >= $this->xls_colmax) {
            return false;
        }
        if ($row < $this->dim_rowmin) {
            $this->dim_rowmin = $row;
        }
        if ($row > $this->dim_rowmax) {
            $this->dim_rowmax = $row;
        }
        if ($col < $this->dim_colmin) {
            $this->dim_colmin = $col;
        }
        if ($col > $this->dim_colmax) {
            $this->dim_colmax = $col;
        }

        return true;
    }

    /**
     * Writes a note associated with the cell given by the row and column.
     * NOTE records don't have a length limit.
     *
     * @access public
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $note   The note to write
     */
    public function writeNote($row, $col, $note)
    {
        $note_length = strlen($note);
        $record = 0x001C; // Record identifier
        $max_length = 2048; // Maximun length for a NOTE record
        //$length      = 0x0006 + $note_length;    // Bytes to follow

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xls_rowmax) {
            return (-2);
        }
        if ($col >= $this->xls_colmax) {
            return (-2);
        }
        if ($row < $this->dim_rowmin) {
            $this->dim_rowmin = $row;
        }
        if ($row > $this->dim_rowmax) {
            $this->dim_rowmax = $row;
        }
        if ($col < $this->dim_colmin) {
            $this->dim_colmin = $col;
        }
        if ($col > $this->dim_colmax) {
            $this->dim_colmax = $col;
        }

        // Length for this record is no more than 2048 + 6
        $length = 0x0006 + min($note_length, 2048);
        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $note_length);
        $this->append($header . $data . substr($note, 0, 2048));

        for ($i = $max_length; $i < $note_length; $i += $max_length) {
            $chunk = substr($note, $i, $max_length);
            $length = 0x0006 + strlen($chunk);
            $header = pack("vv", $record, $length);
            $data = pack("vvv", -1, 0, strlen($chunk));
            $this->append($header . $data . $chunk);
        }

        return 0;
    }

    /**
     * Write a blank cell to the specified row and column (zero indexed).
     * A blank cell is used to specify formatting without adding a string
     * or a number.
     *
     * A blank cell without a format serves no purpose. Therefore, we don't write
     * a BLANK record unless a format is specified.
     *
     * Returns  0 : normal termination (including no format)
     *         -1 : insufficient number of arguments
     *         -2 : row or column out of range
     *
     * @access public
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param mixed $format The XF format
     * @return int
     */
    public function writeBlank($row, $col, $format)
    {
        // Don't write a blank cell unless it has a format
        if (!$format) {
            return 0;
        }

        $record = 0x0201; // Record identifier
        $length = 0x0006; // Number of bytes to follow
        $xf = $this->XF($format); // The cell format

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xls_rowmax) {
            return (-2);
        }
        if ($col >= $this->xls_colmax) {
            return (-2);
        }
        if ($row < $this->dim_rowmin) {
            $this->dim_rowmin = $row;
        }
        if ($row > $this->dim_rowmax) {
            $this->dim_rowmax = $row;
        }
        if ($col < $this->dim_colmin) {
            $this->dim_colmin = $col;
        }
        if ($col > $this->dim_colmax) {
            $this->dim_colmax = $col;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $xf);
        $this->append($header . $data);

        return 0;
    }

    /**
     * Write a formula to the specified row and column (zero indexed).
     * The textual representation of the formula is passed to the parser in
     * Parser.php which returns a packed binary string.
     *
     * Returns  0 : normal termination
     *         -1 : formula errors (bad formula)
     *         -2 : row or column out of range
     *
     * @access public
     * @param integer $row     Zero indexed row
     * @param integer $col     Zero indexed column
     * @param string $formula The formula text string
     * @param mixed $format  The optional XF format
     * @return integer
     */
    public function writeFormula($row, $col, $formula, $format = null)
    {
        $record = 0x0006; // Record identifier

        // Excel normally stores the last calculated value of the formula in $num.
        // Clearly we are not in a position to calculate this a priori. Instead
        // we set $num to zero and set the option flags in $grbit to ensure
        // automatic calculation of the formula when the file is opened.
        //
        $xf = $this->XF($format); // The cell format
        $num = 0x00; // Current value of formula
        $grbit = 0x03; // Option flags
        $unknown = 0x0000; // Must be zero


        // Check that row and col are valid and store max and min values
        if ($this->checkRowCol($row, $col) == false) {
            return -2;
        }

        // Strip the '=' or '@' sign at the beginning of the formula string
        if (preg_match("/^=/", $formula)) {
            $formula = preg_replace("/(^=)/", "", $formula);
        } elseif (preg_match("/^@/", $formula)) {
            $formula = preg_replace("/(^@)/", "", $formula);
        } else {
            // Error handling
            $this->writeString($row, $col, 'Unrecognised character for formula');
            return -1;
        }

        // Parse the formula using the parser in Parser.php
        $this->parser->parse($formula);

        $formula = $this->parser->toReversePolish();

        $formlen = strlen($formula); // Length of the binary string
        $length = 0x16 + $formlen; // Length of the record data

        $header = pack("vv", $record, $length);
        $data = pack(
            "vvvdvVv",
            $row,
            $col,
            $xf,
            $num,
            $grbit,
            $unknown,
            $formlen
        );

        $this->append($header . $data . $formula);

        return 0;
    }

    /**
     * Write a hyperlink.
     * This is comprised of two elements: the visible label and
     * the invisible link. The visible label is the same as the link unless an
     * alternative string is specified. The label is written using the
     * writeString() method. Therefore the 255 characters string limit applies.
     * $string and $format are optional.
     *
     * The hyperlink can be to a http, ftp, mail, internal sheet (not yet), or external
     * directory url.
     *
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *         -3 : long string truncated to 255 chars
     *
     * @access public
     * @param integer $row    Row
     * @param integer $col    Column
     * @param string $url    URL string
     * @param string $string Alternative label
     * @param mixed $format The cell format
     * @return integer
     */
    public function writeUrl($row, $col, $url, $string = '', $format = null)
    {
        // Add start row and col to arg list
        return ($this->writeUrlRange($row, $col, $row, $col, $url, $string, $format));
    }

    /**
     * This is the more general form of writeUrl(). It allows a hyperlink to be
     * written to a range of cells. This function also decides the type of hyperlink
     * to be written. These are either, Web (http, ftp, mailto), Internal
     * (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
     *
     * @access private
     * @see writeUrl()
     * @param integer $row1   Start row
     * @param integer $col1   Start column
     * @param integer $row2   End row
     * @param integer $col2   End column
     * @param string $url    URL string
     * @param string $string Alternative label
     * @param mixed $format The cell format
     * @return integer
     */
    protected function writeUrlRange($row1, $col1, $row2, $col2, $url, $string = '', $format = null)
    {
        // Check for internal/external sheet links or default to web link
        if (preg_match('[^internal:]', $url)) {
            return ($this->writeUrlInternal($row1, $col1, $row2, $col2, $url, $string, $format));
        }

        if (preg_match('[^external:]', $url)) {
            return ($this->writeUrlExternal($row1, $col1, $row2, $col2, $url, $string, $format));
        }

        return ($this->writeUrlWeb($row1, $col1, $row2, $col2, $url, $string, $format));
    }

    /**
     * Used to write http, ftp and mailto hyperlinks.
     * The link type ($options) is 0x03 is the same as absolute dir ref without
     * sheet. However it is differentiated by the $unknown2 data stream.
     *
     * @access private
     * @see writeUrl()
     * @param integer $row1   Start row
     * @param integer $col1   Start column
     * @param integer $row2   End row
     * @param integer $col2   End column
     * @param string $url    URL string
     * @param string $str    Alternative label
     * @param mixed $format The cell format
     * @return integer
     */
    protected function writeUrlWeb($row1, $col1, $row2, $col2, $url, $str, $format = null)
    {
        $record = 0x01B8; // Record identifier
        $length = 0x00000; // Bytes to follow

        if (!$format) {
            $format = $this->url_format;
        }

        // Write the visible label using the writeString() method.
        if ($str == '') {
            $str = $url;
        }
        $str_error = is_numeric($str)
            ? $this->writeNumber($row1, $col1, $str, $format)
            : $this->writeString(
                $row1,
                $col1,
                $str,
                $format
            );
        if (($str_error == -2) || ($str_error == -3)) {
            return $str_error;
        }

        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");
        $unknown2 = pack("H*", "E0C9EA79F9BACE118C8200AA004BA90B");

        // Pack the option flags
        $options = pack("V", 0x03);

        // Convert URL to a null terminated wchar string
        $url = join("\0", preg_split("''", $url, -1, PREG_SPLIT_NO_EMPTY));
        $url = $url . "\0\0\0";

        // Pack the length of the URL
        $url_len = pack("V", strlen($url));

        // Calculate the data length
        $length = 0x34 + strlen($url);

        // Pack the header data
        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row1, $row2, $col1, $col2);

        // Write the packed data
        $this->append(
            $header . $data .
            $unknown1 . $options .
            $unknown2 . $url_len . $url
        );

        return ($str_error);
    }

    /**
     * Used to write internal reference hyperlinks such as "Sheet1!A1".
     *
     * @access private
     * @see writeUrl()
     * @param integer $row1   Start row
     * @param integer $col1   Start column
     * @param integer $row2   End row
     * @param integer $col2   End column
     * @param string $url    URL string
     * @param string $str    Alternative label
     * @param mixed $format The cell format
     * @return integer
     */
    protected function writeUrlInternal($row1, $col1, $row2, $col2, $url, $str, $format = null)
    {
        $record = 0x01B8; // Record identifier
        $length = 0x00000; // Bytes to follow

        if (!$format) {
            $format = $this->url_format;
        }

        // Strip URL type
        $url = preg_replace('/^internal:/', '', $url);

        // Write the visible label
        if ($str == '') {
            $str = $url;
        }
        $str_error = is_numeric($str)
            ? $this->writeNumber($row1, $col1, $str, $format)
            : $this->writeString(
                $row1,
                $col1,
                $str,
                $format
            );
        if (($str_error == -2) || ($str_error == -3)) {
            return $str_error;
        }

        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");

        // Pack the option flags
        $options = pack("V", 0x08);

        // Convert the URL type and to a null terminated wchar string
        $url = join("\0", preg_split("''", $url, -1, PREG_SPLIT_NO_EMPTY));
        $url = $url . "\0\0\0";

        // Pack the length of the URL as chars (not wchars)
        $url_len = pack("V", floor(strlen($url) / 2));

        // Calculate the data length
        $length = 0x24 + strlen($url);

        // Pack the header data
        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row1, $row2, $col1, $col2);

        // Write the packed data
        $this->append(
            $header . $data .
            $unknown1 . $options .
            $url_len . $url
        );

        return ($str_error);
    }

    /**
     * Write links to external directory names such as 'c:\foo.xls',
     * c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
     *
     * Note: Excel writes some relative links with the $dir_long string. We ignore
     * these cases for the sake of simpler code.
     *
     * @access private
     * @see writeUrl()
     * @param integer $row1   Start row
     * @param integer $col1   Start column
     * @param integer $row2   End row
     * @param integer $col2   End column
     * @param string $url    URL string
     * @param string $str    Alternative label
     * @param mixed $format The cell format
     * @return integer
     */
    protected function writeUrlExternal($row1, $col1, $row2, $col2, $url, $str, $format = null)
    {
        // Network drives are different. We will handle them separately
        // MS/Novell network drives and shares start with \\
        if (preg_match('[^external:\\\\]', $url)) {
            return; //($this->writeUrlExternal_net($row1, $col1, $row2, $col2, $url, $str, $format));
        }

        $record = 0x01B8; // Record identifier
        $length = 0x00000; // Bytes to follow

        if (!$format) {
            $format = $this->url_format;
        }

        // Strip URL type and change Unix dir separator to Dos style (if needed)
        //
        $url = preg_replace('/^external:/', '', $url);
        $url = preg_replace('/\//', "\\", $url);

        // Write the visible label
        if ($str == '') {
            $str = preg_replace('/\#/', ' - ', $url);
        }
        $str_error = is_numeric($str)
            ? $this->writeNumber($row1, $col1, $str, $format)
            : $this->writeString(
                $row1,
                $col1,
                $str,
                $format
            );
        if (($str_error == -2) or ($str_error == -3)) {
            return $str_error;
        }

        // Determine if the link is relative or absolute:
        //   relative if link contains no dir separator, "somefile.xls"
        //   relative if link starts with up-dir, "..\..\somefile.xls"
        //   otherwise, absolute

        $absolute = 0x02; // Bit mask
        if (!preg_match("/\\\/", $url)) {
            $absolute = 0x00;
        }
        if (preg_match("/^\.\.\\\/", $url)) {
            $absolute = 0x00;
        }
        $link_type = 0x01 | $absolute;

        // Determine if the link contains a sheet reference and change some of the
        // parameters accordingly.
        // Split the dir name and sheet name (if it exists)
        /*if (preg_match("/\#/", $url)) {
            list($dir_long, $sheet) = split("\#", $url);
        } else {
            $dir_long = $url;
        }

        if (isset($sheet)) {
            $link_type |= 0x08;
            $sheet_len  = pack("V", strlen($sheet) + 0x01);
            $sheet      = join("\0", split('', $sheet));
            $sheet     .= "\0\0\0";
        } else {
            $sheet_len   = '';
            $sheet       = '';
        }*/
        $dir_long = $url;
        if (preg_match("/\#/", $url)) {
            $link_type |= 0x08;
        }


        // Pack the link type
        $link_type = pack("V", $link_type);

        // Calculate the up-level dir count e.g.. (..\..\..\ == 3)
        $up_count = preg_match_all("/\.\.\\\/", $dir_long, $useless);
        $up_count = pack("v", $up_count);

        // Store the short dos dir name (null terminated)
        $dir_short = preg_replace("/\.\.\\\/", '', $dir_long) . "\0";

        // Store the long dir name as a wchar string (non-null terminated)
        //$dir_long       = join("\0", split('', $dir_long));
        $dir_long = $dir_long . "\0";

        // Pack the lengths of the dir strings
        $dir_short_len = pack("V", strlen($dir_short));
        $dir_long_len = pack("V", strlen($dir_long));
        $stream_len = pack("V", 0); //strlen($dir_long) + 0x06);

        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack("H*", 'D0C9EA79F9BACE118C8200AA004BA90B02000000');
        $unknown2 = pack("H*", '0303000000000000C000000000000046');
        $unknown3 = pack("H*", 'FFFFADDE000000000000000000000000000000000000000');
        $unknown4 = pack("v", 0x03);

        // Pack the main data stream
        $data = pack("vvvv", $row1, $row2, $col1, $col2) .
            $unknown1 .
            $link_type .
            $unknown2 .
            $up_count .
            $dir_short_len .
            $dir_short .
            $unknown3 .
            $stream_len;
        /*.
                                  $dir_long_len .
                                  $unknown4     .
                                  $dir_long     .
                                  $sheet_len    .
                                  $sheet        ;*/

        // Pack the header data
        $length = strlen($data);
        $header = pack("vv", $record, $length);

        // Write the packed data
        $this->append($header . $data);

        return ($str_error);
    }


    /**
     * This method is used to set the height and format for a row.
     *
     * @access public
     * @param integer $row    The row to set
     * @param integer $height Height we are giving to the row.
     *                        Use null to set XF without setting height
     * @param mixed $format XF format we are giving to the row
     * @param bool $hidden The optional hidden attribute
     * @param integer $level  The optional outline level for row, in range [0,7]
     */
    public function setRow($row, $height, $format = null, $hidden = false, $level = 0)
    {
        $record = 0x0208; // Record identifier
        $length = 0x0010; // Number of bytes to follow

        $colMic = 0x0000; // First defined column
        $colMac = 0x0000; // Last defined column
        $irwMac = 0x0000; // Used by Excel to optimise loading
        $reserved = 0x0000; // Reserved
        $grbit = 0x0000; // Option flags
        $ixfe = $this->XF($format); // XF index

        // set _row_sizes so _sizeRow() can use it
        $this->row_sizes[$row] = $height;

        // Use setRow($row, null, $XF) to set XF format without setting height
        if ($height != null) {
            $miyRw = $height * 20; // row height
        } else {
            $miyRw = 0xff; // default row height is 256
        }

        $level = max(0, min($level, 7)); // level should be between 0 and 7
        $this->outline_row_level = max($level, $this->outline_row_level);


        // Set the options flags. fUnsynced is used to show that the font and row
        // heights are not compatible. This is usually the case for WriteExcel.
        // The collapsed flag 0x10 doesn't seem to be used to indicate that a row
        // is collapsed. Instead it is used to indicate that the previous row is
        // collapsed. The zero height flag, 0x20, is used to collapse a row.

        $grbit |= $level;
        if ($hidden) {
            $grbit |= 0x0020;
        }
        $grbit |= 0x0040; // fUnsynced
        if ($format) {
            $grbit |= 0x0080;
        }
        $grbit |= 0x0100;

        $header = pack("vv", $record, $length);
        $data = pack(
            "vvvvvvvv",
            $row,
            $colMic,
            $colMac,
            $miyRw,
            $irwMac,
            $reserved,
            $grbit,
            $ixfe
        );
        $this->append($header . $data);
    }

    /**
     * Writes Excel DIMENSIONS to define the area in which there is data.
     *
     * @access private
     * @throw \Exception
     */
    protected function storeDimensions()
    {
        $record = 0x0200; // Record identifier
        $row_min = $this->dim_rowmin; // First row
        $row_max = $this->dim_rowmax + 1; // Last row plus 1
        $col_min = $this->dim_colmin; // First column
        $col_max = $this->dim_colmax + 1; // Last column plus 1
        $reserved = 0x0000; // Reserved by Excel

        if ($this->BIFF_version == 0x0500) {
            $length = 0x000A; // Number of bytes to follow
            $data = pack(
                "vvvvv",
                $row_min,
                $row_max,
                $col_min,
                $col_max,
                $reserved
            );
        } elseif ($this->BIFF_version == 0x0600) {
            $length = 0x000E;
            $data = pack(
                "VVvvv",
                $row_min,
                $row_max,
                $col_min,
                $col_max,
                $reserved
            );
        } else {
            throw new \Exception('Unsupported BIFF version');
        }

        $header = pack("vv", $record, $length);
        $this->prepend($header . $data);
    }

    /**
     * Write BIFF record Window2.
     *
     * @access private
     */
    protected function storeWindow2()
    {
        $record = 0x023E; // Record identifier
        if ($this->BIFF_version == 0x0500) {
            $length = 0x000A; // Number of bytes to follow
        } elseif ($this->BIFF_version == 0x0600) {
            $length = 0x0012;
        }

        $grbit = 0x00B6; // Option flags
        $rwTop = 0x0000; // Top row visible in window
        $colLeft = 0x0000; // Leftmost column visible in window


        // The options flags that comprise $grbit
        $fDspFmla = 0; // 0 - bit
        $fDspGrid = $this->screen_gridlines; // 1
        $fDspRwCol = 1; // 2
        $fFrozen = $this->frozen; // 3
        $fDspZeros = 1; // 4
        $fDefaultHdr = 1; // 5
        $fArabic = $this->Arabic; // 6
        $fDspGuts = $this->outline_on; // 7
        $fFrozenNoSplit = 0; // 0 - bit
        $fSelected = $this->selected; // 1
        $fPaged = 1; // 2

        $grbit = $fDspFmla;
        $grbit |= $fDspGrid << 1;
        $grbit |= $fDspRwCol << 2;
        $grbit |= $fFrozen << 3;
        $grbit |= $fDspZeros << 4;
        $grbit |= $fDefaultHdr << 5;
        $grbit |= $fArabic << 6;
        $grbit |= $fDspGuts << 7;
        $grbit |= $fFrozenNoSplit << 8;
        $grbit |= $fSelected << 9;
        $grbit |= $fPaged << 10;

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $grbit, $rwTop, $colLeft);
        // FIXME !!!
        if ($this->BIFF_version == 0x0500) {
            $rgbHdr = 0x00000000; // Row/column heading and gridline color
            $data .= pack("V", $rgbHdr);
        } elseif ($this->BIFF_version == 0x0600) {
            $rgbHdr = 0x0040; // Row/column heading and gridline color index
            $zoom_factor_page_break = 0x0000;
            $zoom_factor_normal = 0x0000;
            $data .= pack("vvvvV", $rgbHdr, 0x0000, $zoom_factor_page_break, $zoom_factor_normal, 0x00000000);
        }
        $this->append($header . $data);
    }

    /**
     * Write BIFF record DEFCOLWIDTH if COLINFO records are in use.
     *
     * @access private
     */
    protected function storeDefcol()
    {
        $record = 0x0055; // Record identifier
        $length = 0x0002; // Number of bytes to follow
        $colwidth = 0x0008; // Default column width

        $header = pack("vv", $record, $length);
        $data = pack("v", $colwidth);
        $this->prepend($header . $data);
    }

    /**
     * Write BIFF record COLINFO to define column widths
     *
     * Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
     * length record.
     *
     * @access private
     * @param array $col_array This is the only parameter received and is composed of the following:
     *                0 => First formatted column,
     *                1 => Last formatted column,
     *                2 => Col width (8.43 is Excel default),
     *                3 => The optional XF format of the column,
     *                4 => Option flags.
     *                5 => Optional outline level
     */
    protected function storeColinfo($col_array)
    {
        if (isset($col_array[0])) {
            $colFirst = $col_array[0];
        }
        if (isset($col_array[1])) {
            $colLast = $col_array[1];
        }
        if (isset($col_array[2])) {
            $coldx = $col_array[2];
        } else {
            $coldx = 8.43;
        }
        if (isset($col_array[3])) {
            $format = $col_array[3];
        } else {
            $format = 0;
        }
        if (isset($col_array[4])) {
            $grbit = $col_array[4];
        } else {
            $grbit = 0;
        }
        if (isset($col_array[5])) {
            $level = $col_array[5];
        } else {
            $level = 0;
        }
        $record = 0x007D; // Record identifier
        $length = 0x000B; // Number of bytes to follow

        $coldx += 0.72; // Fudge. Excel subtracts 0.72 !?
        $coldx *= 256; // Convert to units of 1/256 of a char

        $ixfe = $this->XF($format);
        $reserved = 0x00; // Reserved

        $level = max(0, min($level, 7));
        $grbit |= $level << 8;

        $header = pack("vv", $record, $length);
        $data = pack(
            "vvvvvC",
            $colFirst,
            $colLast,
            $coldx,
            $ixfe,
            $grbit,
            $reserved
        );
        $this->prepend($header . $data);
    }

    /**
     * Write BIFF record SELECTION.
     *
     * @access private
     * @param array $array array containing ($rwFirst,$colFirst,$rwLast,$colLast)
     * @see setSelection()
     */
    protected function storeSelection($array)
    {
        list($rwFirst, $colFirst, $rwLast, $colLast) = $array;
        $record = 0x001D; // Record identifier
        $length = 0x000F; // Number of bytes to follow

        $pnn = $this->active_pane; // Pane position
        $rwAct = $rwFirst; // Active row
        $colAct = $colFirst; // Active column
        $irefAct = 0; // Active cell ref
        $cref = 1; // Number of refs

        if (!isset($rwLast)) {
            $rwLast = $rwFirst; // Last  row in reference
        }
        if (!isset($colLast)) {
            $colLast = $colFirst; // Last  col in reference
        }

        // Swap last row/col for first row/col as necessary
        if ($rwFirst > $rwLast) {
            list($rwFirst, $rwLast) = array($rwLast, $rwFirst);
        }

        if ($colFirst > $colLast) {
            list($colFirst, $colLast) = array($colLast, $colFirst);
        }

        $header = pack("vv", $record, $length);
        $data = pack(
            "CvvvvvvCC",
            $pnn,
            $rwAct,
            $colAct,
            $irefAct,
            $cref,
            $rwFirst,
            $rwLast,
            $colFirst,
            $colLast
        );
        $this->append($header . $data);
    }

    /**
     * Store the MERGEDCELLS record for all ranges of merged cells
     *
     * @access private
     */
    protected function storeMergedCells()
    {
        // if there are no merged cell ranges set, return
        if (count($this->merged_ranges) == 0) {
            return;
        }
        $record = 0x00E5;
        foreach ($this->merged_ranges as $ranges) {
            $length = 2 + count($ranges) * 8;
            $header = pack('vv', $record, $length);
            $data = pack('v', count($ranges));
            foreach ($ranges as $range) {
                $data .= pack('vvvv', $range[0], $range[2], $range[1], $range[3]);
            }
            $string = $header . $data;
            $this->append($string, true);
        }
    }

    /**
     * Write BIFF record EXTERNCOUNT to indicate the number of external sheet
     * references in a worksheet.
     *
     * Excel only stores references to external sheets that are used in formulas.
     * For simplicity we store references to all the sheets in the workbook
     * regardless of whether they are used or not. This reduces the overall
     * complexity and eliminates the need for a two way dialogue between the formula
     * parser the worksheet objects.
     *
     * @access private
     * @param integer $count The number of external sheet references in this worksheet
     */
    protected function storeExterncount($count)
    {
        $record = 0x0016; // Record identifier
        $length = 0x0002; // Number of bytes to follow

        $header = pack("vv", $record, $length);
        $data = pack("v", $count);
        $this->prepend($header . $data);
    }

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas. A formula references a sheet name via an index. Since we store a
     * reference to all of the external worksheets the EXTERNSHEET index is the same
     * as the worksheet index.
     *
     * @access private
     * @param string $sheetname The name of a external worksheet
     */
    protected function storeExternsheet($sheetname)
    {
        $record = 0x0017; // Record identifier

        // References to the current sheet are encoded differently to references to
        // external sheets.
        //
        if ($this->name == $sheetname) {
            $sheetname = '';
            $length = 0x02; // The following 2 bytes
            $cch = 1; // The following byte
            $rgch = 0x02; // Self reference
        } else {
            $length = 0x02 + strlen($sheetname);
            $cch = strlen($sheetname);
            $rgch = 0x03; // Reference to a sheet in the current workbook
        }

        $header = pack("vv", $record, $length);
        $data = pack("CC", $cch, $rgch);
        $this->prepend($header . $data . $sheetname);
    }

    /**
     * Writes the Excel BIFF PANE record.
     * The panes can either be frozen or thawed (unfrozen).
     * Frozen panes are specified in terms of an integer number of rows and columns.
     * Thawed panes are specified in terms of Excel's units for rows and columns.
     *
     * @access private
     * @param array $panes This is the only parameter received and is composed of the following:
     *                     0 => Vertical split position,
     *                     1 => Horizontal split position
     *                     2 => Top row visible
     *                     3 => Leftmost column visible
     *                     4 => Active pane
     */
    protected function storePanes($panes)
    {
        $y = $panes[0];
        $x = $panes[1];
        $rwTop = $panes[2];
        $colLeft = $panes[3];
        if (count($panes) > 4) { // if Active pane was received
            $pnnAct = $panes[4];
        } else {
            $pnnAct = null;
        }
        $record = 0x0041; // Record identifier
        $length = 0x000A; // Number of bytes to follow

        // Code specific to frozen or thawed panes.
        if ($this->frozen) {
            // Set default values for $rwTop and $colLeft
            if (!isset($rwTop)) {
                $rwTop = $y;
            }
            if (!isset($colLeft)) {
                $colLeft = $x;
            }
        } else {
            // Set default values for $rwTop and $colLeft
            if (!isset($rwTop)) {
                $rwTop = 0;
            }
            if (!isset($colLeft)) {
                $colLeft = 0;
            }

            // Convert Excel's row and column units to the internal units.
            // The default row height is 12.75
            // The default column width is 8.43
            // The following slope and intersection values were interpolated.
            //
            $y = 20 * $y + 255;
            $x = 113.879 * $x + 390;
        }


        // Determine which pane should be active. There is also the undocumented
        // option to override this should it be necessary: may be removed later.
        //
        if (!isset($pnnAct)) {
            if ($x != 0 && $y != 0) {
                $pnnAct = 0; // Bottom right
            }
            if ($x != 0 && $y == 0) {
                $pnnAct = 1; // Top right
            }
            if ($x == 0 && $y != 0) {
                $pnnAct = 2; // Bottom left
            }
            if ($x == 0 && $y == 0) {
                $pnnAct = 3; // Top left
            }
        }

        $this->active_pane = $pnnAct; // Used in _storeSelection

        $header = pack("vv", $record, $length);
        $data = pack("vvvvv", $x, $y, $rwTop, $colLeft, $pnnAct);
        $this->append($header . $data);
    }

    /**
     * Store the page setup SETUP BIFF record.
     *
     * @access private
     */
    protected function storeSetup()
    {
        $record = 0x00A1; // Record identifier
        $length = 0x0022; // Number of bytes to follow

        $iPaperSize = $this->paper_size; // Paper size
        $iScale = $this->print_scale; // Print scaling factor
        $iPageStart = 0x01; // Starting page number
        $iFitWidth = $this->fit_width; // Fit to number of pages wide
        $iFitHeight = $this->fit_height; // Fit to number of pages high
        $grbit = 0x00; // Option flags
        $iRes = 0x0258; // Print resolution
        $iVRes = 0x0258; // Vertical print resolution
        $numHdr = $this->margin_head; // Header Margin
        $numFtr = $this->margin_foot; // Footer Margin
        $iCopies = 0x01; // Number of copies

        $fLeftToRight = 0x0; // Print over then down
        $fLandscape = $this->orientation; // Page orientation
        $fNoPls = 0x0; // Setup not read from printer
        $fNoColor = 0x0; // Print black and white
        $fDraft = 0x0; // Print draft quality
        $fNotes = 0x0; // Print notes
        $fNoOrient = 0x0; // Orientation not set
        $fUsePage = 0x0; // Use custom starting page

        $grbit = $fLeftToRight;
        $grbit |= $fLandscape << 1;
        $grbit |= $fNoPls << 2;
        $grbit |= $fNoColor << 3;
        $grbit |= $fDraft << 4;
        $grbit |= $fNotes << 5;
        $grbit |= $fNoOrient << 6;
        $grbit |= $fUsePage << 7;

        $numHdr = pack("d", $numHdr);
        $numFtr = pack("d", $numFtr);
        if ($this->byte_order) { // if it's Big Endian
            $numHdr = strrev($numHdr);
            $numFtr = strrev($numFtr);
        }

        $header = pack("vv", $record, $length);
        $data1 = pack(
            "vvvvvvvv",
            $iPaperSize,
            $iScale,
            $iPageStart,
            $iFitWidth,
            $iFitHeight,
            $grbit,
            $iRes,
            $iVRes
        );
        $data2 = $numHdr . $numFtr;
        $data3 = pack("v", $iCopies);
        $this->prepend($header . $data1 . $data2 . $data3);
    }

    /**
     * Store the header caption BIFF record.
     *
     * @access private
     */
    protected function storeHeader()
    {
        $record = 0x0014; // Record identifier

        $str = $this->header; // header string
        $cch = strlen($str); // Length of header string
        if ($this->BIFF_version == 0x0600) {
            $encoding = 0x0; // TODO: Unicode support
            $length = 3 + $cch; // Bytes to follow
        } else {
            $length = 1 + $cch; // Bytes to follow
        }

        $header = pack("vv", $record, $length);
        if ($this->BIFF_version == 0x0600) {
            $data = pack("vC", $cch, $encoding);
        } else {
            $data = pack("C", $cch);
        }

        $this->prepend($header . $data . $str);
    }

    /**
     * Store the footer caption BIFF record.
     *
     * @access private
     */
    protected function storeFooter()
    {
        $record = 0x0015; // Record identifier

        $str = $this->footer; // Footer string
        $cch = strlen($str); // Length of footer string
        if ($this->BIFF_version == 0x0600) {
            $encoding = 0x0; // TODO: Unicode support
            $length = 3 + $cch; // Bytes to follow
        } else {
            $length = 1 + $cch;
        }

        $header = pack("vv", $record, $length);
        if ($this->BIFF_version == 0x0600) {
            $data = pack("vC", $cch, $encoding);
        } else {
            $data = pack("C", $cch);
        }

        $this->prepend($header . $data . $str);
    }

    /**
     * Store the horizontal centering HCENTER BIFF record.
     *
     * @access private
     */
    protected function storeHcenter()
    {
        $record = 0x0083; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fHCenter = $this->hcenter; // Horizontal centering

        $header = pack("vv", $record, $length);
        $data = pack("v", $fHCenter);

        $this->prepend($header . $data);
    }

    /**
     * Store the vertical centering VCENTER BIFF record.
     *
     * @access private
     */
    protected function storeVcenter()
    {
        $record = 0x0084; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fVCenter = $this->vcenter; // Horizontal centering

        $header = pack("vv", $record, $length);
        $data = pack("v", $fVCenter);
        $this->prepend($header . $data);
    }

    /**
     * Store the LEFTMARGIN BIFF record.
     *
     * @access private
     */
    protected function storeMarginLeft()
    {
        $record = 0x0026; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->margin_left; // Margin in inches

        $header = pack("vv", $record, $length);
        $data = pack("d", $margin);
        if ($this->byte_order) { // if it's Big Endian
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Store the RIGHTMARGIN BIFF record.
     *
     * @access private
     */
    protected function storeMarginRight()
    {
        $record = 0x0027; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->margin_right; // Margin in inches

        $header = pack("vv", $record, $length);
        $data = pack("d", $margin);
        if ($this->byte_order) { // if it's Big Endian
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Store the TOPMARGIN BIFF record.
     *
     * @access private
     */
    protected function storeMarginTop()
    {
        $record = 0x0028; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->margin_top; // Margin in inches

        $header = pack("vv", $record, $length);
        $data = pack("d", $margin);
        if ($this->byte_order) { // if it's Big Endian
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Store the BOTTOMMARGIN BIFF record.
     *
     * @access private
     */
    protected function storeMarginBottom()
    {
        $record = 0x0029; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->margin_bottom; // Margin in inches

        $header = pack("vv", $record, $length);
        $data = pack("d", $margin);
        if ($this->byte_order) { // if it's Big Endian
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Merges the area given by its arguments.
     * This is an Excel97/2000 method. It is required to perform more complicated
     * merging than the normal setAlign('merge').
     *
     * @access public
     * @param integer $first_row First row of the area to merge
     * @param integer $first_col First column of the area to merge
     * @param integer $last_row  Last row of the area to merge
     * @param integer $last_col  Last column of the area to merge
     */
    public function mergeCells($first_row, $first_col, $last_row, $last_col)
    {
        $record = 0x00E5; // Record identifier
        $length = 0x000A; // Bytes to follow
        $cref = 1; // Number of refs

        // Swap last row/col for first row/col as necessary
        if ($first_row > $last_row) {
            list($first_row, $last_row) = array($last_row, $first_row);
        }

        if ($first_col > $last_col) {
            list($first_col, $last_col) = array($last_col, $first_col);
        }

        $header = pack("vv", $record, $length);
        $data = pack(
            "vvvvv",
            $cref,
            $first_row,
            $last_row,
            $first_col,
            $last_col
        );

        $this->append($header . $data);
    }

    /**
     * Write the PRINTHEADERS BIFF record.
     *
     * @access private
     */
    protected function storePrintHeaders()
    {
        $record = 0x002a; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fPrintRwCol = $this->print_headers; // Boolean flag

        $header = pack("vv", $record, $length);
        $data = pack("v", $fPrintRwCol);
        $this->prepend($header . $data);
    }

    /**
     * Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
     * GRIDSET record.
     *
     * @access private
     */
    protected function storePrintGridlines()
    {
        $record = 0x002b; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fPrintGrid = $this->print_gridlines; // Boolean flag

        $header = pack("vv", $record, $length);
        $data = pack("v", $fPrintGrid);
        $this->prepend($header . $data);
    }

    /**
     * Write the GRIDSET BIFF record. Must be used in conjunction with the
     * PRINTGRIDLINES record.
     *
     * @access private
     */
    protected function storeGridset()
    {
        $record = 0x0082; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fGridSet = !($this->print_gridlines); // Boolean flag

        $header = pack("vv", $record, $length);
        $data = pack("v", $fGridSet);
        $this->prepend($header . $data);
    }

    /**
     * Write the GUTS BIFF record. This is used to configure the gutter margins
     * where Excel outline symbols are displayed. The visibility of the gutters is
     * controlled by a flag in WSBOOL.
     *
     * @see _storeWsbool()
     * @access private
     */
    protected function storeGuts()
    {
        $record = 0x0080; // Record identifier
        $length = 0x0008; // Bytes to follow

        $dxRwGut = 0x0000; // Size of row gutter
        $dxColGut = 0x0000; // Size of col gutter

        $row_level = $this->outline_row_level;
        $col_level = 0;

        // Calculate the maximum column outline level. The equivalent calculation
        // for the row outline level is carried out in setRow().
        $colcount = count($this->colinfo);
        for ($i = 0; $i < $colcount; $i++) {
            // Skip cols without outline level info.
            if (count($this->colinfo[$i]) >= 6) {
                $col_level = max($this->colinfo[$i][5], $col_level);
            }
        }

        // Set the limits for the outline levels (0 <= x <= 7).
        $col_level = max(0, min($col_level, 7));

        // The displayed level is one greater than the max outline levels
        if ($row_level) {
            $row_level++;
        }
        if ($col_level) {
            $col_level++;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $dxRwGut, $dxColGut, $row_level, $col_level);

        $this->prepend($header . $data);
    }


    /**
     * Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
     * with the SETUP record.
     *
     * @access private
     */
    protected function storeWsbool()
    {
        $record = 0x0081; // Record identifier
        $length = 0x0002; // Bytes to follow
        $grbit = 0x0000;

        // The only option that is of interest is the flag for fit to page. So we
        // set all the options in one go.
        //
        /*if ($this->fit_page) {
            $grbit = 0x05c1;
        } else {
            $grbit = 0x04c1;
        }*/
        // Set the option flags
        $grbit |= 0x0001; // Auto page breaks visible
        if ($this->outline_style) {
            $grbit |= 0x0020; // Auto outline styles
        }
        if ($this->outline_below) {
            $grbit |= 0x0040; // Outline summary below
        }
        if ($this->outline_right) {
            $grbit |= 0x0080; // Outline summary right
        }
        if ($this->fit_page) {
            $grbit |= 0x0100; // Page setup fit to page
        }
        if ($this->outline_on) {
            $grbit |= 0x0400; // Outline symbols displayed
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $grbit);
        $this->prepend($header . $data);
    }

    /**
     * Write the HORIZONTALPAGEBREAKS BIFF record.
     *
     * @access private
     */
    protected function storeHbreak()
    {
        // Return if the user hasn't specified pagebreaks
        if (empty($this->hbreaks)) {
            return;
        }

        // Sort and filter array of page breaks
        $breaks = $this->hbreaks;
        sort($breaks, SORT_NUMERIC);
        if ($breaks[0] == 0) { // don't use first break if it's 0
            array_shift($breaks);
        }

        $record = 0x001b; // Record identifier
        $cbrk = count($breaks); // Number of page breaks
        if ($this->BIFF_version == 0x0600) {
            $length = 2 + 6 * $cbrk; // Bytes to follow
        } else {
            $length = 2 + 2 * $cbrk; // Bytes to follow
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $cbrk);

        // Append each page break
        foreach ($breaks as $break) {
            if ($this->BIFF_version == 0x0600) {
                $data .= pack("vvv", $break, 0x0000, 0x00ff);
            } else {
                $data .= pack("v", $break);
            }
        }

        $this->prepend($header . $data);
    }

    /**
     * Write the VERTICALPAGEBREAKS BIFF record.
     *
     * @access private
     */
    protected function storeVbreak()
    {
        // Return if the user hasn't specified pagebreaks
        if (empty($this->vbreaks)) {
            return;
        }

        // 1000 vertical pagebreaks appears to be an internal Excel 5 limit.
        // It is slightly higher in Excel 97/200, approx. 1026
        $breaks = array_slice($this->vbreaks, 0, 1000);

        // Sort and filter array of page breaks
        sort($breaks, SORT_NUMERIC);
        if ($breaks[0] == 0) { // don't use first break if it's 0
            array_shift($breaks);
        }

        $record = 0x001a; // Record identifier
        $cbrk = count($breaks); // Number of page breaks
        if ($this->BIFF_version == 0x0600) {
            $length = 2 + 6 * $cbrk; // Bytes to follow
        } else {
            $length = 2 + 2 * $cbrk; // Bytes to follow
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $cbrk);

        // Append each page break
        foreach ($breaks as $break) {
            if ($this->BIFF_version == 0x0600) {
                $data .= pack("vvv", $break, 0x0000, 0xffff);
            } else {
                $data .= pack("v", $break);
            }
        }

        $this->prepend($header . $data);
    }

    /**
     * Set the Biff PROTECT record to indicate that the worksheet is protected.
     *
     * @access private
     */
    protected function storeProtect()
    {
        // Exit unless sheet protection has been specified
        if ($this->protect == 0) {
            return;
        }

        $record = 0x0012; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fLock = $this->protect; // Worksheet is protected

        $header = pack("vv", $record, $length);
        $data = pack("v", $fLock);

        $this->prepend($header . $data);
    }

    /**
     * Write the worksheet PASSWORD record.
     *
     * @access private
     */
    protected function storePassword()
    {
        // Exit unless sheet protection and password have been specified
        if (($this->protect == 0) || (!isset($this->password))) {
            return;
        }

        $record = 0x0013; // Record identifier
        $length = 0x0002; // Bytes to follow

        $wPassword = $this->password; // Encoded password

        $header = pack("vv", $record, $length);
        $data = pack("v", $wPassword);

        $this->prepend($header . $data);
    }


    /**
     * Insert a 24bit bitmap image in a worksheet.
     *
     * @access public
     * @param integer $row     The row we are going to insert the bitmap into
     * @param integer $col     The column we are going to insert the bitmap into
     * @param string $bitmap  The bitmap filename
     * @param integer $x       The horizontal position (offset) of the image inside the cell.
     * @param integer $y       The vertical position (offset) of the image inside the cell.
     * @param integer $scale_x The horizontal scale
     * @param integer $scale_y The vertical scale
     */
    public function insertBitmap($row, $col, $bitmap, $x = 0, $y = 0, $scale_x = 1, $scale_y = 1)
    {
        $bitmap_array = $this->processBitmap($bitmap);
        list($width, $height, $size, $data) = $bitmap_array;

        // Scale the frame of the image.
        $width *= $scale_x;
        $height *= $scale_y;

        // Calculate the vertices of the image and write the OBJ record
        $this->positionImage($col, $row, $x, $y, $width, $height);

        // Write the IMDATA record to store the bitmap data
        $record = 0x007f;
        $length = 8 + $size;
        $cf = 0x09;
        $env = 0x01;
        $lcb = $size;

        $header = pack("vvvvV", $record, $length, $cf, $env, $lcb);
        $this->append($header . $data);
    }

    /**
     * Calculate the vertices that define the position of the image as required by
     * the OBJ record.
     *
     *         +------------+------------+
     *         |     A      |      B     |
     *   +-----+------------+------------+
     *   |     |(x1,y1)     |            |
     *   |  1  |(A1)._______|______      |
     *   |     |    |              |     |
     *   |     |    |              |     |
     *   +-----+----|    BITMAP    |-----+
     *   |     |    |              |     |
     *   |  2  |    |______________.     |
     *   |     |            |        (B2)|
     *   |     |            |     (x2,y2)|
     *   +---- +------------+------------+
     *
     * Example of a bitmap that covers some of the area from cell A1 to cell B2.
     *
     * Based on the width and height of the bitmap we need to calculate 8 vars:
     *     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
     * The width and height of the cells are also variable and have to be taken into
     * account.
     * The values of $col_start and $row_start are passed in from the calling
     * function. The values of $col_end and $row_end are calculated by subtracting
     * the width and height of the bitmap from the width and height of the
     * underlying cells.
     * The vertices are expressed as a percentage of the underlying cell width as
     * follows (rhs values are in pixels):
     *
     *       x1 = X / W *1024
     *       y1 = Y / H *256
     *       x2 = (X-1) / W *1024
     *       y2 = (Y-1) / H *256
     *
     *       Where:  X is distance from the left side of the underlying cell
     *               Y is distance from the top of the underlying cell
     *               W is the width of the cell
     *               H is the height of the cell
     *
     * @access private
     * @note  the SDK incorrectly states that the height should be expressed as a
     *        percentage of 1024.
     * @param integer $col_start Col containing upper left corner of object
     * @param integer $row_start Row containing top left corner of object
     * @param integer $x1        Distance to left side of object
     * @param integer $y1        Distance to top of object
     * @param integer $width     Width of image frame
     * @param integer $height    Height of image frame
     */
    protected function positionImage($col_start, $row_start, $x1, $y1, $width, $height)
    {
        // Initialise end cell to the same as the start cell
        $col_end = $col_start; // Col containing lower right corner of object
        $row_end = $row_start; // Row containing bottom right corner of object

        // Zero the specified offset if greater than the cell dimensions
        if ($x1 >= $this->sizeCol($col_start)) {
            $x1 = 0;
        }
        if ($y1 >= $this->sizeRow($row_start)) {
            $y1 = 0;
        }

        $width = $width + $x1 - 1;
        $height = $height + $y1 - 1;

        // Subtract the underlying cell widths to find the end cell of the image
        while ($width >= $this->sizeCol($col_end)) {
            $width -= $this->sizeCol($col_end);
            $col_end++;
        }

        // Subtract the underlying cell heights to find the end cell of the image
        while ($height >= $this->sizeRow($row_end)) {
            $height -= $this->sizeRow($row_end);
            $row_end++;
        }

        // Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
        // with zero eight or width.
        //
        if ($this->sizeCol($col_start) == 0) {
            return;
        }
        if ($this->sizeCol($col_end) == 0) {
            return;
        }
        if ($this->sizeRow($row_start) == 0) {
            return;
        }
        if ($this->sizeRow($row_end) == 0) {
            return;
        }

        // Convert the pixel values to the percentage value expected by Excel
        $x1 = $x1 / $this->sizeCol($col_start) * 1024;
        $y1 = $y1 / $this->sizeRow($row_start) * 256;
        $x2 = $width / $this->sizeCol($col_end) * 1024; // Distance to right side of object
        $y2 = $height / $this->sizeRow($row_end) * 256; // Distance to bottom of object

        $this->storeObjPicture(
            $col_start,
            $x1,
            $row_start,
            $y1,
            $col_end,
            $x2,
            $row_end,
            $y2
        );
    }

    /**
     * Convert the width of a cell from user's units to pixels. By interpolation
     * the relationship is: y = 7x +5. If the width hasn't been set by the user we
     * use the default value. If the col is hidden we use a value of zero.
     *
     * @access private
     * @param integer $col The column
     * @return integer The width in pixels
     */
    protected function sizeCol($col)
    {
        // Look up the cell value to see if it has been changed
        if (isset($this->col_sizes[$col])) {
            if ($this->col_sizes[$col] == 0) {
                return (0);
            } else {
                return (floor(7 * $this->col_sizes[$col] + 5));
            }
        } else {
            return (64);
        }
    }

    /**
     * Convert the height of a cell from user's units to pixels. By interpolation
     * the relationship is: y = 4/3x. If the height hasn't been set by the user we
     * use the default value. If the row is hidden we use a value of zero. (Not
     * possible to hide row yet).
     *
     * @access private
     * @param integer $row The row
     * @return integer The width in pixels
     */
    protected function sizeRow($row)
    {
        // Look up the cell value to see if it has been changed
        if (isset($this->row_sizes[$row])) {
            if ($this->row_sizes[$row] == 0) {
                return (0);
            } else {
                return (floor(4 / 3 * $this->row_sizes[$row]));
            }
        } else {
            return (17);
        }
    }

    /**
     * Store the OBJ record that precedes an IMDATA record. This could be generalise
     * to support other Excel objects.
     *
     * @access private
     * @param integer $colL Column containing upper left corner of object
     * @param integer $dxL  Distance from left side of cell
     * @param integer $rwT  Row containing top left corner of object
     * @param integer $dyT  Distance from top of cell
     * @param integer $colR Column containing lower right corner of object
     * @param integer $dxR  Distance from right of cell
     * @param integer $rwB  Row containing bottom right corner of object
     * @param integer $dyB  Distance from bottom of cell
     */
    protected function storeObjPicture($colL, $dxL, $rwT, $dyT, $colR, $dxR, $rwB, $dyB)
    {
        $record = 0x005d; // Record identifier
        $length = 0x003c; // Bytes to follow

        $cObj = 0x0001; // Count of objects in file (set to 1)
        $OT = 0x0008; // Object type. 8 = Picture
        $id = 0x0001; // Object ID
        $grbit = 0x0614; // Option flags

        $cbMacro = 0x0000; // Length of FMLA structure
        $Reserved1 = 0x0000; // Reserved
        $Reserved2 = 0x0000; // Reserved

        $icvBack = 0x09; // Background colour
        $icvFore = 0x09; // Foreground colour
        $fls = 0x00; // Fill pattern
        $fAuto = 0x00; // Automatic fill
        $icv = 0x08; // Line colour
        $lns = 0xff; // Line style
        $lnw = 0x01; // Line weight
        $fAutoB = 0x00; // Automatic border
        $frs = 0x0000; // Frame style
        $cf = 0x0009; // Image format, 9 = bitmap
        $Reserved3 = 0x0000; // Reserved
        $cbPictFmla = 0x0000; // Length of FMLA structure
        $Reserved4 = 0x0000; // Reserved
        $grbit2 = 0x0001; // Option flags
        $Reserved5 = 0x0000; // Reserved


        $header = pack("vv", $record, $length);
        $data = pack("V", $cObj);
        $data .= pack("v", $OT);
        $data .= pack("v", $id);
        $data .= pack("v", $grbit);
        $data .= pack("v", $colL);
        $data .= pack("v", $dxL);
        $data .= pack("v", $rwT);
        $data .= pack("v", $dyT);
        $data .= pack("v", $colR);
        $data .= pack("v", $dxR);
        $data .= pack("v", $rwB);
        $data .= pack("v", $dyB);
        $data .= pack("v", $cbMacro);
        $data .= pack("V", $Reserved1);
        $data .= pack("v", $Reserved2);
        $data .= pack("C", $icvBack);
        $data .= pack("C", $icvFore);
        $data .= pack("C", $fls);
        $data .= pack("C", $fAuto);
        $data .= pack("C", $icv);
        $data .= pack("C", $lns);
        $data .= pack("C", $lnw);
        $data .= pack("C", $fAutoB);
        $data .= pack("v", $frs);
        $data .= pack("V", $cf);
        $data .= pack("v", $Reserved3);
        $data .= pack("v", $cbPictFmla);
        $data .= pack("v", $Reserved4);
        $data .= pack("v", $grbit2);
        $data .= pack("V", $Reserved5);

        $this->append($header . $data);
    }

    /**
     * Convert a 24 bit bitmap into the modified internal format used by Windows.
     * This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
     * MSDN library.
     *
     * @access private
     * @param string $bitmap The bitmap to process
     * @return array Array with data and properties of the bitmap
     */
    protected function processBitmap($bitmap)
    {
        // Open file.
        $bmp_fd = @fopen($bitmap, "rb");
        if (!$bmp_fd) {
            throw new \Exception("Couldn't import $bitmap");
        }

        // Slurp the file into a string.
        $data = fread($bmp_fd, filesize($bitmap));

        // Check that the file is big enough to be a bitmap.
        if (strlen($data) <= 0x36) {
            throw new \Exception("$bitmap doesn't contain enough data.\n");
        }

        // The first 2 bytes are used to identify the bitmap.
        $identity = unpack("A2ident", $data);
        if ($identity['ident'] != "BM") {
            throw new \Exception("$bitmap doesn't appear to be a valid bitmap image.\n");
        }

        // Remove bitmap data: ID.
        $data = substr($data, 2);

        // Read and remove the bitmap size. This is more reliable than reading
        // the data size at offset 0x22.
        //
        $size_array = unpack("Vsa", substr($data, 0, 4));
        $size = $size_array['sa'];
        $data = substr($data, 4);
        $size -= 0x36; // Subtract size of bitmap header.
        $size += 0x0C; // Add size of BIFF header.

        // Remove bitmap data: reserved, offset, header length.
        $data = substr($data, 12);

        // Read and remove the bitmap width and height. Verify the sizes.
        $width_and_height = unpack("V2", substr($data, 0, 8));
        $width = $width_and_height[1];
        $height = $width_and_height[2];
        $data = substr($data, 8);
        if ($width > 0xFFFF) {
            throw new \Exception("$bitmap: largest image width supported is 65k.\n");
        }
        if ($height > 0xFFFF) {
            throw new \Exception("$bitmap: largest image height supported is 65k.\n");
        }

        // Read and remove the bitmap planes and bpp data. Verify them.
        $planes_and_bitcount = unpack("v2", substr($data, 0, 4));
        $data = substr($data, 4);
        if ($planes_and_bitcount[2] != 24) { // Bitcount
            throw new \Exception("$bitmap isn't a 24bit true color bitmap.\n");
        }
        if ($planes_and_bitcount[1] != 1) {
            throw new \Exception("$bitmap: only 1 plane supported in bitmap image.\n");
        }

        // Read and remove the bitmap compression. Verify compression.
        $compression = unpack("Vcomp", substr($data, 0, 4));
        $data = substr($data, 4);

        //$compression = 0;
        if ($compression['comp'] != 0) {
            throw new \Exception("$bitmap: compression not supported in bitmap image.\n");
        }

        // Remove bitmap data: data size, hres, vres, colours, imp. colours.
        $data = substr($data, 20);

        // Add the BITMAPCOREHEADER data
        $header = pack("Vvvvv", 0x000c, $width, $height, 0x01, 0x18);
        $data = $header . $data;

        return array($width, $height, $size, $data);
    }

    /**
     * Store the window zoom factor. This should be a reduced fraction but for
     * simplicity we will store all fractions with a numerator of 100.
     *
     * @access private
     */
    protected function storeZoom()
    {
        // If scale is 100 we don't need to write a record
        if ($this->zoom == 100) {
            return;
        }

        $record = 0x00A0; // Record identifier
        $length = 0x0004; // Bytes to follow

        $header = pack("vv", $record, $length);
        $data = pack("vv", $this->zoom, 100);
        $this->append($header . $data);
    }

    /**
     * FIXME: add comments
     */
    public function setValidation($row1, $col1, $row2, $col2, &$validator)
    {
        $this->dv[] = $validator->getData() .
            pack("vvvvv", 1, $row1, $row2, $col1, $col2);
    }

    /**
     * Store the DVAL and DV records.
     *
     * @access private
     */
    protected function storeDataValidity()
    {
        $record = 0x01b2; // Record identifier
        $length = 0x0012; // Bytes to follow

        $grbit = 0x0002; // Prompt box at cell, no cached validity data at DV records
        $horPos = 0x00000000; // Horizontal position of prompt box, if fixed position
        $verPos = 0x00000000; // Vertical position of prompt box, if fixed position
        $objId = 0xffffffff; // Object identifier of drop down arrow object, or -1 if not visible

        $header = pack('vv', $record, $length);
        $data = pack(
            'vVVVV',
            $grbit,
            $horPos,
            $verPos,
            $objId,
            count($this->dv)
        );
        $this->append($header . $data);

        $record = 0x01be; // Record identifier
        foreach ($this->dv as $dv) {
            $length = strlen($dv); // Bytes to follow
            $header = pack("vv", $record, $length);
            $this->append($header . $dv);
        }
    }
}

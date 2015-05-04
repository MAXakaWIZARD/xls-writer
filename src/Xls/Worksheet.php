<?php

namespace Xls;

/**
 * Class for generating Excel Spreadsheets
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 * @category FileFormats
 * @package  Spreadsheet_Excel_Writer
 */

class Worksheet extends BIFFwriter
{
    const ORIENTATION_PORTRAIT = 1;
    const ORIENTATION_LANDSCAPE = 0;

    /**
     * Name of the Worksheet
     * @var string
     */
    protected $name;

    /**
     * Index for the Worksheet
     * @var integer
     */
    public $index;

    /**
     * Reference to the (default) Format object for URLs
     * @var Format
     */
    public $urlFormat;

    /**
     * Reference to the parser used for parsing formulas
     * @var FormulaParser
     */
    protected $formulaParser;

    /**
     * Filehandle to the temporary file for storing data
     * @var resource
     */
    public $fileHandle;

    /**
     * Boolean indicating if we are using a temporary file for storing data
     * @var bool
     */
    public $usingTmpFile;

    /**
     * Maximum number of rows for an Excel spreadsheet (BIFF5)
     * @var integer
     */
    public $xlsRowmax;

    /**
     * Maximum number of columns for an Excel spreadsheet (BIFF5)
     * @var integer
     */
    public $xlsColmax;

    /**
     * Maximum number of characters for a string (LABEL record in BIFF5)
     * @var integer
     */
    public $xlsStrmax;

    /**
     * First row for the DIMENSIONS record
     * @var integer
     */
    public $dimRowmin;

    /**
     * Last row for the DIMENSIONS record
     * @var integer
     */
    public $dimRowmax;

    /**
     * First column for the DIMENSIONS record
     * @var integer
     */
    public $dimColmin;

    /**
     * Last column for the DIMENSIONS record
     * @var integer
     */
    public $dimColmax;

    /**
     * Array containing format information for columns
     * @var array
     */
    public $colInfo;

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
    public $activePane;

    /**
     * Bit specifying if panes are frozen
     * @var integer
     */
    public $frozen;

    /**
     * Bit specifying if the worksheet is selected
     * @var integer
     */
    protected $selected;

    /**
     * The paper size (for printing) (DOCUMENT!!!)
     * @var integer
     */
    public $paperSize;

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
    public $marginHead;

    /**
     * The margin for the footer
     * @var float
     */
    public $marginFoot;

    /**
     * The left margin for the worksheet in inches
     * @var float
     */
    public $marginLeft;

    /**
     * The right margin for the worksheet in inches
     * @var float
     */
    public $marginRight;

    /**
     * The top margin for the worksheet in inches
     * @var float
     */
    public $marginTop;

    /**
     * The bottom margin for the worksheet in inches
     * @var float
     */
    public $marginBottom;

    /**
     * First row to reapeat on each printed page
     * @var integer
     */
    public $titleRowMin;

    /**
     * Last row to reapeat on each printed page
     * @var integer
     */
    public $titleRowMax;

    /**
     * First column to reapeat on each printed page
     * @var integer
     */
    public $titleColMin;

    /**
     * Last column to reapeat on each printed page
     * @var integer
     */
    public $titleColMax;

    /**
     * First row of the area to print
     * @var integer
     */
    public $printRowMin;

    /**
     * Last row to of the area to print
     * @var integer
     */
    public $printRowMax;

    /**
     * First column of the area to print
     * @var integer
     */
    public $printColMin;

    /**
     * Last column of the area to print
     * @var integer
     */
    public $printColMax;

    /**
     * Whether to display RightToLeft.
     * @var integer
     */
    public $arabic;

    /**
     * Whether to use outline.
     * @var integer
     */
    public $outlineOn;

    /**
     * Auto outline styles.
     * @var bool
     */
    public $outlineStyle;

    /**
     * Whether to have outline summary below.
     * @var bool
     */
    public $outlineBelow;

    /**
     * Whether to have outline summary at the right.
     * @var bool
     */
    public $outlineRight;

    /**
     * Outline row level.
     * @var integer
     */
    public $outlineRowLevel;

    /**
     * Whether to fit to page when printing or not.
     * @var bool
     */
    public $fitPage;

    /**
     * Number of pages to fit wide
     * @var integer
     */
    public $fitWidth;

    /**
     * Number of pages to fit high
     * @var integer
     */
    public $fitHeight;

    /**
     * @var SharedStringsTable
     */
    protected $sst;

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * Number of merged cell ranges in actual record
     *
     * @var int $mergedCellsCounter
     */
    public $mergedCellsCounter = 0;

    /**
     * Number of actual mergedcells record
     *
     * @var int $mergedCellsRecord
     */
    public $mergedCellsRecord = 0;

    /**
     * Merged cell ranges
     * @var array
     */
    public $mergedRanges;

    /**
     * Charset encoding currently used when calling writeString()
     * @var string
     */
    public $inputEncoding;

    /**
     * @var int
     */
    protected $offset;

    protected $printGridLines;
    protected $screenGridLines;
    protected $printHeaders;
    protected $hbreaks;
    protected $vbreaks;
    protected $protect;
    protected $password;
    protected $colSizes;
    protected $rowSizes;
    protected $zoom;
    protected $printScale;
    protected $dv;

    /**
     * Constructor
     *
     * @param string $name         The name of the new worksheet
     * @param integer $index        The index of the new worksheet
     * @param Workbook $workbook Parent workbook
     * @param SharedStringsTable $sst Workbook's shared strings table
     * @param Format $urlFormat  The default format for hyperlinks
     * @param FormulaParser $formulaParser The formula parser created for the Workbook
     */
    public function __construct(
        $name,
        $index,
        $workbook,
        $sst,
        $urlFormat,
        $formulaParser
    ) {
        parent::__construct($workbook->getVersion());

        $this->name = $name;
        $this->index = $index;
        $this->workbook = $workbook;
        $this->sst = $sst;
        $this->urlFormat = $urlFormat;
        $this->formulaParser = $formulaParser;

        $this->usingTmpFile = true;
        $this->xlsRowmax = Biff5::MAX_ROWS;
        $this->xlsColmax = Biff5::MAX_COLS;
        $this->xlsStrmax = Biff5::MAX_STR_LENGTH;
        $this->dimRowmin = $this->xlsRowmax + 1;
        $this->dimRowmax = 0;
        $this->dimColmin = $this->xlsColmax + 1;
        $this->dimColmax = 0;
        $this->colInfo = array();
        $this->selection = array(0, 0, 0, 0);
        $this->panes = array();
        $this->activePane = 3;
        $this->frozen = 0;
        $this->selected = 0;

        $this->paperSize = 0x0;
        $this->orientation = self::ORIENTATION_PORTRAIT;
        $this->header = '';
        $this->footer = '';
        $this->hcenter = 0;
        $this->vcenter = 0;
        $this->marginHead = 0.50;
        $this->marginFoot = 0.50;
        $this->marginLeft = 0.75;
        $this->marginRight = 0.75;
        $this->marginTop = 1.00;
        $this->marginBottom = 1.00;

        $this->titleRowMin = null;
        $this->titleRowMax = null;
        $this->titleColMin = null;
        $this->titleColMax = null;
        $this->printRowMin = null;
        $this->printRowMax = null;
        $this->printColMin = null;
        $this->printColMax = null;

        $this->printGridLines = 1;
        $this->screenGridLines = 1;
        $this->printHeaders = 0;

        $this->fitPage = 0;
        $this->fitWidth = 0;
        $this->fitHeight = 0;

        $this->hbreaks = array();
        $this->vbreaks = array();

        $this->protect = 0;
        $this->password = null;

        $this->colSizes = array();
        $this->rowSizes = array();

        $this->zoom = 100;
        $this->printScale = 100;

        $this->outlineRowLevel = 0;
        $this->outlineStyle = 0;
        $this->outlineBelow = 1;
        $this->outlineRight = 1;
        $this->outlineOn = 1;
        $this->arabic = 0;

        $this->mergedRanges = array();

        $this->inputEncoding = '';

        $this->dv = array();

        $this->init();
    }

    /**
     * Open a tmp file to store the majority of the Worksheet data. If this fails,
     * for example due to write permissions, store the data in memory. This can be
     * slow for large files.
     */
    protected function init()
    {
        if (!$this->usingTmpFile) {
            return;
        }

        $this->tmpFile = tempnam($this->tmpDir, "Spreadsheet_Excel_Writer");
        $fh = @fopen($this->tmpFile, "w+b");

        if ($fh === false) {
            // If tmpfile() fails store data in memory
            $this->usingTmpFile = false;
        } else {
            // Store filehandle
            $this->fileHandle = $fh;
        }
    }

    /**
     * Add data to the beginning of the workbook (note the reverse order)
     * and to the end of the workbook.
     *
     * @see Workbook::storeWorkbook()
     * @param array $sheetnames The array of sheetnames from the Workbook this
     *                          worksheet belongs to
     */
    public function close($sheetnames)
    {
        $numSheets = count($sheetnames);

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

        // Prepend the bottom margin
        $this->storeMarginBottom();

        // Prepend the top margin
        $this->storeMarginTop();

        // Prepend the right margin
        $this->storeMarginRight();

        // Prepend the left margin
        $this->storeMarginLeft();

        // Prepend the page vertical centering
        $this->prependRecord('Vcenter', array($this->vcenter));

        // Prepend the page horizontal centering
        $this->prependRecord('Hcenter', array($this->hcenter));

        $this->prependRecord('Footer', array($this->footer));
        $this->prependRecord('Header', array($this->header));

        // Prepend the vertical page breaks
        $this->storeVbreak();

        // Prepend the horizontal page breaks
        $this->storeHbreak();

        // Prepend WSBOOL
        $this->storeWsbool();

        // Prepend GRIDSET
        $this->storeGridset();

        //  Prepend GUTS
        if ($this->isBiff5()) {
            $this->storeGuts();
        }

        // Prepend PRINTGRIDLINES
        $this->storePrintGridlines();

        // Prepend PRINTHEADERS
        $this->storePrintHeaders();

        // Prepend EXTERNSHEET references
        if ($this->isBiff5()) {
            for ($i = $numSheets; $i > 0; $i--) {
                $sheetname = $sheetnames[$i - 1];
                $this->storeExternsheet($sheetname);
            }
        }

        // Prepend the EXTERNCOUNT of external references.
        if ($this->isBiff5()) {
            $this->prependRecord('Externcount', array($numSheets));
        }

        // Prepend the COLINFO records if they exist
        if (!empty($this->colInfo)) {
            $colcount = count($this->colInfo);
            for ($i = 0; $i < $colcount; $i++) {
                $this->storeColinfo($this->colInfo[$i]);
            }
            $this->prependRecord('Defcolwith');
        }

        $this->prependRecord('Bof', array(self::BOF_TYPE_WORKSHEET));

        /*
        * End of prepend. Read upwards from here.
        ***********************************************/

        // Append
        $this->storeWindow2();
        $this->storeZoom();
        if (!empty($this->panes)) {
            $this->storePanes($this->panes);
        }
        $this->appendRecord('Selection', array($this->selection, $this->activePane));
        $this->storeMergedCells();

        if ($this->isBiff8()) {
            $this->storeDataValidity();
        }

        $this->appendRecord('Eof');

        if ($this->tmpFile != '') {
            @unlink($this->tmpFile);
            $this->tmpFile = '';
            $this->usingTmpFile = true;
        }
    }

    /**
     * Retrieve the worksheet name.
     * This is usefull when creating worksheets without a name.
     *
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
            $fh = $this->fileHandle;
            if ($this->usingTmpFile) {
                fseek($fh, 0);
            }
            return $tmp;
        }
        // Return data stored on disk
        if ($this->usingTmpFile) {
            if ($tmp = fread($this->fileHandle, $buffer)) {
                return $tmp;
            }
        }

        // No data to return
        return '';
    }

    /**
     * Sets a merged cell range
     *
     * @param integer $firstRow First row of the area to merge
     * @param integer $firstCol First column of the area to merge
     * @param integer $lastRow  Last row of the area to merge
     * @param integer $lastCol  Last column of the area to merge
     */
    public function setMerge($firstRow, $firstCol, $lastRow, $lastCol)
    {
        if (($lastRow < $firstRow) || ($lastCol < $firstCol)) {
            return;
        }

        $maxRecordRanges = floor(($this->biff->getLimit() - 6) / 8);
        if ($this->mergedCellsCounter >= $maxRecordRanges) {
            $this->mergedCellsRecord++;
            $this->mergedCellsCounter = 0;
        }

        // don't check rowmin, rowmax, etc... because we don't know when this
        // is going to be called
        $this->mergedRanges[$this->mergedCellsRecord][] = array($firstRow, $firstCol, $lastRow, $lastCol);
        $this->mergedCellsCounter++;
    }

    /**
     * Set this worksheet as a selected worksheet,
     * i.e. the worksheet has its tab highlighted.
     *
     */
    public function select()
    {
        $this->selected = 1;
    }

    /**
     *
     */
    public function unselect()
    {
        $this->selected = 0;
    }

    /**
     * Set this worksheet as the active worksheet,
     * i.e. the worksheet that is displayed when the workbook is opened.
     * Also set it as selected.
     *
     */
    public function activate()
    {
        $this->workbook->setActiveSheet($this->index);
    }

    /**
     * Set this worksheet as the first visible sheet.
     * This is necessary when there are a large number of worksheets and the
     * activated worksheet is not visible on the screen.
     *
     */
    public function setFirstSheet()
    {
        $this->workbook->setFirstSheet($this->index);
    }

    /**
     * Set the worksheet protection flag
     * to prevent accidental modification and to
     * hide formulas if the locked and hidden format properties have been set.
     *
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
     * @param integer $firstcol first column on the range
     * @param integer $lastcol  last column on the range
     * @param integer $width    width to set
     * @param mixed $format   The optional XF format to apply to the columns
     * @param integer $hidden   The optional hidden atribute
     * @param integer $level    The optional outline level
     */
    public function setColumn($firstcol, $lastcol, $width, $format = null, $hidden = 0, $level = 0)
    {
        // look for any ranges this might overlap and remove, size or split where necessary
        foreach ($this->colInfo as $key => $colinfo) {
            $existingStart = $colinfo[0];
            $existingEnd = $colinfo[1];
            // if the new range starts within another range
            if ($firstcol > $existingStart
                && $firstcol < $existingEnd
            ) { // trim the existing range to the beginning of the new range
                $this->colInfo[$key][1] = $firstcol - 1;
                // if the new range lies WITHIN the existing range
                if ($lastcol < $existingEnd) { // split the existing range by adding a range after our new range
                    $this->colInfo[] = array(
                        $lastcol + 1,
                        $existingEnd,
                        $colinfo[2],
                        &$colinfo[3],
                        $colinfo[4],
                        $colinfo[5]
                    );
                }
            } // if the new range ends inside an existing range
            elseif ($lastcol > $existingStart
                && $lastcol < $existingEnd
            ) { // trim the existing range to the end of the new range
                $this->colInfo[$key][0] = $lastcol + 1;
            } // if the new range completely overlaps the existing range
            elseif ($firstcol <= $existingStart && $lastcol >= $existingEnd) {
                unset($this->colInfo[$key]);
            }
        } // added by Dan Lynn <dan@spiderweblabs.com) on 2006-12-06
        // regenerate keys
        $this->colInfo = array_values($this->colInfo);
        $this->colInfo[] = array($firstcol, $lastcol, $width, &$format, $hidden, $level);
        // Set width to zero if column is hidden
        $width = ($hidden) ? 0 : $width;
        for ($col = $firstcol; $col <= $lastcol; $col++) {
            $this->colSizes[$col] = $width;
        }
    }

    /**
     * Set which cell or cells are selected in a worksheet
     *
     * @param integer $firstRow    first row in the selected quadrant
     * @param integer $firstColumn first column in the selected quadrant
     * @param integer $lastRow     last row in the selected quadrant
     * @param integer $lastColumn  last column in the selected quadrant
     */
    public function setSelection($firstRow, $firstColumn, $lastRow, $lastColumn)
    {
        $this->selection = array($firstRow, $firstColumn, $lastRow, $lastColumn);
    }

    /**
     * Set panes and mark them as frozen.
     *
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
     */
    public function setPortrait()
    {
        $this->orientation = self::ORIENTATION_PORTRAIT;
    }

    /**
     * Set the page orientation as landscape.
     */
    public function setLandscape()
    {
        $this->orientation = self::ORIENTATION_LANDSCAPE;
    }

    /**
     * Set the paper type. Ex. 1 = US Letter, 9 = A4
     * @param integer $size The type of paper size to use
     */
    public function setPaper($size = 0)
    {
        $this->paperSize = $size;
    }


    /**
     * Set the page header caption and optional margin.
     * @param string $string The header text
     * @param float $margin optional head margin in inches.
     */
    public function setHeader($string, $margin = 0.50)
    {
        if (strlen($string) > Biff5::MAX_STR_LENGTH) {
            return;
        }
        $this->header = $string;
        $this->marginHead = $margin;
    }

    /**
     * Set the page footer caption and optional margin.
     * @param string $string The footer text
     * @param float $margin optional foot margin in inches.
     */
    public function setFooter($string, $margin = 0.50)
    {
        if (strlen($string) > Biff5::MAX_STR_LENGTH) {
            return;
        }
        $this->footer = $string;
        $this->marginFoot = $margin;
    }

    /**
     * Center the page horinzontally.
     * @param integer $center the optional value for centering. Defaults to 1 (center).
     */
    public function centerHorizontally($center = 1)
    {
        $this->hcenter = $center;
    }

    /**
     * Center the page vertically.
     * @param integer $center the optional value for centering. Defaults to 1 (center).
     */
    public function centerVertically($center = 1)
    {
        $this->vcenter = $center;
    }

    /**
     * Set all the page margins to the same value in inches.
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
     * @param float $margin The margin to set in inches
     */
    public function setMarginsLeftRight($margin)
    {
        $this->setMarginLeft($margin);
        $this->setMarginRight($margin);
    }

    /**
     * Set the top and bottom margins to the same value in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginsTopBottom($margin)
    {
        $this->setMarginTop($margin);
        $this->setMarginBottom($margin);
    }

    /**
     * Set the left margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginLeft($margin = 0.75)
    {
        $this->marginLeft = $margin;
    }

    /**
     * Set the right margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginRight($margin = 0.75)
    {
        $this->marginRight = $margin;
    }

    /**
     * Set the top margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginTop($margin = 1.00)
    {
        $this->marginTop = $margin;
    }

    /**
     * Set the bottom margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginBottom($margin = 1.00)
    {
        $this->marginBottom = $margin;
    }

    /**
     * Set the rows to repeat at the top of each printed page.
     * @param integer $firstRow First row to repeat
     * @param integer $lastRow  Last row to repeat. Optional.
     */
    public function repeatRows($firstRow, $lastRow = null)
    {
        $this->titleRowMin = $firstRow;
        if (isset($lastRow)) { //Second row is optional
            $this->titleRowMax = $lastRow;
        } else {
            $this->titleRowMax = $firstRow;
        }
    }

    /**
     * Set the columns to repeat at the left hand side of each printed page.
     * @param integer $firstCol First column to repeat
     * @param integer $lastCol  Last column to repeat. Optional.
     */
    public function repeatColumns($firstCol, $lastCol = null)
    {
        $this->titleColMin = $firstCol;
        if (isset($lastCol)) { // Second col is optional
            $this->titleColMax = $lastCol;
        } else {
            $this->titleColMax = $firstCol;
        }
    }

    /**
     * Set the area of each worksheet that will be printed.
     * @param integer $firstRow First row of the area to print
     * @param integer $firstCol First column of the area to print
     * @param integer $lastRow  Last row of the area to print
     * @param integer $lastCol  Last column of the area to print
     */
    public function printArea($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $this->printRowMin = $firstRow;
        $this->printColMin = $firstCol;
        $this->printRowMax = $lastRow;
        $this->printColMax = $lastCol;
    }


    /**
     * Set the option to hide gridlines on the printed page.
     */
    public function hideGridlines()
    {
        $this->printGridLines = 0;
    }

    /**
     * Set the option to hide gridlines on the worksheet (as seen on the screen).
     */
    public function hideScreenGridlines()
    {
        $this->screenGridLines = 0;
    }

    /**
     * Set the option to print the row and column headers on the printed page.
     * @param integer $print Whether to print the headers or not. Defaults to 1 (print).
     */
    public function printRowColHeaders($print = 1)
    {
        $this->printHeaders = $print;
    }

    /**
     * Set the vertical and horizontal number of pages that will define the maximum area printed.
     * It doesn't seem to work with OpenOffice.
     * @param  integer $width  Maximun width of printed area in pages
     * @param  integer $height Maximun heigth of printed area in pages
     * @see setPrintScale()
     */
    public function fitToPages($width, $height)
    {
        $this->fitPage = 1;
        $this->fitWidth = $width;
        $this->fitHeight = $height;
    }

    /**
     * Store the horizontal page breaks on a worksheet (for printing).
     * The breaks represent the row after which the break is inserted.
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
     * @param integer $scale The optional scale factor. Defaults to 100
     * @throws \Exception
     */
    public function setPrintScale($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Print scale $scale outside range: 10 <= zoom <= 400");
        }

        // Turn off "fit to page" option
        $this->fitPage = 0;

        $this->printScale = floor($scale);
    }

    /**
     * Map to the appropriate write method acording to the token recieved.
     * @param integer $row    The row of the cell we are writing to
     * @param integer $col    The column of the cell we are writing to
     * @param mixed $token  What we are writing
     * @param mixed $format The optional format to apply to the cell
     * @return mixed
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
     * @param integer $row    The row we are writing to
     * @param integer $col    The first col (leftmost col) we are writing to
     * @param array $val    The array of values to write
     * @param mixed $format The optional format to apply to the cell
     * @throws \Exception
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
     * @param integer $row    The first row (uppermost row) we are writing to
     * @param integer $col    The col we are writing to
     * @param array $val    The array of values to write
     * @param mixed $format The optional format to apply to the cell
     * @throws \Exception
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
     * @param mixed $format The optional XF format
     * @return integer The XF record index
     */
    public function xf($format)
    {
        if ($format) {
            return $format->getXfIndex();
        } else {
            return 0x0F;
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
     * @param string $data The binary data to append
     */
    protected function append($data)
    {
        if ($this->usingTmpFile) {
            $data = $this->addContinueIfNeeded($data);
            fwrite($this->fileHandle, $data);
            $this->datasize += strlen($data);
        } else {
            parent::append($data);
        }
    }

    /**
     * Substitute an Excel cell reference in A1 notation for  zero based row and
     * column values in an argument list.
     *
     * Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
     * @param string $cell The cell reference. Or range of cells.
     * @throws \Exception
     * @return array
     */
    protected function substituteCellref($cell)
    {
        $cell = strtoupper($cell);

        // Convert a column range: 'A:A' or 'B:G'
        if (preg_match("/([A-I]?[A-Z]):([A-I]?[A-Z])/", $cell, $match)) {
            list($noUse, $col1) = $this->cellToRowcol($match[1] . '1'); // Add a dummy row
            list($noUse, $col2) = $this->cellToRowcol($match[2] . '1'); // Add a dummy row
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
            $rotatedBits = $value >> 15; // rotated bits beyond bit 15
            $value &= 0x7fff; // first 15 bits
            $password ^= ($value | $rotatedBits);
            $i++;
        }

        $password ^= strlen($plaintext);
        $password ^= 0xCE4B;

        return $password;
    }

    /**
     * This method sets the properties for outlining and grouping. The defaults
     * correspond to Excel's defaults.
     *
     * @param bool $visible
     * @param bool $symbolsBelow
     * @param bool $symbolsRight
     * @param bool $autoStyle
     */
    public function setOutline(
        $visible = true,
        $symbolsBelow = true,
        $symbolsRight = true,
        $autoStyle = false
    ) {
        $this->outlineOn = $visible;
        $this->outlineBelow = $symbolsBelow;
        $this->outlineRight = $symbolsRight;
        $this->outlineStyle = $autoStyle;

        // Ensure this is a boolean vale for Window2
        if ($this->outlineOn) {
            $this->outlineOn = 1;
        }
    }

    /**
     * This method sets the worksheet direction to right-to-left (RTL)
     *
     * @param bool $rtl
     */
    public function setRTL($rtl = true)
    {
        $this->arabic = ($rtl ? 1 : 0);
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

        $xf = $this->xf($format); // The cell format

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xlsRowmax) {
            return (-2);
        }
        if ($col >= $this->xlsColmax) {
            return (-2);
        }
        if ($row < $this->dimRowmin) {
            $this->dimRowmin = $row;
        }
        if ($row > $this->dimRowmax) {
            $this->dimRowmax = $row;
        }
        if ($col < $this->dimColmin) {
            $this->dimColmin = $col;
        }
        if ($col > $this->dimColmax) {
            $this->dimColmax = $col;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $xf);

        $xlDouble = pack("d", $num);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $xlDouble = strrev($xlDouble);
        }

        $this->append($header . $data . $xlDouble);

        return 0;
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * NOTE: there is an Excel 5 defined limit of 255 characters.
     * $format is optional.
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *         -3 : long string truncated to 255 chars
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     * @return integer
     */
    public function writeString($row, $col, $str, $format = null)
    {
        if ($this->isBiff8()) {
            return $this->writeStringBIFF8($row, $col, $str, $format);
        }

        $strlen = strlen($str);
        $record = 0x0204; // Record identifier
        $length = 0x0008 + $strlen; // Bytes to follow
        $xf = $this->xf($format); // The cell format

        $strError = 0;

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xlsRowmax) {
            return (-2);
        }
        if ($col >= $this->xlsColmax) {
            return (-2);
        }
        if ($row < $this->dimRowmin) {
            $this->dimRowmin = $row;
        }
        if ($row > $this->dimRowmax) {
            $this->dimRowmax = $row;
        }
        if ($col < $this->dimColmin) {
            $this->dimColmin = $col;
        }
        if ($col > $this->dimColmax) {
            $this->dimColmax = $col;
        }

        if ($strlen > $this->xlsStrmax) {
            $str = substr($str, 0, $this->xlsStrmax);
            $length = 0x0008 + $this->xlsStrmax;
            $strlen = $this->xlsStrmax;
            $strError = -3;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row, $col, $xf, $strlen);
        $this->append($header . $data . $str);

        return ($strError);
    }

    /**
     * Sets Input Encoding for writing strings
     * @param string $encoding The encoding. Ex: 'UTF-16LE', 'utf-8', 'ISO-859-7'
     * @throws \Exception
     */
    public function setInputEncoding($encoding)
    {
        if ($encoding != 'UTF-16LE' && !function_exists('iconv')) {
            throw new \Exception("Using an input encoding other than UTF-16LE requires PHP support for iconv");
        }

        $this->inputEncoding = $encoding;
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * This is the BIFF8 version (no 255 chars limit).
     * $format is optional.
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *         -3 : long string truncated to 255 chars
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     * @return integer
     */
    public function writeStringBIFF8($row, $col, $str, $format = null)
    {
        if ($this->inputEncoding == 'UTF-16LE') {
            $strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
            $encoding = 0x1;
        } elseif ($this->inputEncoding != '') {
            $str = iconv($this->inputEncoding, 'UTF-16LE', $str);
            $strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
            $encoding = 0x1;
        } else {
            $strlen = strlen($str);
            $encoding = 0x0;
        }
        $record = 0x00FD; // Record identifier
        $length = 0x000A; // Bytes to follow
        $xf = $this->xf($format); // The cell format

        $strError = 0;

        // Check that row and col are valid and store max and min values
        if (!$this->checkRowCol($row, $col)) {
            return -2;
        }

        $str = pack('vC', $strlen, $encoding) . $str;

        $this->sst->add($str);

        $header = pack('vv', $record, $length);
        $data = pack('vvvV', $row, $col, $xf, $this->sst->getStrIdx($str));
        $this->append($header . $data);

        return $strError;
    }

    /**
     * Check row and col before writing to a cell, and update the sheet's
     * dimensions accordingly
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @return boolean true for success, false if row and/or col are grester
     *                 then maximums allowed.
     */
    protected function checkRowCol($row, $col)
    {
        if ($row >= $this->xlsRowmax) {
            return false;
        }
        if ($col >= $this->xlsColmax) {
            return false;
        }
        if ($row < $this->dimRowmin) {
            $this->dimRowmin = $row;
        }
        if ($row > $this->dimRowmax) {
            $this->dimRowmax = $row;
        }
        if ($col < $this->dimColmin) {
            $this->dimColmin = $col;
        }
        if ($col > $this->dimColmax) {
            $this->dimColmax = $col;
        }

        return true;
    }

    /**
     * Writes a note associated with the cell given by the row and column.
     * NOTE records don't have a length limit.
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $note   The note to write
     * @return mixed
     */
    public function writeNote($row, $col, $note)
    {
        $noteLength = strlen($note);
        $record = 0x001C; // Record identifier
        $maxLength = 2048; // Maximun length for a NOTE record
        //$length      = 0x0006 + $note_length;    // Bytes to follow

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xlsRowmax) {
            return (-2);
        }
        if ($col >= $this->xlsColmax) {
            return (-2);
        }
        if ($row < $this->dimRowmin) {
            $this->dimRowmin = $row;
        }
        if ($row > $this->dimRowmax) {
            $this->dimRowmax = $row;
        }
        if ($col < $this->dimColmin) {
            $this->dimColmin = $col;
        }
        if ($col > $this->dimColmax) {
            $this->dimColmax = $col;
        }

        // Length for this record is no more than 2048 + 6
        $length = 0x0006 + min($noteLength, 2048);
        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $noteLength);
        $this->append($header . $data . substr($note, 0, 2048));

        for ($i = $maxLength; $i < $noteLength; $i += $maxLength) {
            $chunk = substr($note, $i, $maxLength);
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
        $xf = $this->xf($format); // The cell format

        // Check that row and col are valid and store max and min values
        if ($row >= $this->xlsRowmax) {
            return (-2);
        }
        if ($col >= $this->xlsColmax) {
            return (-2);
        }
        if ($row < $this->dimRowmin) {
            $this->dimRowmin = $row;
        }
        if ($row > $this->dimRowmax) {
            $this->dimRowmax = $row;
        }
        if ($col < $this->dimColmin) {
            $this->dimColmin = $col;
        }
        if ($col > $this->dimColmax) {
            $this->dimColmax = $col;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $xf);
        $this->append($header . $data);

        return 0;
    }

    /**
     * Write a formula to the specified row and column (zero indexed).
     * The textual representation of the formula is passed to the formula parser
     * which returns a packed binary string.
     *
     * Returns  0 : normal termination
     *         -1 : formula errors (bad formula)
     *         -2 : row or column out of range
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
        $xf = $this->xf($format); // The cell format
        $num = 0x00; // Current value of formula
        $grbit = 0x03; // Option flags
        $unknown = 0x0000; // Must be zero

        // Check that row and col are valid and store max and min values
        if (!$this->checkRowCol($row, $col)) {
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
        $this->formulaParser->parse($formula);

        $formula = $this->formulaParser->toReversePolish();

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
        return $this->writeUrlRange($row, $col, $row, $col, $url, $string, $format);
    }

    /**
     * This is the more general form of writeUrl(). It allows a hyperlink to be
     * written to a range of cells. This function also decides the type of hyperlink
     * to be written. These are either, Web (http, ftp, mailto), Internal
     * (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
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

        if (!$format) {
            $format = $this->urlFormat;
        }

        // Write the visible label using the writeString() method.
        if ($str == '') {
            $str = $url;
        }
        $strError = is_numeric($str)
            ? $this->writeNumber($row1, $col1, $str, $format)
            : $this->writeString(
                $row1,
                $col1,
                $str,
                $format
            );
        if (($strError == -2) || ($strError == -3)) {
            return $strError;
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
        $urlLen = pack("V", strlen($url));

        // Calculate the data length
        $length = 0x34 + strlen($url);

        // Pack the header data
        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row1, $row2, $col1, $col2);

        // Write the packed data
        $this->append(
            $header . $data .
            $unknown1 . $options .
            $unknown2 . $urlLen . $url
        );

        return ($strError);
    }

    /**
     * Used to write internal reference hyperlinks such as "Sheet1!A1".
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

        if (!$format) {
            $format = $this->urlFormat;
        }

        // Strip URL type
        $url = preg_replace('/^internal:/', '', $url);

        // Write the visible label
        if ($str == '') {
            $str = $url;
        }
        $strError = is_numeric($str)
            ? $this->writeNumber($row1, $col1, $str, $format)
            : $this->writeString(
                $row1,
                $col1,
                $str,
                $format
            );
        if (($strError == -2) || ($strError == -3)) {
            return $strError;
        }

        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");

        // Pack the option flags
        $options = pack("V", 0x08);

        // Convert the URL type and to a null terminated wchar string
        $url = join("\0", preg_split("''", $url, -1, PREG_SPLIT_NO_EMPTY));
        $url = $url . "\0\0\0";

        // Pack the length of the URL as chars (not wchars)
        $urlLen = pack("V", floor(strlen($url) / 2));

        // Calculate the data length
        $length = 0x24 + strlen($url);

        // Pack the header data
        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row1, $row2, $col1, $col2);

        // Write the packed data
        $this->append(
            $header . $data .
            $unknown1 . $options .
            $urlLen . $url
        );

        return ($strError);
    }

    /**
     * Write links to external directory names such as 'c:\foo.xls',
     * c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
     *
     * Note: Excel writes some relative links with the $dir_long string. We ignore
     * these cases for the sake of simpler code.
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

        if (!$format) {
            $format = $this->urlFormat;
        }

        // Strip URL type and change Unix dir separator to Dos style (if needed)
        //
        $url = preg_replace('/^external:/', '', $url);
        $url = preg_replace('/\//', "\\", $url);

        // Write the visible label
        if ($str == '') {
            $str = preg_replace('/\#/', ' - ', $url);
        }
        $strError = is_numeric($str)
            ? $this->writeNumber($row1, $col1, $str, $format)
            : $this->writeString(
                $row1,
                $col1,
                $str,
                $format
            );
        if (($strError == -2) || ($strError == -3)) {
            return $strError;
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
        $linkType = 0x01 | $absolute;

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
        $dirLong = $url;
        if (preg_match("/\#/", $url)) {
            $linkType |= 0x08;
        }

        // Pack the link type
        $linkType = pack("V", $linkType);

        // Calculate the up-level dir count e.g.. (..\..\..\ == 3)
        $upCount = preg_match_all("/\.\.\\\/", $dirLong, $useless);
        $upCount = pack("v", $upCount);

        // Store the short dos dir name (null terminated)
        $dirShort = preg_replace("/\.\.\\\/", '', $dirLong) . "\0";

        // Store the long dir name as a wchar string (non-null terminated)
        //$dir_long       = join("\0", split('', $dir_long));
        $dirLong = $dirLong . "\0";

        // Pack the lengths of the dir strings
        $dirShortLen = pack("V", strlen($dirShort));
        $dirLongLen = pack("V", strlen($dirLong));
        $streamLen = pack("V", 0); //strlen($dir_long) + 0x06);

        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack("H*", 'D0C9EA79F9BACE118C8200AA004BA90B02000000');
        $unknown2 = pack("H*", '0303000000000000C000000000000046');
        $unknown3 = pack("H*", 'FFFFADDE000000000000000000000000000000000000000');
        $unknown4 = pack("v", 0x03);

        // Pack the main data stream
        $data = pack("vvvv", $row1, $row2, $col1, $col2) .
            $unknown1 .
            $linkType .
            $unknown2 .
            $upCount .
            $dirShortLen .
            $dirShort .
            $unknown3 .
            $streamLen;
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

        return ($strError);
    }

    /**
     * This method is used to set the height and format for a row.
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
        $ixfe = $this->xf($format); // XF index

        // set _row_sizes so _sizeRow() can use it
        $this->rowSizes[$row] = $height;

        // Use setRow($row, null, $XF) to set XF format without setting height
        if (!is_null($height)) {
            $miyRw = $height * 20; // row height
        } else {
            $miyRw = 0xff; // default row height is 256
        }

        $level = max(0, min($level, 7)); // level should be between 0 and 7
        $this->outlineRowLevel = max($level, $this->outlineRowLevel);

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
     * @throw \Exception
     */
    protected function storeDimensions()
    {
        $record = 0x0200; // Record identifier
        $rowMin = $this->dimRowmin; // First row
        $rowMax = $this->dimRowmax + 1; // Last row plus 1
        $colMin = $this->dimColmin; // First column
        $colMax = $this->dimColmax + 1; // Last column plus 1
        $reserved = 0x0000; // Reserved by Excel

        if ($this->isBiff5()) {
            $length = 0x000A;
            $data = pack(
                "vvvvv",
                $rowMin,
                $rowMax,
                $colMin,
                $colMax,
                $reserved
            );
        } else {
            $length = 0x000E;
            $data = pack(
                "VVvvv",
                $rowMin,
                $rowMax,
                $colMin,
                $colMax,
                $reserved
            );
        }

        $header = pack("vv", $record, $length);
        $this->prepend($header . $data);
    }

    /**
     * Write BIFF record Window2.
     */
    protected function storeWindow2()
    {
        $record = 0x023E; // Record identifier
        if ($this->isBiff5()) {
            $length = 0x000A; // Number of bytes to follow
        } else {
            $length = 0x0012;
        }

        $grbit = 0x00B6; // Option flags
        $rwTop = 0x0000; // Top row visible in window
        $colLeft = 0x0000; // Leftmost column visible in window

        // The options flags that comprise $grbit
        $fDspFmla = 0; // 0 - bit
        $fDspGrid = $this->screenGridLines; // 1
        $fDspRwCol = 1; // 2
        $fFrozen = $this->frozen; // 3
        $fDspZeros = 1; // 4
        $fDefaultHdr = 1; // 5
        $fArabic = $this->arabic; // 6
        $fDspGuts = $this->outlineOn; // 7
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

        if ($this->isBiff5()) {
            $rgbHdr = 0x00000000; // Row/column heading and gridline color
            $data .= pack("V", $rgbHdr);
        } else {
            $rgbHdr = 0x0040; // Row/column heading and gridline color index
            $zoomFactorPageBreak = 0x0000;
            $zoomFactorNormal = 0x0000;
            $data .= pack("vvvvV", $rgbHdr, 0x0000, $zoomFactorPageBreak, $zoomFactorNormal, 0x00000000);
        }
        $this->append($header . $data);
    }

    /**
     * Write BIFF record COLINFO to define column widths
     *
     * Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
     * length record.
     *
     * @param array $colArray This is the only parameter received and is composed of the following:
     *                0 => First formatted column,
     *                1 => Last formatted column,
     *                2 => Col width (8.43 is Excel default),
     *                3 => The optional XF format of the column,
     *                4 => Option flags.
     *                5 => Optional outline level
     */
    protected function storeColinfo($colArray)
    {
        if (isset($colArray[0])) {
            $colFirst = $colArray[0];
        }
        if (isset($colArray[1])) {
            $colLast = $colArray[1];
        }
        if (isset($colArray[2])) {
            $coldx = $colArray[2];
        } else {
            $coldx = 8.43;
        }
        if (isset($colArray[3])) {
            $format = $colArray[3];
        } else {
            $format = 0;
        }
        if (isset($colArray[4])) {
            $grbit = $colArray[4];
        } else {
            $grbit = 0;
        }
        if (isset($colArray[5])) {
            $level = $colArray[5];
        } else {
            $level = 0;
        }
        $record = 0x007D; // Record identifier
        $length = 0x000B; // Number of bytes to follow

        $coldx += 0.72; // Fudge. Excel subtracts 0.72 !?
        $coldx *= 256; // Convert to units of 1/256 of a char

        $ixfe = $this->xf($format);
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
     * Store the MERGEDCELLS record for all ranges of merged cells
     */
    protected function storeMergedCells()
    {
        // if there are no merged cell ranges set, return
        if (count($this->mergedRanges) == 0) {
            return;
        }
        $record = 0x00E5;
        foreach ($this->mergedRanges as $ranges) {
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
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas. A formula references a sheet name via an index. Since we store a
     * reference to all of the external worksheets the EXTERNSHEET index is the same
     * as the worksheet index.
     *
     * @param string $sheetName The name of a external worksheet
     */
    protected function storeExternsheet($sheetName)
    {
        $record = new Record\Externsheet();
        $this->prepend($record->getDataForCurrentSheet($sheetName, $this->name));
    }

    /**
     * Writes the Excel BIFF PANE record.
     * The panes can either be frozen or thawed (unfrozen).
     * Frozen panes are specified in terms of an integer number of rows and columns.
     * Thawed panes are specified in terms of Excel's units for rows and columns.
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

        $this->activePane = $pnnAct; // Used in _storeSelection

        $header = pack("vv", $record, $length);
        $data = pack("vvvvv", $x, $y, $rwTop, $colLeft, $pnnAct);
        $this->append($header . $data);
    }

    /**
     * Store the page setup SETUP BIFF record.
     */
    protected function storeSetup()
    {
        $record = 0x00A1; // Record identifier
        $length = 0x0022; // Number of bytes to follow

        $iPaperSize = $this->paperSize; // Paper size
        $iScale = $this->printScale; // Print scaling factor
        $iPageStart = 0x01; // Starting page number
        $iFitWidth = $this->fitWidth; // Fit to number of pages wide
        $iFitHeight = $this->fitHeight; // Fit to number of pages high
        $grbit = 0x00; // Option flags
        $iRes = 0x0258; // Print resolution
        $iVRes = 0x0258; // Vertical print resolution
        $numHdr = $this->marginHead; // Header Margin
        $numFtr = $this->marginFoot; // Footer Margin
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
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
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
     * Store the LEFTMARGIN BIFF record.
     */
    protected function storeMarginLeft()
    {
        $record = 0x0026; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->marginLeft; // Margin in inches

        $header = pack("vv", $record, $length);

        $data = pack("d", $margin);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Store the RIGHTMARGIN BIFF record.
     */
    protected function storeMarginRight()
    {
        $record = 0x0027; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->marginRight; // Margin in inches

        $header = pack("vv", $record, $length);

        $data = pack("d", $margin);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Store the TOPMARGIN BIFF record.
     */
    protected function storeMarginTop()
    {
        $record = 0x0028; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->marginTop; // Margin in inches

        $header = pack("vv", $record, $length);

        $data = pack("d", $margin);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Store the BOTTOMMARGIN BIFF record.
     */
    protected function storeMarginBottom()
    {
        $record = 0x0029; // Record identifier
        $length = 0x0008; // Bytes to follow

        $margin = $this->marginBottom; // Margin in inches

        $header = pack("vv", $record, $length);

        $data = pack("d", $margin);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $data = strrev($data);
        }

        $this->prepend($header . $data);
    }

    /**
     * Merges the area given by its arguments.
     * This is an Excel97/2000 method. It is required to perform more complicated
     * merging than the normal setAlign('merge').
     * @param integer $firstRow First row of the area to merge
     * @param integer $firstCol First column of the area to merge
     * @param integer $lastRow  Last row of the area to merge
     * @param integer $lastCol  Last column of the area to merge
     */
    public function mergeCells($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $record = 0x00E5; // Record identifier
        $length = 0x0A; // Bytes to follow
        $cref = 1; // Number of refs

        // Swap last row/col for first row/col as necessary
        if ($firstRow > $lastRow) {
            list($firstRow, $lastRow) = array($lastRow, $firstRow);
        }

        if ($firstCol > $lastCol) {
            list($firstCol, $lastCol) = array($lastCol, $firstCol);
        }

        $header = pack("vv", $record, $length);
        $data = pack(
            "vvvvv",
            $cref,
            $firstRow,
            $lastRow,
            $firstCol,
            $lastCol
        );

        $this->append($header . $data);
    }

    /**
     * Write the PRINTHEADERS BIFF record.
     */
    protected function storePrintHeaders()
    {
        $record = 0x002a; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fPrintRwCol = $this->printHeaders; // Boolean flag

        $header = pack("vv", $record, $length);
        $data = pack("v", $fPrintRwCol);
        $this->prepend($header . $data);
    }

    /**
     * Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
     * GRIDSET record.
     */
    protected function storePrintGridlines()
    {
        $record = 0x002b; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fPrintGrid = $this->printGridLines; // Boolean flag

        $header = pack("vv", $record, $length);
        $data = pack("v", $fPrintGrid);
        $this->prepend($header . $data);
    }

    /**
     * Write the GRIDSET BIFF record. Must be used in conjunction with the
     * PRINTGRIDLINES record.
     */
    protected function storeGridset()
    {
        $record = 0x0082; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fGridSet = !($this->printGridLines); // Boolean flag

        $header = pack("vv", $record, $length);
        $data = pack("v", $fGridSet);
        $this->prepend($header . $data);
    }

    /**
     * Write the GUTS BIFF record. This is used to configure the gutter margins
     * where Excel outline symbols are displayed. The visibility of the gutters is
     * controlled by a flag in WSBOOL.
     */
    protected function storeGuts()
    {
        $this->prependRecord('Guts', array($this->colInfo, $this->outlineRowLevel));
    }

    /**
     * Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
     * with the SETUP record.
     *
     *
     */
    protected function storeWsbool()
    {
        $record = 0x0081; // Record identifier
        $length = 0x0002; // Bytes to follow
        $grbit = 0x0000;

        // Set the option flags
        $grbit |= 0x0001; // Auto page breaks visible
        if ($this->outlineStyle) {
            $grbit |= 0x0020; // Auto outline styles
        }
        if ($this->outlineBelow) {
            $grbit |= 0x0040; // Outline summary below
        }
        if ($this->outlineRight) {
            $grbit |= 0x0080; // Outline summary right
        }
        if ($this->fitPage) {
            $grbit |= 0x0100; // Page setup fit to page
        }
        if ($this->outlineOn) {
            $grbit |= 0x0400; // Outline symbols displayed
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $grbit);
        $this->prepend($header . $data);
    }

    /**
     * Write the HORIZONTALPAGEBREAKS BIFF record.
     *
     *
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
        if ($this->isBiff8()) {
            $length = 2 + 6 * $cbrk; // Bytes to follow
        } else {
            $length = 2 + 2 * $cbrk; // Bytes to follow
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $cbrk);

        // Append each page break
        foreach ($breaks as $break) {
            if ($this->isBiff8()) {
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
     *
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
        if ($this->isBiff8()) {
            $length = 2 + 6 * $cbrk; // Bytes to follow
        } else {
            $length = 2 + 2 * $cbrk; // Bytes to follow
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $cbrk);

        // Append each page break
        foreach ($breaks as $break) {
            if ($this->isBiff8()) {
                $data .= pack("vvv", $break, 0x0000, 0xffff);
            } else {
                $data .= pack("v", $break);
            }
        }

        $this->prepend($header . $data);
    }

    /**
     * Set the Biff PROTECT record to indicate that the worksheet is protected.
     */
    protected function storeProtect()
    {
        // Exit unless sheet protection has been specified
        if ($this->protect == 0) {
            return;
        }

        $this->prependRecord('Protect', array($this->protect));
    }

    /**
     * Write the worksheet PASSWORD record.
     */
    protected function storePassword()
    {
        // Exit unless sheet protection and password have been specified
        if ($this->protect == 0 || !isset($this->password)) {
            return;
        }

        $this->prependRecord('Password', array($this->password));
    }


    /**
     * Insert a 24bit bitmap image in a worksheet.
     *
     * @param integer $row     The row we are going to insert the bitmap into
     * @param integer $col     The column we are going to insert the bitmap into
     * @param string $bitmap  The bitmap filename
     * @param integer $x       The horizontal position (offset) of the image inside the cell.
     * @param integer $y       The vertical position (offset) of the image inside the cell.
     * @param integer $scaleX The horizontal scale
     * @param integer $scaleY The vertical scale
     */
    public function insertBitmap($row, $col, $bitmap, $x = 0, $y = 0, $scaleX = 1, $scaleY = 1)
    {
        $bitmapArray = $this->processBitmap($bitmap);
        list($width, $height, $size, $data) = $bitmapArray;

        // Scale the frame of the image.
        $width *= $scaleX;
        $height *= $scaleY;

        // Calculate the vertices of the image and write the OBJ record
        $this->positionImage($col, $row, $x, $y, $width, $height);

        $this->appendRecord('Imdata', array($size, $data));
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
     *
     * @note  the SDK incorrectly states that the height should be expressed as a
     *        percentage of 1024.
     *
     * @param integer $colStart Col containing upper left corner of object
     * @param integer $rowStart Row containing top left corner of object
     * @param integer $x1        Distance to left side of object
     * @param integer $y1        Distance to top of object
     * @param integer $width     Width of image frame
     * @param integer $height    Height of image frame
     */
    protected function positionImage($colStart, $rowStart, $x1, $y1, $width, $height)
    {
        // Initialise end cell to the same as the start cell
        $colEnd = $colStart; // Col containing lower right corner of object
        $rowEnd = $rowStart; // Row containing bottom right corner of object

        // Zero the specified offset if greater than the cell dimensions
        if ($x1 >= $this->sizeCol($colStart)) {
            $x1 = 0;
        }
        if ($y1 >= $this->sizeRow($rowStart)) {
            $y1 = 0;
        }

        $width = $width + $x1 - 1;
        $height = $height + $y1 - 1;

        // Subtract the underlying cell widths to find the end cell of the image
        while ($width >= $this->sizeCol($colEnd)) {
            $width -= $this->sizeCol($colEnd);
            $colEnd++;
        }

        // Subtract the underlying cell heights to find the end cell of the image
        while ($height >= $this->sizeRow($rowEnd)) {
            $height -= $this->sizeRow($rowEnd);
            $rowEnd++;
        }

        // Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
        // with zero eight or width.
        //
        if ($this->sizeCol($colStart) == 0) {
            return;
        }
        if ($this->sizeCol($colEnd) == 0) {
            return;
        }
        if ($this->sizeRow($rowStart) == 0) {
            return;
        }
        if ($this->sizeRow($rowEnd) == 0) {
            return;
        }

        // Convert the pixel values to the percentage value expected by Excel
        $x1 = $x1 / $this->sizeCol($colStart) * 1024;
        $y1 = $y1 / $this->sizeRow($rowStart) * 256;
        $x2 = $width / $this->sizeCol($colEnd) * 1024; // Distance to right side of object
        $y2 = $height / $this->sizeRow($rowEnd) * 256; // Distance to bottom of object

        $this->appendRecord('Obj', array(
            $colStart,
            $x1,
            $rowStart,
            $y1,
            $colEnd,
            $x2,
            $rowEnd,
            $y2
        ));
    }

    /**
     * Convert the width of a cell from user's units to pixels. By interpolation
     * the relationship is: y = 7x +5. If the width hasn't been set by the user we
     * use the default value. If the col is hidden we use a value of zero.
     *
     *
     * @param integer $col The column
     * @return integer The width in pixels
     */
    protected function sizeCol($col)
    {
        // Look up the cell value to see if it has been changed
        if (isset($this->colSizes[$col])) {
            if ($this->colSizes[$col] == 0) {
                return (0);
            } else {
                return (floor(7 * $this->colSizes[$col] + 5));
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
     *
     * @param integer $row The row
     * @return integer The width in pixels
     */
    protected function sizeRow($row)
    {
        // Look up the cell value to see if it has been changed
        if (isset($this->rowSizes[$row])) {
            if ($this->rowSizes[$row] == 0) {
                return (0);
            } else {
                return (floor(4 / 3 * $this->rowSizes[$row]));
            }
        } else {
            return 17;
        }
    }

    /**
     * Convert a 24 bit bitmap into the modified internal format used by Windows.
     * This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
     * MSDN library.
     *
     * @param string $bitmap The bitmap to process
     * @throws \Exception
     * @return array Array with data and properties of the bitmap
     */
    protected function processBitmap($bitmap)
    {
        // Open file.
        $bmpFd = @fopen($bitmap, "rb");
        if (!$bmpFd) {
            throw new \Exception("Couldn't import $bitmap");
        }

        // Slurp the file into a string.
        $data = fread($bmpFd, filesize($bitmap));

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
        $sizeArray = unpack("Vsa", substr($data, 0, 4));
        $size = $sizeArray['sa'];
        $data = substr($data, 4);
        $size -= 0x36; // Subtract size of bitmap header.
        $size += 0x0C; // Add size of BIFF header.

        // Remove bitmap data: reserved, offset, header length.
        $data = substr($data, 12);

        // Read and remove the bitmap width and height. Verify the sizes.
        $widthAndHeight = unpack("V2", substr($data, 0, 8));
        $width = $widthAndHeight[1];
        $height = $widthAndHeight[2];
        $data = substr($data, 8);
        if ($width > 0xFFFF) {
            throw new \Exception("$bitmap: largest image width supported is 65k.\n");
        }
        if ($height > 0xFFFF) {
            throw new \Exception("$bitmap: largest image height supported is 65k.\n");
        }

        // Read and remove the bitmap planes and bpp data. Verify them.
        $planesAndBitcount = unpack("v2", substr($data, 0, 4));
        $data = substr($data, 4);
        if ($planesAndBitcount[2] != 24) { // Bitcount
            throw new \Exception("$bitmap isn't a 24bit true color bitmap.\n");
        }
        if ($planesAndBitcount[1] != 1) {
            throw new \Exception("$bitmap: only 1 plane supported in bitmap image.\n");
        }

        // Read and remove the bitmap compression. Verify compression.
        $compression = unpack("Vcomp", substr($data, 0, 4));
        $data = substr($data, 4);

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
     */
    protected function storeZoom()
    {
        // If scale is 100 we don't need to write a record
        if ($this->zoom == 100) {
            return;
        }

        $this->appendRecord('Zoom', array($this->zoom));
    }

    /**
     * @param $row1
     * @param $col1
     * @param $row2
     * @param $col2
     * @param Validator $validator
     */
    public function setValidation($row1, $col1, $row2, $col2, $validator)
    {
        $this->dv[] = $validator->getData() .
            pack("vvvvv", 1, $row1, $row2, $col1, $col2);
    }

    /**
     * Store the DVAL and DV records.
     */
    protected function storeDataValidity()
    {
        $this->appendRecord('Dval', array($this->dv));

        foreach ($this->dv as $dv) {
            $this->appendRecord('Dv', array($dv));
        }
    }

    /**
     * @return int
     */
    public function getOffset()
    {
        return $this->offset;
    }

    /**
     * @param int $offset
     */
    public function setOffset($offset)
    {
        $this->offset = $offset;
    }

    /**
     * @return bool
     */
    public function isSelected()
    {
        return (bool)$this->selected;
    }
}

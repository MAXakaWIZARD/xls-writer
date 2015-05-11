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

    const PAPER_CUSTOM = 0;
    const PAPER_US_LETTER = 1;
    const PAPER_A3 = 8;
    const PAPER_A4 = 9;
    const PAPER_A5 = 11;

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
    public $colInfo = array();

    /**
     * Array containing the selected area for the worksheet
     * @var array
     */
    public $selection = array(0, 0, 0, 0);

    /**
     * Array containing the panes for the worksheet
     * @var array
     */
    public $panes = array();

    /**
     * The active pane for the worksheet
     * @var integer
     */
    public $activePane;

    /**
     * Bit specifying if panes are frozen
     * @var integer
     */
    public $frozen = 0;

    /**
     * Bit specifying if the worksheet is selected
     * @var integer
     */
    protected $selected = 0;

    /**
     * The paper size (for printing) (DOCUMENT!!!)
     * @var integer
     */
    public $paperSize = self::PAPER_CUSTOM;

    /**
     * Bit specifying paper orientation (for printing). 0 => landscape, 1 => portrait
     * @var integer
     */
    public $orientation;

    /**
     * The page header caption
     * @var string
     */
    public $header = '';

    /**
     * The page footer caption
     * @var string
     */
    public $footer = '';

    /**
     * The horizontal centering value for the page
     * @var integer
     */
    public $hcenter = 0;

    /**
     * The vertical centering value for the page
     * @var integer
     */
    public $vcenter = 0;

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
    public $titleRowMin = null;

    /**
     * Last row to reapeat on each printed page
     * @var integer
     */
    public $titleRowMax = null;

    /**
     * First column to reapeat on each printed page
     * @var integer
     */
    public $titleColMin = null;

    /**
     * Last column to reapeat on each printed page
     * @var integer
     */
    public $titleColMax = null;

    /**
     * First row of the area to print
     * @var integer
     */
    public $printRowMin = null;

    /**
     * Last row to of the area to print
     * @var integer
     */
    public $printRowMax = null;

    /**
     * First column of the area to print
     * @var integer
     */
    public $printColMin = null;

    /**
     * Last column of the area to print
     * @var integer
     */
    public $printColMax = null;

    /**
     * Whether to display RightToLeft.
     * @var integer
     */
    public $arabic = 0;

    /**
     * Whether to use outline.
     * @var integer
     */
    public $outlineOn = 1;

    /**
     * Auto outline styles.
     * @var bool
     */
    public $outlineStyle = 0;

    /**
     * Whether to have outline summary below.
     * @var bool
     */
    public $outlineBelow = 1;

    /**
     * Whether to have outline summary at the right.
     * @var bool
     */
    public $outlineRight = 1;

    /**
     * Outline row level.
     * @var integer
     */
    public $outlineRowLevel = 0;

    /**
     * Whether to fit to page when printing or not.
     * @var bool
     */
    public $fitPage = 0;

    /**
     * Number of pages to fit wide
     * @var integer
     */
    public $fitWidth = 0;

    /**
     * Number of pages to fit high
     * @var integer
     */
    public $fitHeight = 0;

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
    public $mergedRanges = array();

    /**
     * Charset encoding currently used when calling writeString()
     * @var string
     */
    public $inputEncoding = '';

    /**
     * @var int
     */
    protected $offset;

    protected $printGridLines = 1;
    protected $screenGridLines = 1;
    protected $printRowColHeaders = 0;
    protected $hbreaks = array();
    protected $vbreaks = array();
    protected $protect = 0;
    protected $password = null;
    protected $colSizes = array();
    protected $rowSizes = array();

    /**
     * @var float
     */
    protected $zoom = 100;

    /**
     * @var float
     */
    protected $printScale = 100;

    protected $dv = array();

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

        $this->xlsRowmax = Biff5::MAX_ROWS;
        $this->xlsColmax = Biff5::MAX_COLS;
        $this->dimRowmin = $this->xlsRowmax + 1;
        $this->dimRowmax = 0;
        $this->dimColmin = $this->xlsColmax + 1;
        $this->dimColmax = 0;
        $this->activePane = 3;

        $this->orientation = self::ORIENTATION_PORTRAIT;
        $this->marginHead = 0.50;
        $this->marginFoot = 0.50;
        $this->marginLeft = 0.75;
        $this->marginRight = 0.75;
        $this->marginTop = 1.00;
        $this->marginBottom = 1.00;

        $this->init();
    }

    /**
     * Open a tmp file to store the majority of the Worksheet data. If this fails,
     * for example due to write permissions, store the data in memory. This can be
     * slow for large files.
     */
    protected function init()
    {
        $this->tmpFile = tempnam($this->tmpDir, "Spreadsheet_Excel_Writer");
        $this->fileHandle = @fopen($this->tmpFile, "w+b");

        if ($this->fileHandle === false) {
            throw new \Exception('Unable to create temporary file');
        }
    }

    /**
     * Add data to the beginning of the workbook (note the reverse order)
     * and to the end of the workbook.
     *
     * @see Workbook::save()
     *
     * @param array $sheetNames The array of sheetnames from the Workbook this
     *                          worksheet belongs to
     */
    public function close($sheetNames)
    {
        /***********************************************
         * Prepend in reverse order!!
         */

        // Prepend the sheet dimensions
        $this->storeDimensions();

        // Prepend the sheet password
        $this->storePassword();

        // Prepend the sheet protection
        $this->storeProtect();

        $this->prependRecord('Setup', array($this));

        // Prepend the margins
        $this->storeMargins();

        // Prepend the page vertical centering
        $this->prependRecord('Vcenter', array($this->vcenter));

        // Prepend the page horizontal centering
        $this->prependRecord('Hcenter', array($this->hcenter));

        $this->prependRecord('Footer', array($this->footer));
        $this->prependRecord('Header', array($this->header));

        // Prepend page breaks
        $this->storePageBreaks();

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

        if ($this->isBiff5()) {
            $sheetsCount = count($sheetNames);
            // Prepend EXTERNSHEET references
            for ($i = $sheetsCount; $i > 0; $i--) {
                $sheetName = $sheetNames[$i - 1];
                $this->storeExternsheet($sheetName);
            }

            // Prepend the EXTERNCOUNT of external references.
            $this->prependRecord('Externcount', array($sheetsCount));
        }

        $this->storeColinfo();

        $this->prependRecord('Bof', array(self::BOF_TYPE_WORKSHEET));

        /*
        * End of prepend. Read upwards from here.
        ***********************************************/

        // Append
        $this->appendRecord('Window2', array($this));
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
     * Retrieves data from memory in one chunk (prepended data),
     * or from disk in chunks (appended data).
     *
     * @return string The data
     */
    public function getData()
    {
        // Return data stored in memory
        if (isset($this->data)) {
            $tmp = $this->data;
            $this->data = null;

            fseek($this->fileHandle, 0);

            return $tmp;
        }

        if ($tmp = fread($this->fileHandle, 4096)) {
            return $tmp;
        }

        return '';
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

            if ($firstcol > $existingStart
                && $firstcol < $existingEnd
            ) {
                // if the new range starts within another range
                // trim the existing range to the beginning of the new range
                $this->colInfo[$key][1] = $firstcol - 1;

                if ($lastcol < $existingEnd) {
                    // if the new range lies WITHIN the existing range
                    // split the existing range by adding a range after our new range
                    $this->colInfo[] = array(
                        $lastcol + 1,
                        $existingEnd,
                        $colinfo[2],
                        &$colinfo[3],
                        $colinfo[4],
                        $colinfo[5]
                    );
                }
            } elseif ($lastcol > $existingStart
                && $lastcol < $existingEnd
            ) {
                // if the new range ends inside an existing range
                // trim the existing range to the end of the new range
                $this->colInfo[$key][0] = $lastcol + 1;
            } elseif ($firstcol <= $existingStart && $lastcol >= $existingEnd) {
                // if the new range completely overlaps the existing range
                unset($this->colInfo[$key]);
            }
        }

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
     * Set the paper type
     * @param integer $size The type of paper size to use
     */
    public function setPaper($size = self::PAPER_CUSTOM)
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
        $this->header = $this->truncateStringIfNeeded($string);
        $this->marginHead = $margin;
    }

    /**
     * Set the page footer caption and optional margin.
     * @param string $string The footer text
     * @param float $margin optional foot margin in inches.
     */
    public function setFooter($string, $margin = 0.50)
    {
        $this->footer = $this->truncateStringIfNeeded($string);
        $this->marginFoot = $margin;
    }

    /**
     * @param $string
     *
     * @return string
     */
    protected function truncateStringIfNeeded($string)
    {
        if (strlen($string) > Biff5::MAX_STR_LENGTH) {
            $string = substr($string, 0, Biff5::MAX_STR_LENGTH - 1);
        }

        return $string;
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
        $this->setMarginsLeftRight($margin);
        $this->setMarginsTopBottom($margin);
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
        $this->printRowColHeaders = $print;
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
        if ($this->looksLikeNumber($token)) {
            $this->writeNumber($row, $col, $token, $format);
        } elseif ($this->looksLikeUrl($token)) {
            $this->writeUrl($row, $col, $token, '', $format);
        } elseif ($this->looksLikeFormula($token)) {
            $this->writeFormula($row, $col, $token, $format);
        } else {
            $this->writeString($row, $col, $token, $format);
        }
    }

    /**
     * @param $value
     *
     * @return bool
     */
    protected function looksLikeNumber($value)
    {
        return preg_match("/^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/", $value) === 1;
    }

    /**
     * @param $value
     *
     * @return bool
     */
    protected function looksLikeFormula($value)
    {
        return preg_match("/^=/", $value) === 1;
    }

    /**
     * @param $value
     *
     * @return bool
     */
    protected function looksLikeUrl($value)
    {
        return preg_match("/^[fh]tt?p:\/\//", $value) === 1
            || preg_match("/^mailto:/", $value) === 1
            || preg_match("/^(?:in|ex)ternal:/", $value) === 1;
    }

    /**
     * Write an array of values as a row
     * @param integer $row    The row we are writing to
     * @param integer $col    The first col (leftmost col) we are writing to
     * @param array $val    The array of values to write
     * @param mixed $format The optional format to apply to the cell
     * @throws \Exception
     */
    public function writeRow($row, $col, $val, $format = null)
    {
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
    }

    /**
     * Write an array of values as a column
     * @param integer $row    The first row (uppermost row) we are writing to
     * @param integer $col    The col we are writing to
     * @param array $val    The array of values to write
     * @param mixed $format The optional format to apply to the cell
     * @throws \Exception
     */
    public function writeCol($row, $col, $val, $format = null)
    {
        if (is_array($val)) {
            foreach ($val as $v) {
                $this->write($row, $col, $v, $format);
                $row++;
            }
        } else {
            throw new \Exception('$val needs to be an array');
        }
    }

    /**
     * Returns an index to the XF record in the workbook
     * @param mixed $format The optional XF format
     * @return integer The XF record index
     */
    public function xf($format)
    {
        return ($format) ? $format->getXfIndex(): 0x0F;
    }

    /******************************************************************************
     *******************************************************************************
     *
     * Internal methods
     */


    /**
     * Store Worksheet data to a temporary file.
     * @param string $data The binary data to append
     */
    protected function append($data)
    {
        $data = $this->addContinueIfNeeded($data);
        fwrite($this->fileHandle, $data);
        $this->datasize += strlen($data);
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
            return array($col1, $col2);
        }

        // Convert a cell range: 'A1:B7'
        if (preg_match("/\$?([A-I]?[A-Z]\$?\d+):\$?([A-I]?[A-Z]\$?\d+)/", $cell, $match)) {
            list($row1, $col1) = $this->cellToRowcol($match[1]);
            list($row2, $col2) = $this->cellToRowcol($match[2]);
            return array($row1, $col1, $row2, $col2);
        }

        // Convert a cell reference: 'A1' or 'AD2000'
        if (preg_match("/\$?([A-I]?[A-Z]\$?\d+)/", $cell)) {
            list($row1, $col1) = $this->cellToRowcol($match[1]);
            return array($row1, $col1);
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
        return array_slice(Cell::addressToRowCol($cell), 0, 2);
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
        $this->outlineOn = ($visible) ? 1 : 0;
        $this->outlineBelow = $symbolsBelow;
        $this->outlineRight = $symbolsRight;
        $this->outlineStyle = $autoStyle;
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

    /**
     * @param $row
     *
     * @throws \Exception
     */
    protected function validateRowIndex($row)
    {
        if ($row >= $this->xlsRowmax) {
            throw new \Exception('Row index is beyond max row number');
        }
    }

    /**
     * @param $col
     *
     * @throws \Exception
     */
    protected function validateColIndex($col)
    {
        if ($col >= $this->xlsColmax) {
            throw new \Exception('Col index is beyond max col number');
        }
    }


    /**
     * Write a double to the specified row and column (zero indexed).
     * An integer can be written as a double. Excel will display an
     * integer. $format is optional.
     *
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param float $num    The number to write
     * @param mixed $format The optional XF format
     */
    public function writeNumber($row, $col, $num, $format = null)
    {
        $record = 0x0203; // Record identifier
        $length = 0x000E; // Number of bytes to follow

        $xf = $this->xf($format); // The cell format

        $this->checkRowCol($row, $col);

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $xf);

        $xlDouble = pack("d", $num);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $xlDouble = strrev($xlDouble);
        }

        $this->append($header . $data . $xlDouble);
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * NOTE: there is an Excel 5 defined limit of 255 characters.
     * $format is optional.
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     */
    public function writeString($row, $col, $str, $format = null)
    {
        if ($str === '') {
            $this->writeBlank($row, $col, $format);
            return;
        }

        $this->checkRowCol($row, $col);

        if ($this->isBiff8()) {
            $this->writeStringBIFF8($row, $col, $str, $format);
            return;
        }

        $record = 0x0204; // Record identifier
        $xf = $this->xf($format); // The cell format

        $str = $this->truncateStringIfNeeded($str);
        $strlen = strlen($str);

        $length = 0x0008 + $strlen; // Bytes to follow

        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $row, $col, $xf, $strlen);
        $this->append($header . $data . $str);
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
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
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

        $str = pack('vC', $strlen, $encoding) . $str;

        $this->sst->add($str);

        $header = pack('vv', $record, $length);
        $data = pack('vvvV', $row, $col, $xf, $this->sst->getStrIdx($str));
        $this->append($header . $data);
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
        $this->validateRowIndex($row);
        $this->validateColIndex($col);

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
    }

    /**
     * Writes a note associated with the cell given by the row and column.
     * NOTE records don't have a length limit.
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $note   The note to write
     */
    public function writeNote($row, $col, $note)
    {
        $noteLength = strlen($note);
        $record = 0x001C; // Record identifier
        $maxLength = 2048; // Maximun length for a NOTE record

        $this->checkRowCol($row, $col);

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
    }

    /**
     * Write a blank cell to the specified row and column (zero indexed).
     * A blank cell is used to specify formatting without adding a string
     * or a number.
     *
     * A blank cell without a format serves no purpose. Therefore, we don't write
     * a BLANK record unless a format is specified.
     *
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param mixed $format The XF format
     */
    public function writeBlank($row, $col, $format = null)
    {
        // Don't write a blank cell unless it has a format
        if (!$format) {
            return;
        }

        $record = 0x0201; // Record identifier
        $length = 0x0006; // Number of bytes to follow
        $xf = $this->xf($format); // The cell format

        $this->checkRowCol($row, $col);

        $header = pack("vv", $record, $length);
        $data = pack("vvv", $row, $col, $xf);
        $this->append($header . $data);
    }

    /**
     * Write a formula to the specified row and column (zero indexed).
     * The textual representation of the formula is passed to the formula parser
     * which returns a packed binary string.
     *
     * @param integer $row     Zero indexed row
     * @param integer $col     Zero indexed column
     * @param string $formula The formula text string
     * @param mixed $format  The optional XF format
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

        $this->checkRowCol($row, $col);

        // Strip the '=' or '@' sign at the beginning of the formula string
        if (preg_match("/^=/", $formula)) {
            $formula = preg_replace("/(^=)/", "", $formula);
        } elseif (preg_match("/^@/", $formula)) {
            $formula = preg_replace("/(^@)/", "", $formula);
        } else {
            // Error handling
            $this->writeString($row, $col, 'Unrecognised character for formula');
            return;
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
     * @param integer $row    Row
     * @param integer $col    Column
     * @param string $url    URL string
     * @param string $string Alternative label
     * @param mixed $format The cell format
     */
    public function writeUrl($row, $col, $url, $string = '', $format = null)
    {
        // Add start row and col to arg list
        $this->writeUrlRange($row, $col, $row, $col, $url, $string, $format);
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
     */
    protected function writeUrlRange($row1, $col1, $row2, $col2, $url, $string = '', $format = null)
    {
        // Check for internal/external sheet links or default to web link
        if (preg_match('[^internal:]', $url)) {
            $this->writeUrlInternal($row1, $col1, $row2, $col2, $url, $string, $format);
            return;
        }

        if (preg_match('[^external:]', $url)) {
            $this->writeUrlExternal($row1, $col1, $row2, $col2, $url, $string, $format);
            return;
        }

        $this->writeUrlWeb($row1, $col1, $row2, $col2, $url, $string, $format);
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

        if (is_numeric($str)) {
            $this->writeNumber($row1, $col1, $str, $format);
        } else {
            $this->writeString($row1, $col1, $str, $format);
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

        if (is_numeric($str)) {
            $this->writeNumber($row1, $col1, $str, $format);
        } else {
            $this->writeString($row1, $col1, $str, $format);
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

        if (is_numeric($str)) {
            $this->writeNumber($row1, $col1, $str, $format);
        } else {
            $this->writeString($row1, $col1, $str, $format);
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
        $this->prependRecord(
            'Dimensions',
            array(
                $this->dimRowmin,
                $this->dimRowmax + 1,
                $this->dimColmin,
                $this->dimColmax + 1
            )
        );
    }

    /**
     * Prepend the COLINFO records if they exist
     */
    protected function storeColinfo()
    {
        if (count($this->colInfo) == 0) {
            return;
        }

        foreach ($this->colInfo as $item) {
            $this->prependRecord('Colinfo', array($item));
        }

        $this->prependRecord('Defcolwidth');
    }

    /**
     * Store the MERGECELLS record for all ranges of merged cells
     */
    protected function storeMergedCells()
    {
        foreach ($this->mergedRanges as $ranges) {
            $this->appendRecord('MergeCells', array($ranges));
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
        $record = $this->createRecord('Externsheet');
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
        if ($this->isFrozen()) {
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
     * Store the margins records
     */
    protected function storeMargins()
    {
        $this->prependRecord('BottomMargin', array($this->marginBottom));
        $this->prependRecord('TopMargin', array($this->marginTop));
        $this->prependRecord('RightMargin', array($this->marginRight));
        $this->prependRecord('LeftMargin', array($this->marginLeft));
    }

    /**
     * Merges the area given by its arguments.
     * @param integer $firstRow First row of the area to merge
     * @param integer $firstCol First column of the area to merge
     * @param integer $lastRow  Last row of the area to merge
     * @param integer $lastCol  Last column of the area to merge
     * @throws \Exception
     */
    public function mergeCells($firstRow, $firstCol, $lastRow, $lastCol)
    {
        if ($lastRow < $firstRow || $lastCol < $firstCol) {
            throw new \Exception('Invalid merge range');
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
     * Write the PRINTHEADERS BIFF record.
     */
    protected function storePrintHeaders()
    {
        $record = 0x002a; // Record identifier
        $length = 0x0002; // Bytes to follow

        $fPrintRwCol = $this->printRowColHeaders; // Boolean flag

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

        $fGridSet = !$this->printGridLines; // Boolean flag

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
        if ($this->isOutlineOn()) {
            $grbit |= 0x0400; // Outline symbols displayed
        }

        $header = pack("vv", $record, $length);
        $data = pack("v", $grbit);
        $this->prepend($header . $data);
    }

    /**
     *
     */
    protected function storePageBreaks()
    {
        if (!empty($this->vbreaks)) {
            $this->prependRecord('VerticalPagebreaks', array($this->vbreaks));
        }

        if (!empty($this->hbreaks)) {
            $this->prependRecord('HorizontalPagebreaks', array($this->hbreaks));
        }
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
        list($width, $height, $size, $data) = $this->processBitmap($bitmap);

        // Scale the frame of the image.
        $width *= $scaleX;
        $height *= $scaleY;

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
     * @throws \Exception
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

        if ($this->sizeCol($colStart) == 0
            || $this->sizeCol($colEnd) == 0
            || $this->sizeRow($rowStart) == 0
            || $this->sizeRow($rowEnd) == 0
        ) {
            throw new \Exception('Bitmap isn\'t allowed to start or finish in a hidden cell');
        }

        // Convert the pixel values to the percentage value expected by Excel
        $x1 = $x1 / $this->sizeCol($colStart) * 1024;
        $y1 = $y1 / $this->sizeRow($rowStart) * 256;
        $x2 = $width / $this->sizeCol($colEnd) * 1024; // Distance to right side of object
        $y2 = $height / $this->sizeRow($rowEnd) * 256; // Distance to bottom of object

        $this->appendRecord(
            'Obj',
            array(
                $colStart,
                $x1,
                $rowStart,
                $y1,
                $colEnd,
                $x2,
                $rowEnd,
                $y2
            )
        );
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
                return 0;
            } else {
                return floor(7 * $this->colSizes[$col] + 5);
            }
        }

        return 64;
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
                return 0;
            } else {
                return floor(4 / 3 * $this->rowSizes[$row]);
            }
        }

        return 17;
    }

    /**
     * Convert a 24 bit bitmap into the modified internal format used by Windows.
     * This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
     * MSDN library.
     *
     * @param string $filePath The bitmap to process
     *
     * @throws \Exception
     * @return array Array with data and properties of the bitmap
     */
    protected function processBitmap($filePath)
    {
        $bmp = new Bitmap($filePath);

        $size = $bmp->getSize();
        $size -= Bitmap::HEADER_SIZE; // Subtract size of bitmap header.
        $size += Record\BitmapCoreHeader::LENGTH; // Add size of BIFF header.

        $width = $bmp->getWidth();
        $height = $bmp->getHeight();

        $data = $this->getRecord('BitmapCoreHeader', array($width, $height, $bmp->getDataWithoutHeader()));

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

    /**
     * @return bool
     */
    public function isFrozen()
    {
        return (bool)$this->frozen;
    }

    /**
     * @return bool
     */
    public function isArabic()
    {
        return (bool)$this->arabic;
    }

    /**
     * @return bool
     */
    public function isOutlineOn()
    {
        return (bool)$this->outlineOn;
    }

    /**
     * @return bool
     */
    public function areScreenGridLinesVisible()
    {
        return (bool)$this->screenGridLines;
    }

    /**
     * @return int
     */
    public function getOrientation()
    {
        return $this->orientation;
    }

    /**
     * @return int
     */
    public function getPaperSize()
    {
        return $this->paperSize;
    }

    /**
     * @return float
     */
    public function getMarginHead()
    {
        return $this->marginHead;
    }

    /**
     * @return float
     */
    public function getMarginFoot()
    {
        return $this->marginFoot;
    }

    /**
     * @return int
     */
    public function getFitWidth()
    {
        return $this->fitWidth;
    }

    /**
     * @return int
     */
    public function getFitHeight()
    {
        return $this->fitHeight;
    }

    /**
     * @return float
     */
    public function getPrintScale()
    {
        return $this->printScale;
    }
}

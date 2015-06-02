<?php

namespace Xls;

class Worksheet extends BIFFwriter
{
    const ORIENTATION_PORTRAIT = 1;
    const ORIENTATION_LANDSCAPE = 0;

    const PAPER_CUSTOM = 0;
    const PAPER_US_LETTER = 1;
    const PAPER_A3 = 8;
    const PAPER_A4 = 9;
    const PAPER_A5 = 11;

    const STATE_VISIBLE = 0x00;
    const STATE_HIDDEN = 0x01;
    const STATE_VERYHIDDEN = 0x02;

    const TYPE_SHEET = 0x00;
    const TYPE_CHART = 0x02;
    const TYPE_VB_MODULE = 0x06;

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
     * Maximum number of rows for an Excel spreadsheet
     * @var integer
     */
    public $xlsRowmax = Biff8::MAX_ROWS;

    /**
     * Maximum number of columns for an Excel spreadsheet
     * @var integer
     */
    public $xlsColmax = Biff8::MAX_COLS;

    /**
     * First row for the DIMENSIONS record
     * @var integer
     */
    public $dimRowmin;

    /**
     * Last row for the DIMENSIONS record
     * @var integer
     */
    public $dimRowmax = 0;

    /**
     * First column for the DIMENSIONS record
     * @var integer
     */
    public $dimColmin;

    /**
     * Last column for the DIMENSIONS record
     * @var integer
     */
    public $dimColmax = 0;

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
    public $activePane = 3;

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
    public $orientation = self::ORIENTATION_PORTRAIT;

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
    public $marginHead = 0.50;

    /**
     * The margin for the footer
     * @var float
     */
    public $marginFoot = 0.50;

    /**
     * The left margin for the worksheet in inches
     * @var float
     */
    public $marginLeft = 0.75;

    /**
     * The right margin for the worksheet in inches
     * @var float
     */
    public $marginRight = 0.75;

    /**
     * The top margin for the worksheet in inches
     * @var float
     */
    public $marginTop = 1.00;

    /**
     * The bottom margin for the worksheet in inches
     * @var float
     */
    public $marginBottom = 1.00;

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
     * Holds last OBJ record id
     * @var int
     */
    protected $lastObjectId = 0;

    protected $drawings = array();

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
        parent::__construct();

        $this->name = $name;
        $this->index = $index;
        $this->workbook = $workbook;
        $this->sst = $sst;
        $this->urlFormat = $urlFormat;
        $this->formulaParser = $formulaParser;

        $this->dimRowmin = Biff8::MAX_ROWS + 1;
        $this->dimColmin = Biff8::MAX_COLS + 1;
    }

    /**
     * Add data to the beginning of the workbook (note the reverse order)
     * and to the end of the workbook.
     *
     * @see Workbook::save()
     *
     */
    public function close()
    {
        //save previously written data
        $data = $this->getDataAndFlush();

        $this->appendRecord('Bof', array(self::BOF_TYPE_WORKSHEET));

        $this->storeColinfo();
        $this->storePrintHeaders();
        $this->storeGrid();
        $this->appendRecord('Guts', array($this->colInfo, $this->outlineRowLevel));
        $this->appendRecord('WsBool', array($this));
        $this->storePageBreaks();
        $this->storeHeaderAndFooter();
        $this->storeCentering();
        $this->storeMargins();
        $this->appendRecord('PageSetup', array($this));
        $this->storeProtection();
        $this->storeDimensions();

        $this->appendRaw($data);

        $this->appendRecord('Window2', array($this));
        $this->storeZoom();
        $this->storePanes();
        $this->appendRecord('Selection', array($this->selection, $this->activePane));
        $this->storeMergedCells();
        $this->storeDataValidity();

        $this->appendRecord('Eof');
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
     * Retrieves data from memory in one chunk
     *
     * @return string The data
     */
    public function getData()
    {
        return $this->data;
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
        $this->password = $password;
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
    public function setColumnWidth($firstcol, $lastcol, $width, $format = null, $hidden = 0, $level = 0)
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
    public function setSelection($firstRow, $firstColumn, $lastRow = null, $lastColumn = null)
    {
        if (!isset($lastRow)) {
            $lastRow = $firstRow; // Last row in reference
        }

        if (!isset($lastColumn)) {
            $lastColumn = $firstColumn; // Last col in reference
        }

        // Swap last row/col for first row/col as necessary
        if ($firstRow > $lastRow) {
            list($firstRow, $lastRow) = array($lastRow, $firstRow);
        }

        if ($firstColumn > $lastColumn) {
            list($firstColumn, $lastColumn) = array($lastColumn, $firstColumn);
        }

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

        if (!isset($panes[2])) {
            $panes[2] = $panes[0];
        }

        if (!isset($panes[3])) {
            $panes[3] = $panes[1];
        }

        if (!isset($panes[4])) {
            $panes[4] = null;
        }

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

        // Convert Excel's row and column units to the internal units.
        // The default row height is 12.75
        // The default column width is 8.43
        // The following slope and intersection values were interpolated.
        //
        $panes[0] = 20 * $panes[0] + 255;
        $panes[1] = 113.879 * $panes[1] + 390;

        if (!isset($panes[2])) {
            $panes[2] = 0;
        }

        if (!isset($panes[3])) {
            $panes[3] = 0;
        }

        if (!isset($panes[4])) {
            $panes[4] = null;
        }

        $this->panes = $panes;
    }

    /**
     * Writes the Excel BIFF PANE record.
     * The panes can either be frozen or thawed (unfrozen).
     * Frozen panes are specified in terms of an integer number of rows and columns.
     * Thawed panes are specified in terms of Excel's units for rows and columns.
     */
    protected function storePanes()
    {
        if (empty($this->panes)) {
            return;
        }

        $this->activePane = $this->panes[4];
        if (!isset($this->activePane)) {
            $this->activePane = $this->calculateActivePane($this->panes[0], $this->panes[1]);
            $this->panes[4] = $this->activePane;
        }

        $this->appendRecord('Pane', $this->panes);
    }

    /**
     * Determine which pane should be active. There is also the undocumented
     * option to override this should it be necessary: may be removed later.
     * @param $x
     * @param $y
     *
     * @return int|null
     */
    protected function calculateActivePane($x, $y)
    {
        if ($x != 0 && $y != 0) {
            return 0; // Bottom right
        } elseif ($x != 0 && $y == 0) {
            return 1; // Top right
        } elseif ($x == 0 && $y != 0) {
            return 2; // Bottom left
        } elseif ($x == 0 && $y == 0) {
            return 3; // Top left
        }

        return null;
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
        if (StringUtils::countCharacters($string) > Biff8::MAX_STR_LENGTH) {
            $string = StringUtils::substr($string, 0, Biff8::MAX_STR_LENGTH);
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
    public function printRepeatRows($firstRow, $lastRow = null)
    {
        if (!isset($lastRow)) {
            $lastRow = $firstRow;
        }

        $this->titleRowMin = $firstRow;
        $this->titleRowMax = $lastRow;
    }

    /**
     * Set the columns to repeat at the left hand side of each printed page.
     * @param integer $firstCol First column to repeat
     * @param integer $lastCol  Last column to repeat. Optional.
     */
    public function printRepeatColumns($firstCol, $lastCol = null)
    {
        if (!isset($lastCol)) {
            $lastCol = $firstCol;
        }

        $this->titleColMin = $firstCol;
        $this->titleColMax = $lastCol;
    }

    /**
     * Set the area of each worksheet that will be printed.
     * @param integer $firstRow First row of the area to print
     * @param integer $firstCol First column of the area to print
     * @param integer $lastRow  Last row of the area to print
     * @param integer $lastCol  Last column of the area to print
     */
    public function setPrintArea($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $this->printRowMin = $firstRow;
        $this->printColMin = $firstCol;
        $this->printRowMax = $lastRow;
        $this->printColMax = $lastCol;
    }


    /**
     * Set the option to hide gridlines on the printed page.
     */
    public function hidePrintGridlines()
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
        $this->checkRowCol($row, $col);

        $this->appendRecord('Number', array($row, $col, $num, $format));
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

        $this->writeStringSST($row, $col, $str, $format);
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     */
    public function writeStringSST($row, $col, $str, $format = null)
    {
        $strIdx = $this->sst->add($str);

        $this->appendRecord(
            'LabelSst',
            array(
                $row,
                $col,
                $strIdx,
                $format
            )
        );
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
     * @param string $guid comment guid (only for tests)
     */
    public function writeNote($row, $col, $note, $guid = null)
    {
        $this->checkRowCol($row, $col);

        $objId = $this->getNewObjectId();
        $this->drawings[] = $objId;

        $drawing = '0F 00 02 F0 D4 00 00 00 10 00 08 F0 08 00 00 00 02 00 00 00 01 04 00 00 0F 00 03 F0 BC 00 00 00 0F 00 04 F0 28 00 00 00 01 00 09 F0 10 00 00 00 78 FF 77 A0 00 00 00 00 00 00 00 00 88 FF 77 A0 02 00 0A F0 08 00 00 00 00 04 00 00 05 00 00 00 0F 00 04 F0 84 00 00 00 A2 0C 0A F0 08 00 00 00 01 04 00 00 00 0A 00 00 B3 00 0B F0 42 00 00 00 80 00 98 2C C4 7D BF 00 00 00 08 00 58 01 00 00 00 00 80 01 04 00 00 00 81 01 FB F6 D6 00 83 01 FB FE 82 00 8B 01 00 00 4C FF BF 01 10 00 11 00 C0 01 ED EA A1 00 3F 02 03 00 03 00 BF 03 02 00 0A 00 00 00 10 F0 12 00 00 00 03 00 01 00 EC 00 00 00 22 00 02 00 53 03 04 00 66 00 00 00 11 F0 00 00 00 00';
        $this->appendRecord('MsoDrawing', array($drawing));

        $this->appendRecord('ObjComment', array($objId, $guid));
        $this->appendRecord('MsoDrawing', array('00 00 0D F0 00 00 00 00'));
        $this->appendRecord('Txo', array($note));
        $this->appendRecord('Note', array($row, $col, $objId));
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
     * @throws \Exception
     */
    public function writeBlank($row, $col, $format = null)
    {
        if (!$format) {
            // Don't write a blank cell unless it has a format
            return;
        }

        $this->checkRowCol($row, $col);

        $this->appendRecord('Blank', array($row, $col, $format));
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
     * @throws \Exception
     */
    public function writeFormula($row, $col, $formula, $format = null)
    {
        $this->checkRowCol($row, $col);

        // Strip the '=' or '@' sign at the beginning of the formula string
        if (in_array($formula[0], array('=', '@'), true)) {
            $formula = substr($formula, 1);
        } else {
            throw new \Exception('Invalid formula: should start with = or @');
        }

        $formula = $this->formulaParser->getReversePolish($formula);

        $this->appendRecord('Formula', array($row, $col, $formula, $format));
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
     * @param string $label Alternative label
     * @param mixed $format The cell format
     */
    public function writeUrl($row, $col, $url, $label = '', $format = null)
    {
        $this->checkRowCol($row, $col);

        if (preg_match('[^internal:]', $url)
            || strpos($url, '#') === 0
        ) {
            $this->writeUrlInternal($row, $col, $row, $col, $url, $label, $format);
            return;
        }

        if (preg_match('[^external:]', $url)) {
            $this->writeUrlExternal($row, $col, $row, $col, $url, $label, $format);
            return;
        }

        $this->writeUrlWeb($row, $col, $row, $col, $url, $label, $format);
    }

    /**
     * Used to write http, ftp and mailto hyperlinks.
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
        $this->writeUrlLabel($row1, $col1, $url, $str, $format);
        $this->appendRecord('Hyperlink', array($row1, $row2, $col1, $col2, $url));
    }

    /**
     * Used to write internal reference hyperlinks such as "Sheet1!A1".
     *
     * @param integer $row1   Start row
     * @param integer $col1   Start column
     * @param integer $row2   End row
     * @param integer $col2   End column
     * @param string $url    URL string
     * @param string $label    Alternative label
     * @param mixed $format The cell format
     */
    protected function writeUrlInternal($row1, $col1, $row2, $col2, $url, $label, $format = null)
    {
        // Strip URL type
        $url = preg_replace('/^internal:/', '', $url);

        if (strpos($url, '#') === 0) {
            $url = substr($url, 1);
        }

        $this->writeUrlLabel($row1, $col1, $url, $label, $format);
        $this->appendRecord('HyperlinkInternal', array($row1, $row2, $col1, $col2, $url));
    }

    /**
     * Write links to external directory names such as 'c:\foo.xls',
     * c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
     *
     * @param integer $row1   Start row
     * @param integer $col1   Start column
     * @param integer $row2   End row
     * @param integer $col2   End column
     * @param string $url    URL string
     * @param string $label    Alternative label
     * @param mixed $format The cell format
     */
    protected function writeUrlExternal($row1, $col1, $row2, $col2, $url, $label, $format = null)
    {
        // Strip URL type and change Unix dir separator to Dos style (if needed)
        $url = preg_replace('/^external:/', '', $url);
        $url = preg_replace('/\//', "\\", $url);

        $this->writeUrlLabel($row1, $col1, $url, $label, $format);
        $this->appendRecord('HyperlinkExternal', array($row1, $row2, $col1, $col2, $url));
    }

    /**
     * @param      $row1
     * @param      $col1
     * @param      $url
     * @param      $str
     * @param null $format
     */
    protected function writeUrlLabel($row1, $col1, $url, $str, $format = null)
    {
        if (!$format) {
            $format = $this->urlFormat;
        }

        if ($str == '') {
            $str = $url;
        }

        $this->writeString($row1, $col1, $str, $format);
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
    public function setRowHeight($row, $height, $format = null, $hidden = false, $level = 0)
    {
        $this->rowSizes[$row] = $height;

        $level = max(0, min($level, 7)); // level should be between 0 and 7
        $this->outlineRowLevel = max($level, $this->outlineRowLevel);

        $this->appendRecord('Row', array($row, $height, $format, $hidden, $level));
    }

    /**
     * Writes Excel DIMENSIONS to define the area in which there is data.
     * @throw \Exception
     */
    protected function storeDimensions()
    {
        $this->appendRecord(
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
     * Append the COLINFO records if they exist
     */
    protected function storeColinfo()
    {
        $this->appendRecord('Defcolwidth');

        foreach ($this->colInfo as $item) {
            $this->appendRecord('Colinfo', array($item));
        }
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
     * Store the margins records
     */
    protected function storeMargins()
    {
        $this->appendRecord('LeftMargin', array($this->marginLeft));
        $this->appendRecord('RightMargin', array($this->marginRight));
        $this->appendRecord('TopMargin', array($this->marginTop));
        $this->appendRecord('BottomMargin', array($this->marginBottom));
    }

    protected function storeHeaderAndFooter()
    {
        $this->appendRecord('Header', array($this->header));
        $this->appendRecord('Footer', array($this->footer));
    }

    /**
     *
     */
    protected function storeCentering()
    {
        $this->appendRecord('Hcenter', array($this->hcenter));
        $this->appendRecord('Vcenter', array($this->vcenter));
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

        $maxRecordRanges = floor((Biff8::LIMIT - 6) / 8);
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
        $this->appendRecord('PrintHeaders', array($this->printRowColHeaders));
    }

    /**
     * Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
     * GRIDSET record.
     */
    protected function storeGrid()
    {
        $this->appendRecord('PrintGridLines', array($this->printGridLines));
        $this->appendRecord('Gridset', array(!$this->printGridLines));
    }

    /**
     *
     */
    protected function storePageBreaks()
    {
        if (!empty($this->hbreaks)) {
            $this->appendRecord('HorizontalPagebreaks', array($this->hbreaks));
        }

        if (!empty($this->vbreaks)) {
            $this->appendRecord('VerticalPagebreaks', array($this->vbreaks));
        }
    }

    /**
     * Write sheet protection
     */
    protected function storeProtection()
    {
        if (!$this->protect) {
            return;
        }

        $this->appendRecord('Protect', array($this->protect));

        if (isset($this->password)) {
            $this->appendRecord('Password', array($this->password));
        }
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
            'ObjPicture',
            array(
                $this->getNewObjectId(),
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
        $this->dv[] = $validator->getData($row1, $col1, $row2, $col2);
    }

    /**
     * Store the DVAL and DV records.
     */
    protected function storeDataValidity()
    {
        if (count($this->dv) == 0) {
            return;
        }

        $this->appendRecord('DataValidations', array($this->dv));

        foreach ($this->dv as $dv) {
            $this->appendRecord('DataValidation', array($dv));
        }
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

    /**
     * @return bool
     */
    public function isPrintAreaSet()
    {
        return !is_null($this->printRowMin);
    }

    /**
     * @return int
     */
    public function getIndex()
    {
        return $this->index;
    }

    /**
     * @return integer
     */
    protected function getNewObjectId()
    {
        $this->lastObjectId++;

        return $this->lastObjectId;
    }

    /**
     * @return array
     */
    public function getDrawings()
    {
        return $this->drawings;
    }
}

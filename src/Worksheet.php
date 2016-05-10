<?php

namespace Xls;

class Worksheet extends BIFFwriter
{
    const BOF_TYPE = 0x0010;

    const STATE_VISIBLE = 0x00;
    const STATE_HIDDEN = 0x01;
    const STATE_VERYHIDDEN = 0x02;

    const TYPE_SHEET = 0x00;

    /**
     * Name of the Worksheet
     * @var string
     */
    protected $name;

    /**
     * Index for the Worksheet
     * @var integer
     */
    protected $index;

    /**
     * Reference to the (default) Format object for URLs
     * @var Format
     */
    protected $urlFormat;

    /**
     * Reference to the parser used for parsing formulas
     * @var FormulaParser
     */
    protected $formulaParser;

    /**
     * @var Range
     */
    protected $dimensions;

    /**
     * Array containing format information for columns
     * @var array
     */
    protected $colInfo = array();

    /**
     * Array containing format information for rows
     * @var array
     */
    protected $rowInfo = array();

    /**
     * Range containing the selected area for the worksheet
     * @var Range
     */
    protected $selection = null;

    /**
     * Array containing the panes for the worksheet
     * @var array
     */
    protected $panes = array();

    /**
     * The active pane for the worksheet
     * @var integer
     */
    protected $activePane = 3;

    /**
     * Bit specifying if panes are frozen
     * @var integer
     */
    protected $frozen = 0;

    /**
     * Bit specifying if the worksheet is selected
     * @var integer
     */
    protected $selected = 0;

    /**
     * Whether to display RightToLeft.
     * @var integer
     */
    protected $rtl = 0;

    /**
     * Whether to use outline.
     * @var bool
     */
    protected $outlineOn = true;

    /**
     * Auto outline styles.
     * @var bool
     */
    protected $outlineStyle = false;

    /**
     * Whether to have outline summary below.
     * @var bool
     */
    protected $outlineBelow = true;

    /**
     * Whether to have outline summary at the right.
     * @var bool
     */
    protected $outlineRight = true;

    /**
     * Outline row level.
     * @var integer
     */
    protected $outlineRowLevel = 0;

    /**
     * @var SharedStringsTable
     */
    protected $sst;

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * Merged cell ranges
     * @var array
     */
    protected $mergedRanges = array();

    protected $protect = 0;
    protected $password = null;

    protected $validations = array();

    /**
     * Holds last OBJ record id
     * @var int
     */
    protected $lastObjectId = 0;

    protected $drawings = array();

    /**
     * @var PrintSetup
     */
    protected $printSetup;

    protected $screenGridLines = true;

    /**
     * @var float
     */
    protected $zoom = 100;

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
        $this->name = $name;
        $this->index = $index;
        $this->workbook = $workbook;
        $this->sst = $sst;
        $this->urlFormat = $urlFormat;
        $this->formulaParser = $formulaParser;

        $this->dimensions = new Range();
        $this->selection = new Range();
        $this->printSetup = new PrintSetup();
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

        $this->appendRecord('Bof', array(static::BOF_TYPE));

        $this->storeColsAndRowsInfo();
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
        $this->workbook->setActiveSheetIndex($this->index);
    }

    /**
     * Set this worksheet as the first visible sheet.
     * This is necessary when there are a large number of worksheets and the
     * activated worksheet is not visible on the screen.
     *
     */
    public function setFirstSheet()
    {
        $this->workbook->setFirstSheetIndex($this->index);
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
     * Set the width of a single column
     *
     * @param integer $col Column index
     * @param integer $width    width to set
     * @param mixed $format   The optional XF format to apply to the columns
     */
    public function setColumnWidth($col, $width, $format = null)
    {
        $this->colInfo[$col] = array(
            'col' => $col,
            'col2' => $col,
            'width' => $width,
            'format' => $format,
            'hidden' => $width == 0,
            'level' => 0
        );
    }

    /**
     * This method is used to set the height and format for a row.
     * @param integer $row    The row to set
     * @param integer $height Height we are giving to the row.
     *                        Use null to set XF without setting height
     * @param mixed $format XF format we are giving to the row
     */
    public function setRowHeight($row, $height, $format = null)
    {
        $this->rowInfo[$row] = array(
            'row' => $row,
            'height' => $height,
            'format' => $format,
            'hidden' => $height == 0,
            'level' => 0
        );
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
        $this->selection = new Range($firstRow, $firstColumn, $lastRow, $lastColumn);
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

        $this->setPanes($panes);
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
        $panes[0] = 20 * $panes[0] + 255;
        $panes[1] = 113.879 * $panes[1] + 390;

        if (!isset($panes[2])) {
            $panes[2] = 0;
        }

        if (!isset($panes[3])) {
            $panes[3] = 0;
        }

        $this->setPanes($panes);
    }

    protected function setPanes($panes)
    {
        if (!isset($panes[4])) {
            $panes[4] = $this->calculateActivePane($panes[0], $panes[1]);
        }

        $this->activePane = $panes[4];

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
     * Write value to cell
     *
     * @param integer $row    The row of the cell we are writing to
     * @param integer $col    The column of the cell we are writing to
     * @param mixed $value What we are writing
     * @param mixed $format The optional format to apply to the cell
     *
     */
    public function write($row, $col, $value, $format = null)
    {
        if ($this->looksLikeNumber($value)) {
            $this->writeNumber($row, $col, $value, $format);
        } elseif ($this->looksLikeUrl($value)) {
            $this->writeUrl($row, $col, $value, '', $format);
        } elseif ($this->looksLikeFormula($value)) {
            $this->writeFormula($row, $col, $value, $format);
        } else {
            $this->writeString($row, $col, $value, $format);
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
        $this->outlineOn = $visible;
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
        $this->rtl = ($rtl ? 1 : 0);
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
        $this->addCell($row, $col);

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

        $this->addCell($row, $col);

        $this->writeStringSST($row, $col, $str, $format);
    }

    /**
     * Write a string to the specified row and column (zero indexed).
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @param string $str    The string to write
     * @param mixed $format The XF format for the cell
     */
    protected function writeStringSST($row, $col, $str, $format = null)
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
     * Check row and col before writing to a cell, and update the sheet's
     * dimensions accordingly
     * @param integer $row    Zero indexed row
     * @param integer $col    Zero indexed column
     * @return Cell
     */
    protected function addCell($row, $col)
    {
        $cell = new Cell($row, $col);

        $this->dimensions->expand($cell);

        return $cell;
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
        $this->addCell($row, $col);

        $objId = $this->getNewObjectId();
        $this->drawings[] = $objId;

        $drawing = '0F 00 02 F0 D4 00 00 00 10 00 08 F0 08 00 00 00 02 00 00 00 01 04 00 00 0F 00 03 F0 BC 00 00 00';
        $drawing .= ' 0F 00 04 F0 28 00 00 00 01 00 09 F0 10 00 00 00 78 FF 77 A0 00 00 00 00 00 00 00 00 88 FF 77';
        $drawing .= ' A0 02 00 0A F0 08 00 00 00 00 04 00 00 05 00 00 00 0F 00 04 F0 84 00 00 00 A2 0C 0A F0 08 00';
        $drawing .= ' 00 00 01 04 00 00 00 0A 00 00 B3 00 0B F0 42 00 00 00 80 00 98 2C C4 7D BF 00 00 00 08 00 58';
        $drawing .= ' 01 00 00 00 00 80 01 04 00 00 00 81 01 FB F6 D6 00 83 01 FB FE 82 00 8B 01 00 00 4C FF BF 01';
        $drawing .= ' 10 00 11 00 C0 01 ED EA A1 00 3F 02 03 00 03 00 BF 03 02 00 0A 00 00 00 10 F0 12 00 00 00 03';
        $drawing .= ' 00 01 00 EC 00 00 00 22 00 02 00 53 03 04 00 66 00 00 00 11 F0 00 00 00 00';
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

        $this->addCell($row, $col);

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
        $this->addCell($row, $col);

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
        $this->addCell($row, $col);

        $range = new Range($row, $col);

        if (preg_match('[^internal:]', $url)
            || strpos($url, '#') === 0
        ) {
            $this->writeUrlInternal($range, $url, $label, $format);
            return;
        }

        if (preg_match('[^external:]', $url)) {
            $this->writeUrlExternal($range, $url, $label, $format);
            return;
        }

        $this->writeUrlWeb($range, $url, $label, $format);
    }

    /**
     * Used to write http, ftp and mailto hyperlinks.
     * @param Range $range   Cell range
     * @param string $url    URL string
     * @param string $str    Alternative label
     * @param mixed $format The cell format
     */
    protected function writeUrlWeb(Range $range, $url, $str, $format = null)
    {
        $this->writeUrlLabel($range->getStartCell(), $url, $str, $format);
        $this->appendRecord('Hyperlink', array($range, $url));
    }

    /**
     * Used to write internal reference hyperlinks such as "Sheet1!A1".
     *
     * @param Range $range Cell range
     * @param string $url    URL string
     * @param string $label    Alternative label
     * @param mixed $format The cell format
     */
    protected function writeUrlInternal(Range $range, $url, $label, $format = null)
    {
        // Strip URL type
        $url = preg_replace('/^internal:/', '', $url);

        if (strpos($url, '#') === 0) {
            $url = substr($url, 1);
        }

        $this->writeUrlLabel($range->getStartCell(), $url, $label, $format);
        $this->appendRecord('HyperlinkInternal', array($range, $url));
    }

    /**
     * Write links to external directory names such as 'c:\foo.xls',
     * c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
     *
     * @param Range $range Cell range
     * @param string $url    URL string
     * @param string $label    Alternative label
     * @param mixed $format The cell format
     */
    protected function writeUrlExternal(Range $range, $url, $label, $format = null)
    {
        // Strip URL type and change Unix dir separator to Dos style (if needed)
        $url = preg_replace('/^external:/', '', $url);
        $url = preg_replace('/\//', "\\", $url);

        $this->writeUrlLabel($range->getStartCell(), $url, $label, $format);
        $this->appendRecord('HyperlinkExternal', array($range, $url));
    }

    /**
     * @param Cell $cell
     * @param string $url
     * @param string $str
     * @param null $format
     */
    protected function writeUrlLabel(Cell $cell, $url, $str, $format = null)
    {
        if (!$format) {
            $format = $this->urlFormat;
        }

        if ($str == '') {
            $str = $url;
        }

        $this->writeString($cell->getRow(), $cell->getCol(), $str, $format);
    }

    /**
     * Writes Excel DIMENSIONS to define the area in which there is data.
     * @throw \Exception
     */
    protected function storeDimensions()
    {
        $this->appendRecord('Dimensions', array($this->dimensions));
    }

    /**
     * Append the COLINFO and ROW records if they exist
     */
    protected function storeColsAndRowsInfo()
    {
        $this->appendRecord('Defcolwidth');

        foreach ($this->colInfo as $item) {
            $this->appendRecord('Colinfo', array($item));
        }

        foreach ($this->rowInfo as $item) {
            $this->appendRecord('Row', array($item));
        }
    }

    /**
     * Store the MERGECELLS record for all ranges of merged cells
     */
    protected function storeMergedCells()
    {
        if (count($this->mergedRanges) > 0) {
            $this->appendRecord('MergeCells', array($this->mergedRanges));
        }
    }

    /**
     * Store the margins records
     */
    protected function storeMargins()
    {
        $margin = $this->getPrintSetup()->getMargin();

        $this->appendRecord('LeftMargin', array($margin->getLeft()));
        $this->appendRecord('RightMargin', array($margin->getRight()));
        $this->appendRecord('TopMargin', array($margin->getTop()));
        $this->appendRecord('BottomMargin', array($margin->getBottom()));
    }

    protected function storeHeaderAndFooter()
    {
        $printSetup = $this->getPrintSetup();

        $this->appendRecord('Header', array($printSetup->getHeader()));
        $this->appendRecord('Footer', array($printSetup->getFooter()));
    }

    /**
     *
     */
    protected function storeCentering()
    {
        $printSetup = $this->getPrintSetup();

        $this->appendRecord('Hcenter', array((int)$printSetup->isHcenteringOn()));
        $this->appendRecord('Vcenter', array((int)$printSetup->isVcenteringOn()));
    }

    /**
     * Merges the area given by its arguments.
     * @param integer $firstRow First row of the area to merge
     * @param integer $firstCol First column of the area to merge
     * @param integer $lastRow  Last row of the area to merge
     * @param integer $lastCol  Last column of the area to merge
     */
    public function mergeCells($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $range = new Range($firstRow, $firstCol, $lastRow, $lastCol);
        $this->mergedRanges[] = $range;
    }

    /**
     * Write the PRINTHEADERS BIFF record.
     */
    protected function storePrintHeaders()
    {
        $printHeaders = $this->getPrintSetup()->shouldPrintRowColHeaders();
        $this->appendRecord('PrintHeaders', array($printHeaders));
    }

    /**
     * Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
     * GRIDSET record.
     */
    protected function storeGrid()
    {
        $linesVisible = $this->getPrintSetup()->shouldPrintGridLines();
        $this->appendRecord('PrintGridLines', array($linesVisible));
        $this->appendRecord('Gridset', array(!$linesVisible));
    }

    /**
     *
     */
    protected function storePageBreaks()
    {
        $printSetup = $this->getPrintSetup();

        $hbreaks = $printSetup->getHbreaks();
        if (!empty($hbreaks)) {
            $this->appendRecord('HorizontalPagebreaks', array($hbreaks));
        }

        $vbreaks = $printSetup->getVbreaks();
        if (!empty($vbreaks)) {
            $this->appendRecord('VerticalPagebreaks', array($vbreaks));
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
     * @param string $path  The bitmap filename
     * @param integer $x       The horizontal position (offset) of the image inside the cell.
     * @param integer $y       The vertical position (offset) of the image inside the cell.
     * @param integer $scaleX The horizontal scale
     * @param integer $scaleY The vertical scale
     */
    public function insertBitmap($row, $col, $path, $x = 0, $y = 0, $scaleX = 1, $scaleY = 1)
    {
        $bmp = new Bitmap($path);

        $width = $bmp->getWidth();
        $height = $bmp->getHeight();

        // BITMAPCOREINFO
        $data = $this->getRecord('BitmapCoreHeader', array($width, $height));
        $data .= $bmp->getDataWithoutHeader();

        // Scale the frame of the image.
        $width *= $scaleX;
        $height *= $scaleY;

        $this->positionImage($col, $row, $x, $y, $width, $height);

        $this->appendRecord('Imdata', array($data));
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
        $colStartSize = $this->getColWidth($colStart);
        if ($x1 >= $colStartSize) {
            $x1 = 0;
        }

        $rowStartSize = $this->getRowHeight($rowStart);
        if ($y1 >= $rowStartSize) {
            $y1 = 0;
        }

        $width = $width + $x1 - 1;
        $height = $height + $y1 - 1;

        // Subtract the underlying cell widths to find the end cell of the image
        while ($width >= $this->getColWidth($colEnd)) {
            $width -= $this->getColWidth($colEnd);
            $colEnd++;
        }

        // Subtract the underlying cell heights to find the end cell of the image
        while ($height >= $this->getRowHeight($rowEnd)) {
            $height -= $this->getRowHeight($rowEnd);
            $rowEnd++;
        }

        $colEndSize = $this->getColWidth($colEnd);
        $rowEndSize = $this->getRowHeight($rowEnd);

        if ($colStartSize == 0
            || $colEndSize == 0
            || $rowStartSize == 0
            || $rowEndSize == 0
        ) {
            throw new \Exception('Bitmap isn\'t allowed to start or finish in a hidden cell');
        }

        // Convert the pixel values to the percentage value expected by Excel
        $x1 = $x1 / $colStartSize * 1024;
        $y1 = $y1 / $rowStartSize * 256;
        $x2 = $width / $colEndSize * 1024; // Distance to right side of object
        $y2 = $height / $rowEndSize * 256; // Distance to bottom of object

        $this->appendRecord(
            'ObjPicture',
            array(
                $this->getNewObjectId(),
                new Range($rowStart, $colStart, $rowEnd, $colEnd, false),
                new Margin($x1, $x2, $y1, $y2)
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
    protected function getColWidth($col)
    {
        // Look up the cell value to see if it has been changed
        if (isset($this->colInfo[$col])) {
            $width = $this->colInfo[$col]['width'];
            if ($width == 0) {
                return 0;
            }

            return floor(7 * $width + 5);
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
    protected function getRowHeight($row)
    {
        // Look up the cell value to see if it has been changed
        if (isset($this->rowInfo[$row])) {
            $height = $this->rowInfo[$row]['height'];
            if ($height == 0) {
                return 0;
            }

            return floor(4 / 3 * $height);
        }

        return 17;
    }

    /**
     * Store the window zoom factor. This should be a reduced fraction but for
     * simplicity we will store all fractions with a numerator of 100.
     */
    protected function storeZoom()
    {
        // If scale is 100 we don't need to write a record
        $zoom = $this->getZoom();
        if ($zoom == 100) {
            return;
        }

        $this->appendRecord('Zoom', array($zoom));
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
        $range = new Range($row1, $col1, $row2, $col2);
        $this->validations[] = $validator->getData($range);
    }

    /**
     * Store the DVAL and DV records.
     */
    protected function storeDataValidity()
    {
        if (count($this->validations) == 0) {
            return;
        }

        $this->appendRecord('DataValidations', array($this->validations));

        foreach ($this->validations as $dv) {
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
    public function isRtl()
    {
        return (bool)$this->rtl;
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

    /**
     * @return bool
     */
    public function isOutlineOn()
    {
        return $this->outlineOn;
    }

    /**
     * @return boolean
     */
    public function getOutlineStyle()
    {
        return $this->outlineStyle;
    }

    /**
     * @return boolean
     */
    public function getOutlineBelow()
    {
        return $this->outlineBelow;
    }

    /**
     * @return boolean
     */
    public function getOutlineRight()
    {
        return $this->outlineRight;
    }

    /**
     * @return PrintSetup
     */
    public function getPrintSetup()
    {
        return $this->printSetup;
    }

    /**
     * Set the option to hide gridlines on the worksheet (as seen on the screen).
     *
     * @param bool $visible
     *
     * @return Worksheet
     */
    public function showGridlines($visible = true)
    {
        $this->screenGridLines = $visible;

        return $this;
    }

    /**
     * Set the worksheet zoom factor.
     *
     * @param integer $percents The zoom factor
     *
     * @throws \Exception
     * @return Worksheet
     */
    public function setZoom($percents = 100)
    {
        // Confine the scale to Excel's range
        if ($percents < 10 || $percents > 400) {
            throw new \Exception("Zoom factor $percents outside range: 10 <= zoom <= 400");
        }

        $this->zoom = floor($percents);

        return $this;
    }

    /**
     * @return bool
     */
    public function areGridLinesVisible()
    {
        return (bool)$this->screenGridLines;
    }

    /**
     * @return float
     */
    public function getZoom()
    {
        return $this->zoom;
    }
}

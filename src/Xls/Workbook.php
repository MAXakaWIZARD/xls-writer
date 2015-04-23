<?php

namespace Xls;

use Xls\OLE\OLE;
use Xls\OLE\PPS;

/**
 * Class for generating Excel Spreadsheets
*/

class Workbook extends BIFFwriter
{
    /**
     * Filename for the Workbook
     * @var string
     */
    public $filename;

    /**
     * Formula parser
     * @var object Parser
     */
    public $parser;

    /**
     * Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904)
     * @var integer
     */
    public $f1904;

    /**
     * The active worksheet of the workbook (0 indexed)
     * @var integer
     */
    public $activeSheet;

    /**
     * 1st displayed worksheet in the workbook (0 indexed)
     * @var integer
     */
    public $firstSheet;

    /**
     * Number of workbook tabs selected
     * @var integer
     */
    public $selected;

    /**
     * Index for creating adding new formats to the workbook
     * @var integer
     */
    public $xfIndex;

    /**
     * Flag for preventing close from being called twice.
     * @var boolean
     * @see close()
     */
    public $fileClosed;

    /**
     * The BIFF file size for the workbook.
     * @var integer
     * @see _calcSheetOffsets()
     */
    public $biffSize;

    /**
     * The default sheetname for all sheets created.
     * @var string
     */
    public $sheetName;

    /**
     * The default XF format.
     * @var object Format
     */
    public $tmpFormat;

    /**
     * Array containing references to all of this workbook's worksheets
     * @var Worksheet[]
     */
    public $worksheets;

    /**
     * Array of sheetnames for creating the EXTERNSHEET records
     * @var array
     */
    public $sheetNames;

    /**
     * Array containing references to all of this workbook's formats
     * @var Format[]
     */
    public $formats;

    /**
     * Array containing the colour palette
     * @var array
     */
    public $palette;

    /**
     * The default format for URLs.
     * @var object Format
     */
    public $urlFormat;

    /**
     * The country code used for localization
     * @var integer
     */
    public $countryCode;

    /**
     * @var int
     */
    public $stringSizeinfo;

    /**
     * number of bytes for sizeinfo of strings
     * @var integer
     */
    public $stringSizeinfoSize;

    /**
     * @var
     */
    protected $blockSizes;

    /**
     * @var int
     */
    protected $creationTimestamp;

    /**
     * Class constructor
     *
     * @param string $filename filename for storing the workbook. "-" for writing to stdout.
     * @param int $version
     */
    public function __construct(
        $filename,
        $version = Biff5::VERSION
    ) {
        parent::__construct($version);

        $this->filename = $filename;
        $this->parser = new Parser($this->byteOrder, $this->version);
        $this->f1904 = 0;
        $this->selected = 0;
        $this->xfIndex = 16; // 15 style XF's and 1 cell XF.
        $this->fileClosed = false;
        $this->biffSize = 0;
        $this->sheetName = 'Sheet';

        $this->activeSheet = 0;
        $this->firstSheet = 0;
        $this->worksheets = array();
        $this->sheetNames = array();

        $this->formats = array();
        $this->palette = array();
        $this->countryCode = -1;
        $this->stringSizeinfo = 3;

        $this->tmpFormat = new Format($this->version);

        // Add the default format for hyperlinks
        $this->urlFormat = $this->addFormat(array('color' => 'blue', 'underline' => 1));

        $this->strTotal = 0;
        $this->strUnique = 0;
        $this->strTable = array();
        $this->setPaletteXl97();

        $this->creationTimestamp = time();
    }

    /**
     * @return int
     */
    public function getCreationTimestamp()
    {
        return $this->creationTimestamp;
    }

    /**
     * @param int $creationTime
     */
    public function setCreationTimestamp($creationTime)
    {
        $this->creationTimestamp = $creationTime;
    }

    /**
     * Calls finalization methods.
     * This method should always be the last one to be called on every workbook
     *
     * @throws \Exception
     * @return boolean true on success.
     */
    public function close()
    {
        if ($this->fileClosed) {
            // Prevent close() from being called twice.
            return true;
        }

        $this->storeWorkbook();
        $this->fileClosed = true;

        return true;
    }

    /**
     * An accessor for the _worksheets[] array.
     * Returns an array of the worksheet objects in a workbook
     *
     * @return array
     */
    public function worksheets()
    {
        return $this->worksheets;
    }

    /**
     * Set the country identifier for the workbook
     *
     * @param integer $code Is the international calling country code for the
     *                      chosen country.
     */
    public function setCountry($code)
    {
        $this->countryCode = $code;
    }

    /**
     * Add a new worksheet to the Excel workbook.
     * If no name is given the name of the worksheet will be Sheeti$i, with
     * $i in [1..].
     *
     * @param string $name the optional name of the worksheet
     * @throws \Exception
     * @return Worksheet
     */
    public function addWorksheet($name = '')
    {
        $index = count($this->worksheets);

        if ($name == '') {
            $name = $this->sheetName . ($index + 1);
        }

        $maxLen = $this->biff->getMaxSheetNameLength();
        if (strlen($name) > $maxLen) {
            throw new \Exception(
                "Sheetname $name must be <= $maxLen chars"
            );
        }

        if ($this->isBiff8() && function_exists('iconv')) {
            $name = iconv('UTF-8', 'UTF-16LE', $name);
        }

        if (in_array($name, $this->sheetNames, true)) {
            throw new \Exception("Worksheet '$name' already exists");
        }

        $worksheet = new Worksheet(
            $this->version,
            $name,
            $index,
            $this->activeSheet,
            $this->firstSheet,
            $this->strTotal,
            $this->strUnique,
            $this->strTable,
            $this->urlFormat,
            $this->parser
        );

        $this->worksheets[$index] = $worksheet; // Store ref for iterator
        $this->sheetNames[$index] = $name; // Store EXTERNSHEET names
        $this->parser->setExtSheet($name, $index); // Register worksheet name with parser

        return $worksheet;
    }

    /**
     * Add a new format to the Excel workbook.
     * Also, pass any properties to the Format constructor.
     *
     * @param array $properties array with properties for initializing the format.
     * @return Format reference to an Excel Format
     */
    public function addFormat($properties = array())
    {
        $format = new Format($this->version, $this->xfIndex, $properties);
        $this->xfIndex += 1;
        $this->formats[] = & $format;

        return $format;
    }

    /**
     * Create new validator.
     *
     * @return Validator reference to a Validator
     */
    public function addValidator()
    {
        return new Validator($this->parser);
    }

    /**
     * Change the RGB components of the elements in the colour palette.
     *
     * @param integer $index colour index
     * @param integer $red   red RGB value [0-255]
     * @param integer $green green RGB value [0-255]
     * @param integer $blue  blue RGB value [0-255]
     * @throws \Exception
     *
     * @return integer The palette index for the custom color
     */
    public function setCustomColor($index, $red, $green, $blue)
    {
        // Match a HTML #xxyyzz style parameter
        /*if (defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
            @_ = ($_[0], hex $1, hex $2, hex $3);
        }*/

        // Check that the colour index is the right range
        if ($index < 8 || $index > 64) {
            throw new \Exception("Color index $index outside range: 8 <= index <= 64");
        }

        // Check that the colour components are in the right range
        if (($red < 0 || $red > 255)
            || ($green < 0 || $green > 255)
            || ($blue < 0 || $blue > 255)
        ) {
            throw new \Exception("Color component outside range: 0 <= color <= 255");
        }

        $index -= 8; // Adjust colour index (wingless dragonfly)

        // Set the RGB value
        $this->palette[$index] = array($red, $green, $blue, 0);

        return ($index + 8);
    }

    /**
     * Sets the colour palette to the Excel 97+ default.
     */
    protected function setPaletteXl97()
    {
        $this->palette = array(
            array(0x00, 0x00, 0x00, 0x00), // 8
            array(0xff, 0xff, 0xff, 0x00), // 9
            array(0xff, 0x00, 0x00, 0x00), // 10
            array(0x00, 0xff, 0x00, 0x00), // 11
            array(0x00, 0x00, 0xff, 0x00), // 12
            array(0xff, 0xff, 0x00, 0x00), // 13
            array(0xff, 0x00, 0xff, 0x00), // 14
            array(0x00, 0xff, 0xff, 0x00), // 15
            array(0x80, 0x00, 0x00, 0x00), // 16
            array(0x00, 0x80, 0x00, 0x00), // 17
            array(0x00, 0x00, 0x80, 0x00), // 18
            array(0x80, 0x80, 0x00, 0x00), // 19
            array(0x80, 0x00, 0x80, 0x00), // 20
            array(0x00, 0x80, 0x80, 0x00), // 21
            array(0xc0, 0xc0, 0xc0, 0x00), // 22
            array(0x80, 0x80, 0x80, 0x00), // 23
            array(0x99, 0x99, 0xff, 0x00), // 24
            array(0x99, 0x33, 0x66, 0x00), // 25
            array(0xff, 0xff, 0xcc, 0x00), // 26
            array(0xcc, 0xff, 0xff, 0x00), // 27
            array(0x66, 0x00, 0x66, 0x00), // 28
            array(0xff, 0x80, 0x80, 0x00), // 29
            array(0x00, 0x66, 0xcc, 0x00), // 30
            array(0xcc, 0xcc, 0xff, 0x00), // 31
            array(0x00, 0x00, 0x80, 0x00), // 32
            array(0xff, 0x00, 0xff, 0x00), // 33
            array(0xff, 0xff, 0x00, 0x00), // 34
            array(0x00, 0xff, 0xff, 0x00), // 35
            array(0x80, 0x00, 0x80, 0x00), // 36
            array(0x80, 0x00, 0x00, 0x00), // 37
            array(0x00, 0x80, 0x80, 0x00), // 38
            array(0x00, 0x00, 0xff, 0x00), // 39
            array(0x00, 0xcc, 0xff, 0x00), // 40
            array(0xcc, 0xff, 0xff, 0x00), // 41
            array(0xcc, 0xff, 0xcc, 0x00), // 42
            array(0xff, 0xff, 0x99, 0x00), // 43
            array(0x99, 0xcc, 0xff, 0x00), // 44
            array(0xff, 0x99, 0xcc, 0x00), // 45
            array(0xcc, 0x99, 0xff, 0x00), // 46
            array(0xff, 0xcc, 0x99, 0x00), // 47
            array(0x33, 0x66, 0xff, 0x00), // 48
            array(0x33, 0xcc, 0xcc, 0x00), // 49
            array(0x99, 0xcc, 0x00, 0x00), // 50
            array(0xff, 0xcc, 0x00, 0x00), // 51
            array(0xff, 0x99, 0x00, 0x00), // 52
            array(0xff, 0x66, 0x00, 0x00), // 53
            array(0x66, 0x66, 0x99, 0x00), // 54
            array(0x96, 0x96, 0x96, 0x00), // 55
            array(0x00, 0x33, 0x66, 0x00), // 56
            array(0x33, 0x99, 0x66, 0x00), // 57
            array(0x00, 0x33, 0x00, 0x00), // 58
            array(0x33, 0x33, 0x00, 0x00), // 59
            array(0x99, 0x33, 0x00, 0x00), // 60
            array(0x99, 0x33, 0x66, 0x00), // 61
            array(0x33, 0x33, 0x99, 0x00), // 62
            array(0x33, 0x33, 0x33, 0x00), // 63
        );
    }

    /**
     * Assemble worksheets into a workbook and send the BIFF data to an OLE
     * storage.
     *
     * @throws \Exception
     * @return boolean true on success.
     */
    protected function storeWorkbook()
    {
        if (count($this->worksheets) == 0) {
            return true;
        }

        // Ensure that at least one worksheet has been selected.
        if ($this->activeSheet == 0) {
            $this->worksheets[0]->select();
        }

        // Calculate the number of selected worksheet tabs and call the finalization
        // methods for each worksheet
        foreach ($this->worksheets as $sheet) {
            if ($sheet->isSelected()) {
                $this->selected++;
            }
            $sheet->close($this->sheetNames);
        }

        // Add Workbook globals
        $this->storeBof(self::BOF_TYPE_WORKBOOK);
        $this->storeCodepage();

        if ($this->isBiff8()) {
            $this->storeWindow1();
            $this->storeNames();
        } else {
            $this->storeExterns();
            $this->storeNames();
            $this->storeWindow1();
        }

        $this->storeDatemode();
        $this->storeAllFonts();
        $this->storeAllNumFormats();
        $this->storeAllXfs();
        $this->storeAllStyles();
        $this->storePalette();
        $this->calcSheetOffsets();

        foreach ($this->worksheets as $sheet) {
            $this->storeBoundsheet($sheet->getName(), $sheet->getOffset());
        }

        if ($this->countryCode != -1) {
            $this->storeCountry();
        }

        if ($this->isBiff8()) {
            /* TODO: store external SUPBOOK records and XCT and CRN records
            * in case of external references for BIFF8
            */
            //$this->storeSupbookInternal();
            //$this->storeExternsheetBiff8();
            $this->storeSharedStringsTable();
        }

        // End Workbook globals
        $this->storeEof();

        // Store the workbook in an OLE container
        $this->storeOLEFile();

        return true;
    }

    /**
     * Store the workbook in an OLE container
     *
     * @throws \Exception
     * @return boolean true on success.
     */
    protected function storeOLEFile()
    {
        $ole = new PPS\File(OLE::asc2Ucs($this->biff->getWorkbookName()));

        $ole->init();
        $ole->append($this->data);

        foreach ($this->worksheets as $sheet) {
            while ($tmp = $sheet->getData()) {
                $ole->append($tmp);
            }
        }

        $root = new PPS\Root(
            $this->getCreationTimestamp(),
            $this->getCreationTimestamp(),
            array($ole)
        );

        $root->save($this->filename);

        return true;
    }

    /**
     * Calculate offsets for Worksheet BOF records.
     */
    protected function calcSheetOffsets()
    {
        $boundsheetLength = $this->biff->getBoundsheetLength();
        $EOF = 4;
        $offset = $this->datasize;

        if ($this->isBiff8()) {
            // add the length of the SST
            /* TODO: check if this works for a lot of strings (> 8224 bytes) */
            $offset += $this->calculateSharedStringsSizes();
            if ($this->countryCode != -1) {
                $offset += 8; // adding COUNTRY record
            }
            // add the lenght of SUPBOOK, EXTERNSHEET and NAME records
            //$offset += 8; // TODO: calculate real value when storing the records
        }

        // add the length of the BOUNDSHEET records
        foreach ($this->worksheets as $sheet) {
            $offset += $boundsheetLength + strlen($sheet->getName());
        }
        $offset += $EOF;

        foreach ($this->worksheets as $sheet) {
            $sheet->setOffset($offset);
            $offset += $sheet->getDataSize();
        }

        $this->biffSize = $offset;
    }

    /**
     * Store the Excel FONT records.
     */
    protected function storeAllFonts()
    {
        // tmp_format is added by the constructor. We use this to write the default XF's
        $font = $this->tmpFormat->getFontRecord();

        // Note: Fonts are 0-indexed. According to the SDK there is no index 4,
        // so the following fonts are 0, 1, 2, 3, 5
        for ($i = 1; $i <= 5; $i++) {
            $this->append($font);
        }

        // Iterate through the XF objects and write a FONT record if it isn't the
        // same as the default FONT and if it hasn't already been used.
        $fonts = array();
        $index = 6; // The first user defined FONT
        $key = $this->tmpFormat->getFontKey(); // The default font from _tmp_format
        $fonts[$key] = 0; // Index of the default font

        foreach ($this->formats as $format) {
            $key = $format->getFontKey();
            if (isset($fonts[$key])) {
                // FONT has already been used
                $format->fontIndex = $fonts[$key];
            } else {
                // Add a new FONT record
                $fonts[$key] = $index;
                $format->fontIndex = $index;
                $index++;
                $this->append($format->getFontRecord());
            }
        }
    }

    /**
     * Store user defined numerical formats i.e. FORMAT records
     */
    protected function storeAllNumFormats()
    {
        // Leaning num_format syndrome
        $hashNumFormats = array();
        $numFormats = array();
        $index = 164;

        // Iterate through the XF objects and write a FORMAT record if it isn't a
        // built-in format type and if the FORMAT string hasn't already been used.
        foreach ($this->formats as $format) {
            $numFormat = $format->numFormat;

            // Check if $num_format is an index to a built-in format.
            // Also check for a string of zeros, which is a valid format string
            // but would evaluate to zero.
            //
            if (!preg_match("/^0+\d/", $numFormat)) {
                if (preg_match("/^\d+$/", $numFormat)) { // built-in format
                    continue;
                }
            }

            if (isset($hashNumFormats[$numFormat])) {
                // FORMAT has already been used
                $format->numFormat = $hashNumFormats[$numFormat];
            } else {
                // Add a new FORMAT
                $hashNumFormats[$numFormat] = $index;
                $format->numFormat = $index;
                array_push($numFormats, $numFormat);
                $index++;
            }
        }

        // Write the new FORMAT records starting from 0xA4
        $index = 164;
        foreach ($numFormats as $numFormat) {
            $this->storeNumFormat($numFormat, $index);
            $index++;
        }
    }

    /**
     * Write all XF records.
     */
    protected function storeAllXfs()
    {
        // _tmp_format is added by the constructor. We use this to write the default XF's
        // The default font index is 0
        for ($i = 0; $i <= 14; $i++) {
            $xf = $this->tmpFormat->getXf('style'); // Style XF
            $this->append($xf);
        }

        $xf = $this->tmpFormat->getXf('cell'); // Cell XF
        $this->append($xf);

        // User defined XFs
        foreach ($this->formats as $format) {
            $xf = $format->getXf('cell');
            $this->append($xf);
        }
    }

    /**
     * Write all STYLE records.
     */
    protected function storeAllStyles()
    {
        $this->storeStyle();
    }

    /**
     * Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
     * the NAME records.
     */
    protected function storeExterns()
    {
        // Create EXTERNCOUNT with number of worksheets
        $this->storeExterncount(count($this->worksheets));

        // Create EXTERNSHEET for each worksheet
        foreach ($this->sheetNames as $sheetname) {
            $this->storeExternsheet($sheetname);
        }
    }

    /**
     * Write the NAME record to define the print area and the repeat rows and cols.
     */
    protected function storeNames()
    {
        // Create the print area NAME records
        foreach ($this->worksheets as $sheet) {
            // Write a Name record if the print area has been defined
            if (isset($sheet->printRowMin)) {
                $this->storeNameShort(
                    $sheet->index,
                    0x06, // NAME type
                    $sheet->printRowMin,
                    $sheet->printRowMax,
                    $sheet->printColMin,
                    $sheet->printColMax
                );
            }
        }

        // Create the print title NAME records
        foreach ($this->worksheets as $sheet) {
            $rowmin = $sheet->titleRowMin;
            $rowmax = $sheet->titleRowMax;
            $colmin = $sheet->titleColMin;
            $colmax = $sheet->titleColMax;

            // Determine if row + col, row, col or nothing has been defined
            // and write the appropriate record
            //
            if (isset($rowmin) && isset($colmin)) {
                // Row and column titles have been defined.
                // Row title has been defined.
                $this->storeNameLong(
                    $sheet->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    $colmin,
                    $colmax
                );
            } elseif (isset($rowmin)) {
                // Row title has been defined.
                $this->storeNameShort(
                    $sheet->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    0x00,
                    0xff
                );
            } elseif (isset($colmin)) {
                // Column title has been defined.
                $this->storeNameShort(
                    $sheet->index,
                    0x07, // NAME type
                    0x0000,
                    0x3fff,
                    $colmin,
                    $colmax
                );
            } else {
                // Print title hasn't been defined.
            }
        }
    }

    /**
     * Stores the CODEPAGE biff record.
     */
    protected function storeCodepage()
    {
        $record = new Record\Codepage();
        $this->append($record->getData($this->biff->getCodepage()));
    }

    /**
     * Write Excel BIFF WINDOW1 record.
     */
    protected function storeWindow1()
    {
        $record = 0x003D; // Record identifier
        $length = 0x0012; // Number of bytes to follow

        $xWn = 0x0000; // Horizontal position of window
        $yWn = 0x0000; // Vertical position of window
        $dxWn = 0x25BC; // Width of window
        $dyWn = 0x1572; // Height of window

        $grbit = 0x0038; // Option flags
        $ctabsel = $this->selected; // Number of workbook tabs selected
        $wTabRatio = 0x0258; // Tab to scrollbar ratio

        $itabFirst = $this->firstSheet; // 1st displayed worksheet
        $itabCur = $this->activeSheet; // Active worksheet

        $header = pack("vv", $record, $length);
        $data = pack(
            "vvvvvvvvv",
            $xWn,
            $yWn,
            $dxWn,
            $dyWn,
            $grbit,
            $itabCur,
            $itabFirst,
            $ctabsel,
            $wTabRatio
        );
        $this->append($header . $data);
    }

    /**
     * Writes Excel BIFF BOUNDSHEET record.
     * @param string $sheetname Worksheet name
     * @param integer $offset    Location of worksheet BOF
     */
    protected function storeBoundsheet($sheetname, $offset)
    {
        $record = 0x0085; // Record identifier
        if ($this->isBiff8()) {
            $length = 0x08 + strlen($sheetname); // Number of bytes to follow
        } else {
            $length = 0x07 + strlen($sheetname); // Number of bytes to follow
        }

        $grbit = 0x0000; // Visibility and sheet type
        if ($this->isBiff8()) {
            $cch = mb_strlen($sheetname, 'UTF-16LE'); // Length of sheet name
        } else {
            $cch = strlen($sheetname); // Length of sheet name
        }

        $header = pack("vv", $record, $length);
        if ($this->isBiff8()) {
            $data = pack("VvCC", $offset, $grbit, $cch, 0x1);
        } else {
            $data = pack("VvC", $offset, $grbit, $cch);
        }
        $this->append($header . $data . $sheetname);
    }

    /**
     * Write Internal SUPBOOK record
     */
    protected function storeSupbookInternal()
    {
        $record = 0x01AE; // Record identifier
        $length = 0x0004; // Bytes to follow

        $header = pack("vv", $record, $length);
        $data = pack("vv", count($this->worksheets), 0x0104);
        $this->append($header . $data);
    }

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas.
     */
    protected function storeExternsheetBiff8()
    {
        $refs = $this->parser->getReferences();
        $refCount = count($refs);
        $record = 0x0017; // Record identifier
        $length = 2 + 6 * $refCount; // Number of bytes to follow

        $header = pack("vv", $record, $length);
        $data = pack('v', $refCount);
        for ($i = 0; $i < $refCount; $i++) {
            $data .= $refs[$i];
        }
        $this->append($header . $data);
    }

    /**
     * Write STYLE records.
     */
    protected function storeStyle()
    {
        $record = new Record\Style();
        $this->append($record->getData());
    }

    /**
     * Writes FORMAT record for non "built-in" numerical formats.
     *
     * @param string $format Custom format string
     * @param integer $formatIndex   Format index code
     */
    protected function storeNumFormat($format, $formatIndex)
    {
        $record = new Record\Format();
        $this->append($record->getData($this->version, $format, $formatIndex));
    }

    /**
     * Write DATEMODE record to indicate the date system in use (1904 or 1900).
     */
    protected function storeDatemode()
    {
        $record = new Record\Datemode();
        $this->append($record->getData($this->f1904));
    }


    /**
     * Write BIFF record EXTERNCOUNT to indicate the number of external sheet
     * references in the workbook.
     * @param integer $externalRefsCount Number of external references
     */
    protected function storeExterncount($externalRefsCount)
    {
        $record = new Record\Externcount();
        $this->append($record->getData($externalRefsCount));
    }


    /**
     * Writes the Excel BIFF EXTERNSHEET record.
     *
     * @param string $sheetName Worksheet name
     */
    protected function storeExternsheet($sheetName)
    {
        $record = new Record\Externsheet();
        $this->append($record->getData($sheetName));
    }


    /**
     * Store the NAME record in the short format that is used for storing the print
     * area, repeat rows only and repeat columns only.
     *
     * @param integer $index  Sheet index
     * @param integer $type   Built-in name type
     * @param integer $rowmin Start row
     * @param integer $rowmax End row
     * @param integer $colmin Start colum
     * @param integer $colmax End column
     */
    protected function storeNameShort($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record = new Record\NameShort();
        $this->append($record->getData($index, $type, $rowmin, $rowmax, $colmin, $colmax));
    }

    /**
     * Store the NAME record in the long format that is used for storing the repeat
     * rows and columns when both are specified.
     *
     * @param integer $index Sheet index
     * @param integer $type  Built-in name type
     * @param integer $rowmin Start row
     * @param integer $rowmax End row
     * @param integer $colmin Start colum
     * @param integer $colmax End column
     */
    protected function storeNameLong($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record = new Record\NameLong();
        $this->append($record->getData($index, $type, $rowmin, $rowmax, $colmin, $colmax));
    }

    /**
     * Stores the COUNTRY record for localization
     */
    protected function storeCountry()
    {
        $record = new Record\Country();
        $this->append($record->getData($this->countryCode));
    }

    /**
     * Stores the PALETTE biff record.
     */
    protected function storePalette()
    {
        $record = new Record\Palette();
        $this->append($record->getData($this->palette));
    }

    /**
     * Calculate
     * Handling of the SST continue blocks is complicated by the need to include an
     * additional continuation byte depending on whether the string is split between
     * blocks or whether it starts at the beginning of the block. (There are also
     * additional complications that will arise later when/if Rich Strings are
     * supported).
     */
    protected function calculateSharedStringsSizes()
    {
        /* Iterate through the strings to calculate the CONTINUE block sizes.
           For simplicity we use the same size for the SST and CONTINUE records:
           8228 : Maximum Excel97 block size
             -4 : Length of block header
             -8 : Length of additional SST header information
             -8 : Arbitrary number to keep within _add_continue() limit = 8208
        */
        $continueLimit = 8208;
        $blockLength = 0;
        $written = 0;
        $this->blockSizes = array();
        $continue = 0;

        foreach (array_keys($this->strTable) as $string) {
            $stringLength = strlen($string);
            $headerinfo = unpack("vlength/Cencoding", $string);
            $encoding = $headerinfo["encoding"];
            $splitString = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            $blockLength += $stringLength;

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($blockLength < $continueLimit) {
                $written += $stringLength;
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            while ($blockLength >= $continueLimit) {
                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                $headerLength = 3; // Min string + header size -1
                $spaceRemaining = $continueLimit - $written - $continue;


                /* TODO: Unicode data should only be split on char (2 byte)
                boundaries. Therefore, in some cases we need to reduce the
                amount of available
                */
                $align = 0;

                // Only applies to Unicode strings
                if ($encoding == 1) {
                    // Min string + header size -1
                    $headerLength = 4;

                    if ($spaceRemaining > $headerLength) {
                        // String contains 3 byte header => split on odd boundary
                        if (!$splitString && $spaceRemaining % 2 != 1) {
                            $spaceRemaining--;
                            $align = 1;
                        } // Split section without header => split on even boundary
                        else {
                            if ($splitString && $spaceRemaining % 2 == 1) {
                                $spaceRemaining--;
                                $align = 1;
                            }
                        }

                        $splitString = 1;
                    }
                }


                if ($spaceRemaining > $headerLength) {
                    // Write as much as possible of the string in the current block
                    $written += $spaceRemaining;

                    // Reduce the current block length by the amount written
                    $blockLength -= $continueLimit - $continue - $align;

                    // Store the max size for this block
                    $this->blockSizes[] = $continueLimit - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    if ($blockLength > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Store the max size for this block
                    $this->blockSizes[] = $written + $continue;

                    // Not enough space to start the string in the current block
                    $blockLength -= $continueLimit - $spaceRemaining - $continue;
                    $continue = 0;

                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                if ($blockLength < $continueLimit) {
                    $written = $blockLength;
                } else {
                    $written = 0;
                }
            }
        }

        // Store the max size for the last block unless it is empty
        if ($written + $continue) {
            $this->blockSizes[] = $written + $continue;
        }


        /* Calculate the total length of the SST and associated CONTINUEs (if any).
         The SST record will have a length even if it contains no strings.
         This length is required to set the offsets in the BOUNDSHEET records since
         they must be written before the SST records
        */

        $tmpBlockSizes = $this->blockSizes;

        $length = 12;
        if (!empty($tmpBlockSizes)) {
            $length += array_shift($tmpBlockSizes); // SST
        }
        while (!empty($tmpBlockSizes)) {
            $length += 4 + array_shift($tmpBlockSizes); // CONTINUEs
        }

        return $length;
    }

    /**
     * Write all of the workbooks strings into an indexed array.
     *
     * The Excel documentation says that the SST record should be followed by an
     * EXTSST record. The EXTSST record is a hash table that is used to optimise
     * access to SST. However, despite the documentation it doesn't seem to be
     * required so we will ignore it.
     */
    protected function storeSharedStringsTable()
    {
        $record = 0x00fc; // Record identifier
        $length = 0x0008; // Number of bytes to follow
        $total = 0x0000;

        // Iterate through the strings to calculate the CONTINUE block sizes
        $continueLimit = 8208;
        $blockLength = 0;
        $written = 0;
        $continue = 0;

        // sizes are upside down
        $tmpBlockSizes = $this->blockSizes;
        // $tmp_block_sizes = array_reverse($this->block_sizes);

        // The SST record is required even if it contains no strings. Thus we will
        // always have a length
        //
        if (!empty($tmpBlockSizes)) {
            $length = 8 + array_shift($tmpBlockSizes);
        } else {
            // No strings
            $length = 8;
        }

        // Write the SST block header information
        $header = pack("vv", $record, $length);
        $data = pack("VV", $this->strTotal, $this->strUnique);
        $this->append($header . $data);

        /* TODO: Possible bottleneck */
        foreach (array_keys($this->strTable) as $string) {
            $stringLength = strlen($string);
            $headerinfo = unpack("vlength/Cencoding", $string);
            $encoding = $headerinfo["encoding"];
            $splitString = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            $blockLength += $stringLength;

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($blockLength < $continueLimit) {
                $this->append($string);
                $written += $stringLength;
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            //
            while ($blockLength >= $continueLimit) {
                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                //
                $headerLength = 3; // Min string + header size -1
                $spaceRemaining = $continueLimit - $written - $continue;

                // Unicode data should only be split on char (2 byte) boundaries.
                // Therefore, in some cases we need to reduce the amount of available
                // space by 1 byte to ensure the correct alignment.
                $align = 0;

                // Only applies to Unicode strings
                if ($encoding == 1) {
                    // Min string + header size -1
                    $headerLength = 4;

                    if ($spaceRemaining > $headerLength) {
                        // String contains 3 byte header => split on odd boundary
                        if (!$splitString && $spaceRemaining % 2 != 1) {
                            $spaceRemaining--;
                            $align = 1;
                        } elseif ($splitString && $spaceRemaining % 2 == 1) {
                            // Split section without header => split on even boundary
                            $spaceRemaining--;
                            $align = 1;
                        }

                        $splitString = 1;
                    }
                }

                if ($spaceRemaining > $headerLength) {
                    // Write as much as possible of the string in the current block
                    $tmp = substr($string, 0, $spaceRemaining);
                    $this->append($tmp);

                    // The remainder will be written in the next block(s)
                    $string = substr($string, $spaceRemaining);

                    // Reduce the current block length by the amount written
                    $blockLength -= $continueLimit - $continue - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    //
                    if ($blockLength > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Not enough space to start the string in the current block
                    $blockLength -= $continueLimit - $spaceRemaining - $continue;
                    $continue = 0;
                }

                // Write the CONTINUE block header
                if (!empty($this->blockSizes)) {
                    $record = 0x003C;
                    $length = array_shift($tmpBlockSizes);

                    $header = pack('vv', $record, $length);
                    if ($continue) {
                        $header .= pack('C', $encoding);
                    }
                    $this->append($header);
                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                //
                if ($blockLength < $continueLimit) {
                    $this->append($string);
                    $written = $blockLength;
                } else {
                    $written = 0;
                }
            }
        }
    }
}

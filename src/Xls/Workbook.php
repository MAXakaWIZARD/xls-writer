<?php

namespace Xls;

use Xls\OLE\OLE;
use Xls\OLE\PpsFile;
use Xls\OLE\PpsRoot;

/**
 * Class for generating Excel Spreadsheets
*/

class Workbook extends BIFFwriter
{
    /**
     * Filename for the Workbook
     * @var string
     */
    protected $filename;

    /**
     * Formula parser
     * @var FormulaParser
     */
    protected $formulaParser;

    /**
     * Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904)
     * @var integer
     */
    protected $f1904;

    /**
     * The active worksheet of the workbook (0 indexed)
     * @var integer
     */
    protected $activeSheet;

    /**
     * 1st displayed worksheet in the workbook (0 indexed)
     * @var integer
     */
    protected $firstSheet;

    /**
     * Index for creating adding new formats to the workbook
     * @var integer
     */
    protected $xfIndex;

    /**
     * Flag for preventing close from being called twice.
     * @var boolean
     * @see close()
     */
    protected $fileClosed;

    /**
     * The default XF format.
     * @var object Format
     */
    protected $tmpFormat;

    /**
     * Array containing references to all of this workbook's worksheets
     * @var Worksheet[]
     */
    protected $worksheets;

    /**
     * Array of sheetnames for creating the EXTERNSHEET records
     * @var array
     */
    protected $sheetNames;

    /**
     * Array containing references to all of this workbook's formats
     * @var Format[]
     */
    protected $formats;

    /**
     * Array containing the colour palette
     * @var array
     */
    protected $palette;

    /**
     * The default format for URLs.
     * @var object Format
     */
    protected $urlFormat;

    /**
     * The country code used for localization
     * @var integer
     */
    protected $countryCode;

    /**
     * @var
     */
    protected $blockSizes;

    /**
     * @var int
     */
    protected $creationTimestamp;

    /**
     * @var SharedStringsTable
     */
    protected $sst;

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
        $this->formulaParser = new FormulaParser($this->byteOrder, $this->version);
        $this->f1904 = 0;
        $this->xfIndex = 16; // 15 style XF's and 1 cell XF.
        $this->fileClosed = false;

        $this->activeSheet = 0;
        $this->firstSheet = 0;
        $this->worksheets = array();
        $this->sheetNames = array();

        $this->countryCode = -1;

        $this->palette = Palette::getXl97Palette();

        $this->formats = array();
        $this->tmpFormat = new Format($this->version);
        // Add the default format for hyperlinks
        $this->urlFormat = $this->addFormat(array('color' => 'blue', 'underline' => 1));

        $this->sst = new SharedStringsTable();

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
     * Returns an array of the worksheet objects in a workbook
     *
     * @return Worksheet[]
     */
    public function getWorksheets()
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
            $name = 'Sheet' . ($index + 1);
        }

        $name = $this->processSheetName($name);

        if ($this->hasSheet($name)) {
            throw new \Exception("Worksheet '$name' already exists");
        }

        $worksheet = new Worksheet(
            $name,
            $index,
            $this,
            $this->sst,
            $this->urlFormat,
            $this->formulaParser
        );

        $this->worksheets[$index] = $worksheet;
        $this->sheetNames[$index] = $name;

        if (count($this->worksheets) == 1) {
            $this->setActiveSheet(0);
        }

        // Register worksheet name with parser
        $this->formulaParser->setExtSheet($name, $index);

        return $worksheet;
    }

    /**
     * @param string $name
     *
     * @return string
     * @throws \Exception
     */
    protected function processSheetName($name)
    {
        $maxLen = $this->biff->getMaxSheetNameLength();
        if (strlen($name) > $maxLen) {
            throw new \Exception(
                "Sheet name must be shorter than $maxLen chars"
            );
        }

        if ($this->isBiff8() && function_exists('iconv')) {
            $name = iconv('UTF-8', 'UTF-16LE', $name);
        }

        return $name;
    }

    /**
     * @param int $sheetIndex
     */
    public function setActiveSheet($sheetIndex)
    {
        $this->activeSheet = $sheetIndex;
        foreach ($this->worksheets as $idx => $sheet) {
            if ($idx == $sheetIndex) {
                $sheet->select();
            } else {
                $sheet->unselect();
            }
        }
    }

    /**
     * @param int $sheetIndex
     */
    public function setFirstSheet($sheetIndex)
    {
        $this->firstSheet = $sheetIndex;
    }

    /**
     * @return int
     */
    public function getActiveSheet()
    {
        return $this->activeSheet;
    }

    /**
     * @return int
     */
    public function getFirstSheet()
    {
        return $this->firstSheet;
    }

    /**
     * @param string $name
     *
     * @return bool
     */
    public function hasSheet($name)
    {
        return in_array($name, $this->sheetNames, true);
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
        $this->xfIndex++;
        $this->formats[] = $format;

        return $format;
    }

    /**
     * Create new validator.
     *
     * @return Validator reference to a Validator
     */
    public function addValidator()
    {
        return new Validator($this->formulaParser);
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

        Palette::validateColor($index, $red, $green, $blue);

        // Set the RGB value, adjust colour index (wingless dragonfly)
        $this->palette[$index - 8] = array($red, $green, $blue, 0);

        return $index;
    }

    /**
     * @return int
     */
    protected function getSelectedSheetsCount()
    {
        $selected = 0;

        foreach ($this->worksheets as $sheet) {
            if ($sheet->isSelected()) {
                $selected++;
            }
        }

        return $selected;
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

        // Calculate the number of selected worksheet tabs and call the finalization
        // methods for each worksheet
        foreach ($this->worksheets as $sheet) {
            $sheet->close($this->sheetNames);
        }

        // Add Workbook globals
        $this->prependRecord('Bof', array(self::BOF_TYPE_WORKBOOK));
        $this->appendRecord('Codepage', array($this->biff->getCodepage()));

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
        $this->appendRecord('Palette', array($this->palette));
        $this->calcSheetOffsets();

        foreach ($this->worksheets as $sheet) {
            $this->appendRecord('Boundsheet', array($sheet->getName(), $sheet->getOffset()));
        }

        if ($this->countryCode != -1) {
            $this->appendRecord('Country', array($this->countryCode));
        }

        if ($this->isBiff8()) {
            /* TODO: store external SUPBOOK records and XCT and CRN records
            * in case of external references for BIFF8
            */
            //$this->appendRecord('Supbook', $this->worksheets);
            //$this->storeExternsheetBiff8();
            $this->storeSharedStringsTable();
        }

        $this->appendRecord('Eof');

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
        $ole = new PpsFile(OLE::asc2Ucs($this->biff->getWorkbookName()));

        $ole->init();
        $ole->append($this->data);

        foreach ($this->worksheets as $sheet) {
            while ($tmp = $sheet->getData()) {
                $ole->append($tmp);
            }
        }

        $root = new PpsRoot(
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
        $eof = 4;
        $offset = $this->datasize;

        if ($this->isBiff8()) {
            // add the length of the SST
            /* TODO: check if this works for a lot of strings (> 8224 bytes) */
            $this->blockSizes = $this->sst->getBlocksSizesOrDataToWrite();
            $offset += $this->sst->calcSharedStringsTableLength($this->blockSizes);
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
        $offset += $eof;

        foreach ($this->worksheets as $sheet) {
            $sheet->setOffset($offset);
            $offset += $sheet->getDataSize();
        }
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
            $this->appendRecord('Format', array($numFormat, $index));
            $index++;
        }
    }

    /**
     * Write all XF records.
     */
    protected function storeAllXfs()
    {
        // tmpFormat is added by the constructor. We use this to write the default XF's
        // The default font index is 0
        for ($i = 0; $i <= 14; $i++) {
            $xf = $this->tmpFormat->getXf('style');
            $this->append($xf);
        }

        $xf = $this->tmpFormat->getXf('cell');
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
        $this->appendRecord('Style');
    }

    /**
     * Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
     * the NAME records.
     */
    protected function storeExterns()
    {
        $this->appendRecord('Externcount', array(count($this->worksheets)));

        foreach ($this->sheetNames as $sheetName) {
            $this->appendRecord('Externsheet', array($sheetName));
        }
    }

    /**
     * Create the print area NAME records
     */
    protected function storePrintAreaNames()
    {
        foreach ($this->worksheets as $sheet) {
            // Write a Name record if the print area has been defined
            if (isset($sheet->printRowMin)) {
                $this->appendRecord('NameShort', array(
                    $sheet->index,
                    0x06, // NAME type
                    $sheet->printRowMin,
                    $sheet->printRowMax,
                    $sheet->printColMin,
                    $sheet->printColMax
                ));
            }
        }
    }

    /**
     * Create the print title NAME records
     */
    protected function storePrintTitleNames()
    {
        foreach ($this->worksheets as $sheet) {
            $this->storePrintTitleName($sheet);
        }
    }

    /**
     * @param Worksheet $sheet
     */
    protected function storePrintTitleName(Worksheet $sheet)
    {
        $rowmin = $sheet->titleRowMin;
        $rowmax = $sheet->titleRowMax;
        $colmin = $sheet->titleColMin;
        $colmax = $sheet->titleColMax;

        if (isset($rowmin) && isset($colmin)) {
            $recordType = 'NameLong';
        } elseif (isset($colmin)) {
            $recordType = 'NameShort';
            $rowmin = 0x00;
            $rowmax = 0x3fff;
        } elseif (isset($rowmin)) {
            $recordType = 'NameShort';
            $colmin = 0x00;
            $colmax = 0xff;
        } else {
            return;
        }

        $this->appendRecord(
            $recordType,
            array(
                $sheet->index,
                0x07, // NAME type
                $rowmin,
                $rowmax,
                $colmin,
                $colmax
            )
        );
    }

    /**
     * Write the NAME record to define the print area and the repeat rows and cols.
     */
    protected function storeNames()
    {
        $this->storePrintAreaNames();
        $this->storePrintTitleNames();
    }

    /**
     * Write Excel BIFF WINDOW1 record.
     */
    protected function storeWindow1()
    {
        $this->appendRecord(
            'Window1',
            array(
                $this->getSelectedSheetsCount(),
                $this->firstSheet,
                $this->activeSheet
            )
        );
    }

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas.
     */
    protected function storeExternsheetBiff8()
    {
        $record = new Record\Externsheet();
        $this->append($record->getDataForReferences($this->formulaParser->getReferences()));
    }

    /**
     * Write DATEMODE record to indicate the date system in use (1904 or 1900).
     */
    protected function storeDatemode()
    {
        $this->appendRecord('Datemode', array($this->f1904));
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
        $this->appendRecord(
            'SharedStringsTable',
            array(
                $this->blockSizes,
                $this->sst->getTotalCount(),
                $this->sst->getUniqueCount()
            )
        );

        $data = $this->sst->getBlocksSizesOrDataToWrite($this->blockSizes, true);

        foreach ($data as $item) {
            $this->append($item);
        }
    }
}
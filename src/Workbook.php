<?php
namespace Xls;

use Xls\OLE\PpsFile;
use Xls\OLE\PpsRoot;

/**
 * Class for generating Excel Spreadsheets
*/

class Workbook extends BIFFwriter
{
    const COUNTRY_NONE = -1;
    const COUNTRY_USA = 1;

    const BOF_TYPE = 0x0005;

    /**
     * Formula parser
     * @var FormulaParser
     */
    protected $formulaParser;

    /**
     * The active worksheet of the workbook (0 indexed)
     * @var integer
     */
    protected $activeSheetIndex = 0;

    /**
     * 1st displayed worksheet in the workbook (0 indexed)
     * @var integer
     */
    protected $firstSheetIndex = 0;

    /**
     * Index for creating adding new formats to the workbook
     * 15 style XF's and 1 cell XF
     * @var integer
     */
    protected $xfIndex = 16;

    /**
     * Flag for preventing close from being called twice.
     * @var boolean
     * @see close()
     */
    protected $saved = false;

    /**
     * The default XF format.
     * @var Format
     */
    protected $tmpFormat;

    /**
     * Array containing references to all of this workbook's worksheets
     * @var Worksheet[]
     */
    protected $worksheets = array();

    /**
     * Array of sheetnames for creating the EXTERNSHEET records
     * @var array
     */
    protected $sheetNames = array();

    /**
     * Array containing references to all of this workbook's formats
     * @var Format[]
     */
    protected $formats = array();

    /**
     * Array containing the colour palette
     * @var array
     */
    protected $palette = array();

    /**
     * The default format for URLs.
     * @var Format
     */
    protected $urlFormat;

    /**
     * The country code used for localization
     * @var integer
     */
    protected $countryCode = self::COUNTRY_NONE;

    /**
     * @var int
     */
    protected $creationTimestamp;

    /**
     * @var SharedStringsTable
     */
    protected $sst;

    /**
     *
     */
    public function __construct()
    {
        $this->formulaParser = new FormulaParser();

        $this->palette = Palette::getXl97Palette();

        $this->addDefaultFormats();

        $this->sst = new SharedStringsTable();

        $this->setCreationTimestamp(time());
    }

    /**
     *
     */
    protected function addDefaultFormats()
    {
        $this->tmpFormat = new Format();

        // Add the default format for hyperlinks
        $this->urlFormat = $this->addFormat(
            array(
                'font.color' => 'blue',
                'font.underline' => Font::UNDERLINE_ONCE
            )
        );
    }

    /**
     * @param int $creationTime
     */
    public function setCreationTimestamp($creationTime)
    {
        $this->creationTimestamp = $creationTime;
    }

    /**
     * Assemble worksheets into a workbook and send the BIFF data to an OLE
     * storage.
     * This method should always be the last one to be called on every workbook
     *
     * @param string $filePath File path to save
     *
     * @throws \Exception
     */
    public function save($filePath)
    {
        if ($this->saved) {
            throw new \Exception('Workbook was already saved!');
        }

        if (count($this->worksheets) == 0) {
            throw new \Exception('Cannot save workbook with no sheets');
        }

        $this->appendRecord('Bof', array(static::BOF_TYPE));
        $this->appendRecord('Codepage', array(Biff8::CODEPAGE));
        $this->storeWindow1();
        $this->storeDatemode();

        $this->storeAllFonts();
        $this->storeAllNumFormats();
        $this->storeAllXfs();
        $this->storeAllStyles();
        $this->appendRecord('Palette', array($this->palette));

        $this->startBufferedWrite();
        $this->storeCountry();
        $this->appendRecord('RecalcId');
        $this->storeSupbookInternal();
        $this->storeExternsheet();
        $this->storeDefinedNames();
        $this->storeDrawings();
        $this->storeSharedStringsTable();
        $this->appendRecord('Eof');
        $this->endBufferedWrite();

        $this->storeSheets();

        $this->appendRaw($this->getBuffer());

        $this->saveOleFile($filePath);

        $this->saved = true;
    }

    protected function storeDrawings()
    {
        $totalDrawings = 0;
        foreach ($this->getWorksheets() as $sheet) {
            $totalDrawings += count($sheet->getDrawings());
        }

        if ($totalDrawings == 0) {
            return;
        }

        $data = '0F 00 00 F0 52 00 00 00 00 00 06 F0 18 00 00 00';
        $data .= '01 08';
        $data .= '00 00 02 00 00 00 03 00 00 00 01 00 00 00 01 00 00 00 03 00 00 00 33 00 0B F0 12 00 00 00 BF 00 08';
        $data .= ' 00 08 00 81 01 41 00 00 08 C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00';
        $data .= ' 00 08 F7 00 00 10';
        $this->appendRecord('MsoDrawingGroup', array($data));
    }

    /**
     * Write Internal SUPBOOK record
     */
    protected function storeSupbookInternal()
    {
        $this->appendRecord('ExternalBook', array($this->getSheetsCount()));
    }

    /**
     * Calculate the number of selected worksheet tabs and call the finalization
     * methods for each worksheet
     */
    protected function closeSheets()
    {
        foreach ($this->getWorksheets() as $sheet) {
            $sheet->close();
        }
    }

    /**
     *
     */
    protected function storeSheets()
    {
        $this->closeSheets();

        $offset = $this->getDataSize();
        $offset += $this->calcSheetRecordsTotalSize();
        $offset += $this->getBufferSize();

        foreach ($this->getWorksheets() as $sheet) {
            $this->appendRecord('Sheet', array($sheet->getName(), $offset));
            $offset += $sheet->getDataSize();
        }
    }

    /**
     * @return int
     */
    protected function calcSheetRecordsTotalSize()
    {
        $size = 0;
        foreach ($this->worksheets as $sheet) {
            $recordData = $this->getRecord('Sheet', array($sheet->getName()));
            $size += strlen($recordData);
        }

        return $size;
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
     * @return int
     */
    public function getSheetsCount()
    {
        return count($this->worksheets);
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

        $this->checkSheetName($name);

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
            $this->setActiveSheetIndex(0);
        }

        // Register worksheet name with parser
        $this->formulaParser->addSheet($name, $index);
        $this->formulaParser->addRef($index, $index);

        return $worksheet;
    }

    /**
     * @param string $name
     *
     * @return string
     * @throws \Exception
     */
    protected function checkSheetName($name)
    {
        $maxLen = Biff8::MAX_SHEET_NAME_LENGTH;
        if (strlen($name) > $maxLen) {
            throw new \Exception(
                "Sheet name must be shorter than $maxLen chars"
            );
        }
    }

    /**
     * @param int $sheetIndex
     */
    public function setActiveSheetIndex($sheetIndex)
    {
        $this->activeSheetIndex = $sheetIndex;
        foreach ($this->worksheets as $idx => $sheet) {
            if ($idx == $sheetIndex) {
                $sheet->select();
            } else {
                $sheet->unselect();
            }
        }
    }

    /**
     * @return int
     */
    public function getActiveSheetIndex()
    {
        return $this->activeSheetIndex;
    }

    /**
     * @return int
     */
    public function getFirstSheetIndex()
    {
        return $this->firstSheetIndex;
    }

    /**
     * @param int $firstSheetIndex
     */
    public function setFirstSheetIndex($firstSheetIndex)
    {
        $this->firstSheetIndex = $firstSheetIndex;
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
        $format = new Format($this->xfIndex, $properties);
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
        Palette::validateColor($index, $red, $green, $blue);

        // Set the RGB value, adjust colour index (wingless dragonfly)
        $this->palette[$index - 8] = array($red, $green, $blue, 0);

        return $index;
    }

    /**
     * Store the workbook in an OLE container
     * @param string $filePath
     */
    protected function saveOleFile($filePath)
    {
        $ole = new PpsFile(Biff8::WORKBOOK_NAME);
        $ole->append($this->data);

        foreach ($this->worksheets as $sheet) {
            $ole->append($sheet->getData());
        }

        $root = new PpsRoot(
            $this->creationTimestamp,
            array($ole)
        );

        $root->save($filePath);
    }

    /**
     * Store the Excel FONT records.
     */
    protected function storeAllFonts()
    {
        foreach ($this->getFonts() as $font) {
            $this->appendRecord('Font', array($font));
        }
    }

    /**
     * @return Font[]
     */
    protected function getFonts()
    {
        $fontsMap = array();

        $defaultFont = $this->tmpFormat->getFont();
        $defaultFont->setIndex(0);

        $key = $defaultFont->getKey();
        $fontsMap[$key] = 1;

        //add default font for 5 times
        $fonts = array_fill(0, 5, $defaultFont);

        // Iterate through the XF objects and write a FONT record if it isn't the
        // same as the default FONT and if it hasn't already been used.
        $index = 6; // The first user defined FONT
        foreach ($this->formats as $format) {
            $font = $format->getFont();
            $key = $font->getKey();

            if (!isset($fontsMap[$key])) {
                // Add a new FONT record
                $fontsMap[$key] = 1;
                $font->setIndex($index);
                $fonts[] = $font;
                $index++;
            }
        }

        return $fonts;
    }

    /**
     * Store user defined numerical formats i.e. FORMAT records
     */
    protected function storeAllNumFormats()
    {
        $map = array();
        $index = 164;

        // Iterate through the XF objects and write a FORMAT record if it isn't a
        // built-in format type and if the FORMAT string hasn't already been used.
        foreach ($this->formats as $format) {
            $numFormat = $format->getNumFormat();

            if (NumberFormat::isBuiltIn($numFormat)) {
                $format->setNumFormatIndex($numFormat);
                continue;
            }

            if (!isset($map[$numFormat])) {
                // Add a new FORMAT
                $map[$numFormat] = 1;
                $format->setNumFormatIndex($index);
                $this->appendRecord('Format', array($numFormat, $index));
                $index++;
            }
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
            $xfRecord = $this->tmpFormat->getXf('style');
            $this->append($xfRecord);
        }

        $xfRecord = $this->tmpFormat->getXf('cell');
        $this->append($xfRecord);

        // User defined XFs
        foreach ($this->formats as $format) {
            $xfRecord = $format->getXf('cell');
            $this->append($xfRecord);
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
     *
     */
    protected function storeCountry()
    {
        if ($this->countryCode != self::COUNTRY_NONE) {
            $this->appendRecord('Country', array($this->countryCode));
        }
    }

    /**
     * Write the NAME record to define the print area and the repeat rows and cols.
     */
    protected function storeDefinedNames()
    {
        $this->storePrintAreaNames();
        $this->storePrintTitleNames();
    }

    /**
     * Create the print area NAME records
     */
    protected function storePrintAreaNames()
    {
        foreach ($this->worksheets as $sheet) {
            $printSetup = $sheet->getPrintSetup();
            if ($printSetup->isPrintAreaSet()) {
                $area = $printSetup->getPrintArea();

                $data = $this->getRangeCommonHeader($sheet);
                $data .= \Xls\Subrecord\Range::getData(array($area), false);

                $this->appendRecord('DefinedName', array(
                    Record\DefinedName::BUILTIN_PRINT_AREA,
                    $sheet->getIndex() + 1,
                    $data
                ));
            }
        }
    }

    protected function getRangeCommonHeader(Worksheet $sheet)
    {
        return pack('Cv', 0x3B, $sheet->getIndex());
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
        $printRepeat = $sheet->getPrintSetup()->getPrintRepeat();
        if ($printRepeat->isEmpty()) {
            return;
        }

        $this->appendRecord('DefinedName', array(
            Record\DefinedName::BUILTIN_PRINT_TITLES,
            $sheet->getIndex() + 1,
            $this->getPrintTitleData($sheet)
        ));
    }

    /**
     * @param Worksheet $sheet
     *
     * @return string
     */
    protected function getPrintTitleData(Worksheet $sheet)
    {
        $printRepeat = $sheet->getPrintSetup()->getPrintRepeat();

        $rowmin = $printRepeat->getRowFrom();
        $rowmax = $printRepeat->getRowTo();
        $colmin = $printRepeat->getColFrom();
        $colmax = $printRepeat->getColTo();

        $rangeHeader = $this->getRangeCommonHeader($sheet);

        if ($rowmax !== Biff8::MAX_ROW_IDX && $colmax !== Biff8::MAX_COL_IDX) {
            $data = pack('Cv', 0x29, 0x17); // tMemFunc
            $data .= $rangeHeader;
            $data .= pack('v4', 0, Biff8::MAX_ROW_IDX, $colmin, $colmax); // tArea3d
            $data .= $rangeHeader;
            $data .= pack('v4', $rowmin, $rowmax, 0, Biff8::MAX_COL_IDX); // tArea3d
            $data .= pack('C', 0x10); // tList
        } else {
            $data = $rangeHeader;
            $data .= pack('v4', $rowmin, $rowmax, $colmin, $colmax);
        }

        return $data;
    }

    /**
     * Write Excel BIFF WINDOW1 record.
     */
    protected function storeWindow1()
    {
        $selectedSheetsCount = 1;
        $this->appendRecord(
            'Window1',
            array(
                $selectedSheetsCount,
                $this->firstSheetIndex,
                $this->activeSheetIndex
            )
        );
    }

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas.
     */
    protected function storeExternsheet()
    {
        $this->appendRecord('Externsheet', array($this->formulaParser->getReferences()));
    }

    /**
     * Write DATEMODE record to indicate the date system in use (1904 or 1900)
     * Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904)
     */
    protected function storeDatemode()
    {
        $this->appendRecord('Datemode', array(0));
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
        $this->appendRecord('SharedStringsTable', array($this->sst));

        foreach ($this->sst->getDataForWrite() as $item) {
            $this->append($item);
        }
    }
}

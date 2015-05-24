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

    /**
     * Formula parser
     * @var FormulaParser
     */
    protected $formulaParser;

    /**
     * Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904)
     * @var integer
     */
    protected $f1904 = 0;

    /**
     * The active worksheet of the workbook (0 indexed)
     * @var integer
     */
    protected $activeSheet = 0;

    /**
     * 1st displayed worksheet in the workbook (0 indexed)
     * @var integer
     */
    protected $firstSheet = 0;

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
     * @var object Format
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
     * @var object Format
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
        parent::__construct();

        $this->formulaParser = new FormulaParser($this->byteOrder);

        $this->palette = Palette::getXl97Palette();

        $this->tmpFormat = new Format($this->byteOrder);
        // Add the default format for hyperlinks
        $this->urlFormat = $this->addFormat(array('color' => 'blue', 'underline' => 1));

        $this->sst = new SharedStringsTable();

        $this->setCreationTimestamp(time());
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
     * Assemble worksheets into a workbook and send the BIFF data to an OLE
     * storage.
     * This method should always be the last one to be called on every workbook
     *
     * @param string $filePath File path to save
     *
     * @throws \Exception
     * @return boolean true on success.
     */
    public function save($filePath)
    {
        if ($this->saved) {
            throw new \Exception('Workbook was already saved!');
        }

        if (count($this->worksheets) == 0) {
            throw new \Exception('Cannot save workbook with no sheets');
        }

        $this->appendRecord('Bof', array(self::BOF_TYPE_WORKBOOK));
        $this->appendRecord('Codepage', array($this->biff->getCodepage()));
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
        $this->storeSharedStringsTable();
        $this->appendRecord('Eof');
        $this->endBufferedWrite();

        $this->storeSheets();

        $this->appendRaw($this->getBuffer());

        $this->saveOleFile($filePath);

        $this->saved = true;
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
            $this->setActiveSheet(0);
        }

        // Register worksheet name with parser
        $this->formulaParser->setExtSheet($name, $index);
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
        $maxLen = $this->biff->getMaxSheetNameLength();
        if (strlen($name) > $maxLen) {
            throw new \Exception(
                "Sheet name must be shorter than $maxLen chars"
            );
        }
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
        $format = new Format($this->byteOrder, $this->xfIndex, $properties);
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
     * Store the workbook in an OLE container
     * @param $filePath
     */
    protected function saveOleFile($filePath)
    {
        $ole = new PpsFile($this->biff->getWorkbookName());
        $ole->append($this->data);

        foreach ($this->worksheets as $sheet) {
            $ole->append($sheet->getData());
        }

        $root = new PpsRoot(
            $this->getCreationTimestamp(),
            array($ole)
        );

        $root->save($filePath);
    }

    /**
     * Store the Excel FONT records.
     */
    protected function storeAllFonts()
    {
        // tmp_format is added by the constructor. We use this to write the default XF's
        $fontRecordData = $this->tmpFormat->getFontRecord();

        // Note: Fonts are 0-indexed. According to the SDK there is no index 4,
        // so the following fonts are 0, 1, 2, 3, 5
        for ($i = 1; $i <= 5; $i++) {
            $this->append($fontRecordData);
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
            if ($sheet->isPrintAreaSet()) {
                $data = pack(
                    'Cvvvvv',
                    0x3B,
                    $sheet->getIndex(),
                    $sheet->printRowMin,
                    $sheet->printRowMax,
                    $sheet->printColMin,
                    $sheet->printColMax
                );
                $this->appendRecord('DefinedName', array(
                    Record\DefinedName::BUILTIN_PRINT_AREA,
                    $sheet->getIndex() + 1,
                    $data
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

        if (!isset($rowmin) && !isset($colmin)) {
            return;
        }

        if (isset($rowmin) && isset($colmin)) {
            $data = pack('Cv', 0x29, 0x17); // tMemFunc
            $data .= pack('Cvvvvv', 0x3B, $sheet->getIndex(), 0, 65535, $colmin, $colmax); // tArea3d
            $data .= pack('Cvvvvv', 0x3B, $sheet->getIndex(), $rowmin, $rowmax, 0, Biff8::MAX_COLS - 1); // tArea3d
            $data .= pack('C', 0x10); // tList
        } else {
            if (isset($colmin)) {
                $rowmin = 0;
                $rowmax = 65535;
            } else {
                $colmin = 0;
                $colmax = Biff8::MAX_COLS - 1;
            }

            $data = pack('Cvvvvv', 0x3B, $sheet->getIndex(), $rowmin, $rowmax, $colmin, $colmax);
        }

        $this->appendRecord('DefinedName', array(
            Record\DefinedName::BUILTIN_PRINT_TITLES,
            $sheet->getIndex() + 1,
            $data
        ));
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
    protected function storeExternsheet()
    {
        /** @var Record\Externsheet $record */
        $record = $this->createRecord('Externsheet');
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
                $this->sst->getBlocksSizes(),
                $this->sst->getTotalCount(),
                $this->sst->getUniqueCount()
            )
        );

        $data = $this->sst->getDataForWrite();
        foreach ($data as $item) {
            $this->append($item);
        }
    }
}

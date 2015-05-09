<?php
namespace Xls\Tests;

use Xls\Workbook;
use Xls\Biff5;
use Xls\Biff8;
use Xls\Format;
use Xls\Fill;
use Xls\Cell;

/**
 *
 */
class GeneralTest extends \PHPUnit_Framework_TestCase
{
    const WORKBOOK_TS = 1429042916;

    /**
     * @var string
     */
    protected $testFilePath;

    /**
     *
     */
    public function setUp()
    {
        $this->testFilePath = TEST_DATA_PATH . '/test.xls';
    }

    /**
     *
     */
    public function tearDown()
    {
        //@unlink($this->testFilePath);
    }

    /**
     * @param $prefix
     * @param $suffix
     *
     * @return string
     */
    protected function getFilePath($prefix, $suffix)
    {
        return TEST_DATA_PATH . '/' . $prefix . $suffix . '.xls';
    }

    /**
     * @param $params
     *
     * @return Workbook
     */
    protected function createWorkbook($params)
    {
        $workbook = new Workbook($params['format']);
        $workbook->setCreationTimestamp(self::WORKBOOK_TS);

        return $workbook;
    }

    /**
     *
     */
    public function testUnsupportedVersion()
    {
        $this->setExpectedException('\Exception', 'Unsupported BIFF version');
        new Workbook($this->testFilePath, 0);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testGeneral($params)
    {
        $workbook = $this->createWorkbook($params);

        $worksheet = $workbook->addWorksheet('My first worksheet');

        $worksheet->write(0, 0, 'Name');
        $worksheet->write(0, 1, 'Age');
        $worksheet->write(1, 0, 'John Smith');
        $worksheet->write(1, 1, 30);
        $worksheet->write(2, 0, 'Johann Schmidt');
        $worksheet->write(2, 1, 31);
        $worksheet->write(3, 0, 'Juan Herrera');
        $worksheet->write(3, 1, 32);

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('general', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);

        $this->setExpectedException('\Exception', 'Workbook was already saved!');
        $workbook->save($this->testFilePath);
    }

    /**
     * @return array
     */
    public function providerDifferentBiffVersions()
    {
        return array(
            array(
                array(
                    'format' => Biff5::VERSION,
                    'suffix' => ''
                )
            ),
            array(
                array(
                    'format' => Biff8::VERSION,
                    'suffix' => '_biff8'
                )
            )
        );
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param $params
     *
     * @throws \Exception
     */
    public function testRich($params)
    {
        $workbook = $this->createWorkbook($params);

        $worksheet = $workbook->addWorksheet('New PC');

        $headerFormat = $workbook->addFormat();
        $headerFormat->setBold();
        $headerFormat->setBorder(Format::BORDER_THIN);
        $headerFormat->setBorderColor('navy');
        $headerFormat->setColor('blue');
        $headerFormat->setAlign('center');
        $headerFormat->setPattern(Fill::PATTERN_GRAY50);

        //#ccc
        $workbook->setCustomColor(12, 204, 204, 204);
        $headerFormat->setFgColor(12);

        $cellFormat = $workbook->addFormat();
        $cellFormat->setNormal();
        $cellFormat->setBorder(Format::BORDER_THIN);
        $cellFormat->setBorderColor('navy');
        $cellFormat->setUnLocked();

        $priceFormat = $workbook->addFormat();
        $priceFormat->setBorder(Format::BORDER_THIN);
        $priceFormat->setBorderColor('navy');
        $priceFormat->setNumFormat(2);

        $worksheet->write(0, 0, 'Title', $headerFormat);
        $worksheet->write(0, 1, 'Count', $headerFormat);
        $worksheet->write(0, 2, 'Price', $headerFormat);
        $worksheet->write(0, 3, 'Amount', $headerFormat);

        $worksheet->write(1, 0, 'Intel Core i7 2600K', $cellFormat);
        $worksheet->write(1, 1, 1, $cellFormat);
        $worksheet->write(1, 2, 500, $priceFormat);
        $worksheet->writeFormula(1, 3, '=B2*C2', $priceFormat);

        $worksheet->write(2, 0, 'ASUS P8P67', $cellFormat);
        $worksheet->write(2, 1, 1, $cellFormat);
        $worksheet->write(2, 2, 325, $priceFormat);
        $worksheet->writeFormula(2, 3, '=B3*C3', $priceFormat);

        $worksheet->write(3, 0, 'DDR2-800 8Gb', $cellFormat);
        $worksheet->write(3, 1, 4, $cellFormat);
        $worksheet->write(3, 2, 100.15, $priceFormat);
        $worksheet->writeFormula(3, 3, '=B4*C4', $priceFormat);

        for ($i = 4; $i < 10; $i++) {
            for ($j = 0; $j < 4; $j++) {
                $worksheet->write($i, $j, '', $cellFormat);
            }
        }

        $oldPriceFormat = $workbook->addFormat();
        $oldPriceFormat->setBorder(Format::BORDER_THIN);
        $oldPriceFormat->setBorderColor('navy');
        $oldPriceFormat->setSize(12);
        $oldPriceFormat->setStrikeOut();
        $oldPriceFormat->setOutLine();
        $oldPriceFormat->setItalic();
        $oldPriceFormat->setShadow();
        $oldPriceFormat->setNumFormat(2);
        $oldPriceFormat->setLocked();
        $oldPriceFormat->setTextWrap();
        $oldPriceFormat->setTextRotation(0);

        $grandFormat = $workbook->addFormat();
        $grandFormat->setBold();
        $grandFormat->setBorder(Format::BORDER_THIN);
        $grandFormat->setBorderColor('navy');
        $grandFormat->setSize(12);
        $grandFormat->setFontFamily('Tahoma');
        $grandFormat->setUnderline(Format::UNDERLINE_ONCE);
        $grandFormat->setNumFormat(2);
        $this->assertTrue($grandFormat->isBuiltInNumFormat());

        $worksheet->write(10, 0, 'Total', $grandFormat);
        $worksheet->write(10, 1, '', $grandFormat);
        $worksheet->write(10, 2, '', $grandFormat);
        $worksheet->mergeCells(10, 0, 10, 2);
        $worksheet->writeFormula(10, 3, '=SUM(D2:D10)', $oldPriceFormat);

        $worksheet->write(11, 0, 'Grand total', $grandFormat);
        $worksheet->write(11, 1, '', $grandFormat);
        $worksheet->write(11, 2, '', $grandFormat);
        $worksheet->mergeCells(11, 0, 11, 2);
        $worksheet->writeFormula(11, 3, '=ROUND(D11-D11*0.2, 2)', $grandFormat);

        //
        $discountFormat = $workbook->addFormat();
        $discountFormat->setColor('red');
        $discountFormat->setScript(Format::SCRIPT_SUPER);
        $discountFormat->setSize(14);
        $discountFormat->setFgColor('white');
        $discountFormat->setBgColor('black');
        $worksheet->write(11, 4, '20% discount!', $discountFormat);

        //set some validation
        $countValidator = $workbook->addValidator();
        $countValidator->setPrompt('Enter valid item count');
        $countValidator->setFormula1('INDIRECT(ADDRESS(ROW(), COLUMN())) > 0');
        $worksheet->setValidation(1, 1, 3, 1, $countValidator);

        $anotherSheet = $workbook->addWorksheet('Another sheet');
        $anotherSheet->write(0, 0, 'Test');

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('rich', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     *
     */
    public function testRowColToCellInvalid()
    {
        $this->setExpectedException('\Exception', 'Maximum column value exceeded: 256');
        Cell::getAddress(0, 256);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param $params
     *
     * @throws \Exception
     */
    public function testProtected($params)
    {
        $workbook = $this->createWorkbook($params);

        $worksheet = $workbook->addWorksheet();
        $worksheet->write(0, 0, 'Test');
        $worksheet->protect('1234');

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('protected', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testSelection($params)
    {
        $workbook = $this->createWorkbook($params);

        $worksheet = $workbook->addWorksheet();
        $worksheet->write(0, 0, 'Test');
        $worksheet->setSelection(0, 0, 5, 5);

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('selection', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testMultipleSheets($params)
    {
        $workbook = $this->createWorkbook($params);

        for ($i = 1; $i <= 4; $i++) {
            $s = $workbook->addWorksheet();
            $s->write(0, 0, 'Test' . $i);
        }

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('multiple_sheets', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testDefcols($params)
    {
        $workbook = $this->createWorkbook($params);

        $worksheet = $workbook->addWorksheet();
        $worksheet->write(0, 0, 'Test1');
        $worksheet->write(0, 1, 'Test2');
        $worksheet->write(0, 2, 'Test3');

        $worksheet->setColumn(0, 0, 25);
        $worksheet->setColumn(1, 1, 50);
        $worksheet->setColumn(2, 3, 10, null, 1);

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('defcols', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testCountry($params)
    {
        $workbook = $this->createWorkbook($params);
        $workbook->setCountry($workbook::COUNTRY_USA);

        $worksheet = $workbook->addWorksheet();
        $worksheet->write(0, 0, 'Test1');

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('country', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testPortraitLayout($params)
    {
        $params['format'] = Biff5::VERSION;
        $params['suffix'] = '';

        $workbook = $this->createWorkbook($params);
        $workbook->setCountry($workbook::COUNTRY_USA);

        $sheet = $workbook->addWorksheet();
        $row = array(
            'Portrait layout test',
            '',
            '',
            '',
            '',
            'Test2'
        );
        $sheet->writeRow(0, 0, $row);

        $sheet->setZoom(125);

        $sheet->setPortrait();
        $sheet->setMargins(1.25);
        $sheet->setHeader('Page header');
        $sheet->setFooter('Page footer');
        $sheet->setPrintScale(90);
        $sheet->setPaper($sheet::PAPER_A3);
        $sheet->printArea(0, 0, 10, 10);
        $sheet->hideGridlines();
        $sheet->setHPagebreaks(array(1));
        $sheet->setVPagebreaks(array(5));

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('layout_portrait', $params['suffix']);
        //if ($params['suffix'] == '_biff8') exit;
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerDifferentBiffVersions
     * @param array $params
     */
    public function testLandscapeLayout($params)
    {
        $params['format'] = Biff5::VERSION;
        $params['suffix'] = '';

        $workbook = $this->createWorkbook($params);
        $workbook->setCountry($workbook::COUNTRY_USA);

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Landscape layout test');
        $sheet->writeCol(0, 1, range(1, 10));

        $sheet->setZoom(125);
        $sheet->hideScreenGridlines();

        $sheet->setLandscape();

        //header and footer should be cut to max length (255)
        $sheet->setHeader('Page header' . str_repeat('.', 255));
        $sheet->setFooter('Page footer' . str_repeat('.', 255));

        $sheet->centerHorizontally();
        $sheet->centerVertically();
        $sheet->fitToPages(5, 5);
        $sheet->setPaper($sheet::PAPER_A4);
        $sheet->printRowColHeaders();

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('layout_landscape', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }
}

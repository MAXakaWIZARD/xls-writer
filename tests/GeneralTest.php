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
     * @return array
     */
    public function providerBiff5()
    {
        return array(
            array(
                array(
                    'format' => Biff5::VERSION,
                    'suffix' => ''
                )
            )
        );
    }

    /**
     * @return array
     */
    public function providerBiff8()
    {
        return array(
            array(
                array(
                    'format' => Biff8::VERSION,
                    'suffix' => '_biff8'
                )
            )
        );
    }

    /**
     * @return array
     */
    public function providerBiff5AndBiff8()
    {
        $biff5 = $this->providerBiff5();
        $biff8 = $this->providerBiff8();

        return array(
            $biff5[0],
            $biff8[0]
        );
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
     * @dataProvider providerBiff5AndBiff8
     *
     * @param array $params
     */
    public function testGeneral($params)
    {
        $workbook = $this->createWorkbook($params);

        $sheet = $workbook->addWorksheet('My first worksheet');
        $sheet->writeRow(
            0,
            0,
            array(
                array('Name', 'John Smith', 'Johann Schmidt', 'Juan Herrera'),
                array('Age', 30, 31, 32)
            )
        );

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('general', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);

        $this->setExpectedException('\Exception', 'Workbook was already saved!');
        $workbook->save($this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     *
     * @throws \Exception
     */
    public function testRich($params)
    {
        $workbook = $this->createWorkbook($params);

        $sheet = $workbook->addWorksheet('New PC');

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

        $sheet->writeRow(0, 0, array('Title', 'Count', 'Price', 'Amount'), $headerFormat);

        $sheet->writeRow(1, 0, array('Intel Core i7 2600K', 1), $cellFormat);
        $sheet->write(1, 2, 500, $priceFormat);
        $sheet->writeFormula(1, 3, '=B2*C2', $priceFormat);

        $sheet->writeRow(2, 0, array('ASUS P8P67', 1), $cellFormat);
        $sheet->write(2, 2, 325, $priceFormat);
        $sheet->writeFormula(2, 3, '=B3*C3', $priceFormat);

        $sheet->writeRow(3, 0, array('DDR2-800 8Gb', 4), $cellFormat);
        $sheet->write(3, 2, 100.15, $priceFormat);
        $sheet->writeFormula(3, 3, '=B4*C4', $priceFormat);

        $emptyRow = array_fill(0, 4, '');
        for ($i = 4; $i < 10; $i++) {
            $sheet->writeRow($i, 0, $emptyRow, $cellFormat);
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

        $sheet->writeRow(10, 0, array('Total', '', ''), $grandFormat);
        $sheet->mergeCells(10, 0, 10, 2);
        $sheet->writeFormula(10, 3, '=SUM(D2:D10)', $oldPriceFormat);

        $sheet->writeRow(11, 0, array('Grand total', '', ''), $grandFormat);
        $sheet->mergeCells(11, 0, 11, 2);
        //should be written as formula
        $sheet->write(11, 3, '=ROUND(D11-D11*0.2, 2)', $grandFormat);

        //
        $discountFormat = $workbook->addFormat();
        $discountFormat->setColor('red');
        $discountFormat->setScript(Format::SCRIPT_SUPER);
        $discountFormat->setSize(14);
        $discountFormat->setFgColor('white');
        $discountFormat->setBgColor('black');
        $sheet->write(11, 4, '20% discount!', $discountFormat);

        //set some validation
        $countValidator = $workbook->addValidator();
        $countValidator->setPrompt('Enter valid item count');
        $countValidator->setFormula1('INDIRECT(ADDRESS(ROW(), COLUMN())) > 0');
        $sheet->setValidation(1, 1, 3, 1, $countValidator);

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
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     *
     * @throws \Exception
     */
    public function testProtected($params)
    {
        $workbook = $this->createWorkbook($params);

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->protect('1234');

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('protected', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     */
    public function testSelection($params)
    {
        $workbook = $this->createWorkbook($params);

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->setSelection(0, 0, 5, 5);

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('selection', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
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
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     */
    public function testDefcolsAndRowsizes($params)
    {
        $workbook = $this->createWorkbook($params);

        $sheet = $workbook->addWorksheet();
        $sheet->writeRow(0, 0, array('Test1', 'Test2', 'Test3'));

        $sheet->setColumn(0, 0, 25);
        $sheet->setColumn(1, 1, 50);
        $sheet->setColumn(2, 3, 10, null, 1);

        $sheet->setRow(0, 30);
        $sheet->setRow(1, 15);
        $sheet->setRow(2, 10, null, 1);

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('defcols_rowsizes', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     */
    public function testCountry($params)
    {
        $workbook = $this->createWorkbook($params);
        $workbook->setCountry($workbook::COUNTRY_USA);

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test1');

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('country', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5
     * @param array $params
     */
    public function testPortraitLayout($params)
    {
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
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5
     * @param array $params
     */
    public function testLandscapeLayout($params)
    {
        $workbook = $this->createWorkbook($params);
        $workbook->setCountry($workbook::COUNTRY_USA);

        $sheet = $workbook->addWorksheet();

        $fields = range(1, 15);
        $fieldValues = array();
        $headers = array('ID', 'Name');
        foreach ($fields as $idx) {
            $headers[] = 'Field' . $idx;
            $fieldValues[] = 'Field value ' . $idx;
        }

        $sheet->writeRow(0, 0, $headers);

        $ids = range(1, 65);
        foreach ($ids as $id) {
            $sheet->write($id, 0, $id);
            $sheet->write($id, 1, 'Name' . $id);
            $sheet->writeRow($id, 2, $fieldValues);
        }

        $sheet->repeatRows(0);
        $sheet->repeatRows(0, 0);

        $sheet->repeatColumns(0);
        $sheet->repeatColumns(0, 0);

        $sheet->freezePanes(array(1, 1));

        $sheet->setZoom(125);
        $sheet->hideScreenGridlines();

        $sheet->setLandscape();

        //header and footer should be cut to max length (255)
        $sheet->setHeader('Page header ' . str_repeat('.', 255));
        $sheet->setFooter('Page footer ' . str_repeat('.', 255));

        $sheet->centerHorizontally();
        $sheet->centerVertically();
        $sheet->fitToPages(999, 999);
        $sheet->setPaper($sheet::PAPER_A4);
        $sheet->printRowColHeaders();

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('layout_landscape', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     */
    public function testImage($params)
    {
        $workbook = $this->createWorkbook($params);

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->insertBitmap(2, 2, TEST_DATA_PATH . '/elephpant.bmp');

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('image', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     *
     * @param $params
     */
    public function testMergeCells($params)
    {
        $workbook = $this->createWorkbook($params);
        $sheet = $workbook->addWorksheet();

        $sheet->writeRow(1, 0, array('Merge1', '', ''));
        $sheet->mergeCells(1, 0, 1, 4);
        $sheet->writeRow(2, 1, array('Merge2', '', ''));
        $sheet->mergeCells(2, 1, 2, 4);
        $sheet->writeRow(3, 2, array('Merge3', '', ''));
        $sheet->mergeCells(3, 2, 3, 4);

        $format = $workbook->addFormat();
        $format->setAlign('center');
        $sheet->writeRow(4, 3, array('Merge4', '', ''), $format);
        $sheet->mergeCells(4, 3, 5, 4);

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('merge', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }

    /**
     * @dataProvider providerBiff5AndBiff8
     * @param array $params
     */
    public function testThawPanes($params)
    {
        $workbook = $this->createWorkbook($params);
        $workbook->setCountry($workbook::COUNTRY_USA);

        $sheet = $workbook->addWorksheet();

        $fields = range(1, 15);
        $fieldValues = array();
        $headers = array('ID', 'Name');
        foreach ($fields as $idx) {
            $headers[] = 'Field' . $idx;
            $fieldValues[] = 'Field value ' . $idx;
        }

        $sheet->writeRow(0, 0, $headers);

        $ids = range(1, 65);
        foreach ($ids as $id) {
            $sheet->write($id, 0, $id);
            $sheet->write($id, 1, 'Name' . $id);
            $sheet->writeRow($id, 2, $fieldValues);
        }

        $sheet->thawPanes(array(1, 1));

        $workbook->save($this->testFilePath);

        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath('thaw_panes', $params['suffix']);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }
}

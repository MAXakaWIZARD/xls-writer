<?php
namespace Test;

use Xls\Fill;

/**
 *
 */
class GeneralTest extends TestAbstract
{
    public function testGeneral()
    {
        $workbook = $this->createWorkbook();
        $workbook->setCountry($workbook::COUNTRY_USA);

        $sheet = $workbook->addWorksheet('My first worksheet');
        $sheet->writeRow(
            0,
            0,
            array(
                array('Name', 'John Smith', 'Johann Schmidt', 'Иван Иванов'),
                array('Age', 30, 31, 32)
            )
        );

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('general');

        $this->setExpectedException('\Exception', 'Workbook was already saved!');
        $workbook->save($this->testFilePath);
    }

    /**
     * @throws \Exception
     */
    public function testProtected()
    {
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->protect('1234');

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('protected');
    }

    public function testSelection()
    {
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->setSelection(0, 0, 5, 5);

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('selection');
    }

    public function testMultipleSheets()
    {
        $workbook = $this->createWorkbook();

        $sheetNames = array(
            'First sheet',
            'Второй лист',
            'Third sheet',
            '4th sheet'
        );

        for ($i = 1; $i <= 4; $i++) {
            $s = $workbook->addWorksheet($sheetNames[$i - 1]);
            $s->write(0, 0, 'Test' . $i);
        }

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('multiple_sheets');
    }

    public function testDefcolsAndRowsizes()
    {
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();
        $sheet->writeRow(0, 0, array('Test1', 'Test2', 'Test3'));

        $sheet->setColumn(0, 0, 25);
        $sheet->setColumn(1, 1, 50);
        $sheet->setColumn(2, 3, 10, null, 1);

        $sheet->setRow(0, 30);
        $sheet->setRow(1, 15);
        $sheet->setRow(2, 10, null, 1);

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('defcols_rowsizes');
    }

    public function testImage()
    {
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->insertBitmap(2, 2, TEST_DATA_PATH . '/elephpant.bmp');

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('image');
    }

    public function testMergeCells()
    {
        $workbook = $this->createWorkbook();
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

        $this->checkTestFileIsEqualTo('merge');
    }

    public function testThawPanes()
    {
        $workbook = $this->createWorkbook();
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

        $this->checkTestFileIsEqualTo('thaw_panes');
    }

    public function testLongStrings()
    {
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();

        //keep for full test coverage
        $sheet->write(0, 0, str_repeat('a', 33));
        $sheet->writeFormula(0, 1, '=LEN(A1)');
        //keep for full test coverage
        $sheet->write(5, 0, str_repeat('e', 8200));
        $sheet->writeFormula(5, 1, '=LEN(A6)');

        $sheet->write(1, 0, str_repeat('b', 2048));
        $sheet->writeFormula(1, 1, '=LEN(A2)');

        $sheet->write(2, 0, str_repeat('c', 4096));
        $sheet->writeFormula(2, 1, '=LEN(A3)');

        $sheet->write(3, 0, str_repeat('c', 8192));
        $sheet->writeFormula(3, 1, '=LEN(A4)');

        $sheet->write(4, 0, str_repeat('д', 10240));
        $sheet->writeFormula(4, 1, '=LEN(A5)');

        $anotherSheet = $workbook->addWorksheet();

        $anotherSheet->write(0, 0, str_repeat('f', 9216));
        $anotherSheet->writeFormula(0, 1, '=LEN(A1)');

        $anotherSheet->write(1, 0, str_repeat('g', 10240));
        $anotherSheet->writeFormula(1, 1, '=LEN(A2)');

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('long_strings');
    }

    public function testFill()
    {
        $workbook = $this->createWorkbook();
        $workbook->setCountry($workbook::COUNTRY_NONE);

        $sheet = $workbook->addWorksheet();

        $format = $workbook->addFormat();
        $format->setColor('red');
        $format->setAlign('center');

        //intentionally blank string and number bigger than 63
        $format->setBgColor('');
        $format->setBgColor(75);

        $format->setFgColor('navy');
        $format->setPattern(Fill::PATTERN_DIAGONAL_STRIPE);

        $sheet->setRow(0, 75);
        $sheet->setColumn(0, 0, 50);
        $sheet->write(0, 0, 'Test', $format);

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('fill');
    }

    public function testMultipleSheetsFormulas()
    {
        return;
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();
        $sheet->write(0, 0, 5);

        $anotherSheet = $workbook->addWorksheet();
        $anotherSheet->writeFormula(0, 0, '=Sheet1!A1');

        $workbook->save($this->testFilePath);
        exit;

        $this->checkTestFileIsEqualTo('multiple_sheets_formulas');
    }
}

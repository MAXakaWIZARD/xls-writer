<?php
namespace Test;

use Xls\Fill;
use Xls\Workbook;

/**
 *
 */
class GeneralTest extends TestAbstract
{
    public function testGeneral()
    {
        $this->workbook->setCountry(Workbook::COUNTRY_USA);

        $sheet = $this->workbook->addWorksheet('My first worksheet');
        $sheet->writeRow(
            0,
            0,
            array(
                array('Name', 'John Smith', 'Johann Schmidt', 'Иван Иванов'),
                array('Age', 30, 31, 32)
            )
        );

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('general');

        $this->setExpectedException('\Exception', 'Workbook was already saved!');
        $this->workbook->save($this->testFilePath);
    }

    /**
     * @throws \Exception
     */
    public function testProtected()
    {
        $sheet = $this->workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->protect('1234');

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('protected');
    }

    public function testSelection()
    {
        $sheet = $this->workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');

        $sheet->setSelection(0, 0);
        $sheet->setSelection(5, 5, 0, 0);
        $sheet->setSelection(0, 0, 5, 5);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('selection');
    }

    public function testMultipleSheets()
    {
        $sheetNames = array(
            'First sheet',
            'Второй лист',
            'Third sheet',
            '4th sheet'
        );

        for ($i = 1; $i <= 4; $i++) {
            $s = $this->workbook->addWorksheet($sheetNames[$i - 1]);
            $s->write(0, 0, 'Test' . $i);
        }

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('multiple_sheets');
    }

    public function testDefcolsAndRowsizes()
    {
        $sheet = $this->workbook->addWorksheet();
        $sheet->writeRow(0, 0, array('Test1', 'Test2', 'Test3'));

        $sheet->setColumnWidth(0, 0, 25);
        $sheet->setColumnWidth(1, 1, 50);
        $sheet->setColumnWidth(2, 3, 10, null, 1);

        $sheet->setRowHeight(0, 30);
        $sheet->setRowHeight(1, 15);
        $sheet->setRowHeight(2, 10, null, 1);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('defcols_rowsizes');
    }

    public function testImage()
    {
        $sheet = $this->workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->insertBitmap(2, 2, TEST_DATA_PATH . '/elephpant.bmp');

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('image');
    }

    public function testMergeCells()
    {
        $sheet = $this->workbook->addWorksheet();

        $sheet->writeRow(1, 0, array('Merge1', '', ''));
        $sheet->mergeCells(1, 0, 1, 4);
        $sheet->writeRow(2, 1, array('Merge2', '', ''));
        $sheet->mergeCells(2, 1, 2, 4);
        $sheet->writeRow(3, 2, array('Merge3', '', ''));
        $sheet->mergeCells(3, 2, 3, 4);

        $format = $this->workbook->addFormat();
        $format->setAlign('center');
        $sheet->writeRow(4, 3, array('Merge4', '', ''), $format);
        $sheet->mergeCells(4, 3, 5, 4);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('merge');
    }

    public function testThawPanes()
    {
        $sheet = $this->workbook->addWorksheet();

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

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('thaw_panes');
    }

    public function testFreezePanes()
    {
        $sheet = $this->workbook->addWorksheet();

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

        $sheet->freezePanes(array(1, 1));

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('freeze_panes');
    }

    public function testLongStrings()
    {
        $sheet = $this->workbook->addWorksheet();

        //keep for full test coverage
        $sheet->write(0, 0, str_repeat('a', 41));
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

        $anotherSheet = $this->workbook->addWorksheet();

        $anotherSheet->write(0, 0, str_repeat('f', 9216));
        $anotherSheet->writeFormula(0, 1, '=LEN(A1)');

        $anotherSheet->write(1, 0, str_repeat('g', 10240));
        $anotherSheet->writeFormula(1, 1, '=LEN(A2)');

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('long_strings');
    }

    public function testFill()
    {
        $sheet = $this->workbook->addWorksheet();

        $format = $this->workbook->addFormat();
        $format->setColor('red');
        $format->setAlign('center');

        //intentionally blank string and number bigger than 63
        $format->setBgColor('');
        $format->setBgColor(75);

        $format->setFgColor('navy');
        $format->setPattern(Fill::PATTERN_DIAGONAL_STRIPE);

        $sheet->setRowHeight(0, 75);
        $sheet->setColumnWidth(0, 0, 50);
        $sheet->write(0, 0, 'Test', $format);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('fill');
    }

    public function testMultipleSheetsFormulas()
    {
        $sheet = $this->workbook->addWorksheet();
        $sheet->write(0, 0, 5);

        $anotherSheet = $this->workbook->addWorksheet();
        $anotherSheet->writeFormula(0, 0, '=Sheet1!A1 * 5');

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('cross_sheets_formulas');
    }
}

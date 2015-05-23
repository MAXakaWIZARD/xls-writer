<?php
namespace Test;

/**
 *
 */
class LayoutsTest extends TestAbstract
{
    public function testPortraitLayout()
    {
        $sheet = $this->workbook->addWorksheet();
        $row = array(
            'Portrait layout test',
            '1',
            '2',
            '3',
            '4',
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

        $sheet->setPrintArea(0, 0, 5, 5);

        $sheet->hidePrintGridlines();
        $sheet->setHPagebreaks(array(1));
        $sheet->setVPagebreaks(array(5));

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('layout_portrait');
    }

    public function testPrintRepeatRow()
    {
        $sheet = $this->workbook->addWorksheet();

        $sheet->writeRow(0, 0, array('ID', 'Name'));

        $ids = range(1, 65);
        foreach ($ids as $id) {
            $sheet->write($id, 0, $id);
            $sheet->write($id, 1, 'Name' . $id);
        }

        $sheet->printRepeatRows(0);
        $sheet->printRepeatRows(0, 0);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('print_repeat_row');
    }

    public function testPrintRepeatCol()
    {
        $sheet = $this->workbook->addWorksheet();

        $fields = range(1, 15);
        $fieldValues = array();
        $headers = array('ID', 'Name');
        foreach ($fields as $idx) {
            $headers[] = 'Field' . $idx;
            $fieldValues[] = 'Value ' . $idx;
        }

        $sheet->writeRow(0, 0, $headers);

        $ids = range(1, 10);
        foreach ($ids as $id) {
            $sheet->write($id, 0, $id);
            $sheet->write($id, 1, 'Name' . $id);
            $sheet->writeRow($id, 2, $fieldValues);
        }

        $sheet->printRepeatColumns(0);
        $sheet->printRepeatColumns(0, 0);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('print_repeat_col');
    }

    public function testPrintRepeatRowCol()
    {
        $sheet = $this->workbook->addWorksheet();

        $fields = range(1, 15);
        $fieldValues = array();
        $headers = array('ID', 'Name');
        foreach ($fields as $idx) {
            $headers[] = 'Field' . $idx;
            $fieldValues[] = 'Value ' . $idx;
        }

        $sheet->writeRow(0, 0, $headers);

        $ids = range(1, 65);
        foreach ($ids as $id) {
            $sheet->write($id, 0, $id);
            $sheet->write($id, 1, 'Name' . $id);
            $sheet->writeRow($id, 2, $fieldValues);
        }

        $sheet->printRepeatRows(0);
        $sheet->printRepeatColumns(0);

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('print_repeat_rowcol');
    }

    public function testLandscapeLayout()
    {
        $sheet = $this->workbook->addWorksheet();

        $fields = range(1, 15);
        $fieldValues = array();
        $headers = array('ID', 'Name');
        foreach ($fields as $idx) {
            $headers[] = 'Field' . $idx;
            $fieldValues[] = 'Value ' . $idx;
        }

        $sheet->writeRow(0, 0, $headers);

        $ids = range(1, 65);
        foreach ($ids as $id) {
            $sheet->write($id, 0, $id);
            $sheet->write($id, 1, 'Name' . $id);
            $sheet->writeRow($id, 2, $fieldValues);
        }

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

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('layout_landscape');
    }
}

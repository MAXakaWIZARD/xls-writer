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

        $pageSetup = $sheet->getPageSetup();

        $pageSetup
            ->setZoom(125)
            ->setPortrait()
            ->setHeader('Page header')
            ->setFooter('Page footer')
            ->setPrintScale(90)
            ->setPaper($pageSetup::PAPER_A3)
            ->setPrintArea(0, 0, 5, 5)
            ->printGridlines(false)
            ->setHPagebreaks(array(1))
            ->setVPagebreaks(array(5))
        ;
        $pageSetup->getMargin()->setAll(1.25);

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

        $sheet->getPageSetup()
            ->printRepeatRows(0)
            ->printRepeatRows(0, 0)
        ;

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

        $sheet->getPageSetup()
            ->printRepeatColumns(0)
            ->printRepeatColumns(0, 0)
        ;

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

        $sheet->getPageSetup()
            ->printRepeatRows(0)
            ->printRepeatColumns(0)
        ;

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

        $pageSetup = $sheet->getPageSetup();

        //header and footer should be cut to max length (255)
        $pageSetup
            ->setHeader('Page header ' . str_repeat('.', 255))
            ->setFooter('Page footer ' . str_repeat('.', 255))
            ->showGridlines(false)
            ->setLandscape()
            ->centerHorizontally()
            ->centerVertically()
            ->fitToPages(999, 999)
            ->setPaper($pageSetup::PAPER_A4)
            ->printRowColHeaders()
        ;

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('layout_landscape');
    }
}

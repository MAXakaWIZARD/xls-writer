<?php
namespace Test;

/**
 *
 */
class LayoutsTest extends TestAbstract
{
    public function testPortrait2Layout()
    {
        return;
        $workbook = $this->createWorkbookBiff5();
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

        //check for biff8
        $sheet->printArea(0, 0, 10, 10);

        $sheet->hideGridlines();
        $sheet->setHPagebreaks(array(1));
        $sheet->setVPagebreaks(array(5));

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('layout_portrait');
    }

    public function testLandscapeLayout()
    {
        return;
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

        $this->checkTestFileIsEqualTo('layout_landscape');
    }
}

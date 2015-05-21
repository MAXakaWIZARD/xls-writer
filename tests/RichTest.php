<?php
namespace Test;

use Xls\Format;
use Xls\Fill;

/**
 *
 */
class RichTest extends TestAbstract
{
    /**
     * @throws \Exception
     */
    public function testRich()
    {
        $workbook = $this->createWorkbook();

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
        $sheet->write(11, 4, '20% скидка!', $discountFormat);

        $anotherSheet = $workbook->addWorksheet('Лист2');
        $anotherSheet->write(0, 0, 'Тест');

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('rich');
    }
}

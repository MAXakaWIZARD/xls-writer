<?php
namespace Test;

/**
 *
 */
class ValidationTest extends TestAbstract
{
    public function testValidation()
    {
        $workbook = $this->createWorkbook();

        $sheet = $workbook->addWorksheet();

        $sheet->setColumn(0, 0, 20);

        //positive number
        $validator = $workbook->addValidator();
        $validator->setPrompt('Enter positive number', 'Enter positive number');
        $validator->setError('Invalid number', 'Number should be bigger than zero');
        $validator->setDataType($validator::TYPE_INTEGER);
        $validator->allowBlank();
        $validator->setOperator($validator::OP_GT);
        $validator->setFormula1('0');
        $validator->onInvalidStop();
        $sheet->write(0, 0, 'Enter positive number:');
        $sheet->setValidation(0, 1, 0, 1, $validator);

        //number in range
        $validator = $workbook->addValidator();
        $validator->setPrompt('Enter month number', 'Enter month number');
        $validator->setError('Invalid month', 'Number should be in range from 1 to 12');
        $validator->allowBlank();
        $validator->setOperator($validator::OP_BETWEEN);
        $validator->setFormula1('1');
        $validator->setFormula2('12');
        $validator->onInvalidInfo();
        $sheet->write(1, 0, 'Enter month number:');
        $sheet->setValidation(1, 1, 1, 1, $validator);

        //value from list
        $validator = $workbook->addValidator();
        $validator->setPrompt('Select animal', 'Select animal');
        $validator->setError('Invalid selection', 'Select animal from list');
        $validator->setDataType($validator::TYPE_USER_LIST);
        $validator->setFormula1('"Cat,Dog,Mouse"');
        $validator->setShowDropDown();
        $validator->onInvalidWarn();
        $sheet->write(2, 0, 'Select animal:');
        $sheet->setValidation(2, 1, 2, 1, $validator);

        //value from list
        $validator = $workbook->addValidator();
        $validator->setPrompt('Select animal', 'Select animal');
        $validator->setError('Invalid selection', 'Select animal from list');
        $validator->setDataType($validator::TYPE_USER_LIST);
        $validator->setShowDropDown();
        $validator->setFormula1('$F$2:$F$5');
        $sheet->write(3, 0, 'Select animal:');
        $sheet->writeCol(0, 5, array('Animals:', 'Horse', 'Elephant', 'Whale', 'Squirrell'));
        $sheet->setValidation(3, 1, 3, 1, $validator);

        $workbook->save($this->testFilePath);

        $this->checkTestFileIsEqualTo('validation');
    }
}

<?php

require_once 'testDataFileIterator.php';

class FormulaAsStringTest extends PHPUnit_Framework_TestCase
{
    /**
     * @dataProvider providerFormulas
     */
    public function testFunctionsAsString2()
    {
        $spreadsheet = new PHPExcel();
        $workSheet = $spreadsheet->getActiveSheet();
        $workSheet->setCellValue('A1', 10);
        $workSheet->setCellValue('A2', 20);
        $workSheet->setCellValue('A3', 30);
        $workSheet->setCellValue('A4', 40);
        $spreadsheet->addNamedRange(new PHPExcel_NamedRange('namedCell', $workSheet, 'A4'));
        $workSheet->setCellValue('B1', 'uPPER');
        $workSheet->setCellValue('B2', '=TRUE()');
        $workSheet->setCellValue('B3', '=FALSE()');

        $ws2 = $spreadsheet->createSheet();
        $ws2->setCellValue('A1', 100);
        $ws2->setCellValue('A2', 200);
        $ws2->setTitle('Sheet2');
        $spreadsheet->addNamedRange(new PHPExcel_NamedRange('A2B', $ws2, 'A2'));
        $spreadsheet->setActiveSheetIndex(0);

        $args = func_get_args();
        $expectedResult = array_pop($args);
        $formula = array_pop($args);

        $cell2 = $workSheet->getCell('D1');
        $cell2->setValue($formula);
        $result = $cell2->getCalculatedValue();
        self::assertEquals($expectedResult, $result);
    }

    public function providerFormulas()
    {
        return new testDataFileIterator('rawTestData/Calculation/formulas.data');
    }
}

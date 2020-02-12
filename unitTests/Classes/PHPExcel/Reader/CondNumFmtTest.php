<?php

namespace PHPExcel;

use PHPExcel_Reader_Excel2007 as Xlsx;
use PHPExcel_Style_Conditional as Conditional;
use PHPUnit\Framework\TestCase;

class CondNumFmtTest extends TestCase
{
    public function testLoadCondNumFmt()
    {
        $filename = './data/Reader/XLSX/condfmtnum.xlsx';
        $reader = new Xlsx();
        $spreadsheet = $reader->load($filename);

        $worksheet = $spreadsheet->getActiveSheet();
        $conditionalStyle = $worksheet->getConditionalStyles('A1:A3');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_EXPRESSION, $conditionalRule->getConditionType());
        self::assertEquals('MONTH(A1)=10', $conditions[0]);
        $numfmt = $conditionalRule->getStyle()->getNumberFormat()->getFormatCode();
        self::assertEquals('yyyy/mm/dd', $numfmt);
    }
}

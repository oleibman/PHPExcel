<?php

namespace PHPExcel;

use PHPUnit\Framework\TestCase;

class FormulaErrTest extends TestCase
{
    const FILENAM = 'formulaerr';
    const TYP = 'Xls';
    const TY2 = 'Excel5';

    private static function getfilename()
    {
        return self::FILENAM . '.' . strtolower(self::TYP);
    }

    protected function tearDown(): void
    {
        $out = self::getfilename();
        if (file_exists($out)) {
            unlink($out);
        }
    }

    public function testFormulaError()
    {
        $obj = new \PHPExcel();
        $sheet0 = $obj->setActiveSheetIndex(0);
        $sheet0->setCellValue('A1', 2);
        $obj->addNamedRange(new \PHPExcel_NamedRange('DEFNAM', $sheet0, 'A1'));
        $sheet0->setCellValue('B1', '=2*DEFNAM');
        $outtype = self::TY2;
        $writer = \PHPExcel_IOFactory::createWriter($obj, $outtype);
        $out = self::getfilename();
        $writer->save($out);
        $reader = \PHPExcel_IOFactory::createReader($outtype);
        $robj = $reader->load($out);
        $sheet0 = $robj->setActiveSheetIndex(0);
        $a1 = $sheet0->getCell('A1')->getCalculatedValue();
        self::assertEquals(2, $a1);
        $b1 = $sheet0->getCell('B1')->getCalculatedValue();
        self::assertEquals(4, $b1);
    }
}

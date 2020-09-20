<?php

namespace PHPExcel_Calculation;

use PHPUnit\Framework\TestCase;

class DefinedNameConfusedForCellTest extends TestCase
{
    const FILENAM = 'calcerror';
    const TYP = 'Xlsx';
    const TY2 = 'Excel2007';

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

    public function testDefinedName()
    {
        $obj = new \PHPExcel();
        $sheet0 = $obj->setActiveSheetIndex(0);
        $sheet0->setCellValue('A1', 2);
        $obj->addNamedRange(new \PHPExcel_NamedRange('A1A', $sheet0, 'A1'));
        $sheet0->setCellValue('B1', '=2*A1A');
        $writer = \PHPExcel_IOFactory::createWriter($obj, self::TY2);
        $out = self::getfilename();
        $writer->save($out);
        self::assertTrue(true);
    }
}

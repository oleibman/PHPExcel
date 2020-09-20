<?php

namespace PHPExcel;

use PHPExcel_IOFactory as IOFactory;
use PHPExcel_Style_Color as Color;
use PHPExcel_Style_Conditional as Conditional;
use PHPExcel_Style_Fill as Fill;
use PHPUnit\Framework\TestCase;

class Xlsx2Test extends TestCase
{
    const FILENAM = 'conditionalFormatting2Test';
    const TYP = 'Xlsx';
    const TY2 = 'Excel2007';

    private function inputfilename()
    {
        $typ = strtolower(self::TYP);
        $dir = strtoupper(self::TYP);

        return "./data/Reader/$dir/" . self::FILENAM . ".$typ";
    }

    private function outputfilename()
    {
        $typ = strtolower(self::TYP);

        return self::FILENAM . "2.$typ";
    }

    protected function tearDown(): void
    {
        $out = self::outputfilename();
        if (file_exists($out)) {
            unlink($out);
        }
    }

    public function testLoadXlsxConditionalFormatting2()
    {
        $filename = self::inputfilename();
        $reader = IOFactory::createReader(self::TY2);
        $spreadsheet = $reader->load($filename);
        $worksheet = $spreadsheet->getActiveSheet();

        $conditionalStyle = $worksheet->getConditionalStyles('A2:A8');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_NOTCONTAINSBLANKS, $conditionalRule->getConditionType());
        self::assertEquals('LEN(TRIM(A2))>0', $conditions[0]);

        $conditionalStyle = $worksheet->getConditionalStyles('B2:B8');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_CONTAINSBLANKS, $conditionalRule->getConditionType());
        self::assertEquals('LEN(TRIM(B2))=0', $conditions[0]);

        $conditionalStyle = $worksheet->getConditionalStyles('C2:C8');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_CELLIS, $conditionalRule->getConditionType());
        self::assertEquals(Conditional::OPERATOR_GREATERTHAN, $conditionalRule->getOperatorType());
        self::assertEquals('5', $conditions[0]);
    }

    public function testReloadXlsxConditionalFormatting2()
    {
        $filename = self::inputfilename();
        $outfile = self::outputfilename();
        $reader = IOFactory::createReader(self::TY2);
        $spreadshee1 = $reader->load($filename);
        $writer = IOFactory::createWriter($spreadshee1, self::TY2);
        $writer->save($outfile);
        $spreadsheet = $reader->load($outfile);
        $worksheet = $spreadsheet->getActiveSheet();

        $conditionalStyle = $worksheet->getConditionalStyles('A2:A8');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_NOTCONTAINSBLANKS, $conditionalRule->getConditionType());
        self::assertEquals('LEN(TRIM(A2:A8))>0', $conditions[0]);

        $conditionalStyle = $worksheet->getConditionalStyles('B2:B8');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_CONTAINSBLANKS, $conditionalRule->getConditionType());
        self::assertEquals('LEN(TRIM(B2:B8))=0', $conditions[0]);

        $conditionalStyle = $worksheet->getConditionalStyles('C2:C8');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_CELLIS, $conditionalRule->getConditionType());
        self::assertEquals(Conditional::OPERATOR_GREATERTHAN, $conditionalRule->getOperatorType());
        self::assertEquals('5', $conditions[0]);
    }

    public function testNewXlsxConditionalFormatting2()
    {
        $outfile = self::outputfilename();
        $spreadshee1 = new \PHPExcel();
        $sheet = $spreadshee1->getActiveSheet();
        $sheet->setCellValue('A2', 'a2');
        $sheet->setCellValue('A4', 'a4');
        $sheet->setCellValue('A6', 'a6');
        $cond1 = new Conditional();
        $cond1->setConditionType(Conditional::CONDITION_CONTAINSBLANKS);
        $cond1->getStyle()->getFill()->setFillType(Fill::FILL_SOLID);
        $cond1->getStyle()->getFill()->getEndColor()->setARGB(Color::COLOR_RED);
        $cond = [$cond1];
        $sheet->getStyle('A1:A6')->setConditionalStyles($cond);
        $writer = IOFactory::createWriter($spreadshee1, self::TY2);
        $writer->save($outfile);
        $reader = IOFactory::createReader(self::TY2);
        $spreadsheet = $reader->load($outfile);
        $worksheet = $spreadsheet->getActiveSheet();

        $conditionalStyle = $worksheet->getConditionalStyles('A1:A6');
        self::assertNotEmpty($conditionalStyle);
        $conditionalRule = $conditionalStyle[0];
        $conditions = $conditionalRule->getConditions();
        self::assertNotEmpty($conditions);
        self::assertEquals(Conditional::CONDITION_CONTAINSBLANKS, $conditionalRule->getConditionType());
        self::assertEquals('LEN(TRIM(A1:A6))=0', $conditions[0]);
    }
}

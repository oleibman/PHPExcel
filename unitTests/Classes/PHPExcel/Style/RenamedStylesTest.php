<?php
class RenamedStylesTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
    }

    public function testArrayBorderStyle1()
    {
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_Border::BORDER_THIN;
        $styleThinBlackBorderOutline1 = array(
            'borders' => array(
                'outline' => array(
                    'borderStyle' => $expected,
                    'color' => array('argb' => 'FF000000'),
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleThinBlackBorderOutline1);
        $actual = $sheet->getStyle('A1')->getBorders()->getTop()->getBorderStyle();
        self::assertEquals($expected, $actual);
    }

    public function testArrayBorderStyle2()
    {
        //self::expectException(\PHPUnit\Framework\Error\Warning::class);
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_Border::BORDER_DOTTED;
        $styleThinBlackBorderOutline1 = array(
            'borders' => array(
                'outline' => array(
                    'style' => $expected,
                    'color' => array('argb' => 'FF000000'),
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleThinBlackBorderOutline1);
        $actual = $sheet->getStyle('A1')->getBorders()->getTop()->getBorderStyle();
        self::assertEquals($expected, $actual);
    }

    public function testArrayBorderStyle3()
    {
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_Border::BORDER_THIN;
        $unexpected = PHPExcel_Style_Border::BORDER_DOTTED;
        $styleThinBlackBorderOutline1 = array(
            'borders' => array(
                'outline' => array(
                    'borderStyle' => $expected,
                    'style' => $unexpected,
                    'color' => array('argb' => 'FF000000'),
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleThinBlackBorderOutline1);
        $actual = $sheet->getStyle('A1')->getBorders()->getTop()->getBorderStyle();
        self::assertEquals($expected, $actual);
    }

    public function testNumberFormatStyle1()
    {
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
        $styleNumberFormat1 = array(
            'numberFormat' => array('formatCode' => $expected),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleNumberFormat1);
        $actual = $sheet->getStyle('A1')->getNumberFormat()->getFormatCode();
        self::assertEquals($expected, $actual);
    }

    public function testNumberFormatStyle2()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
        $styleNumberFormat1 = array(
            'numberFormat' => array('code' => $expected),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleNumberFormat1);
        $actual = $sheet->getStyle('A1')->getNumberFormat()->getFormatCode();
        self::assertEquals($expected, $actual);
    }

    public function testNumberFormatStyle3()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
        $styleNumberFormat1 = array(
            'numberformat' => array('formatCode' => $expected),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleNumberFormat1);
        $actual = $sheet->getStyle('A1')->getNumberFormat()->getFormatCode();
        self::assertEquals($expected, $actual);
    }

    public function testNumberFormatStyle4()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expected = PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
        $styleNumberFormat1 = array(
            'numberformat' => array('code' => $expected),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleNumberFormat1);
        $actual = $sheet->getStyle('A1')->getNumberFormat()->getFormatCode();
        self::assertEquals($expected, $actual);
    }

    public function testAlignmentStyle1()
    {
        //self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expectedrot = 5;
        $expectedrdord = PHPExcel_Style_Alignment::READORDER_LTR;
        $expectedwrap = true;
        $styleAlignment1 = array(
            'alignment' => array(
                 'textRotation' => $expectedrot,
                 'readOrder' => $expectedrdord,
                 'wrapText' => $expectedwrap,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleAlignment1);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertEquals($expectedrot, $align->getTextRotation());
        self::assertEquals($expectedrdord, $align->getReadOrder());
        self::assertEquals($expectedwrap, $align->getWrapText());
    }

    public function testAlignmentStyle2()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expectedrot = 5;
        $expectedrdord = PHPExcel_Style_Alignment::READORDER_LTR;
        $expectedwrap = true;
        $styleAlignment1 = array(
            'alignment' => array(
                 'rotation' => $expectedrot,
                 'readOrder' => $expectedrdord,
                 'wrapText' => $expectedwrap,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleAlignment1);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertEquals($expectedrot, $align->getTextRotation());
        self::assertEquals($expectedrdord, $align->getReadOrder());
        self::assertEquals($expectedwrap, $align->getWrapText());
    }

    public function testAlignmentStyle3()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expectedrot = 5;
        $expectedrdord = PHPExcel_Style_Alignment::READORDER_LTR;
        $expectedwrap = true;
        $styleAlignment1 = array(
            'alignment' => array(
                 'textRotation' => $expectedrot,
                 'readorder' => $expectedrdord,
                 'wrapText' => $expectedwrap,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleAlignment1);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertEquals($expectedrot, $align->getTextRotation());
        self::assertEquals($expectedrdord, $align->getReadOrder());
        self::assertEquals($expectedwrap, $align->getWrapText());
    }

    public function testAlignmentStyle4()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $expectedrot = 5;
        $expectedrdord = PHPExcel_Style_Alignment::READORDER_LTR;
        $expectedwrap = true;
        $styleAlignment1 = array(
            'alignment' => array(
                 'textRotation' => $expectedrot,
                 'readOrder' => $expectedrdord,
                 'wrap' => $expectedwrap,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleAlignment1);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertEquals($expectedrot, $align->getTextRotation());
        self::assertEquals($expectedrdord, $align->getReadOrder());
        self::assertEquals($expectedwrap, $align->getWrapText());
    }

    public function testFontStyle1()
    {
        //self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $styleFont1 = array(
            'font' => array(
                 'strikethrough' => true,
             ),
        );
        $styleFont2 = array(
            'font' => array(
                 'superscript' => true,
             ),
        );
        $styleFont3 = array(
            'font' => array(
                 'subscript' => true,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFont1);
        $sheet->setCellValue('A2', 2);
        $sheet->getStyle('A2')->applyFromArray($styleFont2);
        $sheet->setCellValue('A3', 3);
        $sheet->getStyle('A3')->applyFromArray($styleFont3);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertTrue($sheet->getStyle('A1')->getFont()->getStrikethrough());
        self::assertTrue($sheet->getStyle('A2')->getFont()->getSuperscript());
        self::assertTrue($sheet->getStyle('A3')->getFont()->getSubscript());
    }

    public function testFontStyle2()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $styleFont1 = array(
            'font' => array(
                 'strike' => true,
             ),
        );
        $styleFont2 = array(
            'font' => array(
                 'superscript' => true,
             ),
        );
        $styleFont3 = array(
            'font' => array(
                 'subscript' => true,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFont1);
        $sheet->setCellValue('A2', 2);
        $sheet->getStyle('A2')->applyFromArray($styleFont2);
        $sheet->setCellValue('A3', 3);
        $sheet->getStyle('A3')->applyFromArray($styleFont3);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertTrue($sheet->getStyle('A1')->getFont()->getStrikethrough());
        self::assertTrue($sheet->getStyle('A2')->getFont()->getSuperscript());
        self::assertTrue($sheet->getStyle('A3')->getFont()->getSubscript());
    }

    public function testFontStyle3()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $styleFont1 = array(
            'font' => array(
                 'strikethrough' => true,
             ),
        );
        $styleFont2 = array(
            'font' => array(
                 'superScript' => true,
             ),
        );
        $styleFont3 = array(
            'font' => array(
                 'subscript' => true,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFont1);
        $sheet->setCellValue('A2', 2);
        $sheet->getStyle('A2')->applyFromArray($styleFont2);
        $sheet->setCellValue('A3', 3);
        $sheet->getStyle('A3')->applyFromArray($styleFont3);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertTrue($sheet->getStyle('A1')->getFont()->getStrikethrough());
        self::assertTrue($sheet->getStyle('A2')->getFont()->getSuperscript());
        self::assertTrue($sheet->getStyle('A3')->getFont()->getSubscript());
    }

    public function testFontStyle4()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $styleFont1 = array(
            'font' => array(
                 'strikethrough' => true,
             ),
        );
        $styleFont2 = array(
            'font' => array(
                 'superscript' => true,
             ),
        );
        $styleFont3 = array(
            'font' => array(
                 'subScript' => true,
             ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFont1);
        $sheet->setCellValue('A2', 2);
        $sheet->getStyle('A2')->applyFromArray($styleFont2);
        $sheet->setCellValue('A3', 3);
        $sheet->getStyle('A3')->applyFromArray($styleFont3);
        $align = $sheet->getStyle('A1')->getAlignment();
        self::assertTrue($sheet->getStyle('A1')->getFont()->getStrikethrough());
        self::assertTrue($sheet->getStyle('A2')->getFont()->getSuperscript());
        self::assertTrue($sheet->getStyle('A3')->getFont()->getSubscript());
    }

    public function testFillStyle1()
    {
        //self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $ftype = PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR;
        $scolor = PHPExcel_Style_Color::COLOR_RED;
        $ecolor = PHPExcel_Style_Color::COLOR_DARKRED;
        $styleFill1 = array(
            'fill' => array(
                'fillType' => $ftype,
                'rotation' => 90,
                'startColor' => array(
                    'argb' => $scolor,
                ),
                'endColor' => array(
                    'argb' => $ecolor,
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFill1);
        $fill = $sheet->getStyle('A1')->getFill();
        self::assertEquals($ftype, $fill->getFillType());
        self::assertEquals($scolor, $fill->getStartColor()->getARGB());
        self::assertEquals($ecolor, $fill->getEndColor()->getARGB());
    }

    public function testFillStyle2()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $ftype = PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR;
        $scolor = PHPExcel_Style_Color::COLOR_RED;
        $ecolor = PHPExcel_Style_Color::COLOR_DARKRED;
        $styleFill1 = array(
            'fill' => array(
                'type' => $ftype,
                'rotation' => 90,
                'startColor' => array(
                    'argb' => $scolor,
                ),
                'endColor' => array(
                    'argb' => $ecolor,
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFill1);
        $fill = $sheet->getStyle('A1')->getFill();
        self::assertEquals($ftype, $fill->getFillType());
        self::assertEquals($scolor, $fill->getStartColor()->getARGB());
        self::assertEquals($ecolor, $fill->getEndColor()->getARGB());
    }

    public function testFillStyle3()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $ftype = PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR;
        $scolor = PHPExcel_Style_Color::COLOR_RED;
        $ecolor = PHPExcel_Style_Color::COLOR_DARKRED;
        $styleFill1 = array(
            'fill' => array(
                'fillType' => $ftype,
                'rotation' => 90,
                'startcolor' => array(
                    'argb' => $scolor,
                ),
                'endColor' => array(
                    'argb' => $ecolor,
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFill1);
        $fill = $sheet->getStyle('A1')->getFill();
        self::assertEquals($ftype, $fill->getFillType());
        self::assertEquals($scolor, $fill->getStartColor()->getARGB());
        self::assertEquals($ecolor, $fill->getEndColor()->getARGB());
    }

    public function testFillStyle4()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $ftype = PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR;
        $scolor = PHPExcel_Style_Color::COLOR_RED;
        $ecolor = PHPExcel_Style_Color::COLOR_DARKRED;
        $styleFill1 = array(
            'fill' => array(
                'fillType' => $ftype,
                'rotation' => 90,
                'startColor' => array(
                    'argb' => $scolor,
                ),
                'endcolor' => array(
                    'argb' => $ecolor,
                ),
            ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleFill1);
        $fill = $sheet->getStyle('A1')->getFill();
        self::assertEquals($ftype, $fill->getFillType());
        self::assertEquals($scolor, $fill->getStartColor()->getARGB());
        self::assertEquals($ecolor, $fill->getEndColor()->getARGB());
    }

    public function testBordersStyle1()
    {
        //self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $dtype = PHPExcel_Style_Borders::DIAGONAL_BOTH;
        $bcolor = '808080';
        $styleBorders1 = array(
            'borders' => array(
                'diagonalDirection' => $dtype,
                'allBorders' => array(
                    'color' => array('rgb' => $bcolor),
                    ),
                ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleBorders1);
        $borders = $sheet->getStyle('A1')->getBorders();
        self::assertEquals($dtype, $borders->getDiagonalDirection());
        self::assertEquals($bcolor, $borders->getTop()->getColor()->getRGB());
    }

    public function testBordersStyle2()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $dtype = PHPExcel_Style_Borders::DIAGONAL_BOTH;
        $bcolor = '808080';
        $styleBorders1 = array(
            'borders' => array(
                'diagonaldirection' => $dtype,
                'allBorders' => array(
                    'color' => array('rgb' => $bcolor),
                    ),
                ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleBorders1);
        $borders = $sheet->getStyle('A1')->getBorders();
        self::assertEquals($dtype, $borders->getDiagonalDirection());
        self::assertEquals($bcolor, $borders->getTop()->getColor()->getRGB());
    }

    public function testBordersStyle3()
    {
        self::expectException('PHPUnit\\Framework\\Error\\Warning');
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $dtype = PHPExcel_Style_Borders::DIAGONAL_BOTH;
        $bcolor = '808080';
        $styleBorders1 = array(
            'borders' => array(
                'diagonalDirection' => $dtype,
                'allborders' => array(
                    'color' => array('rgb' => $bcolor),
                    ),
                ),
        );
        $sheet->setCellValue('A1', 1);
        $sheet->getStyle('A1')->applyFromArray($styleBorders1);
        $borders = $sheet->getStyle('A1')->getBorders();
        self::assertEquals($dtype, $borders->getDiagonalDirection());
        self::assertEquals($bcolor, $borders->getTop()->getColor()->getRGB());
    }
}

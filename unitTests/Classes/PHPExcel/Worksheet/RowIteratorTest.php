<?php

class RowIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRow;

    protected function setUp(): void
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
        
        $this->mockRow = $this->getMockBuilder('PHPExcel_Worksheet_Row')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('PHPExcel_Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
    }


    public function testIteratorFullRange()
    {
        $iterator = new PHPExcel_Worksheet_RowIterator($this->mockWorksheet);
        $rowIndexResult = 1;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('PHPExcel_Worksheet_Row', $row);
        }
    }

    public function testIteratorStartEndRange()
    {
        $iterator = new PHPExcel_Worksheet_RowIterator($this->mockWorksheet, 2, 4);
        $rowIndexResult = 2;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('PHPExcel_Worksheet_Row', $row);
        }
    }

    public function testIteratorSeekAndPrev()
    {
        $iterator = new PHPExcel_Worksheet_RowIterator($this->mockWorksheet, 2, 4);
        $rowIndexResult = 4;
        $iterator->seek($rowIndexResult);
        $this->assertEquals($rowIndexResult, $iterator->key());

        for ($i = 1; $i < $rowIndexResult-1; $i++) {
            $iterator->prev();
            $this->assertEquals($rowIndexResult - $i, $iterator->key());
        }
    }

    public function testStartOutOfRange()
    {
        $this->expectException(PHPExcel_Exception::class);
        $iterator = new PHPExcel_Worksheet_RowIterator($this->mockWorksheet, 256, 512);
    }

    public function testSeekOutOfRange()
    {
        $this->expectException(PHPExcel_Exception::class);
        $iterator = new PHPExcel_Worksheet_RowIterator($this->mockWorksheet, 2, 4);
        $iterator->seek(1);
    }

    public function testPrevOutOfRange()
    {
        $this->expectException(PHPExcel_Exception::class);
        $iterator = new PHPExcel_Worksheet_RowIterator($this->mockWorksheet, 2, 4);
        $iterator->prev();
    }
}

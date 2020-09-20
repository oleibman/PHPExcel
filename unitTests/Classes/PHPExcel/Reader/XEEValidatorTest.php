<?php


class XEEValidatorTest extends PHPUnit_Framework_TestCase
{

    protected function setUp(): void
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
    }

    /**
     * @dataProvider providerInvalidXML
     */
    public function testInvalidXML($filename)
    {
        $this->expectException(PHPExcel_Reader_Exception::class);
        $reader = $this->getMockForAbstractClass('PHPExcel_Reader_Abstract');
        $expectedResult = 'FAILURE: Should throw an Exception rather than return a value';
        $result = $reader->securityScanFile($filename);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerInvalidXML()
    {
        $tests = array();
        foreach (glob('rawTestData/Reader/XEETestInvalid*.xml') as $file) {
            $tests[] = [realpath($file), true];
        }
        return $tests;
    }

    /**
     * @dataProvider providerValidXML
     */
    public function testValidXML($filename, $expectedResult)
    {
        $reader = $this->getMockForAbstractClass('PHPExcel_Reader_Abstract');
        $result = $reader->securityScanFile($filename);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerValidXML()
    {
        $tests = array();
        foreach (glob('rawTestData/Reader/XEETestValid*.xml') as $file) {
            $tests[] = [realpath($file), file_get_contents($file)];
        }
        return $tests;
    }
}

<?php


class DataTypeTest extends PHPUnit_Framework_TestCase
{

    protected function setUp(): void
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
    }

    public function testGetErrorCodes()
    {
        $result = call_user_func(array('PHPExcel_Cell_DataType','getErrorCodes'));
        if (method_exists(self::class, 'assertIsArray')) {
            self::assertIsArray($result);
        } else {
            self::assertInternalType('array', $result);
        }
        $this->assertGreaterThan(0, count($result));
        $this->assertArrayHasKey('#NULL!', $result);
    }
}

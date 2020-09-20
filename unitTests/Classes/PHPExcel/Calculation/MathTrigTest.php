<?php


require_once 'testDataFileIterator.php';

class MathTrigTest extends PHPUnit_Framework_TestCase
{

    protected function setUp(): void
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');

        PHPExcel_Calculation_Functions::setCompatibilityMode(PHPExcel_Calculation_Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerATAN2
     */
    public function testATAN2()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','ATAN2'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerATAN2()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/ATAN2.data');
    }

    /**
     * @dataProvider providerCEILING
     */
    public function testCEILING()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','CEILING'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerCEILING()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/CEILING.data');
    }

    /**
     * @dataProvider providerCOMBIN
     */
    public function testCOMBIN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','COMBIN'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerCOMBIN()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/COMBIN.data');
    }

    /**
     * @dataProvider providerEVEN
     */
    public function testEVEN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','EVEN'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerEVEN()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/EVEN.data');
    }

    /**
     * @dataProvider providerODD
     */
    public function testODD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','ODD'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerODD()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/ODD.data');
    }

    /**
     * @dataProvider providerFACT
     */
    public function testFACT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','FACT'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerFACT()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/FACT.data');
    }

    /**
     * @dataProvider providerFACTDOUBLE
     */
    public function testFACTDOUBLE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','FACTDOUBLE'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerFACTDOUBLE()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/FACTDOUBLE.data');
    }

    /**
     * @dataProvider providerFLOOR
     */
    public function testFLOOR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','FLOOR'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerFLOOR()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/FLOOR.data');
    }

    /**
     * @dataProvider providerGCD
     */
    public function testGCD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','GCD'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerGCD()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/GCD.data');
    }

    /**
     * @dataProvider providerLCM
     */
    public function testLCM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','LCM'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerLCM()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/LCM.data');
    }

    /**
     * @dataProvider providerINT
     */
    public function testINT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','INT'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerINT()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/INT.data');
    }

    /**
     * @dataProvider providerSIGN
     */
    public function testSIGN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','SIGN'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerSIGN()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/SIGN.data');
    }

    /**
     * @dataProvider providerPOWER
     */
    public function testPOWER()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','POWER'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerPOWER()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/POWER.data');
    }

    /**
     * @dataProvider providerLOG
     */
    public function testLOG()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','LOG_BASE'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerLOG()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/LOG.data');
    }

    /**
     * @dataProvider providerMOD
     */
    public function testMOD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','MOD'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerMOD()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/MOD.data');
    }

    /**
     * @dataProvider providerMDETERM
     */
    public function testMDETERM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','MDETERM'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerMDETERM()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/MDETERM.data');
    }

    /**
     * @dataProvider providerMINVERSE
     */
    public function testMINVERSE()
    {
        //self::markTestSkipped('Owen initial failure');
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','MINVERSE'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerMINVERSE()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/MINVERSE.data');
    }

    /**
     * @dataProvider providerMMULT
     */
    public function testMMULT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','MMULT'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerMMULT()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/MMULT.data');
    }

    /**
     * @dataProvider providerMULTINOMIAL
     */
    public function testMULTINOMIAL()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','MULTINOMIAL'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerMULTINOMIAL()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/MULTINOMIAL.data');
    }

    /**
     * @dataProvider providerMROUND
     */
    public function testMROUND()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        PHPExcel_Calculation::setArrayReturnType(PHPExcel_Calculation::RETURN_ARRAY_AS_VALUE);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','MROUND'), $args);
        PHPExcel_Calculation::setArrayReturnType(PHPExcel_Calculation::RETURN_ARRAY_AS_ARRAY);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerMROUND()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/MROUND.data');
    }

    /**
     * @dataProvider providerPRODUCT
     */
    public function testPRODUCT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','PRODUCT'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerPRODUCT()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/PRODUCT.data');
    }

    /**
     * @dataProvider providerQUOTIENT
     */
    public function testQUOTIENT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','QUOTIENT'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerQUOTIENT()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/QUOTIENT.data');
    }

    /**
     * @dataProvider providerROUNDUP
     */
    public function testROUNDUP()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','ROUNDUP'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerROUNDUP()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/ROUNDUP.data');
    }

    /**
     * @dataProvider providerROUNDDOWN
     */
    public function testROUNDDOWN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','ROUNDDOWN'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerROUNDDOWN()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/ROUNDDOWN.data');
    }

    /**
     * @dataProvider providerSERIESSUM
     */
    public function testSERIESSUM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','SERIESSUM'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerSERIESSUM()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/SERIESSUM.data');
    }

    /**
     * @dataProvider providerSUMIFS
     */
    public function testSUMIFS()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','SUMIFS'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerSUMIFS()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/SUMIFS.data');
    }

    /**
     * @dataProvider providerSUMSQ
     */
    public function testSUMSQ()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','SUMSQ'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerSUMSQ()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/SUMSQ.data');
    }

    /**
     * @dataProvider providerTRUNC
     */
    public function testTRUNC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','TRUNC'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerTRUNC()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/TRUNC.data');
    }

    /**
     * @dataProvider providerROMAN
     */
    public function testROMAN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','ROMAN'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerROMAN()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/ROMAN.data');
    }

    /**
     * @dataProvider providerSQRTPI
     */
    public function testSQRTPI()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig','SQRTPI'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerSQRTPI()
    {
        return new testDataFileIterator('rawTestData/Calculation/MathTrig/SQRTPI.data');
    }

    /**
     * @dataProvider providerSUMIF
     */
    public function testSUMIF()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('PHPExcel_Calculation_MathTrig', 'SUMIF'), $args);
        if (method_exists($this, 'assertEqualsWithDelta')) {
            $this->assertEqualsWithDelta($expectedResult, $result, 1E-12);
        } else {
            $this->assertEquals($expectedResult, $result, null, 1E-12);
        }
    }

    public function providerSUMIF()
    {
        return array(
            array(
                array(
                    array(1),
                    array(5),
                    array(10),
                ),
                '>=5',
                15,
            ),
            array(
                array(
                    array('text'),
                    array(2),
                ),
                '=text',
                array(
                    array(10),
                    array(100),
                ),
                10,
            ),
            array(
                array(
                    array('"text with quotes"'),
                    array(2),
                ),
                '="text with quotes"',
                array(
                    array(10),
                    array(100),
                ),
                10,
            ),
            array(
                array(
                    array('"text with quotes"'),
                    array(''),
                ),
                '>"', // Compare to the single characater " (double quote)
                array(
                    array(10),
                    array(100),
                ),
                10
            ),
            array(
                array(
                    array(''),
                    array('anything'),
                ),
                '>"', // Compare to the single characater " (double quote)
                array(
                    array(10),
                    array(100),
                ),
                100
            ),
        );
    }
}

<?php

namespace Onemedia\SimpleExcelBundle\Lib;
/**
 * Created by JetBrains PhpStorm.
 * User: al
 * Date: 14.06.12
 * Time: 17:18
 * To change this template use File | Settings | File Templates.
 */

class SimpleExcelReader{
    public $skipAtEmptyRow = false;
    public $data = array();

    public static $_controlCharacters = array();

    public static function ControlCharacterOOXML2PHP($value = '') {
        return str_replace( array_keys(self::$_controlCharacters), array_values(self::$_controlCharacters), $value );
    }
    private static function _castToString($c) {
//		echo 'Initial Cast to String<br />';
        return isset($c->v) ? (string) $c->v : null;;
    }	//	function _castToString()
    private static function _buildControlCharacters() {
        for ($i = 0; $i <= 31; ++$i) {
            if ($i != 9 && $i != 10 && $i != 13) {
                $find = '_x' . sprintf('%04s' , strtoupper(dechex($i))) . '_';
                $replace = chr($i);
                self::$_controlCharacters[$find] = $replace;
            }
        }
    }
    public function load($file){
        $this->_buildControlCharacters();
        $zip = new ZipArchive;
        $zip->open($file);
        $rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels")); //~ http://schemas.openxmlformats.org/package/2006/relationships");
        foreach ($rels->Relationship as $rel) {
            if($rel['Type'] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"){
                $dir = dirname($rel["Target"]);
                $relsWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/_rels/" . basename($rel["Target"]) . ".rels"));  //~ http://schemas.openxmlformats.org/package/2006/relationships");
                $relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

                $sharedStrings = array();
                $xpath = self::array_item($relsWorkbook->xpath("rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings']"));
                $xmlStrings = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$xpath[Target]"));  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                if (isset($xmlStrings) && isset($xmlStrings->si)) {
                    foreach ($xmlStrings->si as $val) {
                        if (isset($val->t)) {
                            $sharedStrings[] = self::ControlCharacterOOXML2PHP( (string) $val->t );
                        } elseif (isset($val->r)) {
                            $sharedStrings[] = $this->_parseRichText($val);
                        }
                    }
                }


                $worksheets = array();
                foreach ($relsWorkbook->Relationship as $ele) {
                    if ($ele["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
                        $worksheets[(string) $ele["Id"]] = $ele["Target"];
                    }
                }

                $xmlWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "{$rel['Target']}"));

                foreach ($xmlWorkbook->sheets->sheet as $eleSheet) {
                    $fileWorksheet = $worksheets[(string) self::array_item($eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];

                    $xmlSheet = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$fileWorksheet"));

                    $cleanArray = array();
                    $emptyRowCount = 0;
                    $i = 0;
                    foreach ($xmlSheet->sheetData->row as $row) {

                        if($this->skipAtEmptyRow == false || $this->skipAtEmptyRow  > $emptyRowCount){
                            $cleanArray[$i] = array();
                            foreach ($row->c as $c) {

                                $r 					= (string) $c["r"];
                                $cellDataType 		= (string) $c["t"];
                                $value				= null;
                                $calculatedValue 	= null;

                                switch ($cellDataType) {
                                    case "s":
                                        //											echo 'String<br />';
                                        if ((string)$c->v != '') {
                                            $value = $sharedStrings[intval($c->v)];

                                            if ($value instanceof PHPExcel_RichText) {
                                                $value = clone $value;
                                            }
                                        } else {
                                            $value = '';
                                        }

                                        break;
                                    case "b":
                                        //											echo 'Boolean<br />';
                                        if (!isset($c->f)) {
                                            $value = self::_castToBool($c);
                                        } else {
                                            // Formula
                                            throw new Exception('Formulas not implemented');
                                            /*
                                            $this->_castToFormula($c,$r,$cellDataType,$value,$calculatedValue,$sharedFormulas,'_castToBool');
                                            if (isset($c->f['t'])) {
                                                $att = array();
                                                $att = $c->f;
                                                $docSheet->getCell($r)->setFormulaAttributes($att);
                                            }
                                            */

                                        }
                                        break;
                                    case "inlineStr":
                                        throw new Exception('Richtext not implemented');
                                        /*
                                        $value = $this->_parseRichText($c->is);
                                        */
                                        break;
                                    case "e":
                                        //											echo 'Error<br />';
                                        if (!isset($c->f)) {
                                            $value = self::_castToError($c);
                                        } else {
                                            // Formula
                                            throw new Exception('Formulas not implemented');
                                            /*
                                            $this->_castToFormula($c,$r,$cellDataType,$value,$calculatedValue,$sharedFormulas,'_castToError');
                                            //												echo '$calculatedValue = '.$calculatedValue.'<br />';
                                            */
                                        }

                                        break;

                                    default:
                                        //											echo 'Default<br />';
                                        if (!isset($c->f)) {
                                            //			echo 'Not a Formula<br />';
                                            $value = self::_castToString($c);


                                        } else {
                                            //												echo 'Treat as Formula<br />';
                                            // Formula
                                            throw new Exception('Formulas not implemented');
                                            /*
                                            $this->_castToFormula($c,$r,$cellDataType,$value,$calculatedValue,$sharedFormulas,'_castToString');
                                            //												echo '$calculatedValue = '.$calculatedValue.'<br />';
                                            */
                                        }

                                        break;
                                }
                                $cleanArray[$i][] = $value;

                                if(trim($value) == ""){
                                    $emptyRowCount++;
                                }



                            }
                        }else{
                            break;
                        }
                        $i++;
                    }
                    $this->data[] =  $cleanArray;
                }
            }
        }

        return $this->data;
    }
    private static function array_item($array, $key = 0) {
        return (isset($array[$key]) ? $array[$key] : null);
    }
    public function coordinateFromString($pCoordinateString = 'A1')
    {
        if (preg_match("/^([$]?[A-Z]{1,3})([$]?\d{1,7})$/", $pCoordinateString, $matches)) {
            return array($matches[1],$matches[2]);
        } elseif ((strpos($pCoordinateString,':') !== false) || (strpos($pCoordinateString,',') !== false)) {
            throw new Exception('Cell coordinate string can not be a range of cells.');
        } elseif ($pCoordinateString == '') {
            throw new Exception('Cell coordinate can not be zero-length string.');
        } else {
            throw new Exception('Invalid cell coordinate '.$pCoordinateString);
        }
    }

    public function _getFromZipArchive(ZipArchive $archive, $fileName = '')
    {
        // Root-relative paths
        if (strpos($fileName, '//') !== false)
        {
            $fileName = substr($fileName, strpos($fileName, '//') + 1);
        }
        $fileName = $this->realpath($fileName);

        // Apache POI fixes
        $contents = $archive->getFromName($fileName);
        if ($contents === false)
        {
            $contents = $archive->getFromName(substr($fileName, 1));
        }

        return $contents;
    }

    function realpath($pFilename) {
        // Returnvalue
        $returnValue = '';

        // Try using realpath()
        if (file_exists($pFilename)) {
            $returnValue = realpath($pFilename);
        }

        // Found something?
        if ($returnValue == '' || is_null($returnValue)) {
            $pathArray = explode('/' , $pFilename);
            while(in_array('..', $pathArray) && $pathArray[0] != '..') {
                for ($i = 0; $i < count($pathArray); ++$i) {
                    if ($pathArray[$i] == '..' && $i > 0) {
                        unset($pathArray[$i]);
                        unset($pathArray[$i - 1]);
                        break;
                    }
                }
            }
            $returnValue = implode('/', $pathArray);
        }

        // Return
        return $returnValue;
    }

}
<?php

namespace Onemedia\SimpleExcelBundle\Lib;
/**
 * Created by JetBrains PhpStorm.
 * User: al
 * Date: 14.06.12
 * Time: 17:18
 * To change this template use File | Settings | File Templates.
 */

class Reader2007{
    protected $skipAtEmptyRow = false;
    protected $deleteEmptyRows = false;
    protected $data = array();

    public static $_controlCharacters = array();


    function numtochars($number){
        $abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        $abc_len = strlen($abc);

        $result_len = 1; // how much characters the column's name will have
        $pow = 0;
        while( ( $pow += pow($abc_len, $result_len) ) < $number ){
            $result_len++;
        }

        $result = "";
        $next = false;
        // add each character to the result...
        for ($i = 1; $i<=$result_len; $i++) {
            $index = ($number % $abc_len) - 1; // calculate the module

            // sometimes the index should be decreased by 1
            if ( $next || $next = false ) {
                $index--;
            }

            // this is the point that will be calculated in the next iteration
            $number = floor($number / strlen($abc));

            // if the index is negative, convert it to positive
            if ( $next = ($index < 0) ) {
                $index = $abc_len + $index;
            }
            $result = $abc[$index].$result; // concatenate the letter
        }
        
        return $result;
    }


    public function __construct($configArray) {
        if (isset($configArray['skipAtEmptyRow'])) {
            $this->setSkipAtEmptyRow($configArray['skipAtEmptyRow']);
        }

        if (isset($configArray['deleteEmptyRows'])) {
            $this->setDeleteEmptyRows($configArray['deleteEmptyRows']);
        }
    }

    public function setSkipAtEmptyRow($rowNumber){

        $this->skipAtEmptyRow = intval($rowNumber);
    }

    public function setDeleteEmptyRows($flag){

        $this->deleteEmptyRows =  ($flag == true)? true : false;
    }

    protected function generateAllRowCols($spans){
        $spans = (string) $spans;
        $startEnd = explode(':',$spans);
        $col = array();
        for($i = $startEnd[0]; $i<$startEnd[1]; $i++){
            $col[$this->numtochars($i)] = '';

        }
        return $col;
    }

    private static function _castToError($c) {
        return isset($c->v) ? (string) $c->v : null;;
    }

    private static function _castToBool($c) {
        $value = isset($c->v) ? (string) $c->v : null;
        if ($value == '0') {
            return false;
        } elseif ($value == '1') {
            return true;
        } else {
            return (bool)$c->v;
        }
        return $value;
    }

    private static function _castToString($c) {
        return isset($c->v) ? (string) $c->v : null;;
    }

    private function splitMatrix($matrixId){
        $matrixSplit = str_split($matrixId);
        $row = '';
        $col = '';
        foreach($matrixSplit as $char){
            if(is_numeric($char)){
                $row.= $char;
            }else{
                $col.= $char;
            }
        }
        return array('row' => $row,'col' => $col);
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

    public function _getFromZipArchive(\ZipArchive $archive, $fileName = '')
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

    public static function ControlCharacterOOXML2PHP($value = '') {
        return str_replace( array_keys(self::$_controlCharacters), array_values(self::$_controlCharacters), $value );
    }

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
        $zip = new \ZipArchive;

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
                    if(count($xmlSheet->sheetData->row) > 0){
                        $patternRow = current($xmlSheet->sheetData->row);
                        $colsPattern = $this->generateAllRowCols($patternRow['spans']);
                    }
                    $rowNotEmpty = null;
                    foreach ($xmlSheet->sheetData->row as $row) {
                        if($rowNotEmpty === false){
                            $emptyRowCount++;
                            if($this->deleteEmptyRows === true){
                                unset($cleanArray[$matrix['row']]);
                            }
                        }else{
                            $emptyRowCount = 0;
                        }

                        if($this->skipAtEmptyRow == false || $this->skipAtEmptyRow  > $emptyRowCount){
                            $cleanArray[$row ['r']->__toString()]  = $colsPattern;
                            $rowNotEmpty = false;
                            foreach ($row->c as $c) {
                                $r 					= (string) $c["r"];
                                $cellDataType 		= (string) $c["t"];
                                $value				= null;

                                switch ($cellDataType) {
                                    case "s":
                                        $value = ((string)$c->v != '') ? $sharedStrings[intval($c->v)] : '';
                                        break;
                                    case "b":
                                        if (!isset($c->f)) {
                                            $value = self::_castToBool($c);
                                        } else {
                                            // Formula
                                            throw new Exception('Formulas not implemented');
                                        }
                                        break;
                                    case "inlineStr":
                                        throw new Exception('Richtext not implemented');
                                        break;
                                    case "e":
                                        if (!isset($c->f)) {
                                            $value = self::_castToError($c);
                                        } else {
                                            // Formula
                                            throw new Exception('Formulas not implemented');
                                        }
                                        break;
                                    default:
                                        if (!isset($c->f)) {
                                            $value = self::_castToString($c);
                                        } else {
                                            throw new Exception('Formulas not implemented');
                                        }

                                        break;
                                }
                                $matrix = $this->splitMatrix($r);

                                $cleanArray[$matrix['row']][$matrix['col']] = $value;

                                if(trim($value) != ""){
                                    $rowNotEmpty = true;
                                }
                            }
                        }else{
                            break;
                        }
                    }
                    $this->data[] =  $cleanArray;
                }
            }
        }
        return $this->data;
    }


}
<?php
require_once('PHPExcel/Classes/PHPExcel.php');
require_once('PHPExcel/Classes/PHPExcel/IOFactory.php');

$CLEAR_STATUS_ID = true;
$STOCK_STATUS_ID_DEFAULT = 5;
$STOCK_STATUS_IN_FEW_DAYS = 6;
$MAIN_FILE_NAME = "products.xlsx";
$STOCK_KIEV_FILE_NAME = "остатки Киев.xlsx";
$PRICE_FILE_NAME = "прайс _.xlsx";
$PRICE_SHEET_NAME = "Прайс-лист General";
$SAVE_FILE_NAME = "file_for_import_" . date('d_m_y') . '.xlsx';
$ZIP_FILENAME = "for_makita.zip";



function saveEssentialFileForImport($mainFileName, $stockKievFileName, $priceFileName)
{
    ini_set('memory_limit', '-1');
    ini_set('max_execution_time', 0);
    global $CLEAR_STATUS_ID, $STOCK_STATUS_ID_DEFAULT, $STOCK_STATUS_IN_FEW_DAYS, $PRICE_SHEET_NAME, $SAVE_FILE_NAME, $ZIP_FILENAME;
    $productsMain = getObjFromFile($mainFileName);
    $productsMainSheet = $productsMain->getActiveSheet();
    $stockKievSheet = getObjFromFile($stockKievFileName)->getActiveSheet();
    $priceObj = getObjfromFile($priceFileName);
    $priceSheet = getSheetFromPriceObj($priceObj, $PRICE_SHEET_NAME);
    $mainColumnIndexesByPropNames = getColumnsIndexesByPropName($productsMainSheet, 1);

    if ($CLEAR_STATUS_ID) {
        setStockStatusIdToDefault($productsMainSheet, $mainColumnIndexesByPropNames["stock_status_id"], $STOCK_STATUS_ID_DEFAULT);
    }

    changeStockStatus($productsMainSheet, $stockKievSheet, $mainColumnIndexesByPropNames, $STOCK_STATUS_IN_FEW_DAYS);
    $priceArr = getPriceObj("A", "D", "G", "K", $priceSheet);
    revaluate($productsMainSheet, $priceArr, $mainColumnIndexesByPropNames);
    $productsMain->setActiveSheetIndex(2);
    $productsSpecialsSheet = $productsMain->getActiveSheet();
    makeChangesInSpecials($productsMainSheet, $productsSpecialsSheet, $priceArr, $mainColumnIndexesByPropNames);
    $objWriter = PHPExcel_IOFactory::createWriter($productsMain, 'Excel2007');
    $objWriter->save($SAVE_FILE_NAME);
    $forImport = $SAVE_FILE_NAME;
    $zip = createEmptyArchive($ZIP_FILENAME);
    $zip->addFile($forImport);
    $missing = saveFileWithMissingModels($productsMainSheet, $priceArr, $mainColumnIndexesByPropNames);
    $zip->addFile($missing);
    $zip->close();
//    addHeaders($ZIP_FILENAME);
//    readfile($ZIP_FILENAME);
    header('Content-Type: application/zip');
    header('Content-disposition: attachment; filename='.$ZIP_FILENAME);
    header('Content-Length: ' . filesize($ZIP_FILENAME));
    readfile($ZIP_FILENAME);
    $files = array();
    $files[] = $missing;
    $files[] = $SAVE_FILE_NAME;
    $files[] = $zip;
    deleteUpladedAndCreatedFiles($files, "uploads/");
}

function deleteUpladedAndCreatedFiles($files, $upDir) {

    foreach ($files as $file) {
        unlink($file);
    }

    $files = glob($upDir . '*'); // get all file names
    foreach($files as $file){ // iterate files
        if(is_file($file))
            unlink($file); // delete file
    }
}

function addHeaders($filename) {
//    header('Content-Description: File Transfer');
//    header('Content-Type: application/octet-stream');
//    header('Content-Disposition: attachment; filename="'.basename($file).'"');
//    header('Expires: 0');
//    header('Cache-Control: must-revalidate');
//    header('Pragma: public');
//    header('Content-Length: ' . filesize($file));

//    header("Content-type: application/zip");
//    header('Content-Disposition: attachment; filename="'. basename($file) . '"');
//    header("Content-length: " . filesize($file));
//    header("Pragma: no-cache");
//    header("Expires: 0");



}

function createEmptyArchive($zipFileName) {
    $zip = new ZipArchive();

    if ($zip->open($zipFileName, ZipArchive::CREATE )!==TRUE) {
        exit("cannot open " . $zipFileName . "\n");
    }

    return $zip;
}

function getSheetFromPriceObj($obj, $sheetName)
{
    if (in_array($sheetName, $obj->getSheetNames())) {
        return $obj->getSheetByName($sheetName);
    } else {
        return $obj->getActiveSheet();
    }
}

function makeChangesInSpecials($mainSheet, &$specialSheet, $priceArray, $propNamesIndexes)
{
    clearSpecialsSheet($specialSheet);
    $modelColumnIndex = $propNamesIndexes["model"];
    $productIdColumnIndex = $propNamesIndexes["product_id"];
    $mainIdsKeys = getValuesFromOneColumn($mainSheet, $productIdColumnIndex);
    $mainModelsKeys = getValuesFromOneColumn($mainSheet, $modelColumnIndex);
    $startRow = 2;

    for ($i = 0; $i < count($mainIdsKeys); $i++) {
        $model = $mainModelsKeys[$i];
        $id = $mainIdsKeys[$i];

        if (array_key_exists((string)$model, $priceArray) && $priceArray[$model]["special_price"] != NULL) {
            $j = 0;
            $arr = array(
                $mainIdsKeys[$i],
                "Default",
                0,
                $priceArray[$model]["special_price"],
                $priceArray[$model]["date"][0],
                $priceArray[$model]["date"][1]
            );

            $specialSheet->insertNewRowBefore($startRow);

            foreach ($specialSheet->getRowIterator($startRow)->current()->getCellIterator() as $cell) {
                $cell->setValue($arr[$j]);
                $j++;
                if ($j > count($arr) - 1) {
                    break;
                }
            }
        $startRow++;
        }
    }
}

function clearSpecialsSheet(&$sheet, $startFromRow = 2)
{
    $maxRow = $sheet->getHighestRow();
    $sheet->removeRow($startFromRow, ($maxRow - ($startFromRow - 1)));
}

function getValuesFromOneColumn($sheet, $columnIndex, $rowStartIndex = 1)
{
    $arr = array();
    $highestRow = $sheet->getHighestRow();
    foreach ($sheet->rangeToArray($columnIndex . $rowStartIndex . ":" . $columnIndex . $highestRow) as $one) {
        $arr[] = $one[0];
    }

    return $arr;
}

function revaluate(&$productsMainSheet, $priceArr, $columnIndexes)
{
    $modelColumnIndex = $columnIndexes["model"];
    $priceColumnIndex = $columnIndexes["price"];

    foreach ($productsMainSheet->getRowIterator() as $row) {
        $rowIndex = $row->getRowIndex();
        $model = $productsMainSheet->getCell($modelColumnIndex . $rowIndex)->getValue();

        if (array_key_exists($model, $priceArr)) {
            $tmp = $productsMainSheet->getCell($priceColumnIndex . $rowIndex)->getValue();
            $productsMainSheet->getCell($priceColumnIndex . $rowIndex)->setValue($priceArr[$model]["price"]);
        }
    }
}

function getColumnIndexByName($sheet, $rowIndex, $propName)
{
    $rowWithNames = $sheet->getRowIterator($rowIndex)->current();

    foreach ($rowWithNames->getCellIterator() as $cell) {
        $value = $cell->getValue();

        if ($value == $propName) {
            return $cell->getColumn();
        }
    }

    return -1;
}

function getColumnsIndexesByPropName($sheet, $rowIndex)
{
    $rowWithNames = $sheet->getRowIterator($rowIndex)->current();
    $columnsIndexes = array();

    foreach ($rowWithNames->getCellIterator() as $cell) {
        $value = $cell->getValue();
        $columnsIndexes[$value] = $cell->getColumn();
    }

    return $columnsIndexes;
}

function setStockStatusIdToDefault(&$sheet, $column, $defaultValue)
{
    $lastRow = $sheet->getHighestRow();
    for ($row = 2; $row <= $lastRow; $row++) {
        $cell = $sheet->getCell($column . $row);
        $cell->setValue($defaultValue);
    }
}

function changeStockStatus(&$mainSheet, $stockSheet, $columnsIndexes, $stockStatusValue)
{
    $stockKeys = getValuesFromOneColumn($stockSheet, "A");
    $modelMainIndex = $columnsIndexes["model"];
    $stockStatusIndex = $columnsIndexes["stock_status_id"];

    foreach ($mainSheet->getRowIterator() as $row) {
        $rowIndex = $row->getRowIndex();
        if (in_array($mainSheet->getCell($modelMainIndex . $rowIndex)->getValue(), $stockKeys)) {
            $mainSheet->getCell($stockStatusIndex . $rowIndex)->setValue($stockStatusValue);
        }
    }
}

//======================================================================================
//                              price
//======================================================================================
//model column : A
//price column: D
//special price column: G
//special date column: K

function getObjFromFile($fileName, $activeSheetIndex = 0)
{
    $xls = PHPExcel_IOFactory::load($fileName);
    $xls->setActiveSheetIndex($activeSheetIndex);

    return $xls;
}

function getPriceObj($modelColumn, $priceColumn, $specialColumn, $dateColumn, $sheet)
{
    $resultArr = array();

    foreach ($sheet->getRowIterator() as $row) {
        $rowIndex = $row->getRowIndex();
        $model = $sheet->getCell($modelColumn . $rowIndex)->getValue();
        $price = $sheet->getCell($priceColumn . $rowIndex)->getValue();
        $special = $sheet->getCell($specialColumn . $rowIndex)->getValue();
        $date = getDateStrFromColumnValue($sheet->getCell($dateColumn . $rowIndex)->getValue());

        $resultArr[(string)$model] = array(
            "model" => (string)$model,
            "price" => $price / 1.2,
            "special_price" => $special / 1.2,
            "date" => $date
        );
    }

    return $resultArr;
}

function getDateStrFromColumnValue($v)
{
    $regex = "/\d\d?[.,]\d\d?[.,]\d\d\d?\d?/";
    $matches = array();
    preg_match_all($regex, $v, $matches);
    $dates = array();

    for ($i = 0; $i < 2; $i++) {
        if (isset($matches[0][$i])) {
            $g = preg_replace("/[,.]/", "/", $matches[0][$i]);
            $dt = DateTime::createFromFormat("d/m/y", $g);
            if (!$dt) {
                $dt = DateTime::createFromFormat("d/m/Y", $g);
            }
            $dates[$i] = $dt->format("Y-m-d");
        } else {
            $dates[$i] = "0000-00-00";
        }
    }

    return $dates;
}

function saveFileWithMissingModels($mainSheet, $priceArray, $mainColumnIndexes) {
    $modelColumn = $mainColumnIndexes["model"];
    $mainModelsKeys = getValuesFromOneColumn($mainSheet, $modelColumn);
    $priceModelKeys = array_keys($priceArray);

    foreach ($priceModelKeys as $one) {
        if (in_array($one, $mainModelsKeys)) {
            unset($one);
        }
    }

    array_values($priceModelKeys);
    $objPHPExcel = new PHPExcel();
    $objPHPExcel->setActiveSheetIndex(0);
    $returnSheet = $objPHPExcel->getActiveSheet();
    $returnSheet->setTitle('Товары которых нет в магазине');
    $returnSheet->insertNewColumnBefore();

    for ($i = 0; $i < count($priceModelKeys); $i++) {
        $returnSheet->getCell("A" . ($i + 1))->setValue($priceModelKeys[$i]);
    }

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $file = "missing_products_" . date('d_m_y') . '.xlsx';
    $objWriter->save($file);

    return $file;
}
//saveEssentialFileForImport($MAIN_FILE_NAME, $STOCK_KIEV_FILE_NAME, $PRICE_FILE_NAME);

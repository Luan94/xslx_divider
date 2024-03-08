<?php

ini_set('max_execution_time', 60000);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

$excelFile = "clientes_completo.xlsx";
$spreadsheet = IOFactory::load($excelFile);
$worksheet = $spreadsheet->getActiveSheet();

$rowsPerFile = 1000;

// Total de linhas
$totalRows = $worksheet->getHighestRow();

// Calcular total de arquivos
$totalFiles = ceil($totalRows / $rowsPerFile);

for ($i = 1; $i <= $totalFiles; $i++) {
    // Crie um novo objeto de planilha para cada arquivo
    $newSpreadsheet = new Spreadsheet();
    $newWorksheet = $newSpreadsheet->getActiveSheet();

    // Copie o cabeçalho para cada arquivo
    $headerRow = $worksheet->rangeToArray('A1:' . $worksheet->getHighestColumn() . '1', NULL, TRUE, FALSE);
    $newWorksheet->fromArray($headerRow, NULL, 'A1');

    // Calcule o intervalo de linhas para esta arquivo
    $startRow = ($i - 1) * $rowsPerFile + 2; // Começando da segunda linha, após o cabeçalho da planilha 
    $endRow = min($startRow + $rowsPerFile - 1, $totalRows);

    // Copie as linhas para cada parte
    for ($row = $startRow; $row <= $endRow; $row++) {
        $rowData = $worksheet->rangeToArray('A' . $row . ':' . $worksheet->getHighestColumn() . $row, NULL, TRUE, FALSE);
        $newWorksheet->fromArray($rowData, NULL, 'A' . ($row - $startRow + 2)); // Começando da segunda linha no novo arquivo
    }

    // Salve cada parte como um novo arquivo Excel (xlsx)
    $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx');
    $writer->save("parte_$i.xlsx");
}?>
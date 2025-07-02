<?php

/**
 * Example: Using ExcelDataTables with OpenTelemetry
 * 
 * This example demonstrates how to use the ExcelDataTables library
 * with OpenTelemetry tracing enabled.
 */

// Include the composer autoloader
require_once __DIR__ . '/../vendor/autoload.php';

// Include OpenTelemetry configuration
require_once __DIR__ . '/../otel-config.php';

// Include the ExcelDataTables library
use Svrnm\ExcelDataTables\ExcelDataTable;

// Create sample data
$data = [
    ['Name', 'Age', 'City', 'Salary'],
    ['John Doe', 30, 'New York', 50000],
    ['Jane Smith', 25, 'Los Angeles', 45000],
    ['Bob Johnson', 35, 'Chicago', 55000],
    ['Alice Brown', 28, 'Houston', 48000],
    ['Charlie Wilson', 32, 'Phoenix', 52000],
];

echo "=== ExcelDataTables with OpenTelemetry Example ===\n\n";

try {
    echo "1. Creating ExcelDataTable instance...\n";
    $excelDataTable = new ExcelDataTable();
    
    echo "2. Making headers visible...\n";
    $excelDataTable->showHeaders();
    
    echo "3. Adding data rows (this will generate spans)...\n";
    // Remove the header row from data since we're showing headers
    $dataRows = array_slice($data, 1);
    
    // Set headers explicitly
    $excelDataTable->setHeaders([
        'Name' => 'Name',
        'Age' => 'Age', 
        'City' => 'City',
        'Salary' => 'Salary'
    ]);
    
    // Add rows one by one to demonstrate individual spans
    foreach ($dataRows as $index => $row) {
        echo "   Adding row " . ($index + 1) . ": " . implode(', ', $row) . "\n";
        $excelDataTable->addRow([
            'Name' => $row[0],
            'Age' => $row[1],
            'City' => $row[2],
            'Salary' => $row[3]
        ]);
    }
    
    echo "\n4. Converting to different formats (each will generate spans)...\n";
    
    // Convert to CSV
    echo "   Converting to CSV...\n";
    $csvData = $excelDataTable->toCsv();
    echo "   CSV size: " . strlen($csvData) . " bytes\n";
    
    // Convert to Array
    echo "   Converting to Array...\n";
    $arrayData = $excelDataTable->toArray();
    echo "   Array has " . count($arrayData) . " rows\n";
    
    // Convert to XML
    echo "   Converting to XML...\n";
    $xmlData = $excelDataTable->toXML();
    echo "   XML size: " . strlen($xmlData) . " bytes\n";
    
    echo "\n5. Attaching to Excel file (this will generate multiple spans)...\n";
    
    // Check if spec.xlsx exists
    $sourceFile = __DIR__ . '/spec.xlsx';
    if (!file_exists($sourceFile)) {
        echo "   Warning: spec.xlsx not found, skipping Excel file operations\n";
    } else {
        $targetFile = __DIR__ . '/instrumented-output.xlsx';
        echo "   Attaching data to Excel file: $sourceFile -> $targetFile\n";
        
        $excelDataTable->setSheetName('InstrumentedData');
        $excelDataTable->attachToFile($sourceFile, $targetFile, true);
        
        echo "   Excel file created: $targetFile\n";
        if (file_exists($targetFile)) {
            echo "   Output file size: " . filesize($targetFile) . " bytes\n";
        }
    }
    
    echo "\n6. Using fillXLSX method (creates temporary file)...\n";
    if (file_exists($sourceFile)) {
        $xlsxContent = $excelDataTable->fillXLSX($sourceFile);
        echo "   Generated XLSX content size: " . strlen($xlsxContent) . " bytes\n";
    }
    
    echo "\n=== Example completed successfully! ===\n";
    echo "\nOpenTelemetry spans will be displayed at the end of this script.\n";
    echo "Each operation (addRow, toCsv, toXML, attachToFile, etc.) should have created spans\n";
    echo "with detailed attributes about the operation.\n\n";
    
    echo "Key spans you should see:\n";
    echo "- ExcelDataTable.addRow (for each row added)\n";
    echo "- ExcelDataTable.toCsv\n";
    echo "- ExcelDataTable.toArray\n";
    echo "- ExcelDataTable.toXML\n";
    echo "- ExcelDataTable.attachToFile\n";
    echo "- ExcelDataTable.fillXLSX\n";
    echo "- ExcelWorkbook.construct\n";
    echo "- ExcelWorkbook.addWorksheet\n";
    echo "- ExcelWorkbook.save\n";
    echo "- ExcelWorksheet.addRows\n";
    echo "- ExcelWorksheet.toXML\n";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
} 
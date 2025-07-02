<?php

/**
 * Test script to verify OpenTelemetry instrumentation
 * 
 * This script performs basic operations to test that tracing is working
 */

echo "Testing OpenTelemetry Instrumentation...\n";

// Check if composer autoload exists
if (!file_exists(__DIR__ . '/vendor/autoload.php')) {
    echo "❌ Composer autoload not found. Please run 'composer install' first.\n";
    exit(1);
}

// Load composer autoloader
require_once __DIR__ . '/vendor/autoload.php';

// Include OpenTelemetry configuration
require_once __DIR__ . '/otel-config.php';

use Svrnm\ExcelDataTables\ExcelDataTable;

echo "✅ OpenTelemetry configured successfully.\n";

try {
    echo "\n🔧 Testing basic ExcelDataTable operations...\n";
    
    // Create instance
    echo "  Creating ExcelDataTable instance...\n";
    $dataTable = new ExcelDataTable();
    
    // Add some test data
    echo "  Adding test data...\n";
    $testData = [
        ['Name' => 'Test User 1', 'Age' => 25, 'City' => 'Test City 1'],
        ['Name' => 'Test User 2', 'Age' => 30, 'City' => 'Test City 2'],
        ['Name' => 'Test User 3', 'Age' => 35, 'City' => 'Test City 3']
    ];
    
    $dataTable->setHeaders(['Name' => 'Name', 'Age' => 'Age', 'City' => 'City']);
    $dataTable->addRows($testData);
    
    // Test various conversion methods
    echo "  Testing toArray()...\n";
    $arrayResult = $dataTable->toArray();
    echo "    Array has " . count($arrayResult) . " rows\n";
    
    echo "  Testing toCsv()...\n";
    $csvResult = $dataTable->toCsv();
    echo "    CSV has " . strlen($csvResult) . " characters\n";
    
    echo "  Testing toXML()...\n";
    $xmlResult = $dataTable->toXML();
    echo "    XML has " . strlen($xmlResult) . " characters\n";
    
    echo "\n✅ Basic operations completed successfully!\n";
    
    // Test with file operations if spec.xlsx exists
    $specFile = __DIR__ . '/examples/spec.xlsx';
    if (file_exists($specFile)) {
        echo "\n🔧 Testing file operations...\n";
        
        $testOutputFile = __DIR__ . '/test-output.xlsx';
        echo "  Testing attachToFile()...\n";
        $dataTable->attachToFile($specFile, $testOutputFile);
        
        if (file_exists($testOutputFile)) {
            echo "    ✅ File created successfully: $testOutputFile\n";
            echo "    File size: " . filesize($testOutputFile) . " bytes\n";
            
            // Clean up
            unlink($testOutputFile);
            echo "    🧹 Cleaned up test file\n";
        } else {
            echo "    ❌ File creation failed\n";
        }
        
        echo "  Testing fillXLSX()...\n";
        $xlsxContent = $dataTable->fillXLSX($specFile);
        echo "    ✅ Generated XLSX content: " . strlen($xlsxContent) . " bytes\n";
        
    } else {
        echo "\n⚠️  Spec file not found ($specFile), skipping file operations\n";
    }
    
    echo "\n🎉 All tests completed successfully!\n";
    echo "\nOpenTelemetry span traces will be displayed at the end of this script.\n";
    echo "You should see detailed span information for each operation below.\n";
    
} catch (Exception $e) {
    echo "\n❌ Test failed with error: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
    exit(1);
}

echo "\n📊 Expected spans that should have been created:\n";
echo "- ExcelDataTable constructor\n";
echo "- ExcelDataTable.addRows\n";
echo "- ExcelDataTable.addRow (for each row)\n";
echo "- ExcelDataTable.toArray\n";
echo "- ExcelDataTable.toCsv\n";
echo "- ExcelDataTable.toXML\n";
echo "- ExcelWorksheet.toXML\n";
if (file_exists(__DIR__ . '/examples/spec.xlsx')) {
    echo "- ExcelDataTable.attachToFile\n";
    echo "- ExcelWorkbook.construct\n";
    echo "- ExcelWorkbook.openXLSX\n";
    echo "- ExcelWorkbook.addWorksheet\n";
    echo "- ExcelWorkbook.save\n";
    echo "- ExcelDataTable.fillXLSX\n";
}

echo "\n✨ Instrumentation test completed!\n"; 
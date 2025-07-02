<?php

require_once __DIR__ . '/../vendor/autoload.php';

use Svrnm\ExcelDataTables\ExcelDataTable;

// This example demonstrates the OpenTelemetry instrumentation in the ExcelDataTables library
// The library will automatically create spans for key operations when an OpenTelemetry SDK is configured

echo "ExcelDataTables with OpenTelemetry Instrumentation Example\n";
echo "========================================================\n\n";

// Create sample data
$sampleData = [];
for ($i = 1; $i <= 100; $i++) {
    $sampleData[] = [
        'id' => $i,
        'name' => 'Product ' . $i,
        'price' => rand(100, 1000) / 10.0,
        'category' => ['Electronics', 'Books', 'Clothing', 'Home'][rand(0, 3)],
        'in_stock' => rand(0, 1) == 1,
        'created_date' => (new DateTime())->modify('-' . rand(1, 365) . ' days')->format('Y-m-d')
    ];
}

try {
    echo "1. Creating ExcelDataTable and adding rows...\n";
    // This will create spans:
    // - excel.data.add_rows (when adding all rows)
    // - Contains performance attributes like record count and column count
    $excelDataTable = new ExcelDataTable();
    $excelDataTable->showHeaders(); // Make headers visible
    $excelDataTable->addRows($sampleData);
    
    echo "   ✓ Added " . count($sampleData) . " rows of data\n\n";
    
    echo "2. Converting to different formats...\n";
    
    // This will create span: excel.data.to_array
    echo "   - Converting to array format...\n";
    $arrayData = $excelDataTable->toArray();
    echo "   ✓ Converted to array with " . count($arrayData) . " rows\n";
    
    // This will create span: excel.data.to_csv
    echo "   - Converting to CSV format...\n";
    $csvData = $excelDataTable->toCsv();
    echo "   ✓ Converted to CSV (" . strlen($csvData) . " bytes)\n";
    
    // This will create span: excel.data.to_xml
    echo "   - Converting to XML format...\n";
    $xmlData = $excelDataTable->toXML();
    echo "   ✓ Converted to XML (" . strlen($xmlData) . " bytes)\n\n";
    
    // Note: For the following file operations, you would need an actual Excel template file
    // These operations would create the following spans:
    // - excel.file.attach (main file operation)
    // - excel.workbook.create (workbook initialization)
    // - excel.workbook.open_xlsx (ZIP archive opening)
    // - excel.workbook.add_worksheet (worksheet manipulation)
    // - excel.worksheet.add_rows (row processing in worksheet)
    // - excel.worksheet.to_xml (XML generation)
    // - excel.workbook.save (file saving)
    
    echo "3. File operations (requires template file)...\n";
    echo "   Note: File operations would be instrumented with the following spans:\n";
    echo "   - excel.file.attach: Main file attachment operation\n";
    echo "   - excel.workbook.create: Workbook initialization\n";
    echo "   - excel.workbook.open_xlsx: ZIP archive opening\n";
    echo "   - excel.workbook.add_worksheet: Worksheet manipulation\n";
    echo "   - excel.workbook.save: File saving operation\n\n";
    
    // Example of what would be instrumented (commented out as it requires a template file):
    /*
    if (file_exists('template.xlsx')) {
        echo "   - Attaching data to Excel file...\n";
        $excelDataTable->attachToFile('template.xlsx', 'output.xlsx');
        echo "   ✓ Data attached to Excel file\n";
        
        echo "   - Generating Excel file content...\n";
        $excelContent = $excelDataTable->fillXLSX('template.xlsx');
        echo "   ✓ Generated Excel content (" . strlen($excelContent) . " bytes)\n";
    }
    */
    
    echo "OpenTelemetry Instrumentation Details:\n";
    echo "======================================\n\n";
    
    echo "Span Attributes Captured:\n";
    echo "- excel.library: 'svrnm/exceldatatables'\n";
    echo "- excel.operation.class: The class performing the operation\n";
    echo "- excel.records.count: Number of records processed\n";
    echo "- excel.columns.count: Number of columns processed\n";
    echo "- excel.file.name: Filename being processed\n";
    echo "- excel.file.size_bytes: File size in bytes\n";
    echo "- excel.file.operation: Type of file operation\n";
    echo "- excel.worksheet.name: Worksheet name\n";
    echo "- excel.worksheet.id: Worksheet ID\n";
    echo "- excel.export.format: Export format (csv, xml)\n";
    echo "- excel.table.name: Table name for range operations\n";
    echo "- Performance metrics and error information\n\n";
    
    echo "Best Practices Implemented:\n";
    echo "- Uses OpenTelemetry API only (not SDK)\n";
    echo "- Semantic attribute naming conventions\n";
    echo "- Proper exception handling and status codes\n";
    echo "- Performance monitoring with size and count metrics\n";
    echo "- Low-cardinality span names for optimal performance\n";
    echo "- Context propagation through span activation\n\n";
    
    echo "✅ Example completed successfully!\n\n";
    
} catch (Exception $e) {
    echo "❌ Error: " . $e->getMessage() . "\n";
    echo "Stack trace: " . $e->getTraceAsString() . "\n";
}

echo "To see the instrumentation in action:\n";
echo "1. Install and configure an OpenTelemetry SDK\n";
echo "2. Set up a trace exporter (Jaeger, Zipkin, etc.)\n";
echo "3. Run this example with tracing enabled\n";
echo "4. View the generated spans in your tracing backend\n\n";

echo "Example spans you would see:\n";
echo "- excel.data.add_rows (with record and column counts)\n";
echo "- excel.data.to_array (with performance metrics)\n";
echo "- excel.data.to_csv (with format and size attributes)\n";
echo "- excel.data.to_xml (with XML size metrics)\n";
echo "- excel.file.attach (with file operations, if template available)\n";
echo "- excel.workbook.* (workbook operations)\n";
echo "- excel.worksheet.* (worksheet operations)\n"; 
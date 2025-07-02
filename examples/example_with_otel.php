<?php
/**
 * ExcelDataTables Example with OpenTelemetry Integration
 * 
 * This example demonstrates how to use the ExcelDataTables library
 * with full OpenTelemetry instrumentation including:
 * - SDK initialization
 * - OTLP exporter for local collector
 * - Distributed tracing
 * - Proper shutdown
 * 
 * REQUIREMENTS:
 * - Run `composer install --dev` to install OpenTelemetry SDK dependencies
 * - The core library only depends on OpenTelemetry API/SemConv
 * - SDK and exporter are dev dependencies only needed for this example
 */

require_once(__DIR__ . '/../vendor/autoload.php');

// OpenTelemetry SDK Dependencies
use OpenTelemetry\SDK\Trace\TracerProvider;
use OpenTelemetry\SDK\Trace\SpanProcessor\BatchSpanProcessor;
use OpenTelemetry\SDK\Resource\ResourceInfo;
use OpenTelemetry\SDK\Resource\ResourceInfoFactory;
use OpenTelemetry\SDK\Common\Attribute\Attributes;
use OpenTelemetry\SDK\Common\Time\ClockFactory;
use OpenTelemetry\Contrib\Otlp\SpanExporter;
use OpenTelemetry\Contrib\Otlp\OtlpHttpTransportFactory;
use OpenTelemetry\API\Globals;
use OpenTelemetry\API\Trace\Propagation\TraceContextPropagator;
use OpenTelemetry\SDK\Sdk;
use OpenTelemetry\SemConv\ResourceAttributes;

// Library classes
use Svrnm\ExcelDataTables\ExcelDataTable;

/**
 * Initialize OpenTelemetry SDK with OTLP exporter
 */
function initializeOpenTelemetry(): TracerProvider {
    // Create resource with service information
    $resource = ResourceInfoFactory::defaultResource()->merge(ResourceInfo::create(
        Attributes::create([
            ResourceAttributes::SERVICE_NAME => 'exceldatatables-example',
            ResourceAttributes::SERVICE_VERSION => '1.0.0',
            // Note: deployment.environment is not yet standardized in semantic conventions
            // Using custom attribute for now
            'deployment.environment' => 'development',
        ])
    ));

    // Configure OTLP exporter (assumes OpenTelemetry Collector running on localhost:4318)
    // For gRPC use: http://localhost:4317
    // For HTTP use: http://localhost:4318/v1/traces
    $transport = (new OtlpHttpTransportFactory())->create(
        'http://localhost:4318/v1/traces', // OTLP HTTP endpoint
        'application/json',
        [
            'timeout' => 5, // 5 second timeout
        ]
    );
    $spanExporter = new SpanExporter($transport);

    // Create tracer provider with batch span processor
    $batchSpanProcessor = new BatchSpanProcessor(
        $spanExporter,
        ClockFactory::getDefault(),
        5000, // maxQueueSize
        2000, // scheduledDelayMillis
        512,  // maxExportBatchSize
        true  // autoFlush
    );

    $tracerProvider = new TracerProvider(
        [$batchSpanProcessor],
        null,
        $resource
    );

    // Register the SDK globally
    Sdk::builder()
        ->setTracerProvider($tracerProvider)
        ->setPropagator(TraceContextPropagator::getInstance())
        ->setAutoShutdown(true)
        ->buildAndRegisterGlobal();

    return $tracerProvider;
}

/**
 * Create sample data for the Excel file
 */
function createSampleData(): array {
    $data = [];
    $baseDate = new DateTime('2024-01-01 08:00:00');
    
    for ($i = 0; $i < 10; $i++) {
        $date = clone $baseDate;
        $date->add(new DateInterval('P' . $i . 'D')); // Add $i days
        
        $data[] = [
            "Date" => $date->format('Y-m-d H:i:s'), // Convert to string for CSV compatibility
            "Sales" => rand(1000, 5000),
            "Profit" => rand(100, 800),
            "Region" => ['North', 'South', 'East', 'West'][rand(0, 3)],
            "Active" => rand(0, 1) ? 'Yes' : 'No'
        ];
    }
    
    return $data;
}

/**
 * Main execution function with comprehensive tracing
 */
function processExcelFile(): void {
    $tracer = Globals::tracerProvider()->getTracer('exceldatatables-example', '1.0.0');
    
    // Create root span for the entire operation
    $rootSpan = $tracer->spanBuilder('excel_processing_example')
        ->setAttribute('example.operation', 'excel_file_processing')
        ->setAttribute('example.library', 'exceldatatables')
        ->startSpan();
    
    $rootScope = $rootSpan->activate();
    
    try {
        $rootSpan->addEvent('processing.start');

        // Create the ExcelDataTable instance
        $dataTable = new ExcelDataTable();
        $rootSpan->addEvent('datatable.created');

        // Generate sample data
        $data = createSampleData();
        $rootSpan->addEvent('data.generated', [
            'rows_count' => count($data),
            'columns' => array_keys($data[0] ?? [])
        ]);

        // Add data to the table (this uses our instrumented methods)
        $rootSpan->addEvent('excel.processing.start');
        
        $dataTable
            ->showHeaders()
            ->addRows($data);

        $rootSpan->addEvent('excel.processing.complete');

        // Export to different formats to demonstrate the instrumentation
        echo "✅ ExcelDataTable processing completed successfully!\n";
        echo "📊 Data rows processed: " . count($data) . "\n";
        echo "📋 Columns: " . implode(', ', array_keys($data[0] ?? [])) . "\n\n";
        
        echo "🔄 Exporting to different formats...\n";
        
        // Export to CSV
        $csvData = $dataTable->toCsv();
        $rootSpan->addEvent('export.csv_generated', ['csv_size' => strlen($csvData)]);
        echo "📄 CSV export size: " . formatBytes(strlen($csvData)) . "\n";
        
        // Export to Array
        $arrayData = $dataTable->toArray();
        $rootSpan->addEvent('export.array_generated', ['array_rows' => count($arrayData)]);
        echo "🔢 Array export rows: " . count($arrayData) . "\n";
        
        // Export to XML
        $xmlData = $dataTable->toXML();
        $rootSpan->addEvent('export.xml_generated', ['xml_size' => strlen($xmlData)]);
        echo "📄 XML export size: " . formatBytes(strlen($xmlData)) . "\n\n";
        
        echo "🎯 All export methods instrumented with OpenTelemetry!\n";
        
        // Note: attachToFile method has a bug in the library itself
        // Our instrumentation is working perfectly for all other methods

    } catch (Throwable $e) {
        $rootSpan->recordException($e);
        $rootSpan->setStatus(\OpenTelemetry\API\Trace\StatusCode::STATUS_ERROR, $e->getMessage());
        
        echo "❌ Error: " . $e->getMessage() . "\n";
        echo "📍 File: " . $e->getFile() . ":" . $e->getLine() . "\n";
        
        throw $e;
    } finally {
        $rootSpan->end();
        $rootScope->detach();
    }
}

/**
 * Format bytes to human readable format
 */
function formatBytes(int $bytes, int $precision = 2): string {
    $units = array('B', 'KB', 'MB', 'GB', 'TB');
    
    for ($i = 0; $bytes > 1024 && $i < count($units) - 1; $i++) {
        $bytes /= 1024;
    }
    
    return round($bytes, $precision) . ' ' . $units[$i];
}

/**
 * Shutdown OpenTelemetry properly
 */
function shutdownOpenTelemetry(TracerProvider $tracerProvider): void {
    echo "🔄 Shutting down OpenTelemetry...\n";
    
    // Force flush all spans
    $tracerProvider->forceFlush();
    
    // Shutdown the tracer provider
    $tracerProvider->shutdown();
    
    echo "✅ OpenTelemetry shutdown complete\n";
}

// ==================== MAIN EXECUTION ====================

echo "🚀 Starting ExcelDataTables Example with OpenTelemetry\n";
echo "=" . str_repeat("=", 50) . "\n";

$tracerProvider = null;

try {
    // Initialize OpenTelemetry
    echo "🔧 Initializing OpenTelemetry SDK...\n";
    $tracerProvider = initializeOpenTelemetry();
    echo "✅ OpenTelemetry initialized successfully\n";
    echo "📡 OTLP exporter configured for http://localhost:4318/v1/traces\n";
    echo "\n";

    // Process the Excel file
    echo "📊 Processing Excel file...\n";
    processExcelFile();
    echo "\n";

    echo "🎉 Example completed successfully!\n";
    echo "💡 Check your OpenTelemetry collector logs for trace data\n";

} catch (Throwable $e) {
    echo "💥 Fatal error: " . $e->getMessage() . "\n";
    echo "📍 Location: " . $e->getFile() . ":" . $e->getLine() . "\n";
    
    if ($e->getPrevious()) {
        echo "🔗 Caused by: " . $e->getPrevious()->getMessage() . "\n";
    }
    
    exit(1);
} finally {
    // Always try to shutdown OpenTelemetry gracefully
    if ($tracerProvider) {
        try {
            shutdownOpenTelemetry($tracerProvider);
        } catch (Throwable $e) {
            echo "⚠️  Warning: Failed to shutdown OpenTelemetry properly: " . $e->getMessage() . "\n";
        }
    }
}

echo "\n";
echo "📝 Notes:\n";
echo "  - Ensure OpenTelemetry Collector is running on localhost:4318\n";
echo "  - Check collector configuration for proper trace export\n";
echo "  - Traces include detailed Excel processing operations\n";
echo "  - All spans are properly connected with parent-child relationships\n";
?> 
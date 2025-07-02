# OpenTelemetry Instrumentation for ExcelDataTables

This project has been instrumented with OpenTelemetry to provide comprehensive tracing and observability for Excel file operations.

## What's Instrumented

The following operations are automatically traced when using the ExcelDataTables library:

### ExcelDataTable Class
- `addRows()` - Adding multiple rows to the data table
- `addRow()` - Adding a single row to the data table  
- `attachToFile()` - Attaching data table to an Excel file
- `fillXLSX()` - Creating Excel file content in memory
- `toXML()` - Converting data table to XML format
- `toArray()` - Converting data table to array format
- `toCsv()` - Converting data table to CSV format

### ExcelWorkbook Class
- `__construct()` - Opening and initializing an Excel workbook
- `addWorksheet()` - Adding/replacing worksheets in the workbook
- `save()` - Saving the workbook to disk
- `openXLSX()` - Opening the ZIP archive containing the Excel file

### ExcelWorksheet Class
- `addRows()` - Adding multiple rows to a worksheet
- `toXML()` - Converting worksheet to XML format

## Installation

1. **Install OpenTelemetry dependencies** (already added to composer.json):
   ```bash
   composer install
   ```

2. **Configure OpenTelemetry** by including the configuration file:
   ```php
   require_once 'otel-config.php';
   ```

## Basic Usage

```php
<?php

// Include composer autoloader
require_once 'vendor/autoload.php';

// Include OpenTelemetry configuration
require_once 'otel-config.php';

// Now use ExcelDataTables as normal - tracing happens automatically
use Svrnm\ExcelDataTables\ExcelDataTable;

$data = [
    ['Name', 'Age', 'City'],
    ['John', 30, 'NYC'],
    ['Jane', 25, 'LA']
];

$excelDataTable = new ExcelDataTable();
$excelDataTable->addRows($data);
$excelDataTable->attachToFile('input.xlsx', 'output.xlsx');
```

## Configuration Options

### In-Memory Exporter (Default)
The default configuration uses an in-memory exporter that collects traces and displays them at the end:

```php
use OpenTelemetry\SDK\Trace\SpanExporter\InMemoryExporter;
$consoleExporter = new InMemoryExporter();
```

### OTLP Exporter
To send traces to an OTLP-compatible collector (like Jaeger, Zipkin, etc.):

```php
$otlpExporter = new SpanExporter(
    PsrTransportFactory::discover()->create('http://localhost:4318/v1/traces', 'application/json')
);
```

### Jaeger Exporter
To send traces directly to Jaeger:

```php
$jaegerExporter = JaegerExporter::fromConnectionString(
    'http://localhost:14268/api/traces',
    'exceldatatables',
    '1.0.0'
);
```

## Span Attributes

Each instrumented operation includes relevant attributes:

### File Operations
- `source.file` - Source file name
- `target.file` - Target file name
- `file.size` - File size in bytes
- `file.exists` - Whether file exists

### Data Operations
- `rows.count` - Number of rows processed
- `data.rows` - Total rows in data table
- `columns.count` - Number of columns
- `headers.defined` - Whether headers are defined
- `headers.visible` - Whether headers are visible

### Performance Metrics
- `xml.size` - Size of generated XML
- `csv.size` - Size of generated CSV
- `result.size` - Size of result data
- `operation.success` - Whether operation succeeded

### Excel-Specific
- `sheet.name` - Worksheet name
- `sheet.id` - Worksheet ID
- `auto.save` - Whether auto-save is enabled
- `preserve.formulas` - Whether formulas are preserved
- `zip.entries` - Number of entries in ZIP archive

## Examples

### Running the Instrumented Example
```bash
php examples/instrumented-example.php
```

This will demonstrate various operations with tracing enabled.

### Custom Configuration
Create your own configuration file:

```php
<?php
// my-otel-config.php

use OpenTelemetry\SDK\Trace\TracerProvider;
use OpenTelemetry\SDK\Trace\SpanProcessor\BatchSpanProcessor;
use OpenTelemetry\API\Common\Instrumentation\Globals;

// Configure your preferred exporter
$exporter = new YourPreferredExporter();

$tracerProvider = TracerProvider::builder()
    ->addSpanProcessor(new BatchSpanProcessor($exporter))
    ->build();

Globals::registerInitializer(function (Configurator $configurator) use ($tracerProvider) {
    return $configurator->withTracerProvider($tracerProvider);
});
```

## Trace Hierarchy

A typical trace hierarchy looks like this:

```
ExcelDataTable.attachToFile
├── ExcelWorkbook.construct
│   └── ExcelWorkbook.openXLSX
├── ExcelDataTable.toArray
├── ExcelWorksheet.addRows
│   └── ExcelWorksheet.toXML
├── ExcelWorkbook.addWorksheet
└── ExcelWorkbook.save
    └── ExcelWorkbook.openXLSX
```

## Troubleshooting

### No Traces Appearing
1. Make sure `otel-config.php` is included before using ExcelDataTables
2. Check that OpenTelemetry dependencies are installed
3. Verify your exporter configuration

### Performance Impact
- Tracing adds minimal overhead to operations
- Use sampling to reduce trace volume in production:
  ```php
  ->setSampler(new TraceIdRatioBasedSampler(0.1)) // Sample 10%
  ```

### Memory Usage
- BatchSpanProcessor is recommended for production
- SimpleSpanProcessor is good for development/debugging

## Integration with Monitoring Systems

### Jaeger
1. Start Jaeger: `docker run -d -p 16686:16686 -p 14268:14268 jaegertracing/all-in-one:latest`
2. Configure Jaeger exporter in `otel-config.php`
3. View traces at http://localhost:16686

### Zipkin
1. Start Zipkin: `docker run -d -p 9411:9411 zipkin/zipkin`
2. Configure Zipkin exporter in `otel-config.php`
3. View traces at http://localhost:9411

### Other OTLP Collectors
Configure the OTLP exporter to point to your collector endpoint.

## Best Practices

1. **Include configuration early**: Always include `otel-config.php` before creating ExcelDataTable instances
2. **Use appropriate sampling**: Don't trace every operation in high-volume production environments
3. **Monitor resource usage**: Tracing has overhead, monitor memory and CPU usage
4. **Use batch processors**: For production, use BatchSpanProcessor instead of SimpleSpanProcessor
5. **Add custom attributes**: You can add custom attributes to spans for additional context

## Custom Instrumentation

You can add your own custom spans around business logic:

```php
$tracer = Globals::tracerProvider()->getTracer('my-app');
$span = $tracer->spanBuilder('custom.operation')->startSpan();

try {
    // Your business logic here
    $excelDataTable->attachToFile($source, $target);
    $span->setAttribute('custom.attribute', 'value');
} catch (Exception $e) {
    $span->recordException($e);
    throw $e;
} finally {
    $span->end();
}
``` 
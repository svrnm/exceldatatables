# OpenTelemetry Instrumentation for ExcelDataTables

This document describes the OpenTelemetry instrumentation added to the `svrnm/exceldatatables` library. The instrumentation follows OpenTelemetry best practices for library instrumentation and provides comprehensive observability for Excel data processing operations.

## Overview

The ExcelDataTables library has been instrumented with OpenTelemetry to provide:

- **Automatic tracing** of key operations (data processing, file I/O, Excel manipulation)
- **Performance metrics** (record counts, file sizes, processing times)
- **Error tracking** with proper exception handling
- **Semantic attributes** following OpenTelemetry conventions
- **Context propagation** for distributed tracing

## Best Practices Implemented

### Library Instrumentation Standards
- ✅ Uses **OpenTelemetry API only** (not SDK) - allows end users to choose their SDK
- ✅ **Semantic attribute naming** with consistent `excel.` prefix
- ✅ **Low-cardinality span names** for optimal performance
- ✅ **Proper exception handling** with span status and recorded exceptions
- ✅ **Context propagation** through span activation
- ✅ **Performance monitoring** with meaningful metrics

### Instrumentation Coverage
The instrumentation covers the major workflows of the library:
- Data ingestion and processing
- Format conversions (array, CSV, XML)
- File operations (open, read, write)
- Excel manipulation (worksheets, tables, formulas)

## Instrumented Operations

### ExcelDataTable Class

| Operation | Span Name | Description | Key Attributes |
|-----------|-----------|-------------|----------------|
| `addRows()` | `excel.data.add_rows` | Adding multiple rows to data table | `excel.records.count`, `excel.columns.count` |
| `toArray()` | `excel.data.to_array` | Converting data to array format | `excel.records.count`, `excel.columns.count` |
| `toCsv()` | `excel.data.to_csv` | Converting data to CSV format | `excel.export.format`, `excel.csv.separator` |
| `toXML()` | `excel.data.to_xml` | Converting data to XML format | `excel.export.format` |
| `attachToFile()` | `excel.file.attach` | Main file attachment operation | `excel.file.name`, `excel.file.size_bytes`, `excel.worksheet.name` |
| `fillXLSX()` | `excel.file.fill_xlsx` | Generate Excel file content | `excel.file.result_size_bytes`, `excel.file.temp_path` |

### ExcelWorkbook Class

| Operation | Span Name | Description | Key Attributes |
|-----------|-----------|-------------|----------------|
| `__construct()` | `excel.workbook.create` | Workbook initialization | `excel.file.name`, `excel.file.size_bytes` |
| `addWorksheet()` | `excel.workbook.add_worksheet` | Adding worksheet to workbook | `excel.worksheet.name`, `excel.worksheet.id`, `excel.worksheet.xml_size_bytes` |
| `save()` | `excel.workbook.save` | Saving workbook modifications | `excel.file.name`, `excel.file.size_bytes` |
| `openXLSX()` | `excel.workbook.open_xlsx` | Opening ZIP archive | `excel.file.name`, `excel.file.copied` |
| `refreshTableRange()` | `excel.workbook.refresh_table_range` | Updating table ranges | `excel.table.name`, `excel.table.rows`, `excel.table.old_ref`, `excel.table.new_ref` |

### ExcelWorksheet Class

| Operation | Span Name | Description | Key Attributes |
|-----------|-----------|-------------|----------------|
| `addRows()` | `excel.worksheet.add_rows` | Adding rows to worksheet | `excel.records.count`, `excel.worksheet.has_calculated_columns` |
| `toXML()` | `excel.worksheet.to_xml` | Generating worksheet XML | `excel.records.count`, `excel.worksheet.xml_size_bytes` |

## Span Attributes

### Common Attributes
All spans include these base attributes:
- `excel.library`: Always `"svrnm/exceldatatables"`
- `excel.operation.class`: The PHP class performing the operation

### Performance Attributes
- `excel.records.count`: Number of data records processed
- `excel.columns.count`: Number of data columns processed
- `excel.file.size_bytes`: File size in bytes
- `excel.worksheet.xml_size_bytes`: Generated XML size
- `excel.file.result_size_bytes`: Result content size

### File Attributes
- `excel.file.name`: Base filename being processed
- `excel.file.path`: Full file path
- `excel.file.extension`: File extension
- `excel.file.operation`: Type of file operation (`open`, `read`, `write`, `attach`, `fill`)
- `excel.file.target`: Target filename for operations
- `excel.file.temp_path`: Temporary file path (when applicable)
- `excel.file.copied`: Boolean indicating if file was copied

### Worksheet Attributes
- `excel.worksheet.name`: Worksheet name
- `excel.worksheet.id`: Worksheet ID
- `excel.worksheet.has_calculated_columns`: Boolean for calculated columns presence
- `excel.worksheet.calculated_columns_count`: Number of calculated columns

### Export Attributes
- `excel.export.format`: Export format (`csv`, `xml`, `array`)
- `excel.csv.separator`: CSV separator character

### Table Attributes
- `excel.table.name`: Excel table name
- `excel.table.id`: Excel table ID
- `excel.table.rows`: Number of rows in table
- `excel.table.old_ref`: Original table reference
- `excel.table.new_ref`: Updated table reference

### Workbook Attributes
- `excel.workbook.auto_save`: Boolean indicating auto-save status

## Usage Examples

### Basic Usage
```php
use Svrnm\ExcelDataTables\ExcelDataTable;

// Instrumentation is automatic - no code changes needed
$dataTable = new ExcelDataTable();
$dataTable->addRows($data); // Creates span: excel.data.add_rows
$csv = $dataTable->toCsv(); // Creates span: excel.data.to_csv
$dataTable->attachToFile('template.xlsx', 'output.xlsx'); // Creates span: excel.file.attach
```

### With OpenTelemetry SDK
```php
// Configure OpenTelemetry SDK (example)
use OpenTelemetry\SDK\Trace\TracerProvider;
use OpenTelemetry\SDK\Trace\Span\Processor\SimpleSpanProcessor;
use OpenTelemetry\SDK\Trace\Span\Exporter\ConsoleSpanExporter;
use OpenTelemetry\API\Globals;

$tracerProvider = new TracerProvider(
    new SimpleSpanProcessor(new ConsoleSpanExporter())
);
Globals::registerInitializer(function() use ($tracerProvider) {
    return $tracerProvider;
});

// Your Excel operations will now be traced
$dataTable = new ExcelDataTable();
$dataTable->addRows($data); // Automatically traced
```

## Performance Considerations

The instrumentation is designed with performance in mind:

### Minimal Overhead
- Spans are only created when OpenTelemetry is properly configured
- Attribute collection uses efficient operations
- No expensive operations when spans are not recording

### Smart Attribute Collection
- File size attributes only collected when files exist
- Column counts only calculated when headers are defined
- XML size measured only when content is generated

### Best Practices for Production
1. **Configure sampling** to control trace volume
2. **Use async exporters** to minimize latency impact
3. **Monitor resource usage** with large datasets
4. **Set appropriate span limits** in your SDK configuration

## Troubleshooting

### Common Issues

**No spans are created:**
- Ensure OpenTelemetry SDK is properly configured
- Verify tracer provider is registered with `Globals::registerInitializer()`
- Check that the library can access `OpenTelemetry\API\Globals`

**Missing attributes:**
- Some attributes are conditional (e.g., file size only when file exists)
- Performance attributes depend on data availability
- Calculated column attributes only present when columns exist

**Performance impact:**
- Large datasets may create spans with many attributes
- Consider sampling for high-volume scenarios
- Monitor memory usage with very large Excel files

### Debug Mode
You can enable debug logging by configuring your OpenTelemetry SDK with appropriate log levels:

```php
// Enable debug output to console
$exporter = new ConsoleSpanExporter();
// Configure your tracer provider with this exporter
```

## Integration Examples

### With Jaeger
```php
use OpenTelemetry\Contrib\Jaeger\Exporter as JaegerExporter;
use OpenTelemetry\SDK\Trace\SpanProcessor\BatchSpanProcessor;

$exporter = new JaegerExporter(
    'ExcelDataTables',
    'http://jaeger:14268/api/traces'
);
$spanProcessor = new BatchSpanProcessor($exporter);
$tracerProvider = new TracerProvider($spanProcessor);
```

### With Zipkin
```php
use OpenTelemetry\Contrib\Zipkin\Exporter as ZipkinExporter;

$exporter = new ZipkinExporter(
    'ExcelDataTables',
    'http://zipkin:9411/api/v2/spans'
);
// Configure tracer provider...
```

### Custom Attributes
You can add custom attributes to your application spans that will be correlated with the library spans:

```php
$tracer = Globals::tracerProvider()->getTracer('your-app');
$span = $tracer->spanBuilder('process-excel-data')
    ->setAttribute('app.user_id', $userId)
    ->setAttribute('app.department', $department)
    ->startSpan();

$scope = $span->activate();
try {
    // Library operations will be child spans
    $dataTable = new ExcelDataTable();
    $dataTable->addRows($data);
    $dataTable->attachToFile('template.xlsx');
} finally {
    $scope->detach();
    $span->end();
}
```

## Contributing

When contributing to the instrumentation:

1. **Follow semantic conventions**: Use the `excel.` prefix for attributes
2. **Add meaningful attributes**: Include performance and operational metrics
3. **Handle exceptions properly**: Record exceptions and set span status
4. **Test with different scenarios**: Large datasets, missing files, etc.
5. **Document new spans**: Update this documentation for new instrumentation

## Version Compatibility

- **OpenTelemetry API**: ^1.0
- **PHP**: >=7.0.0 (same as base library)
- **ExcelDataTables**: All versions with instrumentation

The instrumentation is designed to be backward compatible and will gracefully degrade if OpenTelemetry is not available. 
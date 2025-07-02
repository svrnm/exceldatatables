<?php

/**
 * OpenTelemetry Configuration for Excel DataTables
 * 
 * This file provides a basic configuration for OpenTelemetry tracing.
 * Include this file before using the ExcelDataTables library to enable tracing.
 */

use OpenTelemetry\SDK\Trace\TracerProvider;
use OpenTelemetry\SDK\Trace\SpanProcessor\SimpleSpanProcessor;
use OpenTelemetry\SDK\Trace\SpanProcessor\BatchSpanProcessor;
use OpenTelemetry\SDK\Common\Attribute\Attributes;
use OpenTelemetry\SDK\Resource\ResourceInfo;
use OpenTelemetry\SDK\Resource\ResourceInfoFactory;
use OpenTelemetry\API\Instrumentation\Configurator;
use OpenTelemetry\Context\Propagation\ArrayAccessGetterSetter;
use OpenTelemetry\Context\Propagation\PropagationGetterInterface;
use OpenTelemetry\Context\Propagation\PropagationSetterInterface;
use OpenTelemetry\Context\Propagation\TextMapPropagatorInterface;

// Basic configuration - modify as needed
$serviceName = 'exceldatatables';
$serviceVersion = '1.0.0';

// Create resource info
$resource = ResourceInfoFactory::defaultResource()->merge(
    ResourceInfo::create(Attributes::create([
        'service.name' => $serviceName,
        'service.version' => $serviceVersion,
    ]))
);

// Example 1: In-Memory exporter for debugging (simpler alternative)
use OpenTelemetry\SDK\Trace\SpanExporter\InMemoryExporter;
$consoleExporter = new InMemoryExporter();

// Example 2: OTLP exporter (for sending to collectors like Jaeger, Zipkin, etc.)
// Uncomment and configure as needed:
/*
$otlpExporter = new SpanExporter(
    PsrTransportFactory::discover()->create('http://localhost:4318/v1/traces', 'application/json')
);
*/

// Example 3: Jaeger exporter (requires jaeger/jaeger-client-php)
// Uncomment and configure as needed:
/*
$jaegerExporter = JaegerExporter::fromConnectionString(
    'http://localhost:14268/api/traces',
    $serviceName,
    $serviceVersion
);
*/

// Create tracer provider with the exporter
$tracerProvider = TracerProvider::builder()
    ->addSpanProcessor(new SimpleSpanProcessor($consoleExporter))
    // ->addSpanProcessor(new BatchSpanProcessor($otlpExporter)) // Uncomment for OTLP
    // ->addSpanProcessor(new BatchSpanProcessor($jaegerExporter)) // Uncomment for Jaeger
    ->setResource($resource)
    ->build();

// Set the global tracer provider (store as global variable for simplicity)
$GLOBALS['_opentelemetry_tracer_provider'] = $tracerProvider;

// Optional: Set sampling configuration
// You can configure sampling to reduce the amount of traces generated
/*
$tracerProvider = TracerProvider::builder()
    ->setSampler(new TraceIdRatioBasedSampler(0.1)) // Sample 10% of traces
    ->addSpanProcessor(new SimpleSpanProcessor($consoleExporter))
    ->setResource($resource)
    ->build();
*/

echo "OpenTelemetry configured for ExcelDataTables\n";
echo "Service Name: $serviceName\n";
echo "Service Version: $serviceVersion\n";
echo "Exporter: In-Memory (modify otel-config.php to use other exporters)\n";

// Register a shutdown function to display collected spans
register_shutdown_function(function() use ($consoleExporter) {
    $spans = $consoleExporter->getSpans();
    if (!empty($spans)) {
        echo "\n" . str_repeat("=", 50) . "\n";
        echo "CAPTURED OPENTELEMETRY SPANS (" . count($spans) . " total)\n";
        echo str_repeat("=", 50) . "\n";
        
        foreach ($spans as $span) {
            echo "\n📊 Span: " . $span->getName() . "\n";
            
            // Try to get timing information if available
            try {
                if (method_exists($span, 'getEndEpochNanos') && method_exists($span, 'getStartEpochNanos')) {
                    $duration = ($span->getEndEpochNanos() - $span->getStartEpochNanos()) / 1000000;
                    echo "   Duration: " . number_format($duration, 2) . "ms\n";
                }
            } catch (\Exception $e) {
                // Ignore timing errors
            }
            
            $attributes = $span->getAttributes()->toArray();
            if (!empty($attributes)) {
                echo "   Attributes:\n";
                foreach ($attributes as $key => $value) {
                    echo "     - $key: " . (is_bool($value) ? ($value ? 'true' : 'false') : $value) . "\n";
                }
            }
            
            try {
                if ($span->getStatus()->getCode() !== 1) { // Not OK
                    echo "   ⚠️  Status: " . $span->getStatus()->getDescription() . "\n";
                }
            } catch (\Exception $e) {
                // Ignore status errors
            }
        }
        echo "\n" . str_repeat("=", 50) . "\n";
    }
}); 
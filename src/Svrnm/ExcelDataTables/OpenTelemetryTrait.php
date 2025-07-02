<?php

namespace Svrnm\ExcelDataTables;

use OpenTelemetry\API\Globals;
use OpenTelemetry\API\Trace\Span;
use OpenTelemetry\API\Trace\SpanKind;
use OpenTelemetry\API\Trace\StatusCode;
use OpenTelemetry\Context\Context;

/**
 * Trait providing OpenTelemetry instrumentation for ExcelDataTables library
 * 
 * This trait follows OpenTelemetry best practices for library instrumentation:
 * - Uses only the OpenTelemetry API (not SDK)
 * - Creates meaningful spans with semantic attributes
 * - Handles exceptions and sets span status appropriately
 * - Uses low-cardinality span names for performance
 */
trait OpenTelemetryTrait
{
    /**
     * The tracer instance for this library
     */
    private static $tracer = null;

    /**
     * Get the tracer instance for this library
     * 
     * @return \OpenTelemetry\API\Trace\TracerInterface
     */
    protected function getTracer()
    {
        if (self::$tracer === null) {
            self::$tracer = Globals::tracerProvider()->getTracer(
                'svrnm/exceldatatables',
                '1.0.0'
            );
        }
        return self::$tracer;
    }

    /**
     * Create and start a new span for Excel operations
     * 
     * @param string $operationName The operation being performed
     * @param array $attributes Additional span attributes
     * @param int $spanKind The span kind (default: INTERNAL)
     * @return \OpenTelemetry\API\Trace\SpanInterface
     */
    protected function startSpan($operationName, array $attributes = [], $spanKind = SpanKind::KIND_INTERNAL)
    {
        $spanBuilder = $this->getTracer()
            ->spanBuilder($operationName)
            ->setSpanKind($spanKind);

        // Add common attributes
        $spanBuilder->setAttribute('excel.library', 'svrnm/exceldatatables');
        $spanBuilder->setAttribute('excel.operation.class', get_class($this));
        
        // Add custom attributes
        foreach ($attributes as $key => $value) {
            $spanBuilder->setAttribute($key, $value);
        }

        return $spanBuilder->startSpan();
    }

    /**
     * Execute a callable within a span context
     * 
     * @param string $operationName The operation being performed
     * @param callable $operation The operation to execute
     * @param array $attributes Additional span attributes
     * @param int $spanKind The span kind (default: INTERNAL)
     * @return mixed The result of the operation
     * @throws \Exception If the operation throws an exception
     */
    protected function traceOperation($operationName, callable $operation, array $attributes = [], $spanKind = SpanKind::KIND_INTERNAL)
    {
        $span = $this->startSpan($operationName, $attributes, $spanKind);
        $scope = $span->activate();

        try {
            $result = $operation($span);
            $span->setStatus(StatusCode::STATUS_OK);
            return $result;
        } catch (\Exception $exception) {
            $span->recordException($exception);
            $span->setStatus(StatusCode::STATUS_ERROR, $exception->getMessage());
            throw $exception;
        } finally {
            $scope->detach();
            $span->end();
        }
    }

    /**
     * Add performance metrics to a span
     * 
     * @param \OpenTelemetry\API\Trace\SpanInterface $span
     * @param int $recordCount Number of records processed
     * @param int $columnCount Number of columns processed
     * @param int|null $fileSize File size in bytes (if applicable)
     */
    protected function addPerformanceAttributes($span, $recordCount = null, $columnCount = null, $fileSize = null)
    {
        if ($recordCount !== null) {
            $span->setAttribute('excel.records.count', $recordCount);
        }
        
        if ($columnCount !== null) {
            $span->setAttribute('excel.columns.count', $columnCount);
        }
        
        if ($fileSize !== null) {
            $span->setAttribute('excel.file.size_bytes', $fileSize);
        }
    }

    /**
     * Add file-related attributes to a span
     * 
     * @param \OpenTelemetry\API\Trace\SpanInterface $span
     * @param string|null $filename The filename being processed
     * @param string|null $operation The file operation (read, write, etc.)
     */
    protected function addFileAttributes($span, $filename = null, $operation = null)
    {
        if ($filename !== null) {
            $span->setAttribute('excel.file.name', basename($filename));
            $span->setAttribute('excel.file.path', $filename);
            $span->setAttribute('excel.file.extension', pathinfo($filename, PATHINFO_EXTENSION));
        }
        
        if ($operation !== null) {
            $span->setAttribute('excel.file.operation', $operation);
        }
    }

    /**
     * Add worksheet-related attributes to a span
     * 
     * @param \OpenTelemetry\API\Trace\SpanInterface $span
     * @param string|null $sheetName The worksheet name
     * @param int|null $sheetId The worksheet ID
     */
    protected function addWorksheetAttributes($span, $sheetName = null, $sheetId = null)
    {
        if ($sheetName !== null) {
            $span->setAttribute('excel.worksheet.name', $sheetName);
        }
        
        if ($sheetId !== null) {
            $span->setAttribute('excel.worksheet.id', $sheetId);
        }
    }
} 
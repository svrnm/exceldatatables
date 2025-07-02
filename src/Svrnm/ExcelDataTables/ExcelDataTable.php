<?php

namespace Svrnm\ExcelDataTables;

use OpenTelemetry\API\Trace\TracerInterface;
use OpenTelemetry\API\Trace\SpanInterface;

/**
 * The class ExcelDataTable converts given rows (e.g. arrays, objects) into a SpreadsheetML
 * compatible representation which can be attachted to a Exel file (.xlsx) file.
 *
 * Useage example:
 *
 * $data = array(array(1,2,3),array(1,2,3));
 * $excelDataTable = new ExcelDataTable();
 * $excelDataTable->addRows($data)->attachToFile('./example.xlsx');
 *
 * @author Severin Neumann <severin.neumann@altmuehlnet.de>
 * @license Apache-2.0 
 */
class ExcelDataTable
{

		/**
		 * OpenTelemetry tracer instance
		 *
		 * @var TracerInterface
		 */
		private $tracer;

		/**
		 * The internal representation of the data table
		 *
		 * @var array
		 */
		protected $data = array();

		/**
		 * The names/identifiers for the data columns
		 *
		 * @var array
		 */
		protected $headerNames = array();

		/**
		 * The numbers for the data columns, may be sparse
		 *
		 * @var array
		 */
		protected $headerNumbers = array();

		/**
		 * The (optional) labels of the data column. If $headersVisible is true
		 * these are written in the first row ("Row 1" in Excel)
		 *
		 * @var array
		 */
		protected $headerLabels = array();

		/**
		 * A list of cells which couldn't be added during a call of addRow()
		 *
		 * @var array
		 */
		protected $failedCells = array();

		/**
		 * True if the headers are defined.
		 *
		 * @var boolean
		 */
		protected $headersDefined = false;

		/**
		 * True if headers should be displayed during an export
		 *
		 * @var boolean
		 */
		protected $headersVisible = false;

		/**
		 * The sheet which will be overwritten when attachToFile is called
		 *
		 * @var int|null
		 */
		protected $sheetId = null;

		/**
		 * The name of the sheet when attachToFile is called
		 *
		 * @var string
		 */
		protected $sheetName = 'Data';

		/**
		 * If set, regenerates the range in the data table with the specified name
		 *
		 * @var null|string
		 */
		protected $refreshTableRange = null;

		/**
		 * If set, injects column formulas into the output
		 *
		 * @var null|string
		 */
		protected $preserveFormulas = null;

		/**
		 * Variable to hold calculated columns from source
		 *
		 * @var null|array
		 */
		protected $calculatedColumns = null;

		/**
		 * Instantiate a new ExcelDataTable object
		 *
		 */
		public function __construct() {
			$tracerProvider = $GLOBALS['_opentelemetry_tracer_provider'] ?? null;
			if ($tracerProvider) {
				$this->tracer = $tracerProvider->getTracer(
					'exceldatatables',
					'1.0.0',
					'https://github.com/svrnm/exceldatatables'
				);
			}
		}

		/**
		 * Helper method to create a span if tracer is available
		 *
		 * @param string $name
		 * @param array $attributes
		 * @return SpanInterface|null
		 */
		private function createSpan($name, $attributes = []) {
			if (!$this->tracer) return null;
			
			$spanBuilder = $this->tracer->spanBuilder($name);
			foreach ($attributes as $key => $value) {
				$spanBuilder->setAttribute($key, $value);
			}
			return $spanBuilder->startSpan();
		}

		/**
		 * Add multiple rows to the data table. $rows is expected to be an array of
		 * possible rows (array, object).
		 *
		 * @param array $rows
		 * @return $this
		 */
		public function addRows($rows) {
				$span = $this->createSpan('ExcelDataTable.addRows', ['rows.count' => count($rows)]);
				
				try {
					foreach($rows as $row) {
							$this->addRow($row);
					}
					return $this;
				} catch (\Exception $e) {
					if ($span) $span->recordException($e);
					throw $e;
				} finally {
					if ($span) $span->end();
				}
		}

		/**
		 * Add a new row to the data table. $row is expected to be an assocative array
		 * or an object, the keys/property names are used to choose the correct
		 * column in the following manner:
		 *
		 * - If no header property for the data table is specified, the given $row is used to define the header
		 * - If a header property for the date table is specified, the given $row is added accordingly
		 *
		 * @param array $row
		 * @return $this
		 */
		public function addRow($row) {
				$span = $this->tracer->spanBuilder('ExcelDataTable.addRow')
					->setAttribute('row.columns', count($row))
					->setAttribute('headers.defined', $this->headersDefined)
					->startSpan();
				
				try {
					$result = array();

					if(!$this->headersDefined) {
							$headers = array();
							foreach($row as $key => $value) {
									$headers[$key] = $key;
							}
							$this->setHeaders($headers);
							$span->setAttribute('headers.auto_defined', true);
					}


					foreach($row as $name => $value) {
							$number = $this->headerNameToHeaderNumber($name);
							if($number !== false) {
									$result[$number] = $value;
							} else {
									$this->failedCells[count($this->data)] = array($name => $value);
							}
					}
					$this->data[] = $result;
					
					$span->setAttribute('failed.cells', count($this->failedCells));
					$span->setAttribute('total.rows', count($this->data));

					return $this;
				} catch (\Exception $e) {
					$span->recordException($e);
					throw $e;
				} finally {
					$span->end();
				}
		}

		/**
		 * Return a list of cells which couldn't be inserted during all previous calls of "addRow"
		 * The reason for this might be, that the key for a certain value is not defined as header
		 * and therfore the value can't be placed in a cell
		 *
		 * @return array A copy of $this->failedCells
		 */
		public function getFailedCells() {
				return $this->failedCells;
		}

		/**
		 * Convert the name of a header/column into the equivalent number. Returns false if
		 * the given name can't be converted.
		 *
		 * @param string $name
		 * @return int|boolean
		 */
		protected function headerNameToHeaderNumber($name) {
				return isset($this->headerNumbers[$name]) ? $this->headerNumbers[$name] : false;
		}

		/**
		 * Convert the number of a header/column into the equivalent name. Returns false if
		 * the given name can't be converted.
		 *
		 * @param int $number
		 * @return string|boolean
		 */
		protected function headerNumberToHeaderName($number) {
				return isset($this->headerNames[$number]) ? $this->headerNames[$number] : false;
		}

		/**
		 * Retrieve the label for a certain column by header/column number. Retruns false if the
		 * requested label can't be retrieved.
		 *
		 * @param int $number
		 * @return string|boolean
		 */
		protected function getHeaderLabelByNumber($number) {
				return isset($this->headerLabels[$number]) ? $this->headerLabels[$number] : false;
		}

		/**
		 * Add a name with label for the next column. Increments the column count by one.
		 *
		 * @param string name
		 * @param string label
		 * @return this
		 */
		protected function addHeader($name, $label) {
				$this->headerNames[] = $name;
				$this->headerLabels[] = $label;
				$this->headerNumbers[$name] = count($this->headerNames) - 1;
				return $this;
		}

		/**
		 * Set the label of a certain column. The column is selected by its name/identifier.
		 *
		 * @param string name
		 * @param string label
		 * @return this
		 */
		public function setLabel($name, $label) {
				$this->headerLabels[$this->headerNameToHeaderNumber($name)] = $label;
				return $this;
		}

		/**
		 * Show the header row during export
		 *
		 * @return this
		 */
		public function showHeaders() {
				$this->headersVisible = true;
				return $this;
		}

		/**
		 * Do not show the header row during export
		 *
		 * @return this
		 */
		public function hideHeaders() {
				$this->headersVisible = false;
				return $this;
		}

		/**
		 * Check if the headers are shown during export
		 *
		 * @return boolean
		 */
		public function areHeadersVisible() {
				return $this->headersVisible;
		}

		/**
		 * Set headers for the data table. Expects an array or an object, which will
		 * be casted to an array. The keys are used as names for the columns, the values
		 * are used as labels for the columns, if the header is printed (see: showHeader()/hideHeader())
		 *
		 * @param array|object $header
		 * @return this
		 */
		public function setHeaders($header) {
				foreach((array)$header as $name => $label) {
						$this->addHeader($name, $label);
				}
				$this->headersDefined = true;
				return $this;
		}

		/**
		 * Returns a representation of the header row.
		 *
		 * @return array
		 */
		public function getHeaders() {
				$result = array();
				foreach($this->headerNumbers as $number) {
						$result[$this->headerNumberToHeaderName($number)] = $this->getHeaderLabelByNumber($number);
				}
				return $result;
		}

		/**
		 * Change the type of the column identified by $columnKey. $type can be 'string', 'number', 'date', 'datetime'.
		 * $columnKey can be a numeric identifier or a key as specified by addHeader().
		 *
		 * NOT YET IMPLEMENTED
		 *
		 * @param int|string columnKey
		 * @param string type
		 * @return this
		 */
		public function setColumnType($columnKey, $type) {
				throw new \Exception('"setColumnType" is not yet implemented');
				return $this;
		}

		/**
		 * Iterate over a given row and convert it into a dense representation
		 * for export.
		 *
		 * @param array arr
		 * @return array
		 */
		protected function fillRow($arr) {
				$result = array();
				for($i = 0; $i < count($this->headerNumbers); $i++) {
						$result[$i] = isset($arr[$i]) ? $arr[$i] : '';
				}
				return $result;
		}

		/**
		 * Exports the data table into an array representation. If the headers are visible
		 * they are at index 0 of the multidimensional array
		 *
		 * @return array
		 */
		public function toArray() {
				$span = $this->tracer->spanBuilder('ExcelDataTable.toArray')
					->setAttribute('data.rows', count($this->data))
					->setAttribute('headers.visible', $this->areHeadersVisible())
					->startSpan();
				
				try {
					$arr = array();
					if($this->areHeadersVisible()) {
							$arr[] = $this->headerLabels;
					}
					foreach($this->data as $row) {
							$arr[] = $this->fillRow($row);
					}
					$span->setAttribute('result.rows', count($arr));
					return $arr;
				} catch (\Exception $e) {
					$span->recordException($e);
					throw $e;
				} finally {
					$span->end();
				}
		}

		/**
		 * Exports the data table into a csv formatted string. If the headers are visible
		 * they are added as first row
		 *
		 * @return string
		 */
		public function toCsv($separator = ',', $quote = '', $newLine = PHP_EOL) {
				$span = $this->tracer->spanBuilder('ExcelDataTable.toCsv')
					->setAttribute('data.rows', count($this->data))
					->setAttribute('csv.separator', $separator)
					->setAttribute('csv.quote', $quote)
					->startSpan();
				
				try {
					$result = implode(
							$newLine,
							array_map(
									function($elem) use($separator, $quote) {
											$s = $quote.$separator.$quote;
											return $quote.implode($s, $elem).$quote;
									},
											$this->toArray()
									)
							);
					
					$span->setAttribute('csv.size', strlen($result));
					return $result;
				} catch (\Exception $e) {
					$span->recordException($e);
					throw $e;
				} finally {
					$span->end();
				}
		}

		/**
		 * Creates and returns the created worksheet as spreadsheetml
		 *
		 * @return string
		 */
		public function toXML() {
				$span = $this->tracer->spanBuilder('ExcelDataTable.toXML')
					->setAttribute('data.rows', count($this->data))
					->startSpan();
				
				try {
					$worksheet = new ExcelWorksheet();
					$result = $worksheet->addRows($this->toArray())->toXML();
					$span->setAttribute('xml.size', strlen($result));
					return $result;
				} catch (\Exception $e) {
					$span->recordException($e);
					throw $e;
				} finally {
					$span->end();
				}
		}

		/**
		 * Return a string representation of the data table.
		 *
		 * This is currently equivalent to a call of toCsv(), but might change in future releases.
		 *
		 * @return string
		 */
		public function __toString() {
				$r = $this->toCsv();
				return $r;
		}

		/**
		 * Change the id of the sheet which will be overwritten when attachToFile is called.
		 *
		 * @param int id
		 * @return $this
		 */
		public function setSheetId($id) {
				$this->sheetId = $id;
				return $this;
		}

		/**
		 * Change the name of the sheet which will be attached when attachToFile is called
		 *
		 * @param string name
		 * @return $this
		 */
		public function setSheetName($name) {
				$this->sheetName = $name;
				return $this;
		}

		/**
		 * Attach the data table to an existing xlsx file. The file location is given via the
		 * first parameter. If a second parameter is given the source file will not be overwritten
		 * and a new file will be created. The third parameter can be used to force updating the
		 * auto calculation in the excel workbook.
		 *
		 * @param string srcFilename
		 * @param string|null targetFilename
		 * @param bool|null forceAutoCalculation
		 * @return $this
		 */
		public function attachToFile($srcFilename, $targetFilename = null, $forceAutoCalculation = false) {
				$span = $this->tracer->spanBuilder('ExcelDataTable.attachToFile')
					->setAttribute('source.file', basename($srcFilename))
					->setAttribute('target.file', $targetFilename ? basename($targetFilename) : basename($srcFilename))
					->setAttribute('sheet.name', $this->sheetName)
					->setAttribute('sheet.id', $this->sheetId)
					->setAttribute('data.rows', count($this->data))
					->setAttribute('auto.calculation', $forceAutoCalculation)
					->setAttribute('preserve.formulas', !is_null($this->preserveFormulas))
					->setAttribute('refresh.table.range', !is_null($this->refreshTableRange))
					->startSpan();
				
				try {
					$calculatedColumns = null;
					if ($this->preserveFormulas){
							$temp_xlsx = new ExcelWorkbook($srcFilename);
							$calculatedColumns = $temp_xlsx->getCalculatedColumns($this->preserveFormulas);
							unset($temp_xlsx);
					}

					$xlsx = new ExcelWorkbook($srcFilename);
					$worksheet = new ExcelWorksheet();
					if(!is_null($targetFilename)) {
							$xlsx->setFilename($targetFilename);
					}
					$worksheet->addRows($this->toArray(), $calculatedColumns);
					$xlsx->addWorksheet($worksheet, $this->sheetId, $this->sheetName);
					if($forceAutoCalculation) {
						$xlsx->enableAutoCalculation();
					}

					if ($this->refreshTableRange) {
						$xlsx->refreshTableRange($this->refreshTableRange, count($this->data) + 1);
					}

					$xlsx->save();
					unset($xlsx);
					
					$span->setAttribute('operation.success', true);
					return $this;
				} catch (\Exception $e) {
					$span->recordException($e);
					$span->setAttribute('operation.success', false);
					throw $e;
				} finally {
					$span->end();
				}
		}

		/**
		 * This functions takes an XLSX-file and an multidimensional and returns a string representation of the
		 * XLSX file including the data table.
		 * This function is especially useful if the file should be provided as download for a http request.
		 *
		 * @param string srcFilename
		 * @return string
		 */
		public function fillXLSX($srcFilename) {
				$span = $this->tracer->spanBuilder('ExcelDataTable.fillXLSX')
					->setAttribute('source.file', basename($srcFilename))
					->setAttribute('data.rows', count($this->data))
					->startSpan();
				
				try {
					$targetFilename = tempnam(sys_get_temp_dir(), 'exceldatatables-');
					$span->setAttribute('temp.file', basename($targetFilename));
					
					$this->attachToFile($srcFilename, $targetFilename);
					$result = file_get_contents($targetFilename);
					unlink($targetFilename);
					
					$span->setAttribute('result.size', strlen($result));
					return $result;
				} catch (\Exception $e) {
					$span->recordException($e);
					throw $e;
				} finally {
					$span->end();
				}
		}

		/**
		 * This function regenerates the range of the dynamic table to match
		 * the total rows inserted
		 *
		 * @param string $table_name name of the excel table
		 * @return $this
		 */
		public function refreshTableRange($table_name = null)
		{
				$table_name = !is_null($table_name) ? $table_name : $this->sheetName;
				$this->refreshTableRange = $table_name;
				return $this;
		}

		/**
		 * This function extracts the existing column formulas and injects them.
		 *
		 * @param string $table_name name of the excel table
		 * @return $this
		 */
		public function preserveFormulas($table_name)
		{
				$table_name = !is_null($table_name) ? $table_name : $this->sheetName;
				$this->preserveFormulas = $table_name;
				return $this;
		}

}

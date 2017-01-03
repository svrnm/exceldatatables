<?php

namespace Svrnm\ExcelDataTables;

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

		}

		/**
		 * Add multiple rows to the data table. $rows is expected to be an array of
		 * possible rows (array, object).
		 *
		 * @param array $rows
		 * @return $this
		 */
		public function addRows($rows) {
				foreach($rows as $row) {
						$this->addRow($row);
				}
				return $this;
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
				$result = array();

				if(!$this->headersDefined) {
						$headers = array();
						foreach($row as $key => $value) {
								$headers[$key] = $key;
						}
						$this->setHeaders($headers);
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

				return $this;
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
				$arr = array();
				if($this->areHeadersVisible()) {
						$arr[] = $this->headerLabels;
				}
				foreach($this->data as $row) {
						$arr[] = $this->fillRow($row);
				}
				return $arr;
		}

		/**
		 * Exports the data table into a csv formatted string. If the headers are visible
		 * they are added as first row
		 *
		 * @return string
		 */
		public function toCsv($separator = ',', $quote = '', $newLine = PHP_EOL) {
				return implode(
						$newLine,
						array_map(
								function($elem) use($separator, $quote) {
										$s = $quote.$separator.$quote;
										return $quote.implode($s, $elem).$quote;
								},
										$this->toArray()
								)
						);

		}

		/**
		 * Creates and returns the created worksheet as spreadsheetml
		 *
		 * @return string
		 */
		public function toXML() {
				$worksheet = new ExcelWorksheet();
				return $worksheet->addRows($this->toArray())->toXML();
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
				return $this;
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
				$targetFilename = tempnam(sys_get_temp_dir(), 'exceldatatables-');
				$this->attachToFile($srcFilename, $targetFilename);
				$result = file_get_contents($targetFilename);
				unlink($targetFilename);
				return $result;
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

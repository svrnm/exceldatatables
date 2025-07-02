<?php

namespace Svrnm\ExcelDataTables;

use OpenTelemetry\API\Instrumentation\CachedInstrumentation;
use OpenTelemetry\API\Trace\Propagation\TraceContextPropagator;
use OpenTelemetry\API\Trace\SpanKind;
use OpenTelemetry\API\Trace\StatusCode;
use OpenTelemetry\SemConv\TraceAttributes;

/**
 * This class is a simple representation of a excel workbook. It expects
 * a xlsx formatted spreadsheet as paramater and overwrites(!) existing
 * worksheets with new data.
 *
 * @author Severin Neumann <severin.neumann@altmuehlnet.de>
 * @license Apache-2.0
 */
class ExcelWorkbook implements \Countable
{
		/**
		 * @var CachedInstrumentation
		 */
		private static $instrumentation;

		/**
		 * The source filename
		 *
		 * @var string
		 */
		protected $srcFilename;

		/**
		 * The target filename
		 *
		 * @var string
		 */
		protected $targetFilename;


		/**
		 * The ZipArchive representation of the workbook
		 *
		 * @var \ZipArchive
		 */
		protected $xlsx;

		/**
		 * The DomDocument representation of the workbook
		 *
		 * @var \DOMDocument
		 */
		protected $workbook;

		/**
		 * The DOMDocument representation of the styles.xml included in the workbook
		 *
		 * @var \DOMDocument
		 */
		protected $styles;


		/**
		 * If true all operations are closing and reopening the zip archive to store
		 * modifications
		 *
		 * @var boolean
		 */
		protected $autoSave = false;

		/**
		 * If true the workbook is modified to to contain the fullCalcOnLoad attribute
		 *
		 * @var boolean
		 */
		protected $autoCalculation = false;

		/**
		 * The default name of the sheet when attachToFile is called
		 *
		 * @var string
		 */
		protected $sheetName = 'Data';

		/**
		 * Instantiate a new object of the type ExcelWorkbook. Expects a filename which
		 * contains a spreadsheet of type xlsx.
		 *
		 * @param string $filename
		 */
		public function __construct($filename) {
				self::$instrumentation ??= new CachedInstrumentation('svrnm-exceldatatables');
				
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.__construct')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.init')
						->setAttribute('excel.source_file', basename($filename))
						->setAttribute(TraceAttributes::CODE_FUNCTION, '__construct')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$this->srcFilename = $filename;
						$this->targetFilename = $filename;
						
						$span->addEvent('workbook.initialized', [
								'excel.source_filename' => basename($filename),
								'excel.auto_save' => $this->autoSave,
								'excel.auto_calculation' => $this->autoCalculation
						]);
						
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		public function __destruct() {
				if (self::$instrumentation === null) {
						if(!is_null($this->xlsx)) {
							$this->getXLSX()->close();
						}
						return;
				}

				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.__destruct')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.cleanup')
						->setAttribute(TraceAttributes::CODE_FUNCTION, '__destruct')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						if(!is_null($this->xlsx)) {
							$this->getXLSX()->close();
							$span->addEvent('workbook.xlsx_closed');
						}
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 * Turn on auto saving
		 *
		 * @return $this
		 */
		public function enableAutoSave() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.enableAutoSave')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.config')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'enableAutoSave')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$this->autoSave = true;
						$span->addEvent('workbook.auto_save_enabled');
						return $this;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 * Turn on auto calculation
		 *
		 * @return $this
		 */
		public function enableAutoCalculation($value = true) {
			$tracer = self::$instrumentation->tracer();
			$span = $tracer->spanBuilder('ExcelWorkbook.enableAutoCalculation')
					->setSpanKind(SpanKind::KIND_INTERNAL)
					->setAttribute('excel.operation', 'workbook.config')
					->setAttribute('excel.auto_calculation', (bool)$value)
					->setAttribute(TraceAttributes::CODE_FUNCTION, 'enableAutoCalculation')
					->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
					->startSpan();

			$scope = $span->activate();
			
			try {
				$this->autoCalculation = (bool)$value;

				$calcPr = $this->getWorkbook()->getElementsByTagName('calcPr')->item(0);

				if(is_null($calcPr)) {
					$calcPr = $this->getWorkbook()->createElement('calcPr');
					$span->addEvent('workbook.calc_pr_created');
				}

				$calcPr->setAttribute('fullCalcOnLoad', $this->autoCalculation ? '1' : '0');
				$span->addEvent('workbook.calc_pr_updated', [
					'excel.full_calc_on_load' => $this->autoCalculation
				]);

				$this->saveWorkbook();

				return $this;
			} catch (\Throwable $e) {
				$span->recordException($e);
				$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
				throw $e;
			} finally {
				$span->end();
				$scope->detach();
			}
		}

		/**
		 * Turn off auto saving
		 *
		 * @return $this
		 */
		public function disableAutoSave() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.disableAutoSave')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.config')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'disableAutoSave')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$this->autoSave = false;
						$span->addEvent('workbook.auto_save_disabled');
						return $this;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 * Return true if auto saving is enabled
		 *
		 * @return $this
		 */
		public function isAutoSaveEnabled() {
				return $this->autoSave;
		}

		/**
		 * Returns the count of worksheets contained within the workbook
		 *
		 * @return int
		 */
		public function count(): int {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.count')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.count')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'count')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$sheets = $this->getWorkbook()->getElementsByTagName('sheet');
						$count = $sheets->length;
						
						$span->setAttribute('excel.worksheet_count', $count);
						$span->addEvent('workbook.worksheets_counted', ['excel.count' => $count]);
						
						return $count;
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 * Change the filepath where the modified XLSX is stored. This should
		 * be called before a worksheet is added.
		 *
		 * @param string $filename
		 * @return $this
		 */
		public function setFilename($filename) {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.setFilename')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.config')
						->setAttribute('excel.target_file', basename($filename))
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'setFilename')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$oldFilename = $this->targetFilename;
						$this->targetFilename = $filename;
						
						$span->addEvent('workbook.target_filename_changed', [
								'excel.old_filename' => basename($oldFilename),
								'excel.new_filename' => basename($filename)
						]);
						
						return $this;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 * Get the filepath where the modifiex XLSX will be stored.
		 *
		 * @return string
		 */
		public function getFilename() {
				return $this->targetFilename;
		}

		/**
		 * Get the XML representation of the worksheet with id $id. If the second
		 * parameter is set not false the result is returned as DOMDocument.
		 *
		 * @param int id
		 * @param boolean asDocument
		 * @return string|DOMDocument
		 */
		public function getWorksheetById($id, $asDocument = false) {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getWorksheetById')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_worksheet')
						->setAttribute('excel.worksheet_id', $id)
						->setAttribute('excel.as_document', $asDocument)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getWorksheetById')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$r = $this->getXLSX()->getFromName('xl/worksheets/sheet'.$id.'.xml');
						if($asDocument && $r !== false) {
								$dom = new \DOMDocument();
								$dom->loadXML($r);
								$span->addEvent('workbook.worksheet_loaded_as_document', ['excel.worksheet_id' => $id]);
								return $dom;
						} else {
								$span->addEvent('workbook.worksheet_loaded_as_string', [
									'excel.worksheet_id' => $id,
									'excel.success' => $r !== false
								]);
								return $r;
						}
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 *
		 */
		public function getSheetIdByName($name)
		{
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getSheetIdByName')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_sheet_id')
						->setAttribute('excel.sheet_name', $name)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getSheetIdByName')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$sheets = $this->getWorkbook()->getElementsByTagName('sheet');
						foreach ($sheets as $index => $sheet) {
								if ($sheet->getAttribute('name') === $name) {
										$id = $sheet->getAttribute('sheetId');
										$span->addEvent('workbook.sheet_found', [
											'excel.sheet_name' => $name,
											'excel.sheet_id' => $id,
											'excel.sheet_index' => $index
										]);
										return $id;
								}
						}
						
						$span->addEvent('workbook.sheet_not_found', ['excel.sheet_name' => $name]);
						return null;
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		public function getTableIdByName($name)
		{
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getTableIdByName')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_table_id')
						->setAttribute('excel.table_name', $name)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getTableIdByName')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$workbookRels = $this->getXLSX()->getFromName('xl/_rels/workbook.xml.rels');
						if ($workbookRels === false) {
							$span->addEvent('workbook.workbook_rels_not_found');
							return null;
						}
						
						$relsDom = new \DOMDocument();
						$relsDom->loadXML($workbookRels);
						$relationships = $relsDom->getElementsByTagName('Relationship');

						for ($i = 1; $i <= 100; $i++) {
								$tablePath = 'xl/tables/table' . $i . '.xml';
								$tableXML = $this->getXLSX()->getFromName($tablePath);
								if ($tableXML === false) continue;

								$tableDom = new \DOMDocument();
								$tableDom->loadXML($tableXML);
								$tables = $tableDom->getElementsByTagName('table');
								if ($tables->length > 0) {
										$table = $tables->item(0);
										if ($table->getAttribute('displayName') === $name) {
											$span->addEvent('workbook.table_found', [
												'excel.table_name' => $name,
												'excel.table_id' => $i,
												'excel.table_path' => $tablePath
											]);
											return $i;
										}
								}
						}
						
						$span->addEvent('workbook.table_not_found', ['excel.table_name' => $name]);
						return null;
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		protected function uniqName($name) {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.uniqName')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.generate_unique_name')
						->setAttribute('excel.requested_name', $name)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'uniqName')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$sheets = $this->getWorkbook()->getElementsByTagName('sheet');
						$names = array();
						foreach ($sheets as $sheet) {
								$names[] = $sheet->getAttribute('name');
						}
						$i = 0;
						$newName = $name;
						while (in_array($newName, $names)) {
								$newName = $name . ' (' . (++$i) . ')';
						}
						
						$span->addEvent('workbook.unique_name_generated', [
							'excel.requested_name' => $name,
							'excel.unique_name' => $newName,
							'excel.iterations' => $i
						]);
						
						return $newName;
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		protected function dateTimeFormatId() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.dateTimeFormatId')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_datetime_format')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'dateTimeFormatId')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
						$styles = $this->getStyles();

						$cellXfs = $styles->getElementsByTagName('cellXfs')->item(0);
						if (is_null($cellXfs)) {
							$span->addEvent('workbook.no_cellxfs_found');
							return 1;
						}
						
						$numFmts = $styles->getElementsByTagName('numFmts')->item(0);
						if (is_null($numFmts)) {
								$numFmts = $styles->createElement('numFmts');
								$numFmts->setAttribute('count', '1');
								$styles->getElementsByTagName('styleSheet')->item(0)->insertBefore($numFmts, $cellXfs->parentNode);
								$span->addEvent('workbook.numfmts_created');
						}

						$id = 164;
						$numFmt = $styles->createElement('numFmt');
						$numFmt->setAttribute('numFmtId', $id);
						$numFmt->setAttribute('formatCode', 'dd/mm/yyyy hh:mm:ss');

						$numFmts->appendChild($numFmt);
						$numFmts->setAttribute('count', $numFmts->getElementsByTagName('numFmt')->length);

						$xf = $styles->createElement('xf');
						$xf->setAttribute('numFmtId', $id);
						$xf->setAttribute('fontId', '0');
						$xf->setAttribute('fillId', '0');
						$xf->setAttribute('borderId', '0');
						$xf->setAttribute('xfId', '0');
						$xf->setAttribute('applyNumberFormat', '1');

						$cellXfs->appendChild($xf);
						$cellXfs->setAttribute('count', $cellXfs->getElementsByTagName('xf')->length);

						$this->saveStyles();

						$formatId = $cellXfs->getElementsByTagName('xf')->length - 1;
						
						$span->addEvent('workbook.datetime_format_created', [
							'excel.format_id' => $formatId,
							'excel.num_fmt_id' => $id
						]);
						
						return $formatId;
				} catch (\Throwable $e) {
						$span->recordException($e);
						$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
						throw $e;
				} finally {
						$span->end();
						$scope->detach();
				}
		}

		/**
		 * Add a worksheet into the workbook with id $id and name $name. If $id is null the last
		 * worksheet is replaced. If $name is empty, its default value is set to the default.
		 *
		 * Currently this replaces an existing worksheet. Adding new worksheets is not yet supported
		 *
		 * @param ExcelWorksheet $worksheet
		 * @param int $id
		 * @param string $name
		 */
		public function addWorksheet(ExcelWorksheet $worksheet, $id = null, $name = null) {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.addWorksheet')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.add_worksheet')
						->setAttribute('excel.worksheet_name', $name ?: $this->sheetName)
						->setAttribute('excel.worksheet_id', $id)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'addWorksheet')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					$name = !is_null($name) ? $name : $this->sheetName;
					if ($id === null) $id = $this->getSheetIdByName($name);

					$span->addEvent('workbook.worksheet_lookup_complete', [
						'excel.resolved_name' => $name,
						'excel.resolved_id' => $id
					]);

					if(is_null($id) || $id <= 0) {
						$span->addEvent('workbook.worksheet_not_found', ['excel.sheet_name' => $name]);
						throw new \Exception('Sheet with name "'.$name.'" not found in file '.$this->srcFilename.'. Appending is not yet implemented.');
						/*
						// find a unused id in the worksheets
						$id = 1;
						while($this->getXLSX()->statName('xl/worksheets/sheet'.($id++).'.xml') !== false) {}
						*/
					}

					$old = $this->getXLSX()->getFromName('xl/worksheets/sheet'.$id.'.xml');
					if($old === false) {
							$span->addEvent('workbook.worksheet_file_not_found', ['excel.worksheet_id' => $id]);
							throw new \Exception('Appending new sheets is not yet implemented: SheetId:' . $id .', SourceFile:'. $this->srcFilename.', TargetFile:'.$this->targetFilename);
					} else {
							$span->addEvent('workbook.worksheet_replacement_start', ['excel.worksheet_id' => $id]);
							
							$document = new \DOMDocument();
							$document->loadXML($old);
							$oldSheetData = $document->getElementsByTagName('sheetData')->item(0);
							$worksheet->setDateTimeFormatId($this->dateTimeFormatId());
							$newSheetData = $document->importNode( $worksheet->getDocument()->getElementsByTagName('sheetData')->item(0), true );
							$oldSheetData->parentNode->replaceChild($newSheetData, $oldSheetData);
							$xml = $document->saveXML();
							$this->getXLSX()->addFromString('xl/worksheets/sheet'.$id.'.xml', $xml);
							
							$span->addEvent('workbook.worksheet_replacement_complete', [
								'excel.worksheet_id' => $id,
								'excel.xml_size' => strlen($xml)
							]);
					}
					if($this->isAutoSaveEnabled()) {
							$span->addEvent('workbook.auto_save_triggered');
							$this->save();
					}
					return $this;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		/**
		 * Refresh the table range in the excel with the number of rows added
		 *
		 * @param string $tableName Name of the table
		 * @param int $numRows number of rows
		 * @return $this
		 */
		public function refreshTableRange($tableName, $numRows)
		{
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.refreshTableRange')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.refresh_table_range')
						->setAttribute('excel.table_name', $tableName)
						->setAttribute('excel.num_rows', $numRows)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'refreshTableRange')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					$id = $this->getTableIdByName($tableName);
					if (is_null($id)) {
						$span->addEvent('workbook.table_not_found', ['excel.table_name' => $tableName]);
						throw new \Exception('table "' . $tableName . '" not found');
					}

					$span->addEvent('workbook.table_found', ['excel.table_id' => $id]);

					$document = new \DOMDocument();
					$document->loadXML($this->getXLSX()->getFromName('xl/tables/table' . $id . '.xml'));

					$table = $document->getElementsByTagName('table')->item(0);
					if (is_null($table)) {
						$span->addEvent('workbook.table_element_not_found');
						throw new \Exception('could not read "table" from document; '.$document);
					}

					$ref = $table->getAttribute('ref');
					$nref = preg_replace('/^(\w+\:[A-Z]+)(\d+)$/', '${1}' . $numRows, $ref);

					$span->addEvent('workbook.table_range_updated', [
						'excel.old_ref' => $ref,
						'excel.new_ref' => $nref,
						'excel.num_rows' => $numRows
					]);

					$table->setAttribute('ref', $nref);

					$this->getXLSX()->addFromString('xl/tables/table' . $id . '.xml', $document->saveXML());

					return $this;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		public function getCalculatedColumns($tableName)
		{
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getCalculatedColumns')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_calculated_columns')
						->setAttribute('excel.table_name', $tableName)
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getCalculatedColumns')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					$id = $this->getTableIdByName($tableName);
					if(isset($id)){
							$span->addEvent('workbook.calculated_columns_processing_start', ['excel.table_id' => $id]);
							
							$document = new \DOMDocument();
							$document->loadXML($this->getXLSX()->getFromName('xl/tables/table' . $id . '.xml'));
							$columns = $document->getElementsByTagName('tableColumn');
							$calculatedColumns = [];
							$calculatedCount = 0;
							
							foreach($columns as $key => $column) {
									if($column->getElementsByTagName("calculatedColumnFormula")->length){
											$header = $column->getAttribute('name');
											$formula = $column->nodeValue;
											$calculatedColumn = array(
													'index' => $key,
													'header' => $header,
													'content' => array(
															$header => array(
																	'type' => 'formula',
																	'value' => $formula,
															)
													)
											);
											$calculatedColumns[] = $calculatedColumn;
											$calculatedCount++;
									}
							}
							
							$span->addEvent('workbook.calculated_columns_processed', [
								'excel.total_columns' => $columns->length,
								'excel.calculated_columns_count' => $calculatedCount
							]);
							
							return $calculatedColumns;
					}
					
					$span->addEvent('workbook.table_id_not_found', ['excel.table_name' => $tableName]);
					return null;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		/**
		 * Return the ZipArchive representation of the current workbook
		 *
		 * @return ZipArchive
		 */
		public function getXLSX() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getXLSX')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_xlsx')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getXLSX')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					if(is_null($this->xlsx)) {
							$span->addEvent('workbook.xlsx_not_initialized');
							$this->openXLSX();
					} else {
							$span->addEvent('workbook.xlsx_already_available');
					}
					return $this->xlsx;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		/**
		 * Open the excel file and create the ZipArchive representation. If the file
		 * does not exists or is not valid an exception is thrown.
		 *
		 * @throws Excepetion
		 * @return $this
		 */
		protected function openXLSX() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.openXLSX')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.open_xlsx')
						->setAttribute('excel.source_file', basename($this->srcFilename))
						->setAttribute('excel.target_file', basename($this->targetFilename))
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'openXLSX')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					$this->xlsx = new \ZipArchive;
					$span->addEvent('workbook.ziparchive_created');
					
					if(!file_exists($this->srcFilename) || !is_readable($this->srcFilename)) {
						$span->addEvent('workbook.source_file_not_accessible', [
							'excel.file_exists' => file_exists($this->srcFilename),
							'excel.file_readable' => is_readable($this->srcFilename)
						]);
						throw new \Exception('File does not exists: '.$this->srcFilename);
					}
					
					if($this->srcFilename !== $this->targetFilename) {
						$span->addEvent('workbook.copying_source_to_target');
						$sourceSize = filesize($this->srcFilename);
						file_put_contents($this->targetFilename, file_get_contents($this->srcFilename));
						$this->srcFilename = $this->targetFilename;
						$span->addEvent('workbook.file_copied', ['excel.file_size' => $sourceSize]);
					}
					
					$isOpen = $this->xlsx->open($this->targetFilename);
					if($isOpen !== true) {
						$span->addEvent('workbook.xlsx_open_failed', ['excel.error_code' => $isOpen]);
						throw new \Exception('Could not open file: '.$this->targetFilename.' [ZipArchive error code: '.$isOpen.']');
					}
					
					$span->addEvent('workbook.xlsx_opened_successfully');
					return $this;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		/**
		 * Save the modifications
		 *
		 * @return $this
		 */
		public function save() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.save')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.save')
						->setAttribute('excel.target_file', basename($this->targetFilename))
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'save')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					$span->addEvent('workbook.save_start');
					
					$this->getXLSX()->close();
					$span->addEvent('workbook.xlsx_closed');
					
					$this->srcFilename = $this->targetFilename;
					$span->addEvent('workbook.filename_synchronized');
					
					$this->openXLSX();
					$span->addEvent('workbook.xlsx_reopened');
					
					$fileSize = file_exists($this->targetFilename) ? filesize($this->targetFilename) : 0;
					$span->addEvent('workbook.save_complete', ['excel.file_size' => $fileSize]);
					
					return $this;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}


		/**
		 * Return the DOMDocument representation of the current workbook
		 *
		 * @return DOMDocument
		 */
		public function getWorkbook() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getWorkbook')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_workbook_xml')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getWorkbook')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					if(is_null($this->workbook)) {
							$span->addEvent('workbook.workbook_xml_not_loaded');
							
							$this->workbook = new \DOMDocument();
							$workbookFile = $this->getXLSX()->getFromName('xl/workbook.xml');
							if ($workbookFile === false) {
								$span->addEvent('workbook.workbook_xml_not_found');
								throw new \Exception('Could not find xl/workbook.xml in "'.$this->targetFilename.'"');
							}
							$this->workbook->loadXML($workbookFile);
							
							$span->addEvent('workbook.workbook_xml_loaded', [
								'excel.xml_size' => strlen($workbookFile)
							]);
					} else {
							$span->addEvent('workbook.workbook_xml_already_available');
					}
					return $this->workbook;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		/**
		 * Return the DOMDocument representation of the styles.xml included in the current workbook
		 *
		 * @return DOMDocument
		 */
		public function getStyles() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.getStyles')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.get_styles_xml')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'getStyles')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					if(is_null($this->styles)) {
							$span->addEvent('workbook.styles_xml_not_loaded');
							
							$this->styles = new \DOMDocument();
							$stylesFile = $this->getXLSX()->getFromName('xl/styles.xml');
							if ($stylesFile === false) {
								$span->addEvent('workbook.styles_xml_not_found');
								throw new \Exception('Could not find xl/styles.xml');
							}
							$this->styles->loadXML($stylesFile);
							
							$span->addEvent('workbook.styles_xml_loaded', [
								'excel.xml_size' => strlen($stylesFile)
							]);
					} else {
							$span->addEvent('workbook.styles_xml_already_available');
					}
					return $this->styles;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

		/**
		 * Save modifications of workbook.xml
		 *
		 * return @this
		 */
		public function saveWorkbook() {
			$tracer = self::$instrumentation->tracer();
			$span = $tracer->spanBuilder('ExcelWorkbook.saveWorkbook')
					->setSpanKind(SpanKind::KIND_INTERNAL)
					->setAttribute('excel.operation', 'workbook.save_workbook_xml')
					->setAttribute(TraceAttributes::CODE_FUNCTION, 'saveWorkbook')
					->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
					->startSpan();

			$scope = $span->activate();
			
			try {
				$workbookXML = $this->getWorkbook()->saveXML();
				$this->getXLSX()->addFromString('xl/workbook.xml', $workbookXML);
				
				$span->addEvent('workbook.workbook_xml_saved', [
					'excel.xml_size' => strlen($workbookXML)
				]);
				
				if($this->isAutoSaveEnabled()) {
					$span->addEvent('workbook.auto_save_triggered');
					$this->save();
				}
				return $this;
			} catch (\Throwable $e) {
				$span->recordException($e);
				$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
				throw $e;
			} finally {
				$span->end();
				$scope->detach();
			}
		}

		/**
		 * Save modifications of styles.xml
		 *
		 * return @this
		 */
		public function saveStyles() {
				$tracer = self::$instrumentation->tracer();
				$span = $tracer->spanBuilder('ExcelWorkbook.saveStyles')
						->setSpanKind(SpanKind::KIND_INTERNAL)
						->setAttribute('excel.operation', 'workbook.save_styles_xml')
						->setAttribute(TraceAttributes::CODE_FUNCTION, 'saveStyles')
						->setAttribute(TraceAttributes::CODE_NAMESPACE, 'Svrnm\\ExcelDataTables\\ExcelWorkbook')
						->startSpan();

				$scope = $span->activate();
				
				try {
					$stylesXML = $this->getStyles()->saveXML();
					$this->getXLSX()->addFromString('xl/styles.xml', $stylesXML);
					
					$span->addEvent('workbook.styles_xml_saved', [
						'excel.xml_size' => strlen($stylesXML)
					]);
					
					if($this->isAutoSaveEnabled()) {
						$span->addEvent('workbook.auto_save_triggered');
						$this->save();
					}
					return $this;
				} catch (\Throwable $e) {
					$span->recordException($e);
					$span->setStatus(StatusCode::STATUS_ERROR, $e->getMessage());
					throw $e;
				} finally {
					$span->end();
					$scope->detach();
				}
		}

}

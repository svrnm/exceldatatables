<?php

namespace ELearningAG\ExcelDataTables;

/**
 * This class is a simple representation of a excel workbook. It expects
 * a xlsx formatted spreadsheet as paramater and overwrites(!) existing
 * worksheets with new data.
 *
 * @author Severin Neumann <s.neumann@elearning-ag.de>
 * @copyright 2014 die eLearning AG
 * @license GPL-3.0 
 */
class ExcelWorkbook implements \Countable
{

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
		 * @var ZipArchive
		 */
		protected $xlsx;

		/**
		 * The DomDocument representation of the workbook
		 *
		 * @var DOMDocument
		 */
		protected $workbook;

		/**
		 * The DOMDocument representation of the styles.xml included in the workbook
		 *
		 * @var DOMDocument
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
		 * Instantiate a new object of the type ExcelWorkbook. Expects a filename which
		 * contains a spreadsheet of type xlsx.
		 *
		 * @param string $filename
		 */
		public function __construct($filename) {
				$this->srcFilename = $filename;
				$this->targetFilename = $filename;
		}

		public function __destruct() {
				$this->getXLSX()->close();
		}

		/**
		 * Turn on auto saving
		 *
		 * @return $this.
		 */
		public function enableAutoSave() {
				$this->autoSave = true;
				return $this;
		}

		/**
		 * Turn off auto saving
		 *
		 * @return $this
		 */
		public function disableAutoSave() {
				$this->autoSave = false;
				return $this;
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
		public function count() {
				$sheets = $this->getWorkbook()->getElementsByTagName('sheet');
				return $sheets->length;

		}

		/**
		 * Change the filepath where the modified XLSX is stored. This should
		 * be called before a worksheet is added.
		 *
		 * @param string $filename
		 * @return $this
		 */
		public function setFilename($filename) {
				$this->targetFilename = $filename;
				return $this;
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
				$r = $this->getXLSX()->getFromName('xl/worksheets/sheet'.$id.'.xml');
				if($asDocument && $r !== false) {
						$dom = new \DOMDocument();
						$dom->loadXML($r);
						return $dom;
				} else {
						return $r;
				}
		}

		/**
		 * Check if a worksheet with name $name already exists. If not $name is
		 * returned. Otherwise $name_<i> is incremented until the name is unique.
		 *
		 * @param $name
		 * @return string
		 */
		protected function uniqName($name) {
				$sheets = $this->getWorkbook()->getElementsByTagName('sheet');
				$names = array();
				foreach($sheets as $sheet) {
						$names[] = $sheet->getAttribute('name');
				}
				$i = 0;
				$origName = $name;
				while(in_array($name, $names)) {
						$name = $origName.'_'.$i;
				}
				return $name;
		}

		/**
		 * Extract a date time format from the styles.xml. Some guessing is used to find
		 * a matching id. If no format is find, the id of the GENERAL format is returned
		 *
		 * @return int
		 */
		protected function getDateTimeFormatId() {
				$formats = $this->getStyles()->getElementsByTagName('numFmt');
				$exists = false;
				$generalId = 0;
				$foundId = 0;
				$maxScore = 3;
				for($id = 0; $id < $formats->length; ++$id) {
						$format = $formats->item($id);
						$code = strtoupper($format->getAttribute('formatCode'));
						if($code === 'GENERAL') {
								$generalID = $id;
						} elseif($code === "DD/MM/YYYY\ HH:MM:SS") {
								$foundId = $id;
								$exists = true;
						} else {
								// Do some "guessing" if the current format is "good enough"
								$score 	= (strpos($code, 'YY') !== false ? 1 : 0)
										+ (strpos($code, 'YYYY') !== false ? 1 : 0)
										+ (strpos($code, 'MM') !== false ? 1 : 0)
										+ (strpos($code, 'D') !== false ? 1 : 0)
										+ (strpos($code, 'DD') !== false ? 1 : 0)
										+ (strpos($code, 'HH') !== false ? 1 : 0)
										+ (strpos($code, 'HH:MM') !== false ? 1 : 0)
										+ (strpos($code, 'HH:MM:SS') !== false ? 1 : 0);
								if($score > $maxScore) {
										$maxScore = $score;
										$exists = true;
										$foundId = $id;
								}
						}
				}
				if($exists) {
						return $foundId;
				} else {
						return $generalId;
				}
		}

		/**
		 * Add a worksheet into the workbook with id $id and name $name. If $id is null the last
		 * worksheet is replaced. If $name is empty, its default value is 'Data'.
		 *
		 * Currently this replaces an existing worksheet. Adding new worksheets is not yet supported
		 *
		 * @param ExcelWorksheet $worksheet
		 * @param int $id
		 * @param string $name
		 */
		public function addWorksheet(ExcelWorksheet $worksheet, $id = null, $name = 'Data') {
				if(is_null($id) || $id <= 0) {
						$lastId = 0;
						while($this->getXLSX()->statName('xl/worksheets/sheet'.($lastId+1).'.xml') !== false) {
								$lastId++;
						}
						$id = $lastId + $id;
				}
				$old = $this->getXLSX()->getFromName('xl/worksheets/sheet'.$id.'.xml');
				if($old === false) {
						throw new \Exception('Appending new sheets is not yet implemented: ' . $id .', '. $this->srcFilename.', '.$this->targetFilename);
				} else {
						$document = new \DOMDocument();
						$document->loadXML($old);					
						$oldSheetData = $document->getElementsByTagName('sheetData')->item(0);
						$worksheet->setDateTimeFormatId($this->getDateTimeFormatId());
						$newSheetData = $document->importNode( $worksheet->getDocument()->getElementsByTagName('sheetData')->item(0), true );
						$oldSheetData->parentNode->replaceChild($newSheetData, $oldSheetData);
						$xml = $document->saveXML();
						$this->getXLSX()->addFromString('xl/worksheets/sheet'.$id.'.xml', $xml);
				}
				if($this->isAutoSaveEnabled()) {
						$this->save();
				}
				return $this;
		}

		/**
		 * Return the ZipArchive representation of the current workbook
		 *
		 * @return ZipArchive
		 */
		public function getXLSX() {
				if(is_null($this->xlsx)) {
						$this->openXLSX();
				}
				return $this->xlsx;
		}

		/**
		 * Open the excel file and create the ZipArchive representation. If the file
		 * does not exists or is not valid an exception is thrown.
		 *
		 * @throws Excepetion
		 * @return $this
		 */ 
		protected function openXLSX() {
				$this->xlsx = new \ZipArchive;
				if(!file_exists($this->srcFilename) && is_readable($this->srcFilename)) {
						throw new \Exception('File does not exists: '.$this->srcFilename);
				}
				if($this->srcFilename !== $this->targetFilename) {
						file_put_contents($this->targetFilename, file_get_contents($this->srcFilename));
						$this->srcFilename = $this->targetFilename;
				}
				$isOpen = $this->xlsx->open($this->targetFilename);
				if($isOpen !== true) {
						throw new \Exception('File not valid: '.$this->targetFilename);
				}
				return $this;
		}

		/**
		 * Save the modifications 
		 *
		 * @return $this
		 */
		public function save() {
				$this->getXLSX()->close();
				$this->srcFilename = $this->targetFilename;
				$this->openXLSX();
				return $this;
		}


		/**
		 * Return the DOMDocument representation of the current workbook
		 *
		 * @return DOMDocument
		 */
		public function getWorkbook() {
				if(is_null($this->workbook)) {
						$this->workbook = new \DOMDocument();
						$this->workbook->loadXML($this->getXLSX()->getFromName('xl/workbook.xml'));
				}
				return $this->workbook;
		}

		/**
		 * Return the DOMDocument representation of the styles.xml included in the current workbook
		 *
		 * @return DOMDocument
		 */
		public function getStyles() {
				if(is_null($this->styles)) {
						$this->styles = new \DOMDocument();
						$this->styles->loadXML($this->getXLSX()->getFromName('xl/styles.xml'));
				}
				return $this->styles;
		}

		/**
		 * Save modifications of styles.xml
		 *
		 * return @this
		 */
		public function saveStyles() {
				$this->getXLSX()->addFromString('xl/styles.xml', $this->getStyles()->saveXML());
				if($this->isAutoSaveEnabled()) {
						$this->save();
				}
				return $this;
		}

}

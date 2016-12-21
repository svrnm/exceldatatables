<?php

namespace spec\Svrnm\ExcelDataTables;

use PhpSpec\ObjectBehavior;
use Prophecy\Argument;

/**
 * Specification for the class ExcelWorksheet.
 *
 * @author Severin Neumann <severin.neumann@altmuehlnet.de>
 * @license Apache-2.0
 */
class ExcelWorksheetSpec extends ObjectBehavior
{

		static function wrapWorksheet($inner) {
			$xml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>'.PHP_EOL;
			$xml .= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

			$xml .= $inner;

			$xml .= '</worksheet>'.PHP_EOL;

			return $xml;
		}

		static function wrapSheetData($inner) {
			return self::wrapWorksheet('<sheetData>'.$inner.'</sheetData>');
		}

		static function wrapRow($inner, $id) {
			return '<row r="'.$id.'">'.$inner.'</row>';
		}

		function it_is_initializable()
		{
				$this->shouldHaveType('Svrnm\ExcelDataTables\ExcelWorksheet');
		}

		function it_converts_to_xml()
		{
				$this->setupDefaultDocument()->toXML()->shouldReturn(self::wrapWorksheet('<sheetData/>'));
		}

		function it_provides_a_dom_document()
		{
				$this->getDocument()->shouldHaveType('\DOMDocument');
		}

		function it_provides_a_worksheet_root_element()
		{
				$this->getWorksheet()->shouldHaveType('\DOMElement');
		}

		function it_provides_a_sheetdata_element()
		{
				$this->getSheetData()->shouldHaveType('\DOMElement');
		}

		function it_adds_new_rows()
		{
				$this->addRow()->addRow()->toXML()->shouldReturn(self::wrapSheetData('<row r="1"/><row r="2"/>'));;
		}

		function it_adds_strings_to_string_columns()
		{
				$this->addRow(array('a', 'b', 'c'))->toXML()->shouldReturn(self::wrapSheetData(self::wrapRow('<c t="inlineStr"><is><t>a</t></is></c><c t="inlineStr"><is><t>b</t></is></c><c t="inlineStr"><is><t>c</t></is></c>', 1)));
		}

		function it_adds_numbers_to_number_columns()
		{
			$this->addRow(array(1,2,3))->toXML()->shouldReturn(self::wrapSheetData(self::wrapRow('<c><v>1</v></c><c><v>2</v></c><c><v>3</v></c>', 1)));
		}

		function it_adds_numbers_to_string_typed_columns()
		{
			$this->addRow(array(array('type' => 'string', 'value' => 13)))->toXML()->shouldReturn(self::wrapSheetData(self::wrapRow('<c t="inlineStr"><is><t>13</t></is></c>', 1)));
		}

		function it_adds_datetimes_todatetime_columns()
		{
			$this->addRow(array(new \DateTime('2013-04-05')))->toXML()->shouldReturn(self::wrapSheetData(self::wrapRow('<c s="1"><v>41369</v></c>', 1)));
		}

		function it_accepts_multidimensional_arrays_as_rows()
		{
			$this->addRows(array(
				array(1,2,3),
				array(2,3,4),
				array(5,6,7)
			))->toXML()->shouldReturn(self::wrapSheetData('<row r="1"><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row><row r="2"><c><v>2</v></c><c><v>3</v></c><c><v>4</v></c></row><row r="3"><c><v>5</v></c><c><v>6</v></c><c><v>7</v></c></row>'));
		}
}

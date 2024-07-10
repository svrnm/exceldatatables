<?php

namespace spec\Svrnm\ExcelDataTables;

use PhpSpec\ObjectBehavior;
use Prophecy\Argument;

/**
 * Specification for the class ExcelDataTables
 *
 * @author Severin Neumann <severin.neumann@altmuehlnet.de>
 * @license Apache-2.0
 */
class ExcelDataTableSpec extends ObjectBehavior
{
	/**
	 * Default specification: Check if the class can be instantiated.
	 */
	function it_is_initializable()
	{
		$this->shouldHaveType('Svrnm\ExcelDataTables\ExcelDataTable');
	}

	/**
	 * The data table can have a special header row which is set using
	 * setHeaders.
	 */
	function it_has_headers()
	{
		$headers = array("A" => "A", "B" => "B", "C" => "C");
		$this->setHeaders($headers)->getHeaders()->shouldReturn($headers);
	}

	/**
	 * A simple numerical array is converted in a single row.
	 * Tow arrays are converted into two rows.
	 * ...
	 */
	function it_converts_an_numerical_array_to_a_row()
	{
		$array = array(1, 2, 3);
		$this->addRow($array)->toArray()->shouldReturn(array($array));
		$array2 = array(3, 4, 5);
		$this->addRow($array2)->toArray()->shouldReturn(array($array, $array2));
	}

	/**
	 * An assocative array is converted into a single row. The keys of the array
	 * are internaly used as column identifier.
	 *
	 * Adding a second assocative array which does not have entires in all columns
	 * will introduce empty cells.
	 */
	function it_converts_an_assocative_to_a_row()
	{
		$array = array("A" => 1, "B" => 2, "C" => 3);
		$this->addRow($array)->toArray()->shouldReturn(array(array(1, 2, 3)));

		$array2 = array("C" => 3);
		$this->addRow($array2)->toArray()->shouldReturn(array(array(1, 2, 3), array('', '', 3)));
	}

	/**
	 * A simple object is treated like an assocative array
	 */
	function it_converts_an_object_to_row()
	{
		$object = new \stdClass();
		$object->A = 1;
		$object->B = 2;
		$object->C = 3;
		$this->addRow($object)->toCsv()->shouldReturn('1,2,3');
	}

	function it_converts_multidimensional_arrays_to_multiple_rows()
	{
		$array = array(
			array(1, 2, 3),
			array(1, 2, 3)
		);
		$this->addRows($array)->toArray()->shouldReturn($array);
	}

	function it_has_a_fluent_interface()
	{
		$array = array(1);
		$this->addRow($array)->shouldReturn($this);
		$this->setHeaders($array)->shouldReturn($this);
		/* ... */
	}

	function it_has_implicit_headers()
	{
		$array = array("A" => 1, "B" => 2, "C" => 3);
		$array2 = array("C" => 3);
		$this->addRow($array)->addRow($array2)->showHeaders()->toArray()->shouldReturn(array(array("A", "B", "C"), array(1, 2, 3), array("", "", 3)));
	}

	function it_converts_data_to_xml()
	{
		$array = array(
			array("Names" => "Test 1", "Value 1" => 13, "Value 2" => new \DateTime('2013-03-04 13:00:00')),
			array("Names" => "Test 2", "Value 1" => 23, "Value 2" => new \DateTime('2014-04-01 13:12:00')),
			array("Names" => "Test 3", "Value 1" => 33, "Value 2" => new \DateTime('1900-01-01 00:00:00')),
		);

		$this->addRows($array)->showHeaders()->toXML()->shouldReturn('<?xml version="1.0" encoding="utf-8" standalone="yes"?' . '>' . PHP_EOL . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c t="inlineStr"><is><t>Names</t></is></c><c t="inlineStr"><is><t>Value 1</t></is></c><c t="inlineStr"><is><t>Value 2</t></is></c></row><row r="2"><c t="inlineStr"><is><t>Test 1</t></is></c><c><v>13</v></c><c s="1"><v>41337.541666667</v></c></row><row r="3"><c t="inlineStr"><is><t>Test 2</t></is></c><c><v>23</v></c><c s="1"><v>41730.55</v></c></row><row r="4"><c t="inlineStr"><is><t>Test 3</t></is></c><c><v>33</v></c><c s="1"><v>1</v></c></row></sheetData></worksheet>' . PHP_EOL);
	}
}

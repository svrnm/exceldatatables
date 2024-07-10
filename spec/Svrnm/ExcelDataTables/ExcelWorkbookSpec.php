<?php

namespace spec\Svrnm\ExcelDataTables;

use PhpSpec\ObjectBehavior;
use Prophecy\Argument;
use Svrnm\ExcelDataTables\ExcelWorksheet;
use org\bovigo\vfs\vfsStream;

/**
 * Specification for the class ExcelWorkbook
 *
 * @author Severin Neumann <severin.neumann@altmuehlnet.de>
 * @license Apache-2.0
 */

class ExcelWorkbookSpec extends ObjectBehavior
{
	protected $testFilename;

	function let()
	{
		/* vfsStream and ZipArchive are not working together... */
		$this->testFilename = sys_get_temp_dir() . '/exceldatatables-test-spec.xlsx';
		copy('./examples/spec.xlsx', $this->testFilename);
		$this->beConstructedWith($this->testFilename);
	}

	function letGo()
	{
		#		unlink($this->testFilename);
	}

	function it_is_initializable()
	{
		$this->shouldHaveType('Svrnm\ExcelDataTables\ExcelWorkbook');
	}

	function it_provides_a_workbook_xml()
	{
		$this->getWorkbook()->shouldHaveType('\DOMDocument');
	}

	function it_is_a_countable()
	{
		$this->count()->shouldReturn(2);
	}

	function it_has_index_worksheets()
	{
		$this->getWorksheetById(1, true)->shouldHaveType('\DOMDocument');
		$this->getWorksheetById(3, true)->shouldReturn(false);
	}

	function it_provides_a_styles_xml()
	{
		$this->getStyles()->shouldHaveType('\DOMDocument');
	}

	function it_modifies_an_existing_excel_workbook(\ZipArchive $test)
	{
		$worksheet = new ExcelWorksheet();

		$worksheet->addRows(
			array(
				array("Date", "Value 1", "Value 2"),
				array(new \DateTime("2014-07-30"), 13, 4),
				array(new \DateTime("2014-07-31"), 18, 5),
				array(new \DateTime("2014-08-01"), 14, 6),
				array(new \DateTime("2014-08-02"), 9, 7),
				array(new \DateTime("2014-08-03"), 3, 4),
				array(new \DateTime("2014-08-04"), 1, 3),
				array(new \DateTime("2014-08-05"), 4, 2),
				array(new \DateTime("2014-08-09"), 4, 0),
				array(new \DateTime("2014-08-10"), 13, 1),
				array(new \DateTime("2014-08-11"), 23, 3),
				array(new \DateTime("2014-08-12"), 18, 23),
				array(new \DateTime("2014-08-13"), 19, 0),
				array(new \DateTime("2014-08-14"), 21, 13),
			)
		);
		$this->addWorksheet($worksheet, 2)->save();
		/* TODO: Validate?! */

	}

}

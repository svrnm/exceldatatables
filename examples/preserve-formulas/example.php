<?php
	require_once('../../vendor/autoload.php');
	$dataTable = new Svrnm\ExcelDataTables\ExcelDataTable();
	$in = 'spec.xlsx';
	$out = 'test.xlsx';
	$data = array(
			array("Date" => new \DateTime('2014-01-01 13:00:00'), "Value 1" => 0, "Value 2" => 1),
			array("Date" => new \DateTime('2014-01-02 14:00:00'), "Value 1" => 1, "Value 2" => 0),
			array("Date" => new \DateTime('2014-01-03 15:00:00'), "Value 1" => 2, "Value 2" => -1),
			array("Date" => new \DateTime('2014-01-04 16:00:00'), "Value 1" => 3, "Value 2" => -2),
			array("Date" => new \DateTime('2014-01-05 17:00:00'), "Value 1" => 4, "Value 2" => -3),
			array("Date" => new \DateTime('2014-01-03 15:00:00'), "Value 1" => 30, "Value 2" => -1),
			array("Date" => new \DateTime('2014-01-04 16:00:00'), "Value 1" => 3, "Value 2" => -2),
			array("Date" => new \DateTime('2014-01-05 17:00:00'), "Value 1" => 4, "Value 2" => -3),
	);
	$dataTable->showHeaders()->preserveFormulas('Data')->addRows($data)->refreshTableRange('Data')->attachToFile($in, $out, false);
?>

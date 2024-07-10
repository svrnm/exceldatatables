<?php
	ini_set('memory_limit', '2048M');
	require_once('../vendor/autoload.php');
	$in = 'spec.xlsx';
	$out = 'test.xlsx';
	$lastPeak = 0;

	$rows = 64;
	$cols = 20;

	while($lastPeak < 1024*1024*300) {


	$data = array();
	for($i = 0; $i < $rows; $i++) {
		$row = array();
		for($j = 0; $j < $cols; $j++) {
			switch($j%3) {
				case 0:
						$row[] = $i;
						break;
				case 1:
						$row[] = '('.$i.','.$j.')' ;
						break;
				case 2:
						$row[] = new DateTime('2014-01-01');
						break;
			}
		}
		$data[] = $row;
	}

	$start = microtime(true);
	$dataTable = new Svrnm\ExcelDataTables\ExcelDataTable();
	$dataTable->showHeaders();
	$dataTable->addRows($data);
	$time0 = microtime(true) - $start;
	$dataTable->attachToFile($in, $out);
	$time1 = microtime(true)-$start;
	$lastPeak = memory_get_peak_usage();
	echo $rows.' x '.$cols.":\t";
	echo ($time0)." s\t";
	echo ($time1)." s\t";
	echo floor((($rows)/$time1))." rows/s\t";
	echo floor((($rows*$cols)/$time1))." entries/s\t";
	echo ($lastPeak/(1024*1024)).' MB'.PHP_EOL;
	$rows*=2;
	}

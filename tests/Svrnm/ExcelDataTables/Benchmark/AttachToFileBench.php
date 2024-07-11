<?php
namespace Svrnm\ExcelDataTables\Benchmark;

use Svrnm\ExcelDataTables\ExcelDataTable;

class AttachToFileBench
{
	 /**
     * @Revs(10)
     * @Iterations(2)
	 * @ParamProviders("provideAttachToFile")
     */
	public function benchAttachToFile(array $params)
    {
    $dataTable = new ExcelDataTable();
    $in = __DIR__ . '/../../../../examples/spec.xlsx';

    // output file need to be recreated => delete if exists
    $out = __DIR__ . '/../../../../examples/test.xlsx';
    if( file_exists($out) ) {
    	if( !@unlink ( $out ) )
    	{
    		echo "CRITIC! - destination file: $out - has to be deleted, and I can't<br>";
    	    echo "CRITIC! - check directory and file permissions<br>";
    		die();	
    	} 
    }	
	$dataTable->showHeaders()->addRows($params["data"])->attachToFile($in, $out, false);
    }

	private static function generate($rows, $cols) {
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
							$row[] = new \DateTime('2024-01-01');
							break;
				}
			}
			$data[] = $row;
		}
		return $data;
	}

	public function provideAttachToFile() {
		yield '100x100' => ['data' => self::generate(100, 100)];
		yield '200x200' => ['data' => self::generate(200, 200)];
		yield '400x400' => ['data' => self::generate(400, 400)];
	}
}
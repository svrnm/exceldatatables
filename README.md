# ExcelDataTables

Replace a worksheet within an Excel workbook (`.xlsx`) without changing any
other properties of the file.

## License

This program is free software; see [LICENSE](./LICENSE) for more details.

## Details

The main purpose of this library is adding a "data table" to an existing excel
file without modifying any other components. This is especially useful if you
have a template file, e.g. for reporting including "advanced" features like
charts, pivot tables, macros and you'd like to change a base data table within a
PHP application.

## Setup

Use composer to add this repository to your dependencies:

```JSON
{
  "require": {
    "svrnm/exceldatatables": "dev-master"
  }
}
```

If you use the [Laravel framework](https://laravel.io/) you can additionally add
the `ServiceProvider` to your `config/app.php`:

```PHP
...
'providers' => array(
  ...
  'Svrnm\ExcelDataTables\ExcelDataTablesServiceProvider'
)
...
```

## Example

The following example demonstrates how to use ExcelDataTables. The following is
also contained as `example.php` within the folder `examples/`:

```php
<?php
  require_once('../vendor/autoload.php');
  // Create a new instance
  $dataTable = new Svrnm\ExcelDataTables\ExcelDataTable();
  // Specify the source file
  $in = 'spec.xlsx';
  // Specify the output file
  $out = 'test.xlsx';
  // Specify the data for the worksheet.
  $data = array(
    array("Date" => new \DateTime('2014-01-01 13:00:00'), "Value 1" => 0, "Value 2" => 1),
    array("Date" => new \DateTime('2014-01-02 14:00:00'), "Value 1" => 1, "Value 2" => 0),
    array("Date" => new \DateTime('2014-01-03 15:00:00'), "Value 1" => 2, "Value 2" => -1),
    array("Date" => new \DateTime('2014-01-04 16:00:00'), "Value 1" => 3, "Value 2" => -2),
    array("Date" => new \DateTime('2014-01-05 17:00:00'), "Value 1" => 4, "Value 2" => -3),
  );
  // Attach the data table and copy the new xlsx file to the output file.
  $dataTable->showHeaders()->addRows($data)->attachToFile($in, $out);
?>
```

In this example the method `attachToFile` creates a new excel file. If you use
this library within a web application you might prefer the `fillXLSX()` function
which returns a string representation of the excel document. The following
examples demonstrates this case within a Laravel application

```php
<?php
  class ReportController extends Controller {
    protected $dataTable;

     public function __construct(\Svrnm\ExcelDataTables\ExcelDataTable $dataTable) {
         $this->dataTable = $dataTable;
  }

  public function show($month) {
   $data = DB::select('select date,value1,value2 from reporting where MONTH(date) = ?', array($month));
   $path = storage_path() . '/reports/example.xlsx';
   $xlsx = $this->dataTable->showHeaders()->addRows($data)->fillXLSX($path);
   return Response::make($xlsx, 200, array(
    'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition' => 'attachment; filename="report.xlsx"',
    'Content-Length' => strlen($xlsx)
   ));
   // ...
  }
 }
?>
```

## Contact

For any questions you can contact Severin Neumann
<severin.neumann@altmuehlnet.de>.

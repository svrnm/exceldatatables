<?php namespace Svrnm\ExcelDataTables;

use Illuminate\Support\ServiceProvider;

/**
 * Laravel ExcelDataTables ServiceProvider
 *
 * @author Severin Neumann <severin.neumann@altmuehlnet.de>
 * @license Apache-2.0
 */
class ExcelDataTablesServiceProvider extends ServiceProvider {


		/**
		 * Register the service provider.
		 *
		 * @return void
		 */
		public function register()
		{
				$this->app->singleton('exceldatatables', function () {
						return new ExcelDataTable();
				});
		}

		/**
		 * Get the services provided by the provider.
		 *
		 * @return array
		 */
		public function provides()
		{
				return array(
						'exceldatatables'
				);
		}
}

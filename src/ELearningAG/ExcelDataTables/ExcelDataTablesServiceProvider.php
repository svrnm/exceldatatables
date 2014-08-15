<?php namespace ELearningAG\ExcelDataTables;

use Illuminate\Support\ServiceProvider;

/**
 * Laravel ExcelDataTables ServiceProvider
 *
 * @author Severin Neumann <s.neumann@elearning-ag.de>
 * @license GPL-3.0
 * @copyright 2014 die eLearning AG
 */
class ExcelDataTablesServiceProvider extends ServiceProvider {
	

	/**
	 * Register the service provider.
	 *
	 * @return void
	 */
	public function register()
	{
		$this->app['exceldatatables'] = $this->app->share(function($app) {
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

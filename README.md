# Yii 2 Excel Data Exchanger

After installing this package you will be able to export data from DB to an excel2007 file and upload the data back again. 
The package is a module that you need to register in the config file in the module section. 
The module has a widget that you can use with different AR models in the usual way (for widgets). 
In the widget's config you need to specify the model name you want to import/export. Other optional properties are listed below.

If the exporting model has a related table the file will have an additional sheet for every table even the ones which don’t have a model. 
The data in these additional sheets helps to handle related columns. This sheets will be ignored in importing.

Required fields are highlighted in the excel file.

![Required fields](https://github.com/uranum/yii2-excel-exchange/blob/master/docs/img/xls_file.png)

There is a feature to backup the table before importing.

Export

![Export](https://github.com/uranum/yii2-excel-exchange/blob/master/docs/img/export.png)

Import

![Import](https://github.com/uranum/yii2-excel-exchange/blob/master/docs/img/import.png)


## Installation

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist uranum/yii2-excel-exchange
```

or add

```
"uranum/yii2-excel-exchange": "*"
```

to the require section of your `composer.json` file.

###In config, modules section
```php
	'modules' => [
//      ...
		'excel' => [
			'class'  => 'uranum\excel\Module',
			'params' => [
				'uploadPath' => 'uploads', //module's properties
				'fileName'   => 'export',
				'extensions' => 'xls, xlsx',
			],
		],
	],
```

How to use
----------

###In view (for example, index.php)
```php
use uranum\excel\ExcelExchanger;
use uranum\excel\Asset;

//...

echo ExcelExchanger::widget([
	'mainModelName' => $searchModel::className(), // here place model class name
 ]);
```


Widget's properties
-------------------

string `$mainModelName` - the name of the model for export/import. It's mandatory property.

string `$fileNameFrom` - path to a file for export/import. Default is `'uploads/export.xlsx'`. The folder must be exist!

string `$notNullColumnColor` - a value of the color for required columns. Default is `'FFDECC'`

string `$nameOfReserveTable` - prefix for the copied table. Default is `'archive_'`

string `$backupUrl` - url for backup action. Default is `'excel/default/backup'`. You can create your own controller with it's actions and change the urls here.

string `$uploadUrl` - url for upload action. Default is `'excel/default/upload'`

string `$importUrl` - url for import action. Default is `'excel/default/import'`

string `$exportUrl` - url for export action. Default is `'excel/default/export'`

integer `$columnWidthOfStringType` - a value of the column's width that has string type. Default is `35`

integer `$columnWidthOfTextType` - a value of the column's width that has text type. Default is `50`

integer `$columnWidthDefault` - a value of the other column's width. Default is `15`

string `$modalId` - a value of the Modal id. Default is `'excelImportModal'`

string `$fileInputId` - a value of the fileInput field id. Default is `'uploadFile'`

string `$exportBtnClass` is the part of css bootstrap class. Default is `'ExcelExchanger::COLOR_SUCCESS'`

string `$uploadBtnClass` is the part of css bootstrap class. Default is `'ExcelExchanger::COLOR_SUCCESS'`

string `$uploadDataBtnClass` is the part of css bootstrap class. Default is `'ExcelExchanger::COLOR_SUCCESS'`

string `$importPopupBtnClass` is the part of css bootstrap class. Default is `'ExcelExchanger::COLOR_DANGER'`

string `$panelColorClass` is the part of css bootstrap class. Default is `'ExcelExchanger::COLOR_DEFAULT'`

string `$fileInputColorClass` is the part of css bootstrap class. Default is `'ExcelExchanger::COLOR_WARNING'`

# Yii 2 Excel Data Exchanger

После установки пакета Вы получите функционал выгрузки данных из базы в файл Excel2007 и загрузки измененных в файле данных обратно в базу.
Пакет работает как виджет, который встраивается в представление. В настройках виджета указывается, данные какой модели нужно выгружать/загружать.
При выгрузке данных модели подхватываются связанные таблицы (не модели). Для связанных таблиц создаются листы с данными из этих таблиц.
При импорте данные связанных таблиц не учитываются, они нужны только для ознакомления со связанными полями.

В файле, поля обязательные для заполнения выделены цветом (настраивается).
![ss](https://github.com/uranum/yii2-excel-exchange/blob/master/docs/img/xls_file.png)

Есть функция резервного копирования таблицы, в которую импортируются данные.

Выгрузка

![ss](https://github.com/uranum/yii2-excel-exchange/blob/master/docs/img/export.png)

Загрузка

![ss](https://github.com/uranum/yii2-excel-exchange/blob/master/docs/img/import.png)


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

## How to use

In view (for example, index.php)
```php
use uranum\src\ExcelExchanger;
use uranum\src\Asset;

//...

echo ExcelExchanger::widget([
	'mainModelName' => $searchModel::className(), // here place model class name
 ]);
```

In controller
```php
use uranum\src\ExcelExchanger;
use uranum\src\ImportXls;

//...

public function actionBackup()
{
	if (Yii::$app->request->isAjax && Yii::$app->request->isPost) {
		$data = new ExcelExchanger(['mainModelName' => YourModel::className()]);
		$data->copyTable();
		echo 'Данные успешно сохранены!'; // Success backup message
	} else {
		echo 'Копирование завершилось неудачей. Попробуйте позже.'; // Fail backup message
	}
}

public function actionExport()
{
	$data = new ExcelExchanger(['mainModelName' => YourModel::className()]);
	$data->export();
}

public function actionImport()
{
	if (Yii::$app->request->isPost && Yii::$app->request->isAjax) {
		$data = new ExcelExchanger(['mainModelName' => YourModel::className()]);
		echo $data->import();
	}
}

public function actionUpload()
{
	if (Yii::$app->request->isPost && Yii::$app->request->isAjax) {
		$model = new ImportXls();
		$model->file = UploadedFile::getInstance($model, 'file');
		if ($model->upload()) {
			echo 'Файл успешно загружен!'; // Success upload message
		} else {
			echo 'Файл имеет некорректный тип. Правильный файл - Excel2007'; // Fail upload message - wrong file-format
		}
	}
}
```

## Widget's properties

string `$mainModelName` - the name of the model for export/import
string `$fileNameFrom` - path to a file for export/import. Default is `'uploads' . DIRECTORY_SEPARATOR . 'export.xlsx'`.
string `$notNullColumnColor` - a value of the color for required columns. Default is `'FFDECC'`
string `$nameOfReserveTable` - prefix for the copied table. Default is `'archive_'`
string `$backupUrl` - url for backup action. Default is `'backup'`
string `$uploadUrl` - url for upload action. Default is `'upload'`
string `$importUrl` - url for import action. Default is `'import'`
integer `$columnWidthOfStringType` - a value of the column's width that has string type. Default is `35`
integer `$columnWidthOfTextType` - a value of the column's width that has text type. Default is `50`
integer `$columnWidthDefault` - a value of the other column's width. Default is `15`

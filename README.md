# Yii 2 Excel Data Exchanger

После установки пакета Вы получите функционал выгрузки данных из базы в файл Excel2007 и загрузки измененных в файле данных обратно в базу.
Пакет работает как виджет, который встраивается в представление. В настройках виджета указывается, данные какой модели нужно выгружать/загружать.
При выгрузке данных модели подхватываются связанные таблицы (не модели). Для связанных таблиц создаются листы с данными из этих таблиц.
При импорте данные связанных таблиц не учитываются, они нужны только для ознакомления со связанными полями.

В файле, поля обязательные для заполнения выделены цветом (настраивается).
![ss](https://yadi.sk/i/jwsC4lbXvPMBM)

Есть функция резервного копирования таблицы, в которую импортируются данные.

Выгрузка

![ss](https://yadi.sk/i/h4laiZmbvPHit)

Загрузка

![ss](https://yadi.sk/i/727R0e7pvPHwo)


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
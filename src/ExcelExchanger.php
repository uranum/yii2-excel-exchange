<?php

namespace uranum\excel;


use PHPExcel_IOFactory;
use Yii;
use yii\base\InvalidValueException;
use yii\base\Model;
use yii\base\Widget;
use yii\bootstrap\ActiveForm;
use yii\bootstrap\Html;
use yii\bootstrap\Modal;
use yii\db\ActiveRecord;
use yii\helpers\Url;

class ExcelExchanger extends Widget
{
	public $mainModelName;
	public $fileNameFrom;
	public $notNullColumnColor = 'FFDECC';
	public $columnWidthOfStringType = 35;
	public $columnWidthOfTextType = 50;
	public $columnWidthDefault = 15;
	public $nameOfReserveTable = 'archive_';
	public $backupUrl = 'backup';
	public $uploadUrl = 'upload';
	public $importUrl = 'import';

	/**
	 * @var \PHPExcel
	 */
	private $xl;
	private $scheme;
	private $columnsNameArray = [];

	public function init()
	{
		parent::init();

		if(!$this->mainModelName){
			throw new InvalidValueException('Model name must be defined!');
		}
		/** @var ActiveRecord $mainModelName */
		$mainModelName = $this->mainModelName;
		$this->scheme  = $mainModelName::getTableSchema();
		$this->xl      = new \PHPExcel();
		$this->fileNameFrom = 'uploads' . DIRECTORY_SEPARATOR . 'export.xlsx';
		Asset::register($this->getView());
	}

	public function run()
	{
		echo "&emsp;";
		echo Html::a('Export', Url::to(['export']), ['class' => 'btn btn-success']);
		echo "&emsp;";
		$this->renderModal();
	}

	/**
	 * Модальное окно для действий импорта
	 * Modal for import actions
	 */
	private function renderModal()
	{
		$model = new ImportXls();
		Modal::begin([
			'header'       => Html::tag('h3', 'Настройки импорта'),
			'toggleButton' => [
				'tab'   => 'a',
				'label' => 'Загрузить данные в базу',
				'class' => 'btn btn-danger',
			],
		]);
		echo Html::tag('h4', 'Порядок действий');
		echo Html::beginTag('ol');
		echo Html::tag('li', 'Сделать резервную копию данных (можно пропустить)');
		echo Html::tag('li', 'Выбрать файл данных и загрузить его на сервер');
		echo Html::tag('li', 'Сделать импорт данных из файла в БД');
		echo Html::endTag('ol');

		echo Html::beginTag('div', ['class' => 'panel panel-default']);
		echo Html::tag('div', 'Шаг 1', ['class' => 'panel-heading text-center']);
		echo Html::beginTag('p', ['class' => 'panel-body text-center', 'id' => 'step_1']);
		echo Html::a('Запустить резервное копирование', '#', ['class' => 'btn btn-success', 'id' => 'startCopying', 'data' => ['url' => Url::to([$this->backupUrl])]]);
		echo Html::endTag('p');
		echo Html::endTag('div');

		echo Html::beginTag('div', ['class' => 'panel panel-default']);
		echo Html::tag('div', 'Шаг 2', ['class' => 'panel-heading text-center']);
		echo Html::beginTag('div', ['class' => 'panel-body text-center', 'id' => 'step_2']);
		$form = ActiveForm::begin([
			'enableClientValidation' => false,
			'action'                 => Url::to([$this->uploadUrl]),
			'options'                => [
				'enctype' => 'multipart/form-data',

			],
		]);
		echo $form->field($model, 'file', [
			'options' => [
				'class' => 'text-center',
				'style' => 'display:inline-block;margin-right:20px;',
			],
		])->fileInput([
			'class' => 'btn btn-warning',
			'id'    => 'uploadFile',
		])->label(false);
		echo Html::submitButton('Загрузить файл', [
			'id'    => 'uploadFileForm',
			'class' => 'btn btn-success',
			'data'  => [
				'url' => Url::to([$this->uploadUrl]),
			],
		]);
		ActiveForm::end();
		echo Html::endTag('div');
		echo Html::endTag('div');

		echo Html::beginTag('div', ['class' => 'panel panel-default']);
		echo Html::tag('div', 'Шаг 3', ['class' => 'panel-heading text-center']);
		echo Html::beginTag('p', ['class' => 'panel-body text-center', 'id' => 'step_3']);
		echo Html::a('Загрузить данные в базу', '#', ['class' => 'btn btn-success', 'id' => 'startImport', 'data' => ['url' => Url::to([$this->importUrl])]]);
		echo Html::endTag('p');
		echo Html::endTag('div');

		Modal::end();
	}

	/**
	 * Export
	 * Exporting data from DB to .xlsx and ordering to save it locally
	 * Выгружает данные из базы в .xlsx файл и предлагает сохранить его локально
	 * @throws \PHPExcel_Exception
	 * @throws \PHPExcel_Writer_Exception
	 * @throws \yii\base\InvalidConfigException
	 */
	public function export()
	{
		$this->prepareActiveSheetCommonStyles();
		$mainModels = $this->getAllModelsOfMainTable();
		$this->populateActiveSheet($mainModels);
		if ($fk = $this->findForeignKeys()) {
			$this->createRelatedSheets($fk);
		}
		$this->saveExportedFile();
	}

	/**
	 * Export
	 * Creates a sheet for each related table and populates it with a data directly from DB (without AR)
	 * Создает листы для каждой связанной таблицы и заполняет данными напрямую из базы (не через Модель)
	 * @param $fk array of Foreign Key names of the related tables
	 */
	private function createRelatedSheets($fk)
	{
		foreach ($fk as $tableName) {
			$newSheet = new \PHPExcel_Worksheet($this->xl, $tableName);
			$this->xl->addSheet($newSheet);
			$this->xl->setActiveSheetIndexByName($tableName);
			$data = Yii::$app->db->createCommand("SELECT * FROM {$tableName} ORDER BY id")->queryAll();
			foreach ($data as $rowNumber => $rowData) {
				$i = 0;
				foreach ($rowData as $cell) {
					$this->xl->getActiveSheet()->setCellValueByColumnAndRow($i, $rowNumber+1, $cell);
					$i++;
				}
			}
		}
		$this->xl->setActiveSheetIndex(0);
	}

	/**
	 * Export
	 * Returns an array of the names of the related tables or false if ones absent
	 * Возвращает массив с названиями таблиц, на которые ссылается основная модель
	 * или false, если у модели нет внешних ключей
	 * @return array|bool
	 */
	private function findForeignKeys()
	{
		$fk = $this->scheme->foreignKeys;
		if (!empty($fk)) {
			foreach ($fk as $key) {
				$fkTableNames[] = $key[0];
			}

			return $fkTableNames;
		} else {
			return false;
		}
	}

	/**
	 * Export
	 * Prepare active sheet - general design
	 * Подготовка активного листа - основное оформление:
	 *  название листа,
	 *  закрепление панели,
	 *  высота строк заголовков,
	 *  шрифт, выравнивание, границы ячеек
	 * @throws \PHPExcel_Exception
	 * @internal param $columnsNames
	 * @internal param $i
	 */
	private function prepareActiveSheetCommonStyles()
	{
		$columnsNames = $this->getXlsColumnNames();
		$this->xl->getActiveSheet()->setTitle($this->scheme->name);
		$this->xl->getActiveSheet()->freezePane('A2');
		$this->xl->getActiveSheet()->getRowDimension(1)->setRowHeight(40);
		$this->xl->getActiveSheet()->getStyle('A1:' . $columnsNames[count($this->scheme->columns) - 1] . '1')
			->applyFromArray([
				'font'      => [
					'bold' => true,
					'size' => '14',
				],
				'alignment' => [
					'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
					'vertical'   => \PHPExcel_Style_Alignment::VERTICAL_CENTER,
				],
				'borders'   => [
					'bottom' => [
						'style' => \PHPExcel_Style_Border::BORDER_THICK,
					],
				],
			]);
		$this->prepareActiveSheetColumnStyles($columnsNames);
	}

	/**
	 * Export
	 * Returns the names of columns of .xls table
	 * Названия колонок в xls таблице
	 * @return array
	 */
	private function getXlsColumnNames()
	{
		return [
			'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
			'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
			'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK',
			'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV',
			'AW', 'AX', 'AY', 'AZ',
		];
	}

	/**
	 * Export
	 * Decorating of the columns
	 * Оформление колонок:
	 *  выставить ширину колонок в зависиомсти от типа данных
	 *  сделать заголовки (названия) для колонок (полей)
	 *  окрашивание столбцов с обязательными полями кроме id, created_at, updated_at
	 * @param $columnsNames
	 */
	private function prepareActiveSheetColumnStyles($columnsNames)
	{
		$i = 0;
		foreach ($this->scheme->columns as $col) {
			$this->columnsNameArray[] = $col->name;
			$cellName                 = !empty($col->comment) ? $col->comment . ' (' . $col->name . ')' : $col->name;
			switch ($col->type) {
				case 'string':
					$width = $this->columnWidthOfStringType;
					break;
				case 'text':
					$width = $this->columnWidthOfTextType;
					break;
				default:
					$width = $this->columnWidthDefault;
			}
			$this->xl->getActiveSheet()->getColumnDimensionByColumn($i)->setWidth($width);
			$this->xl->getActiveSheet()->setCellValueByColumnAndRow($i, 1, $cellName);
			$i++;

			if (!$col->allowNull && $col->name != 'id' && $col->name != 'created_at' && $col->name != 'updated_at') {
				$this->xl->getActiveSheet()->getStyle($columnsNames[$i - 1] . '1')->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)
					->getStartColor()->setARGB($this->notNullColumnColor);
			}
		}
	}

	/**
	 * Export
	 * Populates the active sheet
	 * Заполнение всех полей таблицы активного листа данными из базы
	 * @param $models
	 */
	private function populateActiveSheet($models)
	{
		$j = 1;
		foreach ($models as $model) {
			$j++;
			$k = 0;
			foreach ($this->columnsNameArray as $attr) {
				$this->xl->getActiveSheet()->setCellValueByColumnAndRow($k, $j, $model->{$attr});
				$k++;
			}
		}
		$this->columnsNameArray = [];
	}

	/**
	 * Saves exported file and send it in browser
	 */
	private function saveExportedFile()
	{
		$writer = new \PHPExcel_Writer_Excel2007($this->xl);
		$writer->save($this->fileNameFrom);
		$response = Yii::$app->response;
		$response->sendFile($this->fileNameFrom)->send();
	}

	/**
	 * Finds all models of the main table
	 * @return array|\yii\db\ActiveRecord[]
	 */
	private function getAllModelsOfMainTable()
	{
		/** @var ActiveRecord $mainModelName */
		$mainModelName = $this->mainModelName;
		$mainModels    = $mainModelName::find()->all();

		return $mainModels;
	}

	/**
	 * Import from file to DB
	 * Загружает данные из файла на сервере в базу
	 * @throws \Exception
	 * @throws \PHPExcel_Reader_Exception
	 * @throws \yii\base\InvalidConfigException
	 */
	public function import()
	{
		/** @var ActiveRecord $mainModelName */
		$mainModelName    = $this->mainModelName;
		$columnsNameArray = $this->scheme->getColumnNames();
		$models           = $this->prepareRecordToImport($columnsNameArray, $mainModelName);
		$headers          = $models[$mainModelName::className()]['id'];
		/** если заголовки в файле соответствуют полям модели, то импортировать данные файла */
		if (isset($headers) && empty(array_diff_key(array_flip($columnsNameArray), $headers))) {
			unset($models[$mainModelName::className()]['id']);

			return ($this->populateDB($mainModelName, $models));
		} else {
			return "Данные в файле несовместимы с таблицей в БД!";
		}
	}

	/**
	 * Data preparing: sampling from the file and convert data into an array of model attributes
	 * Подготовка данных: выборка из файла и преобразование в массив атрибутов модели
	 * @param $columnsNameArray
	 * @param $modelName ActiveRecord
	 * @return array Массив атрибутов модели
	 * @throws \PHPExcel_Reader_Exception
	 */
	private function prepareRecordToImport($columnsNameArray, $modelName)
	{
		/** @var \PHPExcel_Reader_Excel2007 $rider */
		$rider = PHPExcel_IOFactory::createReader('Excel2007');
		$rider->setReadDataOnly(true);
		$xl     = $rider->load($this->fileNameFrom);
		$xl->setActiveSheetIndex(0);
		$sheet  = $xl->getActiveSheet();
		$models = [];
		foreach ($sheet->getRowIterator() as $row) {
			$model        = [];
			$cellIterator = $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(false);
			$i = 0;
			foreach ($cellIterator as $cell) {
				if ($i == 0) {
					$mainModelId = $cell->getValue();
				}
				$model[$columnsNameArray[$i]] = $cell->getValue();
				$i++;
			}
			$models[$modelName::className()][$mainModelId] = $model;
		}

		return $models;
	}

	/**
	 * Populates DB
	 * Заполнение БД записями из файла
	 * @param $modelName ActiveRecord
	 * @param $models
	 * @return string
	 */
	private function populateDB($modelName, $models)
	{
		$modelsMain = $modelName::find()->indexBy('id')->all();
		if (Model::loadMultiple($modelsMain, $models, $modelName) && Model::validateMultiple($modelsMain)) {
			$this->saveChangedRecords($modelsMain);
			$deleteResult = $this->deleteAbsentRecordsFromDB($models, $modelsMain, $modelName);
			$insertResult = $this->insertNewRecordsToDB($modelName, $models, $modelsMain);

			return 'Существующие записи обновлены...<br>' . $deleteResult . $insertResult;
		} else {
			return 'Импортирование провалилось!';
		}
	}

	/**
	 * Saves records (without new ones)
	 * Сохранят записи: измененные, неизмененные, удаленные (удаленные записи будут удалены следующим запросом)
	 * @param $modelsMain
	 */
	private function saveChangedRecords($modelsMain)
	{
		/** @var ActiveRecord $modelsMain */
		foreach ($modelsMain as $model) {
			$model->save(false);
		}
	}

	/**
	 * Deletes records from DB that were deleted from file
	 * Удаляет из БД записи, которые были удалены в файле
	 * @param $models
	 * @param $modelsMain
	 * @param $mainModelName ActiveRecord
	 * @return string
	 * @throws \Exception
	 */
	private function deleteAbsentRecordsFromDB($models, $modelsMain, $mainModelName)
	{
		$itemsToDelete = array_diff_key($modelsMain, $models[$mainModelName::className()]);
		if (!empty($itemsToDelete)) {
			/** @var ActiveRecord $item */
			foreach ($itemsToDelete as $item) {
				$item->delete();
			}

			return 'Исключенные записи удалены...<br>';
		} else {
			return '';
		}
	}

	/**
	 * Inserts new records to DB that were added to file
	 * Вставляет в БД новые записи из файла, которых не было до импорта
	 * @param $modelName
	 * @param $models
	 * @param $modelsMain
	 * @return string
	 */
	private function insertNewRecordsToDB($modelName, $models, $modelsMain)
	{
		$itemsToInsert = array_diff_key($models[$modelName], $modelsMain);
		$importFalse = true;
		if (!empty($itemsToInsert)) {
			foreach ($itemsToInsert as $newItem) {
				/** @var ActiveRecord $newModel */
				$newModel             = new $modelName;
				$newModel->attributes = $newItem;
				if ($newModel->validate()) {
					$newModel->save(false);
				} else {
					$importFalse = false;
					return "Новые записи не вставлены, есть ошибки:<br>" . "<code>" . implode(', ', $newModel->getFirstErrors()) . "</code>" .
					"<br>Исправьте ошибки, загрузите файл заново и повторите импорт.";
				}
			}
			if ($importFalse) {
				return "Новые записи добавлены успешно!<br>";
			}
		}
	}

	/**
	 * Backup updated table
	 * Методы для резервного копирования таблиц
	 * Копирует таблицу для резервной копии
	 * @return array
	 * @throws \yii\base\InvalidConfigException
	 * @throws \yii\base\NotSupportedException
	 */
	public function copyTable()
	{
		/** @var ActiveRecord $mainModelName */
		$mainModelName = $this->mainModelName;
		$connection    = Yii::$app->db;
		$tableName     = $mainModelName::tableName();
		$columns       = $this->createColumns($this->scheme);
		$time          = time();
		$newTableName  = $this->nameOfReserveTable . $time . '_' . $tableName;
		$tableOptions  = 'CHARACTER SET utf8 COLLATE utf8_unicode_ci ENGINE=InnoDB';
		$connection->createCommand()->createTable($newTableName, $columns, $tableOptions)->execute();
		$this->createUniqueIndexes($this->scheme, $newTableName);
		$connection->createCommand('INSERT INTO ' . $newTableName . ' SELECT * FROM ' . $tableName)->query();
	}

	/**
	 * Creates an array of columns copied table
	 * Создает массив колонок таблицы для создания резервной копии таблицы
	 * @param $scheme
	 * @return array
	 */
	private function createColumns($scheme)
	{
		$columns = [];
		foreach ($scheme->columns as $col) {
			$pk       = ($col->isPrimaryKey) ? ' PRIMARY KEY' : '';
			$ai       = ($col->autoIncrement) ? ' AUTO_INCREMENT' : '';
			$unsigned = ($col->unsigned) ? ' UNSIGNED' : '';
			$comment  = ($col->comment) ? ' COMMENT ' . Yii::$app->db->quoteValue($col->comment) : '';
			$notNull  = ($col->allowNull) ? '' : ' NOT NULL';
			if ($col->defaultValue === null && $col->allowNull) {
				$default = ' DEFAULT NULL';
			} elseif ($col->defaultValue === null && !$col->allowNull) {
				$default = '';
			} else {
				$default = ' DEFAULT ' . $col->defaultValue;
			}
			$columns[$col->name] = $col->dbType . $unsigned . $notNull . $default . $ai . $pk . $comment;
		}

		return $columns;
	}


	/**
	 * Creates unique indexes for a name of the copied table
	 * Создает для таблицы уникальные индексы при резервном копировании
	 * @param $scheme
	 * @param $newTableName
	 * @throws \yii\base\NotSupportedException
	 * @throws \yii\db\Exception
	 */
	private function createUniqueIndexes($scheme, $newTableName)
	{
		$db            = Yii::$app->db;
		$uniqueIndexes = $db->schema->findUniqueIndexes($scheme);
		if (!empty($uniqueIndexes)) {
			foreach ($uniqueIndexes as $name => $index) {
				$db->createCommand()->createIndex($name, $newTableName, $db->queryBuilder->buildColumns($index), true)->execute();
			}
		}
	}
}
<?php

namespace uranum\excel;


use PHPExcel_IOFactory;
use uranum\excel\models\ImportXls;
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
	const COLOR_SUCCESS = 'success';
	const COLOR_DEFAULT = 'default';
	const COLOR_WARNING = 'warning';
	const COLOR_PRIMARY = 'primary';
	const COLOR_DANGER = 'danger';
	const COLOR_INFO = 'info';
	
	public $columnWidthOfStringType = 35;
	public $columnWidthOfTextType = 50;
	public $notNullColumnColor = 'FFDECC';
	public $columnWidthDefault = 15;
	public $nameOfReserveTable = 'archive_';
	public $backupUrl = '/excel/default/backup';
	public $uploadUrl = '/excel/default/upload';
	public $importUrl = '/excel/default/import';
	public $exportUrl = '/excel/default/export';
	public $modalId = 'excelImportModal';
	public $fileInputId = 'uploadFile';
	public $fullFileNameFrom;
	public $mainModelName;
	public $uploadPath;
	public $fileName;
	public $exportBtnClass = self::COLOR_SUCCESS;
	public $importPopupBtnClass = self::COLOR_DANGER;
	public $panelColorClass = self::COLOR_DEFAULT;
	public $fileInputColorClass = self::COLOR_WARNING;
	public $uploadBtnClass = self::COLOR_SUCCESS;
	public $uploadDataBtnClass = self::COLOR_SUCCESS;
	/**
	 * @var \PHPExcel
	 */
	private $xl;
	private $scheme;
	private $columnsNameArray = [];
	
	/**
	 * @inheritdoc
	 */
	public function init()
	{
		parent::init();
		
		$module                 = \Yii::$app->getModule('excel');
		$this->uploadPath       = $module->params['uploadPath'];
		$this->fileName         = $module->params['fileName'];
		$this->fullFileNameFrom = $this->getFullFileNameFrom();
		if (!$this->mainModelName) {
			throw new InvalidValueException('Model name must be defined!');
		}
		/** @var ActiveRecord $mainModelName */
		$mainModelName = $this->mainModelName;
		$this->scheme  = $mainModelName::getTableSchema();
		$this->xl      = new \PHPExcel();
		Asset::register($this->getView());
	}
	
	/**
	 * @inheritdoc
	 */
	public function run()
	{
		echo "&emsp;";
		echo Html::a(Yii::t('excel', 'Выгрузка из базы'), Url::to([$this->exportUrl, 'className' => $this->mainModelName]), ['class' => 'btn btn-' . $this->exportBtnClass]);
		echo "&emsp;";
		$this->renderModal();
	}
	
	/**
	 * Return the path to uploaded file
	 * @return string full file name
	 * @throws InvalidValueException
	 */
	protected function getFullFileNameFrom()
	{
		if (self::checkFileExist($this->uploadPath)) {
			$extension = (empty($this->getFileExtension())) ? 'xlsx' : $this->getFileExtension();
			
			return $this->uploadPath . DIRECTORY_SEPARATOR . $this->fileName . '.' . $extension;
		} else {
			throw new InvalidValueException('Can not create a destination directory.');
		}
	}
	
	public static function checkFileExist($file)
	{
		if (!file_exists($file)) {
			if (!mkdir($file, 0777, true)) {
				return false;
			}
		}
		
		return true;
	}
	
	/**
	 * Find uploaded file and return it's extension
	 * @return string file extension
	 */
	protected function getFileExtension()
	{
		$alias = Yii::getAlias('@webroot');
		$path  = ($alias . DIRECTORY_SEPARATOR . $this->uploadPath . DIRECTORY_SEPARATOR . $this->fileName . '.*');
		$files = glob($path);
		
		return (is_array($files) && !empty($files)) ? pathinfo($files[0], 4) : false;
	}
	
	/**
	 * Модальное окно для действий импорта
	 * Modal for import actions
	 */
	private function renderModal()
	{
		$model = new ImportXls();
		Modal::begin([
			'header'       => Html::tag('h3', Yii::t('excel', 'Настройки импорта')),
			'id'           => $this->modalId,
			'toggleButton' => [
				'tab'   => 'a',
				'label' => Yii::t('excel', 'Загрузить данные в базу'),
				'class' => 'btn btn-' . $this->importPopupBtnClass,
			],
		]);
		echo Html::tag('h4', Yii::t('excel', 'Порядок действий'));
		echo Html::beginTag('ol');
		echo Html::tag('li', Yii::t('excel', 'Сделать резервную копию данных (можно пропустить)'));
		echo Html::tag('li', Yii::t('excel', 'Выбрать файл данных и загрузить его на сервер'));
		echo Html::tag('li', Yii::t('excel', 'Сделать импорт данных из файла в БД'));
		echo Html::endTag('ol');
		
		echo Html::beginTag('div', ['class' => 'panel panel-' . $this->panelColorClass]);
		echo Html::tag('div', Yii::t('excel', 'Шаг {0, number}', 1), ['class' => 'panel-heading text-center']);
		echo Html::beginTag('div', ['class' => 'panel-body text-center', 'id' => 'step_1']);
		echo Html::beginTag('div', ['class' => 'step-body']);
		echo Html::a(Yii::t('excel', 'Запустить резервное копирование'), '#', [
			'class' => 'btn btn-' . self::COLOR_SUCCESS,
			'id'    => 'startCopying',
			'data'  => [
				'url' => Url::to([
					$this->backupUrl,
					'className' => $this->mainModelName,
				]),
			],
		]);
		echo Html::endTag('div');
		echo Html::endTag('div');
		echo Html::endTag('div');
		
		echo Html::beginTag('div', ['class' => 'panel panel-' . $this->panelColorClass]);
		echo Html::tag('div', Yii::t('excel', 'Шаг {0, number}', 2), ['class' => 'panel-heading text-center']);
		echo Html::beginTag('div', ['class' => 'panel-body text-center', 'id' => 'step_2']);
		echo Html::beginTag('div', ['class' => 'step-body']);
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
			'class' => 'btn btn-' . $this->fileInputColorClass,
			'id'    => $this->fileInputId,
		])->label(false);
		echo Html::submitButton(Yii::t('excel', 'Загрузить файл'), [
			'id'    => 'uploadFileForm',
			'class' => 'btn btn-' . $this->uploadBtnClass,
			'data'  => [
				'url' => Url::to([$this->uploadUrl]),
			],
		]);
		ActiveForm::end();
		echo Html::endTag('div');
		echo Html::endTag('div');
		echo Html::endTag('div');
		
		echo Html::beginTag('div', ['class' => 'panel panel-' . $this->panelColorClass]);
		echo Html::tag('div', Yii::t('excel', 'Шаг {0, number}', 3), ['class' => 'panel-heading text-center']);
		echo Html::beginTag('div', ['class' => 'panel-body text-center', 'id' => 'step_3']);
		echo Html::beginTag('div', ['class' => 'step-body']);
		echo Html::a(Yii::t('excel', 'Загрузить данные в базу'), '#',
			['class' => 'btn btn-' . $this->uploadDataBtnClass, 'id' => 'startImport', 'data' => ['url' => Url::to([$this->importUrl, 'className' => $this->mainModelName])]]);
		echo Html::endTag('div');
		echo Html::endTag('div');
		echo Html::endTag('div');
		
		Modal::end();
	}
	
	/**
	 * Export
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
	 * Finds all models of the main table
	 * @return array|\yii\db\ActiveRecord[]
	 */
	private function getAllModelsOfMainTable()
	{
		/** @var ActiveRecord $mainModelName */
		$mainModelName = $this->mainModelName;
		
		return $mainModelName::find()->all();
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
					$this->xl->getActiveSheet()->setCellValueByColumnAndRow($i, $rowNumber + 1, $cell);
					$i++;
				}
			}
		}
		$this->xl->setActiveSheetIndex(0);
	}
	
	/**
	 * Saves exported file and send it in browser
	 */
	private function saveExportedFile()
	{
		$writer = new \PHPExcel_Writer_Excel2007($this->xl);
		$writer->save($this->fullFileNameFrom);
		$response = Yii::$app->response;
		$response->sendFile($this->fullFileNameFrom)->send();
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
			return Yii::t('excel', "Данные в файле несовместимы с таблицей в БД!");
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
		$xl = $rider->load($this->fullFileNameFrom);
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
			$deleteResult  = $this->deleteAbsentRecordsFromDB($models, $modelsMain, $modelName);
			$insertResult  = $this->insertNewRecordsToDB($modelName, $models, $modelsMain);
			$resultMessage = (!$insertResult && !$deleteResult) ? Yii::t('excel', 'Импортирование завершено!') : '';
			
			return Yii::t('excel', 'Существующие записи обновлены') . '...<br>' . $deleteResult . $insertResult . $resultMessage;
		} else {
			return Yii::t('excel', 'Импортирование провалилось!');
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
			
			return Yii::t('excel', 'Исключенные записи удалены') . '...<br>';
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
		$importFalse   = true;
		if (!empty($itemsToInsert)) {
			foreach ($itemsToInsert as $newItem) {
				/** @var ActiveRecord $newModel */
				$newModel             = new $modelName;
				$newModel->attributes = $newItem;
				if ($newModel->validate()) {
					$newModel->save(false);
				} else {
					$importFalse = false;
					
					return Yii::t('excel', "Новые записи не вставлены, есть ошибки:") . "<br>" . "<code>" . implode(', ', $newModel->getFirstErrors()) . "</code>" .
					"<br>" . Yii::t('excel', "Исправьте ошибки, загрузите файл заново и повторите импорт.");
				}
			}
			if ($importFalse) {
				return Yii::t('excel', "Новые записи добавлены успешно!") . "<br>";
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
		$tableName     = ltrim($mainModelName::tableName(), '{{%');
		$tableName     = rtrim($tableName, '}}');
		$columns       = $this->createColumns($this->scheme);
		$time          = time();
		$newTableName  = $this->nameOfReserveTable . $time . '_' . $tableName;
		$tableOptions  = 'CHARACTER SET utf8 COLLATE utf8_unicode_ci ENGINE=InnoDB';
		$connection->createCommand()->createTable($newTableName, $columns, $tableOptions)->execute();
		$this->createUniqueIndexes($this->scheme, $newTableName);
		$connection->createCommand('INSERT INTO ' . $newTableName . ' SELECT * FROM ' . $mainModelName::tableName())->query();
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
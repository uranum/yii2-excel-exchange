<?php

namespace uranum\excel\controllers;

use uranum\excel\ExcelExchanger;
use uranum\excel\models\ImportXls;
use Yii;
use yii\web\Controller;
use yii\web\UploadedFile;

/**
 * Default controller for the `excel` module
 */
class DefaultController extends Controller
{
	public function actionBackup($className)
	{
		if (Yii::$app->request->isAjax && Yii::$app->request->isPost) {
			$data = new ExcelExchanger(['mainModelName' => $className]);
			$data->copyTable();
			echo Yii::t('excel', 'Данные успешно сохранены!');
		} else {
			echo Yii::t('excel', 'Копирование завершилось неудачей. Попробуйте позже.');
		}
	}
	
	public function actionExport($className)
	{
		$data = new ExcelExchanger(['mainModelName' => $className]);
		$data->export();
	}
	
	public function actionImport($className)
	{
		if (Yii::$app->request->isPost && Yii::$app->request->isAjax) {
			$data = new ExcelExchanger(['mainModelName' => $className]);
			echo $data->import();
		} else {
			echo Yii::t('excel', 'Неправильный запрос.');
		}
	}
	
	/**
	 * Загружает xlsx-файл в папку uploads на сервере с именем export.xlsx
	 * @return string
	 */
	public function actionUpload()
	{
		if (Yii::$app->request->isPost && Yii::$app->request->isAjax) {
			$model = new ImportXls();
			$model->file = UploadedFile::getInstance($model, 'file');
			if ($model->upload()) {
				echo Yii::t('excel', 'Файл успешно загружен!');
			} else {
				echo Yii::t('excel', 'Файл имеет некорректный тип или не удалось создать папку назначения.');
			}
		}
	}
}

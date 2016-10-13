<?php

namespace uranum\excel\models;

use uranum\excel\ExcelExchanger;
use yii\base\Model;
use yii\web\UploadedFile;

class ImportXls extends Model
{
	public $extensions;
	/**
	 * @var UploadedFile
	 */
	public $file;
	public $fileName;
	public $path;
	
	public function init()
	{
		parent::init();
		
		$module           = \Yii::$app->getModule('excel');
		$this->extensions = $module->params['extensions'];
		$this->path       = $module->params['uploadPath'] . DIRECTORY_SEPARATOR;
		$this->fileName   = $module->params['fileName'];
	}
	
	public function rules()
	{
		return [
			[['file'], 'file', 'extensions' => $this->extensions, 'skipOnEmpty' => false,],
		];
	}
	
	public function upload()
	{
		if (!ExcelExchanger::checkFileExist($this->path)) {
			return false;
		}
		if ($this->validate()) {
			$this->file->saveAs($this->path . $this->fileName . '.' . $this->file->extension);
			
			return true;
		} else {
			return false;
		}
	}
}
<?php

namespace uranum\src;

use yii\web\UploadedFile;
use yii\base\Model;

class ImportXls extends Model
{
	/**
	 * @var UploadedFile
	 */
	public $file;

	public function rules()
	{
		return [
			[['file'], 'file', 'extensions' => 'xlsx', 'skipOnEmpty' => false,]
		];
	}

	public function upload()
	{
		if ($this->validate()) {
			$this->file->saveAs('uploads' . DIRECTORY_SEPARATOR . 'export.' . $this->file->extension);
			return true;
		} else {
			return false;
		}
	}
}
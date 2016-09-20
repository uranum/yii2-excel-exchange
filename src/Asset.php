<?php

namespace uranum\excel;


use yii\web\AssetBundle;

class Asset extends AssetBundle
{
	public $sourcePath = '@uranum/excel/assets';
	public $js = [
		'scripts.js'
	];
	public $depends = [
		'yii\web\JqueryAsset'
	];
}
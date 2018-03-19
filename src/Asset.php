<?php

namespace uranum\excel;


use yii\web\AssetBundle;

class Asset extends AssetBundle
{
	public $sourcePath = '@vendor/uranum/yii2-excel-exchange/src/assets/';
	public $js = [
		'js/scripts.js'
	];
	public $css = [
		'css/style.css'
	];
	public $depends = [
		'yii\web\JqueryAsset'
	];
}
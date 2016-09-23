<?php

namespace uranum\excel;


use yii\web\AssetBundle;

class Asset extends AssetBundle
{
	public $sourcePath = '@uranum/excel/assets';
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
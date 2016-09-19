<?php

namespace uranum\src;


use yii\web\AssetBundle;

class Asset extends AssetBundle
{
	public $sourcePath = '@uranum/src/assets';
	public $js = [
		'scripts.js'
	];
	public $depends = [
		'yii\web\JqueryAsset'
	];
}
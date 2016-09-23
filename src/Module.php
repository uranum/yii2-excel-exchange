<?php

namespace uranum\excel;

/**
 * excel module definition class
 */
class Module extends \yii\base\Module
{
    /**
     * @inheritdoc
     */
    public $controllerNamespace = 'uranum\excel\controllers';

    /**
     * @inheritdoc
     */
    public function init()
    {
        parent::init();
	    $this->params['uploadPath'] = ltrim($this->params['uploadPath'], "\x2F\x5C");
	    $this->registerTranslations();
    }
	
	private function registerTranslations()
	{
		\Yii::$app->i18n->translations['excel'] = [
			'class'          => 'yii\i18n\PhpMessageSource',
			'sourceLanguage' => 'ru-RU',
			'basePath'       => __DIR__ . '/messages',
		];
    }
}

<?php

namespace uranum\excel;

use yii\helpers\ArrayHelper;

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
        $params = [
            'fileName'   => $this->params['fileName'] ?? 'export',
            'extensions' => $this->params['extensions'] ?? 'xls xlsx',
            'uploadPath' => ltrim(($this->params['uploadPath'] ?? 'uploads'), "\x2F\x5C"),
        ];
        $this->params = ArrayHelper::merge($this->params, $params);
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

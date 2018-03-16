<?php

namespace Gevman\Yii2Excel;

use Gevman\Yii2Excel\Exception\ImporterException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use Yii;
use yii\base\BaseObject;
use yii\db\ActiveRecord;
use yii\helpers\ArrayHelper;

class Importer extends BaseObject
{
    /**
     * @var string
     */
    public $filePath;

    /**
     * @var string
     */
    public $activeRecord;

    /**
     * @var array
     */
    public $fields;

    /**
     * @var bool
     */
    public $skipFirstRow = false;

    /**
     * @var string|null
     */
    public $scenario;

    /**
     * @var array
     */
    protected $rows;

    /**
     * @var ActiveRecord[]
     */
    protected $models;
    /**
     * @var bool
     */
    protected $isValidated = false;

    /**
     * @var array
     */
    protected $errors;

    public function init()
    {
        $this->filePath = Yii::getAlias($this->filePath);

        try {
            $spreadsheet = IOFactory::load($this->filePath);
            $this->rows = $spreadsheet->getActiveSheet()->toArray();
            if ($this->skipFirstRow) {
                array_shift($this->rows);
            }
            $this->process();
        } catch (Exception $e) {
            throw new ImporterException($e->getMessage(), $e->getCode(), $e);
        }
    }

    public function validate()
    {
        $this->errors = [];
        foreach ($this->models as $index => $model) {
            if (!$model->validate()) {
                $this->errors[$index] = $model->getFirstErrors();
            }
        }
        $this->isValidated = true;

        return empty($this->errors);
    }

    public function save()
    {
        $savedRows = [];

        if (!$this->isValidated) {
            $this->validate();
        }

        foreach ($this->models as $model) {
            if ($model->save()) {
                $savedRows[] = $model->getPrimaryKey();
            }
        }

        return $savedRows;
    }

    public function getErrors()
    {
        return $this->errors;
    }

    public function getModels()
    {
        return $this->models;
    }

    /**
     * @throws ImporterException
     */
    protected function _validate()
    {
        if (!file_exists($this->filePath)) {
            throw new ImporterException("filePath `{$this->filePath}` dones not exist");
        }

        if (!(new $this->activeRecord) instanceof ActiveRecord) {
            throw new ImporterException(sprintf('activeRecord must be instance `%s`'), ActiveRecord::class);
        }
    }

    protected function process()
    {
        foreach ($this->rows as $index => $row) {
            /** @var ActiveRecord $model */
            $model = (new $this->activeRecord);
            if ($this->scenario) {
                $model->setScenario($this->scenario);
            }
            $attributes = [];
            foreach ($this->fields as $field) {
                if (!($attribute = ArrayHelper::getValue($field, 'attribute'))) {
                    throw new ImporterException('attribute missing from one of your fields');
                }
                if (!($value = ArrayHelper::getValue($field, 'value'))) {
                    throw new ImporterException('value missing from one of your fields');
                }
                if (!is_callable($value) && !array_key_exists($value, $row)) {
                    throw new ImporterException("index `$value` not found in row");
                }
                if (is_callable($value)) {
                    $value = $value($row);
                } else {
                    $value = $row[$value];
                }

                $attributes[$attribute] = $value;
            }
            $model->setAttributes($attributes);
            $this->models[$index] = $model;
        }
    }
}
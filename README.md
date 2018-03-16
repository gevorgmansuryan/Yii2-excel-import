# Yii2 Excel Importer
Import excel to ActiveRecord models

## installation
using composer
```bash
composer require gevman/yii2-excel-import
```

## methods
#### __constructor
supports array parameter with following attributes
- `filePath` - full path to excel file to be imported
- `activeRecord` - ActiveRecord class name where imported data should be saved
- `scenario` - ActiveRecord scenario, if leave empty no scenario will be set
- `skipFirstRow` - if true will skip excel's first row (eg. heading row), otherwise will try save first row also
- `fields[]` - array of field definitions

##### fields[]
- `attribute` - attribute name from Your ActieRecord class
- `value` - if callable passed it will receive current row, and return value will be saved to AR, otherwise it will find element which key is passed value in current row

#### validate
validates each populated AR model, and returns false if there's any error, otherwise it will return true

#### save
Saves populated AR models, and returns an array of each saved AR model's primary key
if models is not validated yet, it will validate all models before save 

#### getErrors
Will return array of AR model errors indexed by row's index

#### getModels
Will return array of populated AR models

## examples

- Define Fields

```php
$uploadedFile = \yii\web\UploadedFile::getInstanceByName('file');

$importer = new \Gevman\Yii2Excel\Importer([
    'filePath' => $uploadedFile->tempName,
    'activeRecord' => Product::class,
    'scenario' => Product::SCENARIO_IMPORT,
    'skipFirstRow' => true,
    'fields' => [
        [
            'attribute' => 'keywords',
            'value' => 1,
        ],
        [
            'attribute' => 'itemTitle',
            'value' => 2,
        ],
        [
            'attribute' => 'marketplaceTitle',
            'value' => 3,
        ],
        [
            'attribute' => 'brand',
            'value' => function ($row) {
                return strval($row[4]);
            },
        ],
        [
            'attribute' => 'category',
            'value' => function ($row) {
                return strval($row[4]);
            },
        ],
        [
            'attribute' => 'mpn',
            'value' => function ($row) {
                return strval($row[6]);
            },
        ],
        [
            'attribute' => 'ean',
            'value' => function ($row) {
                return strval($row[7]);
            },
        ],
        [
            'attribute' => 'targetPrice',
            'value' => 8,
        ],
        [
            'attribute' => 'photos',
            'value' => function ($row) {
                $photos = [];
                foreach (StringHelper::explode(strval($row[11]), ',', true, true) as $photo) {
                    if (filter_var($photo, FILTER_VALIDATE_URL)) {
                        $file = @file_get_contents($photo);
                        if ($file) {
                            $filename = md5($file) . '.jpg';
                            file_put_contents(Yii::getAlias("@webroot/gallery/$filename"), $file);
                            $photos[] = $filename;
                        }
                    } else {
                        $photos[] = $photo;
                    }
                }

                return implode(',', $photos);
            }
        ],
        [
            'attribute' => 'currency',
            'value' => 13,
        ],
    ],
]);
```
- Validate, Save, and show errors
```php
if (!$importer->validate()) {
    foreach($importer->getErrors() as $rowNumber => $errors) {
        echo "$rowNumber errors <br>" . implode('<br>', $errors);
    }
} else {
    $importer->save();
}
```

of just

```php
$importer->save();

foreach($importer->getErrors() as $rowNumber => $errors) {
    echo "$rowNumber errors <br>" . implode('<br>', $errors);
}
```

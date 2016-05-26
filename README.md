# PhalconDocs
Library for fast handling of template with phpword

## Requirements

* PHP >= 5.3.9
* Phalcon >= 2.0.0

## Installing via Composer

Install composer in a common location or in your project:

```bash
curl -s http://getcomposer.org/installer | php
```

Create the composer.json file as follows:

```json
{
    "require": {
        "Mctekk/PhalconDocs":"dev-master"
    }
}
```

## Usage

TemplateProcessor

```php
<?php
	$model = YorModels::find(1);
	$path = '/path/to/you/template/';
	$docs = new PhalconDocs\TemplateProcessor($path . "file.docx", $model);
    $docs->saveAs($path . "newfile.docx");

```
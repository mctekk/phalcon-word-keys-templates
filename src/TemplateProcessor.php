<?php namespace PhalconDocs;

use Phalcon\DI\Injectable;
use Phalcon\Mvc\Model;
use PhpOffice\PhpWord\TemplateProcessor as TProcessor;

/**
 *
 */
class TemplateProcessor extends Injectable
{
    public $templatePath;
    public $models;
    public $extras;

    public function __construct($templatePath, Model $models, array $extras = [])
    {
        $this->templatePath = $templatePath;
        $this->models = $models;
        $this->extras = $extras;
    }

    public function make($name = null, $newPath = null)
    {
        $path = "";
        if (is_null($name)) {
            $path = "unknown.docx";
        } else {
            $path = is_null($newPath) ? $name : $newPath . $name;
        }
        $templateProcessor = new TProcessor($this->templatePath);

        if (count($this->extras)) {
            foreach ($this->extras as $key => $value) {
                $templateProcessor->setValue($key, $value);
            }
        }

        $values = $this->getValues();

        foreach ($values as $key => $value) {
            $templateProcessor->setValue($key, $value);
        }

        return $templateProcessor->saveAs($path);
    }

    public function getValues()
    {
        $values = [];
        if (!$this->models) {
            throw new Exception("no data model", 1);
        }

        if ($base = $this->toArray()) {
            foreach ($variable as $key => $value) {
                $values[$key] = $value;
            }
        } else {
            throw new Exception("model empty", 1);
        }

        foreach ($this->modelsManager->getRelations(get_class($this->models)) as $r) {
            $options = $r->getOptions();
            $relacion[$r->getFields()] = $options["alias"];
        }
        print_r($relacion);
    }
}

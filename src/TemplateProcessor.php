<?php namespace PhalconDocs;

use Phalcon\DI\Injectable;
use Phalcon\Mvc\Model;
use PhpOffice\PhpWord\TemplateProcessor as TProcessor;

/**
 *  Class to handle the management of phpdocx template for the keys with sumplantar model and relationships Phalcon
 *  Clase para manejar los template de phpdocx para sumplantar las keys con un modelo de phalcon y sus relaciones
 *  @author Daniel Sánchez <dasdo1@mctekk.com>
 *  @version 1.0.0
 */
class TemplateProcessor extends Injectable {
	/**
	 * Path template file
	 * @var string
	 * @required true
	 */
	public $templatePath;

	/**
	 * Phalcon models to map
	 * @var Phalcon\Mvc\Model
	 */
	public $models;

	/**
	 * array keys and values outside the model
	 * @var array
	 */
	public $extras;

	/**
	 * [__construct description]
	 * @param string             $templatePath Path template file
	 * @param Phalcon\Mvc\Model  $models       Phalcon models to map
	 * @param array              $extras       array keys and values outside the model
	 */
	public function __construct($templatePath, Model $models, array $extras = []) {
		$this->templatePath = $templatePath;
		$this->models = $models;
		$this->extras = $extras;
	}

	/**
	 * Method to create a new document based on the template
	 * @param  string $name    path for the new file (path/to/file.docx)
	 * @return void
	 */
	public function saveAs($path = null) {
		$path = is_null($path) ? 'unknown.docx' : $path;

		/**
		 * new instance of PhpOffice\PhpWord\TemplateProcessor
		 * @var PhpOffice\PhpWord\TemplateProcessor
		 */
		$templateProcessor = new TProcessor($this->templatePath);

		/**
		 * Extra parameters change first
		 */
		if (count($this->extras)) {
			foreach ($this->extras as $key => $value) {
				/**
				 * set value extras
				 */
				$templateProcessor->setValue($key, $value);
			}
		}

		/**
		 * get models values formated
		 * @var array
		 */
		$modelsValues = $this->getModelsValues();

		/**
		 * We iterate the values obtained to apply the values to the templates
		 */
		foreach ($modelsValues as $key => $value) {
			/**
			 * If an array has a different management to flat values
			 */
			if (is_array($value)) {
				try {
					/**
					 * Clone rows based on the primary key of the relationship model
					 */
					$templateProcessor->cloneRow($key, count($value));

					foreach ($value as $k => $next) {
						/**
						 * Set a clone value
						 */
						foreach ($next as $ke => $v) {
							$ke = $key . '.' . $ke;
							$templateProcessor->setValue($ke, $v);
						}
						/**
						 * remove unnecessary keys
						 */
						$r = $key . '#' . ($k + 1);
						$templateProcessor->setValue($r, '');
					}
				} catch (\Exception $e) {
					/**
					 * If a relationship of many is not used the library throws an exception key not found, continue with the capture and execution.
					 * Si una relacion de muchos no es usado la libreria arroja una excepcion de llave no encontrada, la capturamos y seguimos con la ejecucion.
					 */
					continue;
				}
			} else {
				$templateProcessor->setValue($key, $value);
			}
		}

		/**
		 * Save the new file
		 */
		$templateProcessor->saveAs($path);
		return;
	}

	/**
	 * Method for relationships and values model
	 * @return Array [description]
	 */
	public function getModelsValues() {
		$values = [];

		if (!$this->models) {
			throw new Exception('no data model', 1);
		}

		/**
		 * get class name
		 * @var string
		 */
		$namespaceModelsName = get_class($this->models);

		/**
		 * avoid php warning
		 * @var array
		 */
		$explode = explode('\\', $namespaceModelsName);

		/**
		 * get base models name for the key
		 * @var string
		 */
		$modelName = strtolower(end($explode));

		if ($base = $this->models->toArray()) {
			foreach ($base as $key => $value) {
				$key = $modelName . '.' . $key;
				/**
				 * based information model
				 * example.value = value
				 */
				$values[$key] = $value;
			}
		} else {
			throw new Exception('model empty', 1);
		}

		/**
		 * we obtain relations model
		 */
		foreach ($this->modelsManager->getRelations($namespaceModelsName) as $r) {
			/**
			 * relations internal information model
			 * @var array
			 */
			$options = $r->getOptions();

			/**
			 * Name of relationship
			 * @var string
			 */
			$relationshipName = $options['alias'];

			/**
			 * relationship information
			 * @var Phalcon\Mvc\Model Or Phalcon\Mvc\Model\Resultset\Simple
			 */
			$relationshipAlias = 'get' . ucfirst($relationshipName);
			if ($relationship = $this->models->{$relationshipAlias}($options)) {
				foreach ($relationship->toArray() as $k => $value) {
					$key = $modelName . '.' . $relationshipName;
					if (is_array($value)) {
						foreach ($value as $indice => $valor) {
							$values[$key][$k][$indice . '#' . ($k + 1)] = $valor;
						}
					} else {
						$values[$key . '.' . $k] = $value;
					}
				}
			}
		}
		return $values;
	}

	/**
	 * Gets the Path template file.
	 *
	 * @return string
	 */
	public function getTemplatePath() {
		return $this->templatePath;
	}

	/**
	 * Sets the Path template file.
	 *
	 * @param string $templatePath the template path
	 *
	 * @return self
	 */
	public function setTemplatePath($templatePath) {
		$this->templatePath = $templatePath;

		return $this;
	}

	/**
	 * Gets the Phalcon models to map.
	 *
	 * @return Phalcon\Mvc\Model
	 */
	public function getModels() {
		return $this->models;
	}

	/**
	 * Sets the Phalcon models to map.
	 *
	 * @param Phalcon\Mvc\Model $models the models
	 *
	 * @return self
	 */
	public function setModels(Model $models) {
		$this->models = $models;

		return $this;
	}

	/**
	 * Gets the array keys and values outside the model.
	 *
	 * @return array
	 */
	public function getExtras() {
		return $this->extras;
	}

	/**
	 * Sets the array keys and values outside the model.
	 *
	 * @param array $extras the extras
	 *
	 * @return self
	 */
	public function setExtras(array $extras) {
		$this->extras = $extras;

		return $this;
	}
}

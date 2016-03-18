<?php

/**
 * @file
 * Contains \Drupal\excel_serialization\Encoder\ExcelEncoder.
 */

namespace Drupal\excel_serialization\Encoder;

use Symfony\Component\Serializer\Encoder\EncoderInterface;
use Drupal\Component\Serialization\Exception\InvalidDataTypeException;
use Drupal\Core\Entity\EntityFieldManagerInterface;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\Locale\CountryManagerInterface;

/**
 * Adds CSV encoder support for the Serialization API.
 */
class ExcelEncoder implements EncoderInterface {

  /**
   * The entity field manager.
   *
   * @var \Drupal\Core\Entity\EntityFieldManagerInterface
   */
  protected $entityFieldManager;

  /**
   * The entity type manager.
   *
   * @var \Drupal\Core\Entity\EntityTypeManagerInterface
   */
  protected $entityTypeManager;

  /**
   * The country list
   *
   * @var array
   */
  protected $countryList;

  /**
   * Constructs a new ExcelEncoder
   *
   * @param \Drupal\Core\Entity\EntityFieldManagerInterface $entity_field_manager
   *   The entity field manager.
   */
  public function __construct(EntityFieldManagerInterface $entity_field_manager, EntityTypeManagerInterface $entity_type_manager, CountryManagerInterface $country_manager) {
    $this->entityFieldManager = $entity_field_manager;
    $this->entityTypeManager = $entity_type_manager;
    $this->countryList = $country_manager->getList();
  }

  /**
   * The format that this encoder supports.
   *
   * @var string
   */
  protected static $format = 'xlsx';

  /**
   * Column definition
   *
   * @var array
   */
  protected $columns = [];

  /**
   * {@inheritdoc}
   */
  public function supportsEncoding($format) {
    return $format == static::$format;
  }

  /**
   * {@inheritdoc}
   *
   * Uses HTML-safe strings, with several characters escaped.
   */
  public function encode($data, $format, array $context = array()) {
    // Deny any page caching on the current request.
    \Drupal::service('page_cache_kill_switch')->trigger();

    switch (gettype($data)) {
      case "array":
        break;

      case 'object':
        $data = (array) $data;
        break;

      // May be bool, integer, double, string, resource, NULL, or unknown.
      default:
        $data = array($data);
        break;
    }

    try {
      // Bump up memory limit
      ini_set('memory_limit', '1024M');

      // Create a new PHPExcel Object
      $objPHPExcel = new \PHPExcel();

      $workSheet = $objPHPExcel->getActiveSheet();

      if (!empty($data)) {
        $this->extractColumns($data[0]);

        foreach ($this->columns as $key => $column) {
          $workSheet->setCellValueByColumnAndRow($key, 1, $column['label']);
        }

        $row = 2;
        foreach ($data as $rawdata) {
          $cellcontents = $this->formatRow($rawdata);
          foreach ($cellcontents as $column => $cellcentent) {
            $workSheet->setCellValueByColumnAndRow($column, $row, $this->formatValue($cellcentent));
          }
          $row++;
        }
      }
      else {
        $workSheet->setCellValueByColumnAndRow(0, 1, t('No alumni for the selected query.'));
      }

      $objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);

      ob_start();
      $objWriter->save('php://output');
      $excelOutput = ob_get_clean();

      return $excelOutput;
    } catch (\Exception $e) {
      throw new InvalidDataTypeException($e->getMessage(), $e->getCode(), $e);
    }
  }

  /**
   * Formats all cells in a given row.
   *
   * This flattens complex data structures into a string, and formats
   * the string.
   *
   * @param $row
   * @return array
   */
  public function formatRow($row) {
    $formatted_row = array();

    $entity_managers = [
      'paragraph' => $this->entityTypeManager->getStorage('paragraph'),
      'taxonomy_term' => $this->entityTypeManager->getStorage('taxonomy_term')
    ];

    foreach ($this->columns as $column) {
      $value = '';
      // Special handling for paragraph
      if (isset($column['paragraph_field'])) {
        // Create empty pseudo paragraph if no data
        if (!isset($row[$column['paragraph_field']][$column['delta']])) {
          $row[$column['paragraph_field']][$column['delta']] = [
            'paragraph' => FALSE
          ];
        }
        // Load paragraph if not done already
        elseif (!isset($row[$column['paragraph_field']][$column['delta']]['paragraph'])) {
          $row[$column['paragraph_field']][$column['delta']]['paragraph'] =
            $entity_managers['paragraph']->load($row[$column['paragraph_field']][$column['delta']]['target_id']);
        }
        if ($row[$column['paragraph_field']][$column['delta']]['paragraph'] !== FALSE) {
          $field = $row[$column['paragraph_field']][$column['delta']]['paragraph']->get($column['name']);
          if ($field) {
            if ($item = $field->first()) {
              if (isset($column['taxonomy_term'])) {
                $id = $item->get('target_id')->getValue();
                $term = $entity_managers['taxonomy_term']->load($id);
                if (!empty($term)) {
                  $value = $term->getName();
                }
              }
              elseif (isset($column['country'])) {
                $key = $item->get('value')->getValue();
                $value = $this->countryList[$key];
              }
              elseif (method_exists($item, 'toString')) {
                $value = $item->toString();
              }
              else {
                $value = $item->getValue();
                if (is_array($value) && isset($value['value'])) {
                  $value = $value['value'];
                }
              }
            }
          }
        }
      }
      elseif (isset($row[$column['name']][0]['value'])) {
        $value = $row[$column['name']][0]['value'];
        if (isset($column['country'])) {
          $value = $this->countryList[$value];
        }
        elseif (isset($column['date']) && !empty($value)) {
          $value = date(DATE_ISO8601, $value);
        }
      }
      $formatted_row[] = $value;
    }

    return $formatted_row;
  }

  /**
   * Formats a single value for a given CSV cell.
   *
   * @param string $value
   *   The raw value to be formatted.
   *
   * @return string
   *   The formatted value.
   *
   */
  protected function formatValue($value) {
    $value = $value;

    return $value;
  }


  /**
   * Extracts the headers using the first row of values.
   *
   * @param array $data
   *   The array of data to be converted to a Excel 2007.
   *
   * We make the assumption that each row shares the same set of headers
   * will all other rows.
   *
   * @return array
   *   An array of headers.
   */
  protected function extractColumns($first_row) {
    // TODO make this configurable and make use of view configurations of what fields to show

    // Give priority to user name
    unset($first_row['field_name']);
    // And last priority to statistics
    unset($first_row['status']);
    unset($first_row['created']);
    unset($first_row['changed']);
    unset($first_row['access']);
    unset($first_row['login']);
    $fieldnames = array_merge(
      ['field_name_given', 'field_name'],
      array_keys($first_row),
      ['status', 'created','changed', 'access', 'login']
    );

    // Remove unwanted user fields
    $remove_fields = [
      'uuid',
      'langcode',
      'default_langcode',
      'preferred_langcode',
      'preferred_admin_langcode',
      'timezone',
      'init',
      'roles',
      'ds_switch',
      'path',
      'user_picture',
      'field_institution' // This is due to e view relationship for the main view
    ];

    $fieldnames = array_diff($fieldnames, $remove_fields);

    $paragraph_fields = [
      'field_jobs' => 'job',
      'field_scholarships' => 'scholarship'
    ];

    $datetime_fields = [
      'created',
      'changed',
      'access',
      'login'
    ];

    $country_fields = [
      'field_country',
      'field_country_residency'
    ];

    $term_fields = [
      'field_institution',
      'field_qualification',
      'field_sector',
      'field_area'
    ];

    // We hardcode how many paragraph items to show
    $paragraph_count = 2;

    foreach ($fieldnames as $field_name) {
      // Special handling for paragraphs
      if (isset($paragraph_fields[$field_name])) {
        // Add the regular fields of the paragraph
        $paragraph_subfields = $this->entityFieldManager->getFieldDefinitions('paragraph', $paragraph_fields[$field_name]);
        for ($i = 0; $i < $paragraph_count; $i++) {
          foreach ($paragraph_subfields as $paragraph_subfield_name => $paragraph_subfield) {
            if (strpos($paragraph_subfield_name, 'field_') === 0) {
              $label = $paragraph_subfield_name . '_' . ($i + 1);
              $this->columns[] = [
                'name' => $paragraph_subfield_name,
                'label' => $label,
                'paragraph_field' => $field_name,
                'delta' => $i
              ];
            }
          }
        }
      }
      else {
        $this->columns[] = ['name' => $field_name, 'label' => $field_name];
      }
    }

    // Mark date and country fields
    foreach ($this->columns as &$column) {
      if (in_array($column['name'], $datetime_fields)) {
        $column['date'] = TRUE;
      }
      elseif (in_array($column['name'], $country_fields)) {
        $column['country'] = TRUE;
      }
      elseif (in_array($column['name'], $term_fields)) {
        $column['taxonomy_term'] = TRUE;
      }
    }
  }
}

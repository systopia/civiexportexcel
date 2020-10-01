<?php

require_once 'civiexportexcel.civix.php';
use CRM_CiviExportExcel_ExtensionUtil as E;

require_once(__DIR__ . '/vendor/autoload.php');

/**
 * Implementation of hook_civicrm_config
 */
function civiexportexcel_civicrm_config(&$config) {
  _civiexportexcel_civix_civicrm_config($config);
}

/**
 * Implementation of hook_civicrm_xmlMenu
 *
 * @param $files array(string)
 */
function civiexportexcel_civicrm_xmlMenu(&$files) {
  _civiexportexcel_civix_civicrm_xmlMenu($files);
}

/**
 * Implementation of hook_civicrm_install
 */
function civiexportexcel_civicrm_install() {
  return _civiexportexcel_civix_civicrm_install();
}

/**
 * Implementation of hook_civicrm_uninstall
 */
function civiexportexcel_civicrm_uninstall() {
  return _civiexportexcel_civix_civicrm_uninstall();
}

/**
 * Implementation of hook_civicrm_enable
 */
function civiexportexcel_civicrm_enable() {
  return _civiexportexcel_civix_civicrm_enable();
}

/**
 * Implementation of hook_civicrm_disable
 */
function civiexportexcel_civicrm_disable() {
  return _civiexportexcel_civix_civicrm_disable();
}

/**
 * Implementation of hook_civicrm_upgrade
 *
 * @param $op string, the type of operation being performed; 'check' or 'enqueue'
 * @param $queue CRM_Queue_Queue, (for 'enqueue') the modifiable list of pending up upgrade tasks
 *
 * @return mixed  based on op. for 'check', returns array(boolean) (TRUE if upgrades are pending)
 *                for 'enqueue', returns void
 */
function civiexportexcel_civicrm_upgrade($op, CRM_Queue_Queue $queue = NULL) {
  return _civiexportexcel_civix_civicrm_upgrade($op, $queue);
}

/**
 * Implementation of hook_civicrm_managed
 *
 * Generate a list of entities to create/deactivate/delete when this module
 * is installed, disabled, uninstalled.
 */
function civiexportexcel_civicrm_managed(&$entities) {
  return _civiexportexcel_civix_civicrm_managed($entities);
}

/**
 * Implements hook_civicrm_alterSettingsFolders().
 */
function civiexportexcel_civicrm_alterSettingsFolders(&$metaDataFolders = NULL) {
  _civiexportexcel_civix_civicrm_alterSettingsFolders($metaDataFolders);
}

/**
 * Implementation of hook_civicrm_buildForm().
 *
 * Used to add a 'Export to Excel' button in the Report forms.
 */
function civiexportexcel_civicrm_buildForm($formName, &$form) {
  // Reports extend the CRM_Report_Form class.
  // We use that to check whether we should inject the Excel export buttons.
  if (!is_subclass_of($form, 'CRM_Report_Form')) {
    return;
  }

  // Keeps the legacy 4.6 behaviour of displaying the "Export to Excel" button
  // @todo Add a setting and disable it by default in the future (but keep for upgrades)
  if ($form->elementExists('task')) {
    $smarty = CRM_Core_Smarty::singleton();
    $vars = $smarty->get_template_vars();

    $form->_excelButtonName = $form->getButtonName('submit', 'excel');

    $label = (! empty($vars['instanceId']) ? E::ts('Export to Excel') : E::ts('Preview Excel'));
    $form->addElement('submit', $form->_excelButtonName, $label);

    CRM_Core_Region::instance('page-body')->add(array(
      'template' => 'CRM/Report/Form/Actions-civiexportexcel.tpl',
    ));
  }
}

/**
 * Implements hook_civicrm_export().
 *
 * Called mostly to export search results.
 */
function civiexportexcel_civicrm_export($exportTempTable, $headerRows, $sqlColumns, $exportMode) {
  $writeHeader = true;

  $rows = array();

  $query = "SELECT * FROM $exportTempTable";
  $dao = CRM_Core_DAO::executeQuery($query);

  while ($dao->fetch()) {
    $row = array();
    foreach ($sqlColumns as $column => $dontCare) {
      $row[$column] = $dao->$column;
    }

    $rows[] = $row;
  }

  $dao->free();

  CRM_CiviExportExcel_Utils_SearchExport::export2excel2007($headerRows, $sqlColumns, $rows);
}

/**
 * Implements hook_civicrm_alterMailParams().
 *
 * Intercepts outgoing report emails, in order to attach the
 * excel2007 version of the report.
 *
 * TODO: we should really propose a patch to CRM_Report_Form::endPostProcess().
 */
function civiexportexcel_attach_to_email(&$form, &$rows, &$attachments) {
  $config = CRM_Core_Config::singleton();

  $filename = 'CiviReport.xlsx';
  $fullname = $config->templateCompileDir . CRM_Utils_File::makeFileName($filename);

  CRM_CiviExportExcel_Utils_Report::generateFile($form, $rows, $fullname);

  $attachments[] = array(
    'fullPath' => $fullname,
    'mime_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'cleanName' => $filename,
  );
}

/**
 * Implements hook_civicrm_searchTasks().
 */
function civiexportexcel_civicrm_searchTasks($objectType, &$tasks) {
  $path = CRM_Utils_System::currentPath();

  if ($objectType == 'contact' && $path == 'civicrm/contact/search/custom') {
    $id = CRM_Core_Task::TASK_EXPORT;
    $tasks[$id]['class'] = 'CRM_CiviExportExcel_Form_Task_Export';
  }
}

/**
 * Implements hook_civicrm_alterReportVar().
 */
function civiexportexcel_civicrm_alterReportVar($type, &$vars, $form) {
  switch ($type) {
  case 'outputhandlers':
    $vars['\Civi\Report\Civiexportexcel\Excel2007'] = '\Civi\Report\Civiexportexcel\Excel2007';
    break;
  case 'actions':
    // Make sure to add before "Export to CSV"
    $new = [];

    foreach ($vars as $key => $val) {
      if ($key == 'report_instance.csv') {
        $new['report_instance.excel2007'] = ['title' => E::ts('Export to Excel')];
      }

      $new[$key] = $val;
    }

    $vars = $new;

    break;
  }
}

/**
 * Implements hook_civicrm_links().
 */
function civiexportexcel_civicrm_links($op, $objectName, $objectId, &$links, &$mask, &$values) {
  if ($op == 'view.report.links') {
    // These don't seem to quite follow the docs for hook_civicrm_links.
    // See CRM/Report/Page/InstanceList.php
    $links['excel2007'] = [
      'label' => E::ts('Export to Excel'),
      'url' => CRM_Utils_System::url("civicrm/report/instance/{$objectId}", 'reset=1&force=1&output=excel2007'),
    ];
  }
}

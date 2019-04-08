<?php

class CRM_CiviExportExcel_Form_Task_Export extends CRM_Contact_Form_Task {

  /**
   * Build the form
   *
   * @access public
   *
   * @return void
   */
  public function buildQuickForm() {
    $this->addDefaultButtons(ts('Export'));
  }

  /**
   * Process the form after the input has been submitted and validated.
   *
   * @access public
   *
   * @return None
   */
  public function postProcess() {
    $values = $this->controller->exportValues();
    $csid = $values['csid'];

    // Fetch the Custom Search class name
    $classname = civicrm_api3('OptionValue', 'getsingle', [
      'option_group_id' => 'custom_search',
      'value' => $csid,
    ])['name'];

    $cs = new $classname($values);

    $headers = $cs->columns();
    $headers = array_flip($headers);

    // The includeIds (4th argument) makes sure that we only export the selected IDs
    $sql = $cs->all(0, 0, NULL, TRUE);

    $dao = CRM_Core_DAO::executeQuery($sql);
    $rows = [];

    while ($dao->fetch()) {
      $row = [];

      foreach ($headers as $key => $val) {
        $row[$key] = $dao->{$key};
      }

      if (method_exists($cs, 'alterRow')) {
        $cs->alterRow($row);
      }

      $rows[] = $row;
    }

    CRM_CiviExportExcel_Utils_SearchExport::export2excel2007($headers, $headers, $rows);
  }


}

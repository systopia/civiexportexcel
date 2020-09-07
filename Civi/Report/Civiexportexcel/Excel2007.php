<?php

namespace Civi\Report\Civiexportexcel;

use Civi\Report\OutputHandlerInterface;
use Civi\Report\OutputHandlerBase;
use CRM_CiviExportExcel_ExtensionUtil as E;

/**
 * Excel2007 Report Output Handler
 *
 * Based on examplereportoutputhandler by DaveD
 * https://lab.civicrm.org/DaveD/examplereportoutputhandler/-/blob/master/Civi/Report/Examplereportoutputhandler/Emoji.php
 */
class Excel2007 extends OutputHandlerBase implements OutputHandlerInterface {

  /**
   * Are we a suitable output handler based on the given form?
   *
   * The class member $form isn't set yet at this point since we don't
   * even know if we're in play yet, so the form is a parameter.
   *
   * @param \CRM_Report_Form $form
   *
   * @return bool
   */
  public function isOutputHandlerFor(\CRM_Report_Form $form):bool {
    return ($form->getOutputMode() === 'excel2007');
  }

  /**
   * Return the download filename. This should be the "clean" name, not
   * a munged temporary filename.
   *
   * @return string
   */
  public function getFileName():string {
    $datetime = date('Y-m-d H:i');
    return $this->getForm()->getTitle() . ' - ' . $datetime . '.xlsx';
  }

  /**
   * Return the html body of the email.
   *
   * @return string
   */
  public function getMailBody():string {
    return $this->getForm()->getReportHeader()
      . '<p>' . E::ts('See attached.') . '</p>'
      . $this->getForm()->getReportFooter();
  }

  /**
   * Return the report contents as a string.
   *
   * @return string
   */
  public function getOutputString():string {
    // See note in CRM_Report_Form::endPostProcess about $rows.
    // Copied from CRM/Report/OutputHandler/Csv.php
    $rows = $this->getForm()->getTemplate()->get_template_vars('rows');

    // avoid pass-by-ref warning
    $form = $this->getForm();

    return \CRM_CiviExportExcel_Utils_Report::generateFile($form, $rows);
  }

  /**
   * Set headers as appropriate and send the output to the browser.
   */
  public function download() {
    \CRM_Utils_System::setHttpHeader('Content-Type', $this->getMimeType());
    \CRM_Utils_System::setHttpHeader('Content-Disposition', 'attachment; filename=' . $this->getFileName());
    \CRM_Utils_System::setHttpHeader('Content-Description', $this->getFileName());
    \CRM_Utils_System::setHttpHeader('Content-Transfer-Encoding', 'binary');
    // Being over-zealous to make sure that the file is not cached
    \CRM_Utils_System::setHttpHeader('Cache-Control', 'no-cache, must-revalidate, post-check=0, pre-check=0');
    \CRM_Utils_System::setHttpHeader('Last-Modified', gmdate('D, d M Y H:i:s') . ' GMT');

    echo $this->getOutputString();
  }

  /**
   * Mime type of the attachment.
   *
   * @return string
   */
  public function getMimeType():string {
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }

  /**
   * Charset of the attachment.
   *
   * @return string
   */
  public function getCharset():string {
    return 'utf-8';
  }

}

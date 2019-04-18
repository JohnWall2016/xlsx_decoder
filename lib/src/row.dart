import './attached_xml_element.dart';
import './sheet.dart';
import './workbook.dart';

class Row extends AttachedXmlElement {
  Sheet _sheet;
  Sheet get sheet => _sheet;
  Workbook get workbook => _sheet?.workbook;
  
  Row(this._sheet, XmlElement node) : super(node) {

  }
}
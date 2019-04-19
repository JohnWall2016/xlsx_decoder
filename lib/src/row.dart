import './attached_xml_element.dart';
import './sheet.dart';
import './workbook.dart';
import './cell.dart';
import './xml_utils.dart';

class Row extends AttachedXmlElement {
  Sheet _sheet;
  Sheet get sheet => _sheet;
  Workbook get workbook => _sheet?.workbook;

  Map<int, Cell> _cells = {};
  
  Row(this._sheet, XmlElement node) : super(node) {
    elements.forEach((cellNode) {
      var cell = Cell(this, cellNode);
      _cells[cell.column] = cell;
    });
  }

  int get row => getAttribute(thisNode, 'r');
}
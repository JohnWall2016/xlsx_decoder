import './attached_xml_element.dart';
import './sheet.dart';
import './workbook.dart';
import './cell.dart';
import './xml_utils.dart';
import './address_converter.dart';

class Row extends AttachedXmlElement {
  Sheet _sheet;
  Sheet get sheet => _sheet;
  Workbook get workbook => _sheet?.workbook;

  Map<int, Cell> _cells = {};
  
  Row(this._sheet, XmlElement node) : super(node) {
    elements.forEach((cellNode) {
      var cell = Cell(this, cellNode);
      _cells[cell.columnIndex] = cell;
    });
  }

  int get index => getAttribute(thisNode, 'r');

  Cell cellAt(int index) {
    var cell = _cells[index];
    if (cell != null) return cell;

    int styleId;
    int rowStyleId = getAttribute(thisNode, 's');
    int columnStyleId = sheet.existingColumnStyleId(index);

    if (rowStyleId != null) styleId = rowStyleId;
    else if (columnStyleId != null) styleId = columnStyleId;

    cell = Cell.create(this, index, styleId);
    _cells[index] = cell;

    return cell;
  }

  Cell cell(String columnName) {
    int index = columnNameToNumber(columnName);
    return cellAt(index);
  }

  XmlElement toXml() {
    thisNode.children.clear();
    var cellIndexes = _cells.keys.toList()..sort();
    cellIndexes.forEach((index) {
      thisNode.children.add(_cells[index].toXml());
    });
    return thisNode;
  }
}
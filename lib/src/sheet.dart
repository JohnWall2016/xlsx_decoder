import './workbook.dart';
import './root_element.dart';

class Sheet {
  Workbook _workbook;
  XmlElement _sheetIdNode;
  XmlElement _sheetNode, _sheetRelationshipsNode;

  Sheet(Workbook workbook, XmlElement idNode, XmlElement node,
      XmlElement relationshipsNode) {
    if (node == null) {}
  }
}

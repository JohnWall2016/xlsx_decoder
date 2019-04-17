import './workbook.dart';
import './document.dart';

class Sheet {
  Workbook _workbook;
  XmlElement _sheetIdNode;
  XmlDocument _sheetNode, _sheetRelationshipsNode;

  Sheet(this._workbook, this._sheetIdNode, this._sheetNode,
      this._sheetRelationshipsNode);
}

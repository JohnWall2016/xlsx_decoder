import './workbook.dart';
import './attached_xml_element.dart';
import './nodes.dart';
import './range.dart';
import './relationships.dart';
import './xml_utils.dart';
import './row.dart';
import './column.dart';

class Sheet extends AttachedXmlElement {
  Workbook _workbook;
  Workbook get workbook => _workbook;

  XmlElement _IdNode;
  XmlElement _sheetRelationshipsNode;

  int _maxSharedFormulaId = -1;

  Map<String, XmlElement> _mergeCells = {};
  Map<String, XmlElement> _dataValidations = {};
  Map<String, XmlElement> _hyperlinks = {};

  Range _autoFilter = null;

  Relationships _relationships;

  List<Row> _rows = [];
  List<Column> _columns = [];

  Sheet(this._workbook, _idNode, XmlElement node, XmlElement relationshipsNode)
      : super(node ??
            (Element('worksheet', {
              'xmlns':
                  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
              'xmlns:r':
                  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'xmlns:mc':
                  'http://schemas.openxmlformats.org/markup-compatibility/2006',
              'mc:Ignorable': 'x14ac',
              'xmlns:x14ac':
                  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
            })
                  ..children.add(Element('sheetData')))
                .toXmlNode()) {
    _relationships = Relationships(relationshipsNode);

    removeChild(thisNode, 'dimension');

    var sheetDataNode = findChild(thisNode, 'sheetData');
    sheetDataNode.children.forEach((rowNode) {
      var row = Row(this, rowNode);
      _rows[row.row] = row;
    });


  }

  updateMaxSharedFormulaId(int sharedFormulaId) {
    if (sharedFormulaId > _maxSharedFormulaId) {
      _maxSharedFormulaId = sharedFormulaId;
    }
  }

}

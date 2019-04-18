import './workbook.dart';
import './root_element.dart';
import './nodes.dart';
import './range.dart';
import './relationships.dart';

class Sheet extends RootElement {
  Workbook _workbook;
  XmlElement _IdNode;
  XmlElement _sheetRelationshipsNode;

  int _maxSharedFormulaId = -1;

  Map<String, XmlElement> _mergeCells = {};
  Map<String, XmlElement> _dataValidations = {};
  Map<String, XmlElement> _hyperlinks = {};

  Range _autoFilter = null;

  Relationships _relationships;

  Sheet(Workbook workbook, XmlElement idNode, XmlElement node,
      XmlElement relationshipsNode)
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
    _workbook = workbook;
    _IdNode = idNode;
    
    _relationships = Relationships(relationshipsNode);
  }
}

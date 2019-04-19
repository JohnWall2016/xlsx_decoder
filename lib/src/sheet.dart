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

  XmlElement _idNode;

  int _maxSharedFormulaId = -1;

  Range _autoFilter = null;

  Relationships _relationships;

  Map<int, Row> _rows = {};
  List<Column> _columns = [];

  XmlElement _colsNode;
  Map<int, XmlElement> _colNodes = {};

  XmlElement _sheetPrNode;

  XmlElement _mergeCellsNode;
  Map<String, XmlElement> _mergeCells = {};

  XmlElement _dataValidationsNode;
  Map<String, XmlElement> _dataValidations = {};

  XmlElement _hyperlinksNode;
  Map<String, XmlElement> _hyperlinks = {};

  Sheet(this._workbook, this._idNode, XmlElement node, XmlElement relationshipsNode)
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

    _colsNode = findChild(thisNode, 'cols');
    if (_colsNode != null) {
      removeChild(thisNode, _colsNode);
    } else {
      _colsNode = Element('cols').toXmlNode();
    }

    _colsNode.children.whereType<XmlElement>().forEach((colNode) {
      int min = getAttribute(colNode, 'min');
      int max = getAttribute(colNode, 'max');
      for (var i = min; i <= max; i++) {
        _colNodes[i] = colNode;
      }
    });

    _sheetPrNode = findChild(thisNode, 'sheetPr');
    if (_sheetPrNode == null) {
      _sheetPrNode = Element('sheetPr').toXmlNode();
      insertInOrder(thisNode, _sheetPrNode, nodeOrder);
    }

    _mergeCellsNode = findChild(thisNode, 'mergeCells');
    if (_mergeCellsNode != null) {
      removeChild(thisNode, _mergeCellsNode);
    } else {
      _mergeCellsNode = Element('mergeCells').toXmlNode();
    }

    _mergeCellsNode.children.whereType<XmlElement>().forEach((mergeCellNode) {
      _mergeCells[getAttribute(mergeCellNode, 'ref')] = mergeCellNode;
    });
    _mergeCellsNode.children.clear();

    _dataValidationsNode = findChild(thisNode, 'dataValidations');
    if (_dataValidationsNode != null) {
      removeChild(thisNode, _dataValidationsNode);
    } else {
      _dataValidationsNode = Element('dataValidations').toXmlNode();
    }

    _dataValidationsNode.children.whereType<XmlElement>().forEach((dataValidationNode) {
      _dataValidations[getAttribute(dataValidationNode, 'sqref')] = dataValidationNode;
    });
    _dataValidationsNode.children.clear();

    _hyperlinksNode = findChild(thisNode, 'hyperlinks');
    if (_hyperlinksNode != null) {
      removeChild(thisNode, _hyperlinksNode);
    } else {
      _hyperlinksNode = Element('hyperlinks').toXmlNode();
    }

    _hyperlinksNode.children.whereType<XmlElement>().forEach((hyperlinkNode) {
      _hyperlinks[getAttribute(hyperlinkNode, 'ref')] = hyperlinkNode;
    });
    _hyperlinksNode.children.clear();
  }

  updateMaxSharedFormulaId(int sharedFormulaId) {
    if (sharedFormulaId > _maxSharedFormulaId) {
      _maxSharedFormulaId = sharedFormulaId;
    }
  }


  static const nodeOrder = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "autoFilter", "protectedRanges", "scenarios", "autoFilter",
    "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
    "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
    "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
    "extLst"
  ];
}

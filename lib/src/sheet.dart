import './workbook.dart';
import './attached_xml_element.dart';
import './nodes.dart';
import './range.dart';
import './relationships.dart';
import './xml_utils.dart';
import './row.dart';
import './column.dart';
import './cell.dart';
import './address_converter.dart';
import './splay_tree.dart';

class Sheet extends AttachedXmlElement {
  Workbook _workbook;
  Workbook get workbook => _workbook;

  XmlElement _idNode;

  int _maxSharedFormulaId = -1;

  Range _autoFilter;
  Range get autoFilter => _autoFilter;

  Relationships _relationships;

  int _lastRowIndex = -1;
  int get lastRowIndex => _lastRowIndex;

  XmlElement _sheetDataNode;

  SplayTreeMap<int, Row> _rows = SplayTreeMap();
  List<Column> _columns = [];

  XmlElement _colsNode;
  Map<int, XmlElement> _colNodes = {};

  XmlElement _sheetPrNode;

  XmlElement _mergeCellsNode;
  Map<RangeRef, XmlElement> _mergeCells = {};

  XmlElement _dataValidationsNode;
  Map<String, XmlElement> _dataValidations = {};

  XmlElement _hyperlinksNode;
  Map<String, XmlElement> _hyperlinks = {};

  Sheet(this._workbook, this._idNode, XmlElement node,
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
    _relationships = Relationships(relationshipsNode);

    removeChild(thisNode, 'dimension');

    _sheetDataNode = findChild(thisNode, 'sheetData');
    if (_sheetDataNode != null) {
      removeChild(thisNode, _sheetDataNode);
    }
    _sheetDataNode.children.forEach((rowNode) {
      var row = Row(this, rowNode);
      var index = row.index;
      if (index > _lastRowIndex) _lastRowIndex = index;
      _rows[index] = row;
    });
    _sheetDataNode.children.clear();

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
    _colsNode.children.clear();

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
      _mergeCells[RangeRef.fromAddress(getAttribute(mergeCellNode, 'ref'))] =
          mergeCellNode;
    });
    _mergeCellsNode.children.clear();

    _dataValidationsNode = findChild(thisNode, 'dataValidations');
    if (_dataValidationsNode != null) {
      removeChild(thisNode, _dataValidationsNode);
    } else {
      _dataValidationsNode = Element('dataValidations').toXmlNode();
    }

    _dataValidationsNode.children
        .whereType<XmlElement>()
        .forEach((dataValidationNode) {
      _dataValidations[getAttribute(dataValidationNode, 'sqref')] =
          dataValidationNode;
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

  String get name => getAttribute(_idNode, 'name');

  void set name(String name) => setAttribute(_idNode, 'name', name);

  Row rowAt(int index) {
    var row = _rows[index];
    if (row != null) return row;

    var rowNode = Element('row', {'r': index}).toXmlNode();

    row = Row(this, rowNode);
    _rows[index] = row;
    if (index > _lastRowIndex) _lastRowIndex = index;
    return row;
  }

  Cell cellAt(int row, int column) => rowAt(row).cellAt(column);

  Cell cell(String cellName) {
    var ref = CellRef.fromAddress(cellName);
    if (ref == null) throw 'Invalid cell name';

    return cellAt(ref.row, ref.column);
  }

  int existingColumnStyleId(int column) {
    var colNode = _colNodes[column];
    return getAttribute(colNode, 'style');
  }

  updateMaxSharedFormulaId(int sharedFormulaId) {
    if (sharedFormulaId > _maxSharedFormulaId) {
      _maxSharedFormulaId = sharedFormulaId;
    }
  }

  Row insertRow(int index, [XmlElement rowNode]) {
    rowNode ??= Element('row', {'r': index}).toXmlNode();
    var row = Row(this, rowNode);
    _rows.nodesFrom(index, (node) {
      var r = node.key + 1;
      node.key = r;
      setAttribute(node.value.thisNode, 'r', r);
    });
    _rows[index] = row;

    _mergeCells.forEach((ref, node) {
      if (index <= ref.start.row || index <= ref.end.row) {
        if (index <= ref.start.row) ref.start.row++;
        if (index <= ref.end.row) ref.end.row++;
        setAttribute(node, 'ref', ref.toAddress());
      }
    });
    return row;
  }

  Row insertRowCopyFrom(int index, int copyIndex, {bool clearValue = false}) {
    var copyRow = _rows[copyIndex];
    if (copyRow == null) return insertRow(index);
    return insertRow(
        index, copyRow.toXml(rowIndex: index, clearValue: clearValue));
  }

  toXmls() {
    var node = thisNode.copy();

    var colNodes = <XmlElement>[];
    _colNodes.keys.forEach((i) {
      var colNode = _colNodes[i];
      if (i == getAttribute<int>(colNode, 'min') &&
          colNode.attributes.length > 2) {
        colNodes.add(colNode.copy());
      }
    });

    _colsNode.children.clear();
    var colsNode = _colsNode.copy();
    colsNode.children.addAll(colNodes);
    if (colsNode.children.length > 0) {
      insertInOrder(node, colsNode, nodeOrder);
    }

    _sheetDataNode.children.clear();
    var sheetDataNode = _sheetDataNode.copy();
    sheetDataNode.children.addAll(_rows.values.map((r) => r.toXml()));
    if (sheetDataNode.children.length > 0) {
      insertInOrder(node, sheetDataNode, nodeOrder);
    }

    _hyperlinksNode.children.clear();
    var hyperlinksNode = _hyperlinksNode.copy();
    hyperlinksNode.children.addAll(_hyperlinks.values.map((h) => h.copy()));
    if (hyperlinksNode.children.length > 0) {
      insertInOrder(node, hyperlinksNode, nodeOrder);
    }

    _mergeCellsNode.children.clear();
    var mergeCellsNode = _mergeCellsNode.copy();
    mergeCellsNode.children.addAll(_mergeCells.values.map((m) => m.copy()));
    if (mergeCellsNode.children.length > 0) {
      insertInOrder(node, mergeCellsNode, nodeOrder);
    }

    _dataValidationsNode.children.clear();
    var dataValidationsNode = _dataValidationsNode.copy();
    dataValidationsNode.children
        .addAll(_dataValidations.values.map((d) => d.copy()));
    if (dataValidationsNode.children.length > 0) {
      insertInOrder(node, dataValidationsNode, nodeOrder);
    }

    if (_autoFilter != null) {
      insertInOrder(
          node,
          Element('autoFilter', {'ref': _autoFilter.address()}).toXmlNode(),
          nodeOrder);
    }

    return {
      'id': _idNode,
      'sheet': node,
      'relationships': _relationships.thisNode
    };
  }

  static const nodeOrder = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "autoFilter",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ];
}

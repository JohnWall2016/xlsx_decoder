import './attached_xml_element.dart';
import './row.dart';
import './address_converter.dart';
import './xml_utils.dart';
import './workbook.dart';
import './formula_error.dart';
import './sheet.dart';
import './nodes.dart';

class Cell {
  Row _row;
  Row get row => _row;
  Sheet get sheet => _row?.sheet;
  Workbook get workbook => _row?.workbook;

  int get rowIndex => _row?.index;

  int _columnIndex;
  int get columnIndex => _columnIndex;
  int _styleId;

  String _type;
  dynamic _value;

  void setValue<T>(T value) {
    _value = value;
  }

  List<XmlAttribute> _remainingAttributes = [];

  String _formulaType;
  String _formulaRef;
  String _formula;
  int _sharedFormulaId;
  List<XmlAttribute> _remainingFormulaAttributes = [];

  List<XmlNode> _remainingChildren;

  T value<T>() {
    if (_value == null)
      return null;
    else if (_value is T)
      return _value;
    else if (T == String)
      return _value.toString() as T;
    else if (T == double && _value is int)
      return _value.toDouble();
    else if (T == Object || T == dynamic) return _value;
    return null;
  }

  Cell(this._row, XmlElement node) {
    _parseNode(node);
  }

  Cell.create(this._row, this._columnIndex, [this._styleId]);

  void _parseNode(XmlElement node) {
    String type;
    node.attributes.forEach((attr) {
      switch (attr.name.local) {
        case 'r':
          var ref = CellRef.fromAddress(attr.value);
          _columnIndex = ref.column;
          break;
        case 's':
          _styleId = int.parse(attr.value);
          break;
        case 't':
          type = attr.value;
          break;
        default:
          _remainingAttributes.add(attr);
      }
    });

    // Parse the value.
    _type = type;
    switch (type) {
      case 's':
        // String value.
        var sharedIndex = int.parse(findChild(node, 'v').children[0].text);
        _value = workbook.sharedStrings.getStringByIndex(sharedIndex);
        break;
      case 'str':
        // Simple string value.
        _value = findChild(node, 'v').children[0].text;
        break;
      case 'inlineStr':
        // Inline string value: can be simple text or rich text.
        var isNode = findChild(node, 'is');
        XmlElement tNode = isNode.children[0];
        if (tNode.name.local == 't') {
          _value = tNode.children[0].text;
        } else {
          _value = isNode.children;
        }
        break;
      case 'b':
        // Boolean value.
        _value = findChild(node, 'v').children[0].text == '1';
        break;
      case 'e':
        // Error value.
        _value = FormulaError.getError(findChild(node, 'v').children[0].text);
        break;
      default:
        // Number value.
        var vNode = findChild(node, 'v');
        _value = vNode != null ? num.parse(vNode.children[0].text) : null;
        break;
    }

    // Parse the formula if present..
    var fNode = findChild(node, 'f');
    if (fNode != null) {
      fNode.attributes.forEach((attr) {
        switch (attr.name.local) {
          case 't':
            _formulaType = attr.value ?? 'normal';
            break;
          case 'ref':
            _formulaRef = attr.value;
            break;
          case 'si':
            _sharedFormulaId = int.tryParse(attr.value);
            if (_sharedFormulaId != null) {
              sheet?.updateMaxSharedFormulaId(_sharedFormulaId);
            }
            break;
          default:
            _remainingFormulaAttributes.add(attr);
        }
      });
      _formula = fNode.children[0].text;
    }

    // If any unknown children are still present, store them for later output.
    removeChild(node, 'f');
    removeChild(node, 'v');
    removeChild(node, 'is');

    _remainingChildren = node.children;
  }

  String address() => CellRef(rowIndex, columnIndex).toAddress();

  XmlElement toXml() {
    var node = Element('c').toXmlNode()
      ..attributes.addAll(_remainingAttributes ?? []);

    setAttribute(node, 'r', address());

    if (_formulaType != null) {
      var fNode = Element('f').toXmlNode()
        ..attributes.addAll(_remainingFormulaAttributes ?? []);

      if (_formulaType != null && _formulaType != 'normal') {
        setAttribute(fNode, 't', _formulaType);
      }
      if (_formulaRef != null) {
        setAttribute(fNode, 'ref', _formulaRef);
      }
      if (_sharedFormulaId != null) {
        setAttribute(fNode, 'si', _sharedFormulaId);
      }
      if (_formula != null) {
        fNode.children.add(Text(_formula).toXmlNode());
      }
      node.children.add(fNode);
    } else if (_value != null) {
      String type, text;
      if (_type == 's' || _value is String || _value is List) {
        // TODO(wj): Rich text is array for now
        type = 's';
        text = workbook.sharedStrings.getIndexForString(_value).toString();
      } else if (_value is bool) {
        type = 'b';
        text = _value ? '1' : '0';
      } else if (_value is num) {
        type = '';
        text = _value.toString();
      }
      if (type != null) {
        if (type.isNotEmpty) setAttribute(node, 't', type);
        node.children.add(
            Element('v').toXmlNode()..children.add(Text(text).toXmlNode()));
      }
    }

    if (_styleId != null) {
      setAttribute(node, 's', _styleId);
    }

    if (_remainingChildren != null) {
      node.children.addAll(_remainingChildren.map((c) => c.copy()));
    }

    return node;
  }
}

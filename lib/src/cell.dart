import './attached_xml_element.dart';
import './row.dart';
import './address_converter.dart';
import './xml_utils.dart';
import './workbook.dart';
import './formula_error.dart';
import './sheet.dart';

class Cell {
  Row _row;
  Row get row => _row;
  Sheet get sheet => _row?.sheet;
  Workbook get workbook => _row?.workbook;

  int _column;
  int get column => _column;
  int _styleId;
  dynamic _value;
  List<XmlAttribute> _remainingAttributes = [];

  String _formulaType;
  String _formulaRef;
  String _formula;
  int _sharedFormulaId;
  List<XmlAttribute> _remainingFormulaAttributes = [];

  List<XmlNode> _remainingChildren;

  T value<T>() {
    if (_value is T) return _value as T;
    return null;
  }

  Cell(this._row, XmlElement node) {
    _parseNode(node);
  }

  Cell.create(this._row, this._column, [this._styleId]);

  void _parseNode(XmlElement node) {
    String type;
    node.attributes.forEach((attr) {
      switch (attr.name.local) {
        case 'r':
          var ref = CellRef.fromAddress(attr.value);
          _column = ref.column;
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
        _value = vNode != null ? double.parse(vNode.children[0].text) : null;
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
}

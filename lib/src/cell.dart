import './attached_xml_element.dart';
import './row.dart';
import './address_converter.dart';
import './xml_utils.dart';
import './workbook.dart';
import './formula_error.dart';

class Cell {
  Row _row;
  Row get row => _row;

  Workbook get workbook => _row?.workbook;

  int _column;
  int _styleId;

  dynamic _value;

  T value<T>() {
    if (_value is T) return _value as T;
    return null;
  }

  List<XmlAttribute> _remainingAttributes = [];

  Cell(this._row, XmlElement node) {
    _parseNode(node);
  }

  Cell.create(this._row, this._column, [this._styleId]);

  void _parseNode(XmlElement node) {
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
          switch (attr.value) {
            // Parse the value.
            case 's':
              // String value.
              var sharedIndex =
                  int.parse(findChild(node, 'v').children[0].text);
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
              _value =
                  FormulaError.getError(findChild(node, 'v').children[0].text);
              break;
            default:
              // Number value.
              _value = double.parse(findChild(node, 'v').children[0].text);
              break;
          }
          break;
        default:
          _remainingAttributes.add(attr);
      }
    });

    // TODO(WJ): Parse the formula if present..

    // TODO(WJ): If any unknown children are still present, store them for later output.
  }
}

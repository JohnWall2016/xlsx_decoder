import './style_sheet.dart';
import 'package:xml/xml.dart';
import './nodes.dart';

class Color {
  String rgb;
  int theme;
  String tint;

  bool get isEmpty => rgb == null || theme == null || tint == null;

  static const colors = [
    "000000",
    "FFFFFF",
    "FF0000",
    "00FF00",
    "0000FF",
    "FFFF00",
    "FF00FF",
    "00FFFF",
    "000000",
    "FFFFFF",
    "FF0000",
    "00FF00",
    "0000FF",
    "FFFF00",
    "FF00FF",
    "00FFFF",
    "800000",
    "008000",
    "000080",
    "808000",
    "800080",
    "008080",
    "C0C0C0",
    "808080",
    "9999FF",
    "993366",
    "FFFFCC",
    "CCFFFF",
    "660066",
    "FF8080",
    "0066CC",
    "CCCCFF",
    "000080",
    "FF00FF",
    "FFFF00",
    "00FFFF",
    "800080",
    "800000",
    "008080",
    "0000FF",
    "00CCFF",
    "CCFFFF",
    "CCFFCC",
    "FFFF99",
    "99CCFF",
    "FF99CC",
    "CC99FF",
    "FFCC99",
    "3366FF",
    "33CCCC",
    "99CC00",
    "FFCC00",
    "FF9900",
    "FF6600",
    "666699",
    "969696",
    "003366",
    "339966",
    "003300",
    "333300",
    "993300",
    "993366",
    "333399",
    "333333",
    "System Foreground",
    "System Background"
  ];
}

class Style {
  StyleSheet _styleSheet;
  int _id;
  XmlElement _xfNode;
  XmlElement _fontNode;
  XmlElement _fillNode;
  XmlElement _borderNode;

  Style(this._styleSheet, this._id, this._xfNode, this._fontNode,
      this._fillNode, this._borderNode) {}

  int get id => _id;

  XmlElement _findChild(XmlElement node, name) {
    return node.children
        .whereType<XmlElement>()
        .firstWhere((node) => node.name.local == name, orElse: () => null);
  }

  void _setChildAttributes(XmlElement node, String name, Attributes attributes) {
    var child = _findChild(node, name);
    if (child != null) {
      child.attributes.clear();
      attributes.removeEmptyAttributes();
      child.attributes.addAll(attributes.toList());
    }
  }

  void _removeChildIfEmpty(XmlElement node, name) {
    var child = _findChild(node, name);
    if (child != null && child.children.isEmpty && child.attributes.isEmpty)
      node.children.remove(child);
  }

  Color _getColor(XmlElement node, String name) {
    var child = _findChild(node, name);
    if (child == null || child.attributes.isEmpty) return null;

    var color = Color();
    child.attributes.forEach((attr) {
      switch (attr.name.local) {
        case 'rgb':
          color.rgb = attr.value;
          break;
        case 'theme':
          color.theme = int.parse(attr.value);
          break;
        case 'indexed':
          color.rgb = Color.colors[int.parse(attr.value)];
          break;
        case 'tint':
          color.tint = attr.value;
      }
    });
    if (color.isEmpty) return null;

    return color;
  }

  void _setColor(XmlElement node, String name, dynamic color) {
    var clr = Color();
    if (color is String) clr.rgb = color;
    else if (color is int) clr.theme = color;

    _setChildAttributes(node, name, Attributes({
      'rgb': clr.rgb?.toUpperCase(),
      'indexed': null,
      'theme': clr.theme?.toString(),
      'tint': clr.tint
    }));

    _removeChildIfEmpty(node, 'color');
  }
}

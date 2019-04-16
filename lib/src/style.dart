import 'package:xml/xml.dart';

import './style_sheet.dart';
import './nodes.dart';
import './xml_utils.dart';

class Color {
  String rgb;
  int theme;
  String tint;

  bool get isEmpty => rgb == null && theme == null && tint == null;

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

  Color _getColor(XmlElement node, String name) {
    var child = findChild(node, name);
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
    if (color is String)
      color = Color()..rgb = color;
    else if (color is int) color = Color()..theme = color;

    setChildAttributes(
        node,
        name,
        Attributes({
          'rgb': color.rgb?.toUpperCase(),
          'indexed': null,
          'theme': color.theme?.toString(),
          'tint': color.tint
        }));

    removeChildIfEmpty(node, 'color');
  }

  bool get bold => findChild(_fontNode, 'b') != null;

  void set bold(bool value) {
    if (value)
      appendChildIfNotFound(_fontNode, 'b');
    else
      removeChild(_fontNode, 'b');
  }

  get underline {
    var node = findChild(_fontNode, 'u');
    return node == null ? false : getAttribute(node, 'val') ?? true;
  }

  set underline(value) {
    if (value is bool) {
      if (value)
        appendChildIfNotFound(_fontNode, 'u');
      else
        removeChild(_fontNode, 'u');
    } else if (value is String) {
      var node = appendChildIfNotFound(_fontNode, 'u');
      setAttributes(node, Attributes({'val': value}));
    }
  }

  bool get strikethrough => findChild(_fontNode, 'strike') != null;

  void set strikethrough(bool value) {
    if (value)
      appendChildIfNotFound(_fontNode, 'strike');
    else
      removeChild(_fontNode, 'strike');
  }

  String _getFontVerticalAlignment() =>
      getChildAttribute(this._fontNode, 'vertAlign', "val");

  void _setFontVerticalAlignment(String alignment) {
    setChildAttributes(_fontNode, 'vertAlign', Attributes({'val': alignment}));
    removeChildIfEmpty(_fontNode, 'vertAlign');
  }

  bool get subscript => _getFontVerticalAlignment() == "subscript";

  void set subscript(bool value) =>
      _setFontVerticalAlignment(value ? "subscript" : null);

  bool get superscript => _getFontVerticalAlignment() == "superscript";

  void set superscript(bool value) =>
      _setFontVerticalAlignment(value ? "superscript" : null);
}

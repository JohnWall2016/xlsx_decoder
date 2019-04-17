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

class Stop {
  String position;
  Color color;
}

abstract class Fill {}

class SolidFill extends Fill {
  Color color;
}

class PatternFill extends Fill {
  String type;
  Color foreground;
  Color background;
}

class GradientFill extends Fill {
  String type;
  List<Stop> stops;
  String angle;
  String left, right, top, bottom;
}

class Side {
  String style;
  Color color;
  String direction;

  bool get isEmpty => style == null && color == null && direction == null;
}

class Border {
  Side left, right, top, bottom, diagonal;
}

class Style {
  StyleSheet _styleSheet;
  int _id;
  XmlElement _xfNode;
  XmlElement _fontNode;
  XmlElement _fillNode;
  XmlElement _borderNode;

  Style(_styleSheet, _id, _xfNode, _fontNode, _fillNode, _borderNode) {}

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

  /// Set [color] to the child of [node] whose name is [name].
  /// [color] can be a [Color] | [String] | [int] | [null].
  void _setColor(XmlElement node, String name, dynamic color) {
    if (color is String)
      color = Color()..rgb = color;
    else if (color is int) color = Color()..theme = color;

    setChildAttributes(node, name, {
      'rgb': color?.rgb?.toUpperCase(),
      'indexed': null,
      'theme': color?.theme?.toString(),
      'tint': color?.tint
    });

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
      setAttributes(node, {'val': value});
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
      getChildAttribute(_fontNode, 'vertAlign', "val");

  void _setFontVerticalAlignment(String alignment) {
    setChildAttributes(_fontNode, 'vertAlign', {'val': alignment});
    removeChildIfEmpty(_fontNode, 'vertAlign');
  }

  bool get subscript => _getFontVerticalAlignment() == "subscript";

  void set subscript(bool value) =>
      _setFontVerticalAlignment(value ? "subscript" : null);

  bool get superscript => _getFontVerticalAlignment() == "superscript";

  void set superscript(bool value) =>
      _setFontVerticalAlignment(value ? "superscript" : null);

  int get fontSize => int.parse(getChildAttribute(_fontNode, 'sz', "val"));

  void set fontSize(int size) {
    setChildAttributes(_fontNode, 'sz', {'val': '$size'});
    removeChildIfEmpty(_fontNode, 'sz');
  }

  String get fontFamily => getChildAttribute(_fontNode, 'name', "val");

  void set fontFamily(String family) {
    setChildAttributes(_fontNode, 'name', {'val': family});
    removeChildIfEmpty(_fontNode, 'name');
  }

  Color get fontColor => _getColor(_fontNode, "color");

  void set fontColor(color) => _setColor(_fontNode, "color", color);

  String get horizontalAlignment =>
      getChildAttribute(_xfNode, 'alignment', 'horizontal');

  void set horizontalAlignment(String alignment) {
    setChildAttributes(_xfNode, 'alignment', {'horizontal': alignment});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  bool get justifyLastLine =>
      getChildAttribute(_xfNode, 'alignment', 'justifyLastLine') == '1';

  void set justifyLastLine(bool value) {
    setChildAttributes(
        _xfNode, 'alignment', {'justifyLastLine': value ? '1' : null});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  String get indent => getChildAttribute(_xfNode, 'alignment', 'indent');

  void set indent(String value) {
    setChildAttributes(_xfNode, 'alignment', {'indent': value});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  String get verticalAlignment =>
      getChildAttribute(_xfNode, 'alignment', 'vertical');

  void set verticalAlignment(String value) {
    setChildAttributes(_xfNode, 'alignment', {'vertical': value});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  bool get wrapText =>
      getChildAttribute(_xfNode, 'alignment', 'wrapText') == '1';

  void set wrapText(bool value) {
    setChildAttributes(_xfNode, 'alignment', {'wrapText': value ? '1' : null});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  bool get shrinkToFit =>
      getChildAttribute(_xfNode, 'alignment', 'shrinkToFit') == '1';

  void set shrinkToFit(bool value) {
    setChildAttributes(
        _xfNode, 'alignment', {'shrinkToFit': value ? '1' : null});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  String get textDirection {
    var readingOrder = getChildAttribute(_xfNode, 'alignment', 'readingOrder');
    if (readingOrder == '1') return 'left-to-right';
    if (readingOrder == '2') return 'right-to-left';
    return readingOrder;
  }

  void set textDirection(String value) {
    String readingOrder;
    if (value == "left-to-right")
      readingOrder = '1';
    else if (textDirection == "right-to-left") readingOrder = '2';
    setChildAttributes(_xfNode, 'alignment', {'readingOrder': readingOrder});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  double _getTextRotation() =>
      double.parse(getChildAttribute(_xfNode, 'alignment', "textRotation"));

  void _setTextRotation(double textRotation) {
    setChildAttributes(_xfNode, 'alignment', {'textRotation': '$textRotation'});
    removeChildIfEmpty(_xfNode, 'alignment');
  }

  double get textRotation {
    var textRotation = _getTextRotation();

    // Negative angles in Excel correspond to values > 90 in OOXML.
    if (textRotation > 90) textRotation = 90 - textRotation;
    return textRotation;
  }

  void set textRotation(double textRotation) {
    // Negative angles in Excel correspond to values > 90 in OOXML.
    if (textRotation < 0) textRotation = 90 - textRotation;
    _setTextRotation(textRotation);
  }

  bool get angleTextCounterclockwise => _getTextRotation() == 45;

  void set angleTextCounterclockwise(bool value) {
    _setTextRotation(value ? 45 : null);
  }

  bool get angleTextClockwise => _getTextRotation() == 135;

  void set angleTextClockwise(bool value) =>
      _setTextRotation(value ? 135 : null);

  bool get rotateTextUp => _getTextRotation() == 90;

  void set rotateTextUp(bool value) => _setTextRotation(value ? 90 : null);

  bool get rotateTextDown => _getTextRotation() == 180;

  void set rotateTextDown(bool value) => _setTextRotation(value ? 180 : null);

  bool get verticalText => _getTextRotation() == 255;

  void set verticalText(bool value) => _setTextRotation(value ? 255 : null);

  Fill get fill {
    var patternFillNode = findChild(_fillNode, 'patternFill');
    var gradientFillNode = findChild(this._fillNode, 'gradientFill');
    var patternType = getAttribute(patternFillNode, 'patternType');

    if (patternType == "solid") {
      return SolidFill()..color = _getColor(patternFillNode, 'fgColor');
    } else if (patternType != null) {
      return PatternFill()
        ..type = patternType
        ..foreground = _getColor(patternFillNode, 'fgColor')
        ..background = _getColor(patternFillNode, 'bgColor');
    } else if (gradientFillNode != null) {
      var gradientType = getAttribute(gradientFillNode, 'type') ?? 'linear';
      var stops = <Stop>[];
      gradientFillNode.children.forEach((node) {
        if (node is XmlElement) {
          stops.add(Stop()
            ..position = getAttribute(node, 'position')
            ..color = _getColor(node, 'color'));
        }
      });
      var fill = GradientFill()
        ..type = gradientType
        ..stops = stops;

      if (gradientType == "linear") {
        fill.angle = getAttribute(gradientFillNode, 'degree');
      } else {
        fill.left = getAttribute(gradientFillNode, 'left');
        fill.right = getAttribute(gradientFillNode, 'right');
        fill.top = getAttribute(gradientFillNode, 'top');
        fill.bottom = getAttribute(gradientFillNode, 'bottom');
      }

      return fill;
    }

    return null;
  }

  void set fill(Fill fill) {
    _fillNode.children.clear();

    if (fill == null) return;

    if (fill is SolidFill) {
      var patternFill =
          Element('patternFill', Attributes({'patternType': 'solid'}))
              .toXmlNode();
      _setColor(patternFill, 'fgColor', fill.color);
      _fillNode.children.add(patternFill);
      return;
    } else if (fill is PatternFill) {
      var patternFill =
          Element('patternFill', Attributes({'patternType': fill.type}))
              .toXmlNode();
      _setColor(patternFill, 'fgColor', fill.foreground);
      _setColor(patternFill, 'bgColor', fill.background);
      _fillNode.children.add(patternFill);
      return;
    } else if (fill is GradientFill) {
      var gradientFill = Element('gradientFill').toXmlNode();
      setAttributes(gradientFill, {
        'type': fill.type == 'path' ? 'path' : null,
        'left': fill.left,
        'right': fill.right,
        'top': fill.top,
        'bottom': fill.bottom,
        'degree': fill.angle
      });
      fill.stops.forEach((stop) {
        var node = Element('stop', Attributes({'position': stop.position}))
            .toXmlNode();
        _setColor(node, 'color', stop.color);
        gradientFill.children.add(node);
      });
      _fillNode.children.add(gradientFill);
      return;
    }
  }

  Border _getBorder() {
    var border = Border();
    _borderNode.children.forEach((node) {
      if (node is XmlElement) {
        Side getSide() => Side()
            ..style = getAttribute(node, 'style')
            ..color = _getColor(node, 'color');

        switch (node.name.local) {
          case 'left':
            border.left = getSide();
            break;
          case 'right':
            border.right = getSide();
            break;
          case 'top':
            border.top = getSide();
            break;
          case 'bottom':
            border.bottom = getSide();
            break;
          case 'diagonal':
            var up = getAttribute(_borderNode, 'diagonalUp');
            var down = getAttribute(_borderNode, 'diagonalDown');
            String direction;
            if (up != null && down != null)
              direction = 'both';
            else if (up != null)
              direction = 'up';
            else if (down != null) direction = 'down';
            border.diagonal = getSide()..direction = direction;
            break;
        }
      }
    });

    return border;
  }

  void _setBorder(Border border) {
    border ??= Border();
    ({
      'left': border.left,
      'right': border.right,
      'top': border.top,
      'bottom': border.bottom,
      'diagonal': border.diagonal
    }).forEach((name, side) {
      var node = findChild(_borderNode, name);
      if (node != null) return;
      
      setAttributes(node, { 'style': side.style });
      _setColor(node, 'color', side.color);

      if (name == 'diagonal') {
        setAttributes(node, {
          'diagonalUp': side.direction == 'up' || side.direction == 'both' ? '1' : null,
          'diagonalDown': side.direction == 'down' || side.direction == 'both' ? '1' : null
        });
      }
    });
  }

  Border get border => _getBorder();

  void set border(Border border) => _setBorder(border);

  
}

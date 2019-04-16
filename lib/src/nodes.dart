import 'package:xml/xml.dart';

abstract class Node {
  XmlNode toXmlNode();
}

class Text extends Node {
  String _text;

  Text(this._text);

  @override
  XmlNode toXmlNode() => XmlText(_text);
}

class Attributes {
  Map<String, String> _attributes;

  operator []=(key, value) => _attributes[key] = value;

  Attributes([this._attributes]) {
    if (_attributes == null) _attributes = {};
  }

  List<XmlAttribute> toList() {
    var list = <XmlAttribute>[];
    _attributes.entries.forEach((entry) {
      var attr = XmlAttribute(XmlName(entry.key), entry.value);
      list.add(attr);
    });
    return list;
  }

  void removeEmptyAttributes() {
    _attributes.removeWhere(
        (k, v) => k == null || v == null || k.isEmpty || v.isEmpty);
  }
}

class Nodes {
  List<Node> _children;

  Nodes([this._children]) {
    if (_children == null) _children = [];
  }

  void add(Node node) => _children.add(node);

  List<XmlNode> toList() => _children.map((e) => e.toXmlNode()).toList();
}

class Element extends Node {
  String _name;
  Attributes _attributes = Attributes();
  Nodes _children = Nodes();

  Element(this._name);

  Attributes get attributes => _attributes;

  Nodes get children => _children;

  XmlElement toXmlNode() =>
      XmlElement(XmlName(_name), _attributes.toList(), _children.toList());
}

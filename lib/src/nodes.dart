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

class Element extends Node {
  String _name;
  Map<String, String> _attributes = {};
  List<Node> _children = [];

  Element(this._name);

  void addAttribute(String key, String value) => _attributes[key] = value;

  void addChild(Node child) => _children.add(child);

  XmlElement toXmlNode() {
    var node = XmlElement(XmlName(_name));

    _attributes.entries.forEach((entry) {
      var attr = XmlAttribute(XmlName(entry.key), entry.value);
      node.attributes.add(attr);
    });

    _children.forEach((e) {
      node.children.add(e.toXmlNode());
    });

    return node;
  }
}

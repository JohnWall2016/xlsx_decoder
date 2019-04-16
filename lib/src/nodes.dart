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

  void operator []=(String key, String value) => _attributes[key] = value;

  String operator [](String key) => _attributes[key];

  Attributes([this._attributes]) {
    if (_attributes == null) _attributes = {};
  }

  Iterable<XmlAttribute> toXml() => _attributes.entries
      .map((entry) => XmlAttribute(XmlName(entry.key), entry.value));

  Iterable<String> get keys => _attributes.keys;

  bool containKey(String key) => _attributes.containsKey(key);
}

class NodeList {
  List<Node> _children;

  NodeList([this._children]) {
    if (_children == null) _children = [];
  }

  void add(Node node) => _children.add(node);

  Iterable<XmlNode> toXml() => _children.map((e) => e.toXmlNode());
}

class Element extends Node {
  String _name;
  Attributes _attributes = Attributes();
  NodeList _children = NodeList();

  Element(this._name);

  Attributes get attributes => _attributes;

  NodeList get children => _children;

  XmlElement toXmlNode() =>
      XmlElement(XmlName(_name), _attributes.toXml(), _children.toXml());
}

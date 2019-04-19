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

class Attributes<T> {
  Map<String, T> _attributes;

  void operator []=(String key, T value) => _attributes[key] = value;

  String operator [](String key) => _attributes[key].toString();

  Attributes([this._attributes]) {
    if (_attributes == null) _attributes = {};
  }

  Iterable<XmlAttribute> toXml() => _attributes.entries
      .map((entry) => XmlAttribute(XmlName(entry.key), entry.value.toString()));

  Iterable<String> get keys => _attributes.keys;

  bool containsKey(String key) => _attributes.containsKey(key);
}

class NodeList {
  List<Node> _children;

  NodeList([this._children]) {
    if (_children == null) _children = [];
  }

  void add(Node node) => _children.add(node);

  void addAll(Iterable<Node> nodes) => _children.addAll(nodes);

  Iterable<XmlNode> toXml() => _children.map((e) => e.toXmlNode());
}

class Element extends Node {
  String _name;
  Attributes _attributes = Attributes();
  NodeList _children = NodeList();

  Element(String name, [Map<String, String> attributes]) {
    _name = name;
    _attributes = Attributes(attributes);
  }

  Attributes get attributes => _attributes;

  NodeList get children => _children;

  XmlElement toXmlNode() =>
      XmlElement(XmlName(_name), _attributes.toXml(), _children.toXml());
}

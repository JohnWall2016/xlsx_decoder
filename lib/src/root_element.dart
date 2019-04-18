import 'package:xml/xml.dart';
export 'package:xml/xml.dart';

class RootElement {
  XmlElement _root;

  RootElement(this._root);

  XmlElement get root => _root;

  List<XmlAttribute> get attributes => _root.attributes;

  Iterable<XmlElement> get elements => _root.children.whereType<XmlElement>();

  void addNode(XmlNode node) => _root.children.add(node);

  void insertNode(int index, XmlNode node) =>
      _root.children.insert(index, node);
}

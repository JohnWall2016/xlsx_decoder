import 'package:xml/xml.dart';
export 'package:xml/xml.dart';

class AttachedXmlElement {
  XmlElement _node;

  AttachedXmlElement(this._node);

  XmlElement get thisNode => _node;

  List<XmlAttribute> get attributes => _node.attributes;

  Iterable<XmlElement> get elements => _node.children.whereType<XmlElement>();

  void addNode(XmlNode node) => _node.children.add(node);

  void insertNode(int index, XmlNode node) =>
      _node.children.insert(index, node);
}

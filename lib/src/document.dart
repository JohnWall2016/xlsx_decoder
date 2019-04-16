import 'package:xml/xml.dart';
export 'package:xml/xml.dart';

abstract class Document {
  XmlDocument _document;

  void load(XmlDocument document) => _document = document;

  Iterable<XmlElement> get elements =>
      _document.rootElement.children.whereType<XmlElement>();

  XmlDocument get document => _document;

  void addNode(XmlNode node) => _document.rootElement.children.add(node);

  void insertNode(int index, XmlNode node) =>
      _document.rootElement.children.insert(index, node);

  String get id;
}

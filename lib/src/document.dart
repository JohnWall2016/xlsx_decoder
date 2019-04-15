import 'package:xml/xml.dart';
export 'package:xml/xml.dart';

class Document {
  XmlDocument _document;

  Document(this._document);

  Iterable<XmlElement> get elements =>
      _document.rootElement.children.whereType<XmlElement>();

  XmlDocument get document => _document;

  void addNode(XmlNode node) => _document.rootElement.children.add(node);
}

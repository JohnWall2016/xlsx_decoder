import 'package:xml/xml.dart';

class CoreProperties {
  XmlDocument _document;
  Map<String, String> _properties = {};

  CoreProperties(this._document);
}
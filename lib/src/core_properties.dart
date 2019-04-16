import './document.dart';
import './nodes.dart';


class CoreProperties extends Document {
  @override
  String get id => 'docProps/core.xml';

  static const allowedProperties = {
    "title": "dc:title",
    "subject": "dc:subject",
    "author": "dc:creator",
    "creator": "dc:creator",
    "description": "dc:description",
    "keywords": "cp:keywords",
    "category": "cp:category"
  };

  Map<String, String> _properties = {};

  void operator []=(String name, String value) {
    var key = name.toLowerCase();

    if (!allowedProperties.containsKey(key)) {
      throw 'Unknown property name: "$name"';
    }

    var eName = allowedProperties[key];
    if (_properties.containsKey(key)) {
      var element = super.elements.firstWhere((e) => e.name == XmlName(eName));
      if (element != null) {
        element.children
          ..clear()
          ..add(XmlText(value));
      }
    } else {
      _properties[key] = value;
      super.addNode((Element(eName)..addChild(Text(value))).toXmlNode());
    }
  }

  String operator [](String name) {
    var key = name.toLowerCase();

    if (!allowedProperties.containsKey(key)) {
      throw 'Unknown property name: "$name"';
    }

    return _properties[key];
  }

}

/*
docProps/core.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Title</dc:title>
<dc:subject>Subject</dc:subject>
<dc:creator>Creator</dc:creator>
<cp:keywords>Keywords</cp:keywords>
<dc:description>Description</dc:description>
<cp:category>Category</cp:category>
</cp:coreProperties>
 */

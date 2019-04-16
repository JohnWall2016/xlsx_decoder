import 'package:xml/xml.dart';

import './nodes.dart';

XmlElement findChild(XmlElement node, String name) {
  return node.children
      .whereType<XmlElement>()
      .firstWhere((node) => node.name.local == name, orElse: () => null);
}

void setChildAttributes(XmlElement node, String name, Attributes attributes) {
  var child = findChild(node, name);
  if (child == null) {
    child = Element(name).toXmlNode();
    node.children.add(child);
  }
  List<XmlAttribute> xmlAttrs = [];
  child.attributes.forEach((xmlAttr) {
    if (!attributes.containKey(xmlAttr.name.local)) {
      xmlAttrs.add(xmlAttr);
    } else {
      var value = attributes[xmlAttr.name.local];
      if (value != null) {
        xmlAttr.value = value;
        xmlAttrs.add(xmlAttr);
      }
    }
  });
  child.attributes.clear();
  child.attributes.addAll(xmlAttrs);
}

void removeChildIfEmpty(XmlElement node, String name) {
  var child = findChild(node, name);
  if (child != null && isEmpty(child)) node.children.remove(child);
}

bool isEmpty(XmlElement node) =>
    node.children.isEmpty && node.attributes.isEmpty;

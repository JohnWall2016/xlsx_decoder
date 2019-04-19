import 'package:xml/xml.dart';

import './nodes.dart';

XmlElement findChild(XmlNode node, String name) {
  return node.children
      .whereType<XmlElement>()
      .firstWhere((node) => node.name.local == name, orElse: () => null);
}

void setAttributes<T>(XmlElement node, Map<String, T> attributes) {
  List<XmlAttribute> xmlAttrs = [];
  node.attributes.forEach((xmlAttr) {
    if (!attributes.containsKey(xmlAttr.name.local)) {
      xmlAttrs.add(xmlAttr);
    } else {
      var value = attributes[xmlAttr.name.local];
      if (value != null) {
        xmlAttr.value = value.toString();
        xmlAttrs.add(xmlAttr);
      }
    }
  });
  node.attributes.clear();
  node.attributes.addAll(xmlAttrs);
}

void setChildAttributes<T>(XmlElement node, String name, Map<String, T> attributes) {
  var child = appendChildIfNotFound(node, name);
  setAttributes<T>(child, attributes);
}

T getAttribute<T>(XmlElement node, String attribute) {
  var attr = node?.attributes?.firstWhere(
      (xmlAttr) => xmlAttr.name.local == attribute,
      orElse: () => null);
  if (attr != null && attr.value != null) {
    if (T == int) return int.tryParse(attr.value) as T;
    else if (T == double) return double.tryParse(attr.value) as T;
    else if (T == String) return attr.value as T;
    else if (T == Object || T == dynamic) {
      return attr.value as T;
    }
  }
  return null;
}

T getChildAttribute<T>(XmlElement node, String name, String attribute) {
  var child = findChild(node, name);
  if (child != null) return getAttribute<T>(child, attribute);
  return null;
}

void removeChildIfEmpty(XmlElement node, String name) {
  node.children.removeWhere(
      (node) => node is XmlElement && node.name.local == name && isEmpty(node));
}

bool isEmpty(XmlElement node) =>
    node.children.isEmpty && node.attributes.isEmpty;

XmlElement appendChildIfNotFound(XmlElement node, String name) {
  var child = findChild(node, name);
  if (child == null) {
    child = Element(name).toXmlNode();
    node.children.add(child);
  }
  return child;
}

void removeChild(XmlElement node, String name) {
  node.children
      .removeWhere((node) => node is XmlElement && node.name.local == name);
}

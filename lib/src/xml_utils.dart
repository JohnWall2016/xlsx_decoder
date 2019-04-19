import 'package:xml/xml.dart';

import './nodes.dart';

XmlElement findChild(XmlNode node, String name) {
  return node.children
      .whereType<XmlElement>()
      .firstWhere((node) => node.name.local == name, orElse: () => null);
}

int findChildIndex(XmlNode node, String name) {
  for (var i = 0; i < node.children.length; i++) {
    var child = node.children[i];
    if (child is XmlElement && child.name.local == name) {
      return i;
    }
  }
  return -1;
}

void setAttributes<T>(XmlElement node, Map<String, T> attributes) {
  var attrs = Map.from(attributes);
  List<XmlAttribute> xmlAttrs = [];
  node.attributes.forEach((xmlAttr) {
    var key = xmlAttr.name.local;
    if (!attrs.containsKey(key)) {
      xmlAttrs.add(xmlAttr);
    } else {
      var value = attrs[key];
      if (value != null) {
        xmlAttr.value = value.toString();
        xmlAttrs.add(xmlAttr);
      }
      attrs.remove(key);
    }
  });
  node.attributes.clear();
  node.attributes.addAll(xmlAttrs);
  node.attributes.addAll(Attributes(attrs).toXml());
}

void setChildAttributes<T>(
    XmlElement node, String name, Map<String, T> attributes) {
  var child = appendChildIfNotFound(node, name);
  setAttributes<T>(child, attributes);
}

T getAttribute<T>(XmlElement node, String attribute) {
  if (node == null) return null;

  var attr = node.attributes.firstWhere(
      (xmlAttr) => xmlAttr.name.local == attribute,
      orElse: () => null);
  if (attr != null && attr.value != null) {
    if (T == int)
      return int.tryParse(attr.value) as T;
    else if (T == double)
      return double.tryParse(attr.value) as T;
    else if (T == String)
      return attr.value as T;
    else if (T == Object || T == dynamic) {
      return attr.value as T;
    }
  }
  return null;
}

XmlAttribute toXmlAttribute<T>(String key, T value) =>
    XmlAttribute(XmlName(key), value.toString());

void setAttribute<T>(XmlElement node, String attribute, T value) {
  if (node == null) return;
  var attr = node.attributes.firstWhere(
      (xmlAttr) => xmlAttr.name.local == attribute,
      orElse: () => null);
  if (attr != null)
    attr.value = value.toString();
  else
    node.attributes.add(toXmlAttribute(attribute, value));
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

/// [String]|[XmlNode] [name]
void removeChild(XmlElement node, dynamic nameOrNode) {
  if (nameOrNode is String) {
    node.children.removeWhere(
        (node) => node is XmlElement && node.name.local == nameOrNode);
  } else if (nameOrNode is XmlNode) {
    node.children.remove(nameOrNode);
  }
}

void insertInOrder(XmlElement node, XmlElement child, List<String> nodeOrder) {
  var index = nodeOrder.indexOf(child.name.local);
  if (index >= 0) {
    for (var i = index + 1; i < nodeOrder.length; i++) {
      var name = nodeOrder[i];
      var siblingIndex = findChildIndex(node, name);
      if (siblingIndex >= 0) {
        node.children.insert(siblingIndex, child);
        return;
      }
    }
  }
  node.children.add(child);
}

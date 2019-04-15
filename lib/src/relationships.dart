import './document.dart';

class Relationships extends Document {
  @override
  String get id => 'xl/_rels/workbook.xml.rels';
  
  int _nextId = 1;

  @override
  void load(XmlDocument document) {
    super.load(document ?? parse(emptyXml));
    elements.forEach((node) {
      var id = int.parse(node.getAttribute('Id').substring(3));
      if (id >= _nextId) _nextId = id + 1;
    });
  }

  XmlNode add(String type, String target, [String targetMode]) {
    var node = XmlElement(XmlName('Relationship'), [
      XmlAttribute(XmlName('Id'), 'rId${_nextId++}'),
      XmlAttribute(XmlName('Type'), '$relationshipSchemaPrefix$type'),
      XmlAttribute(XmlName('Target'), target)
    ]);
    if (targetMode != null)
      node.attributes.add(XmlAttribute(XmlName('TargetMode'), targetMode));
    addNode(node);
    return node;
  }

  XmlNode findById(String id) => elements
      .firstWhere((node) => node.getAttribute('Id') == id, orElse: () => null);

  XmlNode findByType(String type) => elements.firstWhere(
      (node) => node.getAttribute('Type') == '$relationshipSchemaPrefix$type',
      orElse: () => null);

  static const relationshipSchemaPrefix =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

  static const emptyXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""";
}

/*
xl/_rels/workbook.xml.rels

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
*/

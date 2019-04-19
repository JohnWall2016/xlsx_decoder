import 'package:xlsx_decoder/src/relationships.dart';
import 'package:xlsx_decoder/xlsx_decoder.dart';
import 'package:xml/xml.dart';
import 'package:xlsx_decoder/src/xml_utils.dart';

void main(List<String> args) {
  testRelationships();

  var workbook = Workbook.fromFile(args[0]);
  testSharedString(workbook);
  //testXml();
}

void testCoreProperties(Workbook workbook) {
  print(workbook.coreProperties.thisNode);
  workbook.coreProperties['keywords'] = 'XLSX';
  print(workbook.coreProperties.thisNode);
  workbook.coreProperties['keywords'] = 'DART';
  print(workbook.coreProperties.thisNode);
}

void testRelationships() {
  var rels = Relationships(null);
  rels.add('styles', 'styles.xml');
  rels.add('theme', 'theme1.xml', 'UNKNOWN');
  print(rels.thisNode);
  print(rels.findById('rId1'));
  print(rels.findByType('theme'));
}

void testSharedString(Workbook workbook) {
  //print(workbook.sharedStrings.document);
  //print(workbook.sharedStrings.getStringByIndex(0));
  workbook.sharedStrings.getIndexForString('刘德华');
  print(workbook.sharedStrings.thisNode.toXmlString(pretty: false));
}

void testXml() {
  var node = parse(
'''
<c r="A7" s="8" t="s">
  <v>21</v>
</c>
''').rootElement;
  print(node);
  print(findChild(node, 'v').children[0].text);
  double s = getAttribute(node, 's');
  print(s);
  String t = getAttribute(node, 't');
  print(t);
  print(getAttribute(node, 't') == null);
  getAttribute(node, 't');
  print(getAttribute(node, 't'));
}

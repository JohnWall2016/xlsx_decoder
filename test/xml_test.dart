import 'package:xlsx_decoder/src/relationships.dart';
import 'package:xlsx_decoder/xlsx_decoder.dart';

void main(List<String> args) {
  testRelationships();

  var workbook = Workbook.fromFile(args[0]);
  testSharedString(workbook);
}

void testCoreProperties(Workbook workbook) {
  print(workbook.coreProperties.root);
  workbook.coreProperties['keywords'] = 'XLSX';
  print(workbook.coreProperties.root);
  workbook.coreProperties['keywords'] = 'DART';
  print(workbook.coreProperties.root);
}

void testRelationships() {
  var rels = Relationships(null);
  rels.add('styles', 'styles.xml');
  rels.add('theme', 'theme1.xml', 'UNKNOWN');
  print(rels.root);
  print(rels.findById('rId1'));
  print(rels.findByType('theme'));
}

void testSharedString(Workbook workbook) {
  //print(workbook.sharedStrings.document);
  //print(workbook.sharedStrings.getStringByIndex(0));
  workbook.sharedStrings.getIndexForString('刘德华');
  print(workbook.sharedStrings.root.toXmlString(pretty: false));
}
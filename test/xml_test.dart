import 'package:xlsx_decoder/src/relationships.dart';
import 'package:xlsx_decoder/xlsx_decoder.dart';

void main(List<String> args) {
  //testRelationships();

  var workbook = Workbook.fromFile(args[0]);
  /*print(workbook.contentTypes.document);
  print(workbook.appProperties.document);
  print(workbook.coreProperties.document);
  print(workbook.relationships.document);
  print(workbook.sharedStrings.document);
  print(workbook.styleSheet.document);
  print(workbook.document);*/
  print(workbook.coreProperties.document);
  workbook.coreProperties['keywords'] = 'XLSX';
  print(workbook.coreProperties.document);
  workbook.coreProperties['keywords'] = 'DART';
  print(workbook.coreProperties.document);

  print(workbook.coreProperties.document.rootElement.children.toString());
}

void testRelationships() {
  var rels = Relationships()..load(null);
  rels.add('styles', 'styles.xml');
  rels.add('theme', 'theme1.xml', 'UNKNOWN');
  print(rels.document);
  print(rels.findById('rId1'));
  print(rels.findByType('theme'));
}
import 'package:xlsx_decoder/src/relationships.dart';
import 'package:xlsx_decoder/xlsx_decoder.dart';
import 'package:xml/xml.dart';
import 'package:xlsx_decoder/src/xml_utils.dart';
import 'package:xlsx_decoder/src/address_converter.dart';

void main(List<String> args) {
  //testRelationships();
  //testAddressConvert();

  var workbook = Workbook.fromFile(args[0]);
  //testSharedString(workbook);
  //testXml();
  //testValue(workbook);
  testToXml(workbook);
}

void testToXml(Workbook workbook) {
  var sheet = workbook.sheetAt(0);
  print(sheet.cell('A1').toXml());
  print(sheet.rowAt(1).toXml());
  print(sheet.toXmls()['sheet']);

  sheet.cell('A1').setValue('刘德华的演唱会2');
  //sheet.insertRow(3);
  sheet.insertRowCopyFrom(5, 4);
  sheet.rowAt(4).cell('B').setValue('刘德华1');
  sheet.rowAt(5).cell('B').setValue('刘德华2');
  workbook.toFile('e:/aaa.xlsx');
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

void testValue(Workbook workbook) {
  var sheet = workbook.sheetAt(0);
  String title = sheet.cell('A1').value();
  print(title);
  String name = sheet.cell('B4').value();
  print(name);
  double money = sheet.cell('E4').value();
  print(money);
  String date = sheet.cell('F4').value();
  print(date);
  int date2 = sheet.cell('F4').value();
  print(date2);
  double date3 = sheet.cell('F4').value();
  print(date3);
  print(sheet.cell('J6').value());
  print(sheet.cell('X12').value());

  print(sheet.lastRowIndex);

  for (int i = 1; i <= sheet.lastRowIndex; i++) {
    print(sheet.rowAt(i).cellAt(1).value());
  }
}

testAddressConvert() {
  print(columnNameToNumber('AB'));
  print(columnNumberToName(27));
  print(CellRef.fromAddress('A1'));
  print(CellRef.fromAddress('AB27'));

  var cell = CellRef.fromAddress('A1');
  print(cell.row); print(cell.column); print(cell.toAddress());

  var range = RangeRef.fromAddress('AB27:AC33');
  print(range); print(range.toAddress());
  range = RangeRef.fromAddress(r'$AB27:$AC33');
  print(range); print(range.toAddress());
}
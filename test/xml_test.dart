import 'package:xml/xml.dart';
import 'package:xlsx_decoder/src/relationships.dart';

void main(List<String> args) {
  var rels = Relationships()..load(null);
  rels.add('styles', 'styles.xml');
  rels.add('theme', 'theme1.xml', 'UNKNOWN');
  print(rels.document);
  print(rels.findById('rId1'));
  print(rels.findByType('theme'));
}
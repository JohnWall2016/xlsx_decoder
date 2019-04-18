import 'package:xlsx_decoder/src/address_converter.dart';

void main(List<String> args) {
  print(columnNameToNumber('AB'));
  print(columnNumberToName(27));
  print(CellRef.fromAddress('A1'));
  print(CellRef.fromAddress('AB27'));
}
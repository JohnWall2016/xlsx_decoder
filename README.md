A library for reading and writing a xlsx file.

## Usage

A simple usage example:

```dart
import 'package:xlsx_decoder/xlsx_decoder.dart';

main() {
  // load a xlsx file.
  var workbook = Workbook.fromFile('a.xlsx');
  var sheet = workbook.sheetAt(0);
  
  // read a cell's value.
  String s = sheet.cell('A1').value();
  print(s);
  int i = sheet.rowAt(2).cell('B').value(); // B2
  print(i);

  // write a cell's value.
  sheet.cell('A1').setValue('a string');
  sheet.rowAt(2).cell('B').setValue(100);
  
  // save to another xlsx file.
  workbook.toFile('b.xlsx');
}
```
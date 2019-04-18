class CellRef {
  int row;
  int column;
  String columnName;

  CellRef();

  factory CellRef.fromAddress(String address) {
    var match = RegExp(r'([A-Z]+)(\d+)').firstMatch(address);
    if (match != null) {
      return CellRef()
        ..columnName = match.group(1)
        ..column = columnNameToNumber(match.group(1))
        ..row = int.parse(match.group(2));
    }
    return null;
  }

  String toString() => 'CellRef { row: $row, column: $column, columnName: $columnName }';
}

int columnNameToNumber(String name) {
  name = name.toUpperCase();
  var sum = 0;
  for (var i = 0; i < name.length; i++) {
    sum = sum * 26;
    sum = sum + (name.codeUnitAt(i) - 65 + 1); // A
  }
  return sum;
}

String columnNumberToName(int number) {
  var dividend = number;
  var name = '';
  var modulo = 0;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    dividend = (dividend - modulo) ~/ 26;
  }

  return name;
}

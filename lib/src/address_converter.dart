const _cellRegex = r'(\$?)([A-Z]+)(\$?)(\d+)';
const _rangeRegex = '($_cellRegex):($_cellRegex)';

class CellRef {
  int _row;
  int get row => _row;
  void set row(int r) => _row = r;

  bool _rowAnchored;

  int _column;
  int get column => _column;

  bool _columnAnchored;
  String _columnName;

  CellRef(this._row, int this._column,
      {bool anchored = false,
      bool rowAnchored = false,
      bool columnAnchored = false}) {
    _rowAnchored = anchored || rowAnchored;
    _columnAnchored = anchored || columnAnchored;

    _columnName = columnNumberToName(_column);
  }

  CellRef._new();

  factory CellRef.fromAddress(String address) {
    var match = RegExp(_cellRegex).firstMatch(address);
    if (match != null) {
      return CellRef._new()
        .._columnAnchored = match.group(1).isNotEmpty
        .._columnName = match.group(2)
        .._column = columnNameToNumber(match.group(2))
        .._rowAnchored = match.group(3).isNotEmpty
        .._row = int.parse(match.group(4));
    }
    return null;
  }

  String toAddress() {
    String address = '';
    if (_columnAnchored) address += '\$';
    address += _columnName;
    if (_rowAnchored) address += '\$';
    address += '$_row';
    return address;
  }

  String toString() => 'CellRef { row: $_row, column: $column }';
}

class RangeRef {
  CellRef _start;
  CellRef get start => _start;

  CellRef _end;
  CellRef get end => _end;

  RangeRef(int startRow, int startColumn, int endRow, int endColumn,
      {bool anchored = false,
      bool startRowAnchored = false,
      bool startColumnAnchored = false,
      bool endRowAnchored = false,
      bool endColumnAnchored = false}) {
    _start = CellRef(startRow, startColumn,
        anchored: anchored,
        rowAnchored: startRowAnchored,
        columnAnchored: startColumnAnchored);
    _end = CellRef(endRow, endColumn,
        anchored: anchored,
        rowAnchored: endRowAnchored,
        columnAnchored: endColumnAnchored);
  }

  RangeRef._new();

  factory RangeRef.fromAddress(String address) {
    var match = RegExp(_rangeRegex).firstMatch(address);
    if (match != null) {
      return RangeRef._new()
        .._start = CellRef.fromAddress(match.group(1))
        .._end = CellRef.fromAddress(match.group(6));
    }
    return null;
  }

  String toAddress() => _start.toAddress() + ':' + _end.toAddress();

  String toString() => 'RangeRef { start: ${start}, end: ${end} }';
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

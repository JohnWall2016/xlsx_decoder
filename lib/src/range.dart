import 'dart:math' as math;

import './sheet.dart';
import './cell.dart';
import './address_converter.dart';

class Range {
  Cell _startCell;
  Cell _endCell;

  Sheet get sheet => _startCell.sheet;

  int _minRowIndex, _maxRowIndex;
  int _minColumnIndex, _maxColumnIndex;
  int _numRows, _numColumns;

  Range(this._startCell, this._endCell) {
    _findRangeExtent();
  }

  void _findRangeExtent() {
    _minRowIndex = math.min(_startCell.rowIndex, _endCell.rowIndex);
    _maxRowIndex = math.max(_startCell.rowIndex, _endCell.rowIndex);
    _minColumnIndex = math.min(_startCell.columnIndex, _endCell.columnIndex);
    _maxColumnIndex = math.max(_startCell.columnIndex, _endCell.columnIndex);
    _numRows = _maxRowIndex - _minRowIndex + 1;
    _numColumns = _maxColumnIndex - _minColumnIndex + 1;
  }

  String address(
      {bool anchored = false,
      bool startRowAnchored = false,
      bool startColumnAnchored = false,
      bool endRowAnchored = false,
      bool endColumnAnchored = false,
      bool includeSheetName = false}) {
    var range = RangeRef(_startCell.rowIndex, _startCell.columnIndex,
        _endCell.rowIndex, _endCell.columnIndex,
        anchored: anchored,
        startRowAnchored: startRowAnchored,
        startColumnAnchored: startColumnAnchored,
        endRowAnchored: endRowAnchored,
        endColumnAnchored: endColumnAnchored);
    var address = '';
    if (includeSheetName) {
      address += "'${sheet.name.replaceAll("'", "''")}'!";
    }
    address += range.toAddress();
    return address;
  }
}



class FormulaError {
  final String _error;

  FormulaError(this._error);

  String get error => _error;

  static final DIV0 = FormulaError("#DIV/0!");
  static final NA = FormulaError("#N/A");
  static final NAME = FormulaError("#NAME?");
  static final NULL = FormulaError("#NULL!");
  static final NUM = FormulaError("#NUM!");
  static final REF = FormulaError("#REF!");
  static final VALUE = FormulaError("#VALUE!");

  static FormulaError getError(String error) {
    for (var e in [DIV0, NA, NAME, NULL, NUM, REF, VALUE]) {
      if (e.error == error) return e;
    }
    return FormulaError(error);
  }
}
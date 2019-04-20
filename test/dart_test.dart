void main(List<String> args) {
  /*var address = r'$A$1';
  var match = RegExp(r'(\$?)([A-Z]+)(\$?)(\d+)').firstMatch(address);
  print(match.group(1).isEmpty);
  print(match.group(2));
  print(match.group(3));
  print(match.group(4));*/
  /*bool a = null;
  print(a);
  if (a ?? false) print(true);
  else print(false);*/
  /*var s = 'aaa';
  print('$s');*/

  const _cellRegex = r'(\$?)([A-Z]+)(\$?)(\d+)';
  const _rangeRegex = '($_cellRegex):($_cellRegex)';

  var address = 'A1:B2';
  var match = RegExp(_rangeRegex).firstMatch(address);
  print(match.group(1));
  print(match.group(6));
}
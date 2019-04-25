import 'package:xlsx_decoder/src/splay_tree.dart';

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

  /*const _cellRegex = r'(\$?)([A-Z]+)(\$?)(\d+)';
  const _rangeRegex = '($_cellRegex):($_cellRegex)';

  var address = 'A1:B2';
  var match = RegExp(_rangeRegex).firstMatch(address);
  print(match.group(1));
  print(match.group(6));*/
  //LinkedList
  //int
  //SplayTreeMap

  /*
  num i = 1.1;
  int j = i as int;
  print(j);
  */

  var map = SplayTreeMap();
  map[2] = 'two';
  map[1] = 'one';
  map[100] = 'one hundred';
  map[50] = 'fifty';

  map.keys.forEach((k) => print(k));
  map.values.forEach((v) => print(v));
  map.nodesFrom(1, (node) {
    node.key += 1;
    node.value += 'new';
  });
  map[1] = 'one';
  map.entries.forEach((e) => print(e));
}

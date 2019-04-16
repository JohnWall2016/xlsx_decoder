import './document.dart';
import './nodes.dart';

class SharedStrings extends Document {
  @override
  String get id => 'xl/sharedStrings.xml';

  List<dynamic> _stringArray = [];

  Map<dynamic, int> _indexMap = {};

  @override
  void load(XmlDocument document) {
    super.load(document ?? parse(emptyXml));
    document.rootElement.attributes.removeWhere((attr) {
      return attr.name == XmlName('count') ||
          attr.name == XmlName('uniqueCount');
    });

    _cacheExistingSharedStrings();
  }

  int getIndexForString(string) {
    var key = string.toString();
    var index = _indexMap[key];
    if (index != null) return index;

    index = _stringArray.length;
    _stringArray.add(string);
    _indexMap[key] = index;

    var element = Element('si');
    if (string is String) {
      var node = (element
            ..addChild(Element('t')
              ..addAttribute('xml:space', 'preserve')
              ..addChild(Text(string))))
          .toXmlNode();
      addNode(node);
    } else {
      var node = element.toXmlNode()..children.add(string);
      addNode(node);
    }
    return index;
  }

  String getStringByIndex(int index) => _stringArray[index] as String;

  void _cacheExistingSharedStrings() {
    var i = 0;
    elements.forEach((node) {
      var content = node.children[0];
      if (content is XmlElement) {
        if (content.name == XmlName('t')) {
          var string = content.children[0].text;
          _stringArray.add(string);
          _indexMap[string] = i++;
        } else {
          // TODO(wj): A dirty hack
          var string = '';
          node.children.forEach((cnode) {
            if (cnode is XmlElement && cnode.name == XmlName('r')) {
              cnode.children.forEach((ccnode) {
                if (ccnode is XmlElement && ccnode.name == XmlName("t"))
                  string += ccnode.text;
              });
            }
          });
          if (string.isNotEmpty) {
            _stringArray.add(string);
            _indexMap[string] = i++;
          } else {
            // TODO: Properly support rich text nodes in the future. For now just store the object as a placeholder.
            _stringArray.add(node.children);
            _indexMap[node.children.toString()] = i++;
          }
        }
      }
    });
  }

  static const emptyXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>""";
}

/*
xl/sharedStrings.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="13" uniqueCount="4">
	<si>
		<t>Foo</t>
	</si>
	<si>
		<t>Bar</t>
	</si>
	<si>
		<t>Goo</t>
	</si>
	<si>
		<r>
			<t>s</t>
		</r><r>
			<rPr>
				<b/>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr><t>d;</t>
		</r><r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr><t>lfk;l</t>
		</r>
	</si>
</sst>
*/

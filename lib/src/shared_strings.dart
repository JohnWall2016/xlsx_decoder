import './attached_xml_element.dart';
import './nodes.dart';

class SharedStrings extends AttachedXmlElement {
  SharedStrings(XmlElement node)
      : super(node ??
            Element('sst', {
              'xmlns':
                  'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            })) {
    attributes.removeWhere((attr) {
      return attr.name.local == 'count' || attr.name.local == 'uniqueCount';
    });

    _cacheExistingSharedStrings();
  }

  List<dynamic> _nodeArray = [];

  Map<dynamic, int> _indexMap = {};

  int getIndexForString(string) {
    var key = string.toString();
    var index = _indexMap[key];
    if (index != null) return index;

    index = _nodeArray.length;
    _nodeArray.add(string);
    _indexMap[key] = index;

    var element = Element('si');
    if (string is String) {
      var node = (element
            ..children.add(Element('t')
              ..attributes['xml:space'] = 'preserve'
              ..children.add(Text(string))))
          .toXmlNode();
      addNode(node);
    } else {
      var node = element.toXmlNode()..children.addAll(string);
      addNode(node);
    }
    return index;
  }

  String getStringByIndex(int index) => _nodeArray[index].toString();

  void _cacheExistingSharedStrings() {
    var i = 0;
    elements.forEach((node) {
      var content = node.children[0];
      if (content is XmlElement) {
        if (content.name.local == 't') {
          if (content.children.isNotEmpty) {
            var string = content.children[0].text;
            _nodeArray.add(string);
            _indexMap[string] = i++;
          } else {
            _nodeArray.add('');
            _indexMap[''] = i++;
          }
        } else {
          // TODO(wj): A dirty hack
          var string = '';
          node.children.forEach((cnode) {
            if (cnode is XmlElement && cnode.name.local == 'r') {
              cnode.children.forEach((ccnode) {
                if (ccnode is XmlElement && ccnode.name.local == 't')
                  string += ccnode.text;
              });
            }
          });
          if (string.isNotEmpty) {
            _nodeArray.add(string);
            _indexMap[string] = i++;
          } else {
            // TODO: Properly support rich text nodes in the future. For now just store the object as a placeholder.
            _nodeArray.add(node.children);
            _indexMap[node.children.toString()] = i++;
          }
        }
      }
    });
  }
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

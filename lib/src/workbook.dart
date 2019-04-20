import 'dart:io' show File;
import 'dart:convert';

import 'package:archive/archive.dart';

import './attached_xml_element.dart';
import './content_types.dart';
import './app_properties.dart';
import './core_properties.dart';
import './relationships.dart';
import './shared_strings.dart';
import './style_sheet.dart';
import './xml_utils.dart';
import './sheet.dart';
import './nodes.dart';

class Workbook extends AttachedXmlElement {
  ContentTypes _contentTypes;
  AppProperties _appProperties;
  CoreProperties _coreProperties;
  Relationships _relationships;
  SharedStrings _sharedStrings;
  StyleSheet _styleSheet;

  int _maxSheetId;
  XmlElement _sheetsNode;
  List<Sheet> _sheets;
  Sheet _activeSheet;

  ContentTypes get contentTypes => _contentTypes;
  AppProperties get appProperties => _appProperties;
  CoreProperties get coreProperties => _coreProperties;
  Relationships get relationships => _relationships;
  SharedStrings get sharedStrings => _sharedStrings;
  StyleSheet get styleSheet => _styleSheet;

  Workbook(XmlElement node) : super(node);

  factory Workbook.fromFile(String path) {
    var data = File(path).readAsBytesSync();
    return Workbook.fromData(data);
  }

  factory Workbook.fromData(List<int> data) {
    var archive = ZipDecoder().decodeBytes(data);

    XmlElement getRoot(String path) {
      var file = archive.findFile(path);
      if (file == null) return null;
      file.decompress();
      return parse(utf8.decode(file.content))?.rootElement;
    }

    var workbook = Workbook(getRoot('xl/workbook.xml')).._init(getRoot);

    return workbook;
  }

  List<int> toData() {
    _setSheetRefs();
/*
    var definedNamesNode = findChild(thisNode, "definedNames");
    for (int i = 0; i < _sheets.length; i++) {
      var sheet = _sheets[i];
      if (sheet.autoFilter != null) {
        if (definedNamesNode == null) {
          definedNamesNode = Element('definedNames').toXmlNode();
          insertInOrder(thisNode, definedNamesNode, nodeOrder);
        }

        definedNamesNode.children.add(Element('definedName', {
          'name': '_xlnm._FilterDatabase',
          'localSheetId': i,
          'hidden': '1'
        }));
      }
    }

        this._sheets.forEach((sheet, i) => {
            xmlq.appendChild(definedNamesNode, {
                name: "definedName",
                attributes: {
                    name: "_xlnm._FilterDatabase",
                    localSheetId: i,
                    hidden: "1"
                },
                children: [sheet._autoFilter.address({ includeSheetName: true, anchored: true })]
            });
        });
*/
  }

  void toFile(String path) {

  }

  void _init(XmlElement getRoot(String path)) {
    _maxSheetId = 0;
    _sheets = [];

    _contentTypes = ContentTypes(getRoot('[Content_Types].xml'));
    _appProperties = AppProperties(getRoot('docProps/app.xml'));
    _coreProperties = CoreProperties(getRoot('docProps/core.xml'));
    _relationships = Relationships(getRoot('xl/_rels/workbook.xml.rels'));
    _sharedStrings = SharedStrings(getRoot('xl/sharedStrings.xml'));
    _styleSheet = StyleSheet(getRoot('xl/styles.xml'));

    if (_relationships.findByType("sharedStrings") == null) {
      _relationships.add("sharedStrings", "sharedStrings.xml");
    }

    if (_contentTypes.findByPartName("/xl/sharedStrings.xml") == null) {
      _contentTypes.add("/xl/sharedStrings.xml",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    }

    _sheetsNode = findChild(thisNode, "sheets");
    for (var i = 0; i < _sheetsNode.children.length; i++) {
      var sheetIdNode = _sheetsNode.children[i] as XmlElement;
      int sheetId = getAttribute(sheetIdNode, 'sheetId');
      if (sheetId > _maxSheetId) _maxSheetId = sheetId;

      var sheetNode = getRoot('xl/worksheets/sheet${i + 1}.xml');
      var sheetRelationshipsNode =
          getRoot('xl/worksheets/_rels/sheet${i + 1}.xml.rels');

      _sheets.add(Sheet(this, sheetIdNode, sheetNode, sheetRelationshipsNode));
    }

    _parseSheetRefs();
  }

  Sheet sheet(String name) => _sheets.firstWhere((sheet) => sheet.name == name);

  Sheet sheetAt(int index) => _sheets[index];

  Map<XmlElement, Sheet> _definedNameNode_localSheet = {};

  void _parseSheetRefs() {
    var bookViewsNode = findChild(thisNode, 'bookViews');
    var workbookViewNode = findChild(bookViewsNode, 'workbookView');
    int activeTabId = getAttribute(workbookViewNode, 'activeTab') ?? 0;
    _activeSheet = _sheets[activeTabId];

    var definedNamesNode = findChild(thisNode, "definedNames");
    if (definedNamesNode != null) {
      definedNamesNode.children.forEach((definedNameNode) {
        int localSheetId = getAttribute(definedNameNode, 'localSheetId');
        if (localSheetId != null) {
          _definedNameNode_localSheet[definedNameNode] = _sheets[localSheetId];
        }
      });
    }
  }

  void _setSheetRefs() {
    var bookViewsNode = findChild(thisNode, 'bookViews');
    if (bookViewsNode == null) {
      bookViewsNode = Element('bookViews').toXmlNode();
      insertInOrder(thisNode, bookViewsNode, nodeOrder);
    }

    var workbookViewNode = findChild(bookViewsNode, 'workbookView');
    if (workbookViewNode == null) {
      workbookViewNode = Element('workbookView').toXmlNode();
      bookViewsNode.children.add(workbookViewNode);
    }

    setAttribute(workbookViewNode, 'activeTab', _sheets.indexOf(_activeSheet));

    var definedNamesNode = findChild(thisNode, 'definedNames');
    if (definedNamesNode != null) {
      definedNamesNode.children.forEach((definedNameNode) {
        var localSheet = _definedNameNode_localSheet[definedNameNode];
        if (localSheet != null) {
          setAttribute(definedNameNode, 'localSheetId', _sheets.indexOf(localSheet));
        }
      });
    }
  }

  static const nodeOrder = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
    "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
  ];
}

/*
xl/workbook.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
	<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16925"/>
	<workbookPr defaultThemeVersion="164011"/>
	<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
		<mc:Choice Requires="x15">
			<x15ac:absPath url="\path\to\file" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>
		</mc:Choice>
	</mc:AlternateContent>
	<bookViews>
		<workbookView xWindow="3720" yWindow="0" windowWidth="27870" windowHeight="12795"/>
	</bookViews>
	<sheets>
		<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
	</sheets>
	<calcPr calcId="171027"/>
	<extLst>
		<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
			<x15:workbookPr chartTrackingRefBase="1"/>
		</ext>
	</extLst>
</workbook>
// */

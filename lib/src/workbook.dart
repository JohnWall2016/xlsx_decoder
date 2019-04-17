import 'dart:io' show File;
import 'dart:convert';

import 'package:archive/archive.dart';

import './document.dart';
import './content_types.dart';
import './app_properties.dart';
import './core_properties.dart';
import './relationships.dart';
import './shared_strings.dart';
import './style_sheet.dart';
import './xml_utils.dart';
import './sheet.dart';

class Workbook extends Document {
  @override
  String get id => 'xl/workbook.xml';

  Workbook._new();

  factory Workbook.fromFile(String path) {
    var data = File(path).readAsBytesSync();
    return Workbook.fromData(data);
  }

  factory Workbook.fromData(List<int> data) => Workbook._new().._init(data);

  Archive _archive;
  ContentTypes _contentTypes;
  AppProperties _appProperties;
  CoreProperties _coreProperties;
  Relationships _relationships;
  SharedStrings _sharedStrings;
  StyleSheet _styleSheet;

  int _maxSheetId;
  XmlElement _sheetsNode;
  List<Sheet> _sheets;

  ContentTypes get contentTypes => _contentTypes;
  AppProperties get appProperties => _appProperties;
  CoreProperties get coreProperties => _coreProperties;
  Relationships get relationships => _relationships;
  SharedStrings get sharedStrings => _sharedStrings;
  StyleSheet get styleSheet => _styleSheet;

  void _init(data) {
    _maxSheetId = 0;
    _sheets = [];

    _archive = ZipDecoder().decodeBytes(data);
    _contentTypes = _loadDocument(ContentTypes());
    _appProperties = _loadDocument(AppProperties());
    _coreProperties = _loadDocument(CoreProperties());
    _relationships = _loadDocument(Relationships());
    _sharedStrings = _loadDocument(SharedStrings());
    _styleSheet = _loadDocument(StyleSheet());
    _loadDocument(this);

    if (_relationships.findByType("sharedStrings") == null) {
      _relationships.add("sharedStrings", "sharedStrings.xml");
    }

    if (_contentTypes.findByPartName("/xl/sharedStrings.xml") == null) {
      _contentTypes.add("/xl/sharedStrings.xml",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    }

    _sheetsNode = findChild(document.rootElement, "sheets");
    for (var i = 0; i < _sheetsNode.children.length; i++) {
      var sheetIdNode = _sheetsNode.children[i] as XmlElement;
      var sheetId = int.parse(getAttribute(sheetIdNode, 'sheetId'));
      if (sheetId > _maxSheetId) _maxSheetId = sheetId;

      var sheetNode = parseDocument('xl/worksheets/sheet${i + 1}.xml');
      var sheetRelationshipsNode =
          parseDocument('xl/worksheets/_rels/sheet${i + 1}.xml.rels');

      _sheets.add(Sheet(this, sheetIdNode, sheetNode, sheetRelationshipsNode));
    }
  }

  T _loadDocument<T extends Document>(T doc) =>
      doc..load(parseDocument(doc.id));

  XmlDocument parseDocument(String path) {
    var file = _archive.findFile(path);
    if (file == null) return null;
    file.decompress();
    return parse(utf8.decode(file.content));
  }
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

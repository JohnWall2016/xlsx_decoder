import 'dart:io' show File;
import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

import './content_types.dart';
import './app_properties.dart';
import './core_properties.dart';
import './relationships.dart';
import './shared_strings.dart';
import './style_sheet.dart';

class Workbook {
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
  XmlDocument _document;

  void _init(data) {
    _archive = ZipDecoder().decodeBytes(data);
    _contentTypes = ContentTypes(_parseDocument('[Content_Types].xml'));
    _appProperties = AppProperties(_parseDocument('docProps/app.xml'));
    _coreProperties = CoreProperties(_parseDocument('docProps/core.xml'));
    _relationships = Relationships(_parseDocument('xl/_rels/workbook.xml.rels'));
    _sharedStrings = SharedStrings(_parseDocument('xl/sharedStrings.xml'));
    _styleSheet = StyleSheet(_parseDocument('xl/styles.xml'));
    _document = _parseDocument('xl/workbook.xml');
  }

  XmlDocument _parseDocument(String name) {
    var file = _archive.findFile(name);
    if (file == null) return null;
    file.decompress();
    return parse(utf8.decode(file.content));
  }
}

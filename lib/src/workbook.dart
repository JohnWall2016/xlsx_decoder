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
  XmlDocument _document;

  void _init(data) {
    _archive = ZipDecoder().decodeBytes(data);
    _contentTypes = _parseDocument(ContentTypes());
    _appProperties = _parseDocument(AppProperties());
    _coreProperties = _parseDocument(CoreProperties());
    _relationships = _parseDocument(Relationships());
    _sharedStrings = _parseDocument(SharedStrings());
    _styleSheet = _parseDocument(StyleSheet());
    _document = _parseDocument(this);
  }

  T _parseDocument<T extends Document>(T doc) {
    var file = _archive.findFile(doc.id);
    if (file == null) 
      doc.load(null);
    else {
      file.decompress();
      doc.load(parse(utf8.decode(file.content)));
    }
    return doc;
  }
}

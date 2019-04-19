import 'package:xml/xml.dart';

import './sheet.dart';
import './attached_xml_element.dart';

class Column extends AttachedXmlElement {
  Sheet _sheet;

  Column(this._sheet, XmlElement node) : super(node) {

  }
}
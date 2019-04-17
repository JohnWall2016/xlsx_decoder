import './document.dart';
import './nodes.dart';
import './style.dart';
import './xml_utils.dart';

class StyleSheet extends Document {
  @override
  String get id => 'xl/styles.xml';

  static const standardCodes = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@'
  };

  static const startingCustomNumberFormatId = 164;

  XmlElement _numFmtsNode;
  XmlElement _fontsNode;
  XmlElement _fillsNode;
  XmlElement _bordersNode;
  XmlElement _cellXfsNode;

  @override
  void load(XmlDocument document) {
    super.load(document);

    elements.forEach((node) {
      switch (node.name.local) {
        case 'numFmts':
          _numFmtsNode = node;
          break;
        case 'fonts':
          _fontsNode = node;
          break;
        case 'fills':
          _fillsNode = node;
          break;
        case 'borders':
          _bordersNode = node;
          break;
        case 'cellXfs':
          _cellXfsNode = node;
          break;
      }
    });

    if (_numFmtsNode == null) {
      _numFmtsNode = Element('numFmts').toXmlNode();
      insertNode(0, _numFmtsNode);
    }

    [_numFmtsNode, _fontsNode, _fillsNode, _bordersNode, _cellXfsNode]
        .forEach((node) {
      node.attributes.removeWhere((attr) => attr.name.local == 'count');
    });

    _cacheNumberFormats();
  }

  Map<int, String> _numberFormatCodesById = {};
  Map<String, int> _numberFormatIdsByCode = {};

  int _nextNumFormatId = startingCustomNumberFormatId;

  String getNumberFormatCode(int id) => _numberFormatCodesById[id];

  int getNumberFormatId(String code) {
    var id = _numberFormatIdsByCode[code];
    if (id == null) {
      id = _nextNumFormatId++;
      _numberFormatCodesById[id] = code;
      _numberFormatIdsByCode[code] = id;

      var node = (Element('numFmt')
            ..attributes['numFmtId'] = '$id'
            ..attributes['formatCode'] = code)
          .toXmlNode();

      addNode(node);
    }
    return id;
  }

  void _cacheNumberFormats() {
    for (var id in standardCodes.keys) {
      var code = standardCodes[id];
      _numberFormatCodesById[id] = code;
      _numberFormatIdsByCode[code] = id;
    }

    _nextNumFormatId = startingCustomNumberFormatId;

    _numFmtsNode.children.whereType<XmlElement>().forEach((node) {
      var id = int.parse(node.getAttribute('numFmtId'));
      var code = node.getAttribute('formatCode');
      if (id != null && code != null) {
        _numberFormatCodesById[id] = code;
        _numberFormatIdsByCode[code] = id;
        if (id >= _nextNumFormatId) _nextNumFormatId = id + 1;
      }
    });
  }

  Style createStyle(int sourceId) {
    XmlElement fontNode, fillNode, borderNode, xfNode;

    if (sourceId >= 0) {
      var sourceXfNode = _cellXfsNode.children[sourceId] as XmlElement;
      xfNode = sourceXfNode.copy();

      if (getAttribute(sourceXfNode, 'applyFont') != null) {
        var fontId = int.parse(getAttribute(sourceXfNode, 'fontId'));
        fontNode = _fontsNode.children[fontId].copy();
      }

      if (getAttribute(sourceXfNode, 'applyFill') != null) {
        var fillId = int.parse(getAttribute(sourceXfNode, 'fillId'));
        fillNode = _fillsNode.children[fillId].copy();
      }

      if (getAttribute(sourceXfNode, 'applyBorder') != null) {
        var borderId = int.parse(getAttribute(sourceXfNode, 'borderId'));
        borderNode = _bordersNode.children[borderId].copy();
      }
    }

    if (fontNode == null) fontNode = Element('font').toXmlNode();
    _fontsNode.children.add(fontNode);

    if (fillNode == null) fillNode = Element('fill').toXmlNode();
    _fillsNode.children.add(fillNode);

    // The border sides must be in order
    if (borderNode == null) {
      borderNode = (Element('border')
            ..children.addAll([
              Element('left'),
              Element('right'),
              Element('top'),
              Element('bottom'),
              Element('diagonal'),
            ]))
          .toXmlNode();
    }
    _bordersNode.children.add(borderNode);

    if (xfNode == null) xfNode = Element('xf').toXmlNode();
    setAttributes(xfNode, {
      'fontId': '${_fontsNode.children.length - 1}',
      'fillId': '${_fillsNode.children.length - 1}',
      'borderId': '${_bordersNode.children.length - 1}',
      'applyFont': '1',
      'applyFill': '1',
      'applyBorder': '1'
    });
    _cellXfsNode.children.add(xfNode);

    return Style(this, _cellXfsNode.children.length - 1, xfNode, fontNode,
        fillNode, borderNode);
  }
}

/*
xl/styles.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0_);[Red]\(#,##0\)\)"/>
    </numFmts>
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="11">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFC00000"/>
                <bgColor indexed="64"/>
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="lightDown">
                <fgColor theme="4"/>
                <bgColor rgb="FFC00000"/>
            </patternFill>
        </fill>
        <fill>
            <gradientFill degree="90">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill>
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="45">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="135">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="270">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
    </fills>
    <borders count="10">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="hair">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dotted">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashed">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="slantDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashed">
                <color auto="1"/>
            </diagonal>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="19">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="10" borderId="0" xfId="0" applyFill="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
        </ext>
    </extLst>
</styleSheet>
*/

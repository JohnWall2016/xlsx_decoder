import './document.dart';

class CoreProperties extends Document {
  @override
  String get id => 'docProps/core.xml';
}

/*
docProps/core.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Title</dc:title>
<dc:subject>Subject</dc:subject>
<dc:creator>Creator</dc:creator>
<cp:keywords>Keywords</cp:keywords>
<dc:description>Description</dc:description>
<cp:category>Category</cp:category>
</cp:coreProperties>
 */
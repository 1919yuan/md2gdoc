function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  DocumentApp.getUi().createMenu('md2gdoc')
    .addItem('Import Markdown', 'showImport')
    .addToUi();
}

function showImport() {
  DocumentApp.getUi().showDialog(
    HtmlService.createTemplateFromFile('markdownit').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setTitle('md2gdoc')
      .setWidth(720)
      .setHeight(400));
}

function include(file) {
  return HtmlService.createTemplateFromFile(file).evaluate().getContent();
}

function preprocess(markdown) {
  markdown = markdown.replace(/<hr>/g, '<hr></hr>');
  markdown = markdown.replace(/<br>/g, '<br></br>');
  prev = markdown;
  markdown = prev.replace(/(\/?(?:div|p|h1|h2|h3|h4|h5|h6|blockquote|ul|ol|li))>\s+<(\/?(?:div|p|h1|h2|h3|h4|h5|h6|blockquote|ul|ol|li))/g, '$1><$2');
  while (markdown != prev) {
    prev = markdown;
    markdown = prev.replace(/(\/?(?:div|p|h1|h2|h3|h4|h5|h6|blockquote|ul|ol|li))>\s+<(\/?(?:div|p|h1|h2|h3|h4|h5|h6|blockquote|ul|ol|li))/g, '$1><$2');
  }
  Logger.log(markdown);
  return markdown;
}

function importMarkdown(markdown) {
  markdown = preprocess(markdown);
  var document = XmlService.parse(markdown);
  var root = document.getRootElement();
  PropertiesService.getScriptProperties().setProperty(
    'listNestingLevel', JSON.stringify(-1));
  PropertiesService.getScriptProperties().setProperty(
    'prevListId', JSON.stringify([]));
  PropertiesService.getScriptProperties().setProperty(
    'glyphType', JSON.stringify([]));
  PropertiesService.getScriptProperties().setProperty(
    'inBlockquote', JSON.stringify(false));
  processDescendents(root.getDescendants(), root);
}

function processDescendents(ds, parent) {
  for (var i = 0; i < ds.length; i++) {
    if (xmlGetPath(ds[i].getParentElement()) == xmlGetPath(parent)) {
      processDescendent(ds[i]);
    }
  }
}

function processDescendent(d) {
  switch (d.getType()) {
    case XmlService.ContentTypes.ELEMENT:
      var e = d.asElement();
      switch (e.getName()) {
        case 'h1':
          h1(d);
          break;
        case 'h2':
          h2(d);
          break;
        case 'h3':
          h3(d);
          break;
        case 'h4':
          h4(d);
          break;
        case 'h5':
          h5(d);
          break;
        case 'h6':
          h6(d);
          break;
        case 'p':
          p(d);
          break;
        case 'a':
          a(d);
          break;
        case 'hr':
          hr(d);
          break;
        case 'br':
          br(d);
          break;
        case 'em':
          em(d);
          break;
        case 'strong':
          strong(d);
          break;
        case 'blockquote':
          blockquote(d);
          break;
        case 'ul':
          ul(d);
          break;
        case 'ol':
          ol(d);
          break;
        case 'li':
          li(d);
          break;
        default:
          break;
      }
      break;
    case XmlService.ContentTypes.TEXT:
      var e = d.asText();
      var item = getLastItemToAppend();
      var text = e.getText();
      if (text && item) {
        var t = item.appendText(text);
        t.setBold(false);
        t.setItalic(false);
        t.setLinkUrl(null);
      }
      break;
    default:
      break;
  };
}

function processParagraph(e, heading) {
  var paragraph = getLastParagraph(null);
  if (paragraph) {
    if (paragraph.getText() != "") {
      var body = DocumentApp.getActiveDocument().getBody();
      paragraph = body.appendParagraph("a");
      paragraph = paragraph.clear();
    } else {
      paragraph.clear();
    }
  } else {
    var body = DocumentApp.getActiveDocument().getBody();
    paragraph = body.appendParagraph("a");
    paragraph = paragraph.clear();
  }
  paragraph.setHeading(heading);
  processDescendents(e.getDescendants(), e);
}

function h1(e) {
  processParagraph(e, DocumentApp.ParagraphHeading.HEADING1);
}

function h2(e) {
  processParagraph(e, DocumentApp.ParagraphHeading.HEADING2);
}

function h3(e) {
  processParagraph(e, DocumentApp.ParagraphHeading.HEADING3);
}

function h4(e) {
  processParagraph(e, DocumentApp.ParagraphHeading.HEADING4);
}

function h5(e) {
  processParagraph(e, DocumentApp.ParagraphHeading.HEADING5);
}

function h6(e) {
  processParagraph(e, DocumentApp.ParagraphHeading.HEADING6);
}

function p(e) {
  var level = JSON.parse(PropertiesService.getScriptProperties().getProperty(
    'listNestingLevel'));
  var blockquote = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('inBlockquote'));
  if (level != -1) {
    processDescendents(e.getDescendants(), e);
  } else if (blockquote) {
    var cell = getLastTableCell();
    var paragraph = getLastParagraph(cell);
    if (paragraph) {
      if (paragraph.getText() != "") {
        paragraph = cell.appendParagraph("a");
        paragraph = paragraph.clear();
      } else {
        paragraph.clear();
      }
    } else {
      paragraph = cell.appendParagraph("a");
      paragraph = paragraph.clear();
    }
    processDescendents(e.getDescendants(), e);
  } else {
    processParagraph(e, DocumentApp.ParagraphHeading.NORMAL);
  }
}

function hr(e) {
  var body = DocumentApp.getActiveDocument().getBody();
  body.appendHorizontalRule();
}

function br(e) {
  var item = getLastItemToAppend();
  if (item) {
    item.appendText('\n');
  }
}

function blockquote(e) {
  var body = DocumentApp.getActiveDocument().getBody();
  var t = body.appendTable([['a']]);
  t.clear();
  var row = t.appendTableRow();
  var cell = row.appendTableCell();
  var body = DocumentApp.getActiveDocument().getBody();
  var inBlockquote = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('inBlockquote'));
  if (inBlockquote) {
    throw new Error("Nested levels of Blockquote is not supported.");
  }
  PropertiesService.getScriptProperties().setProperty(
    'inBlockquote', JSON.stringify(true));
  processDescendents(e.getDescendants(), e);
  PropertiesService.getScriptProperties().setProperty(
    'inBlockquote', JSON.stringify(false));
}

function a(e) {
  var href = e.getAttribute('href').getValue();
  var text = e.getText().trim();
  var item = getLastItemToAppend();
  if (text && item) {
    var t = item.appendText(text.trim());
    var len = t.getText().length;
    var newlen = text.length;
    var start = len - newlen;
    var end = len - 1;
    t.setLinkUrl(start, end, href);
  }
}

function em(e) {
  var text = e.getText();
  var item = getLastItemToAppend();
  if (text && item) {
    var t = item.appendText(text);
    var len = t.getText().length;
    var newlen = text.length;
    var start = len - newlen;
    var end = len - 1;
    t.setItalic(start, end, true);
  }
}

function strong(e) {
  var text = e.getText();
  var item = getLastItemToAppend();
  if (text) {
    var t = item.appendText(text);
    var len = t.getText().length;
    var newlen = text.length;
    var start = len - newlen;
    var end = len - 1;
    Logger.log(start);
    Logger.log(end);
    Logger.log(t.getText());
    t.setBold(start, end, true);
  }
}


function ul(e) {
  var level = JSON.parse(PropertiesService.getScriptProperties().getProperty(
    'listNestingLevel')) + 1;
  PropertiesService.getScriptProperties().setProperty(
    'listNestingLevel', JSON.stringify(level));
  var prevListIds = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'));
  prevListIds.push(null);
  PropertiesService.getScriptProperties().setProperty(
    'prevListId', JSON.stringify(prevListIds));
  var glyphTypes = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('glyphType'));
  glyphTypes.push(DocumentApp.GlyphType.BULLET);
  PropertiesService.getScriptProperties().setProperty(
    'glyphType', JSON.stringify(glyphTypes));

  processDescendents(e.getDescendants(), e);

  close_ul(e);
}

function close_ul(e) {
  var level = JSON.parse(PropertiesService.getScriptProperties().getProperty(
    'listNestingLevel'));
  PropertiesService.getScriptProperties().setProperty(
    'listNestingLevel', JSON.stringify(level - 1));
  var prevListIds = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'));
  prevListIds.pop();
  PropertiesService.getScriptProperties().setProperty(
    'prevListId', JSON.stringify(prevListIds));
  var glyphTypes = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('glyphType'));
  glyphTypes.pop();
  PropertiesService.getScriptProperties().setProperty(
    'glyphType', JSON.stringify(glyphTypes));
}

function ol(e) {
  var level = JSON.parse(PropertiesService.getScriptProperties().getProperty(
    'listNestingLevel')) + 1;
  PropertiesService.getScriptProperties().setProperty(
    'listNestingLevel', JSON.stringify(level));
  var prevListIds = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'));
  prevListIds.push(null);
  PropertiesService.getScriptProperties().setProperty(
    'prevListId', JSON.stringify(prevListIds));
  var glyphTypes = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('glyphType'));
  glyphTypes.push(DocumentApp.GlyphType.NUMBER);
  PropertiesService.getScriptProperties().setProperty(
    'glyphType', JSON.stringify(glyphTypes));

  processDescendents(e.getDescendants(), e);

  close_ol(e);
}

function close_ol(e) {
  var level = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('listNestingLevel'));
  PropertiesService.getScriptProperties().setProperty(
    'listNestingLevel', JSON.stringify(level - 1));
  var prevListIds = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'));
  prevListIds.pop();
  PropertiesService.getScriptProperties().setProperty(
    'prevListId', JSON.stringify(prevListIds));
  var glyphTypes = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('glyphType'));
  glyphTypes.pop();
  PropertiesService.getScriptProperties().setProperty(
    'glyphType', JSON.stringify(glyphTypes));
}

function li(e) {
  var body = DocumentApp.getActiveDocument().getBody();
  var level = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('listNestingLevel'));
  var glyphType = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('glyphType'))[level];
  var prevListId = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'))[level];
  var item = body.appendListItem("a");
  item = item.clear();
  item = item.setNestingLevel(level).setGlyphType(
    DocumentApp.GlyphType[glyphType]);
  item = item.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  if (prevListId) {
    var prevItem = getLastListItemById(prevListId);
    item = item.setListId(prevItem);
  }
  var prevListIds = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'));
  prevListIds[level] = item.getListId();
  PropertiesService.getScriptProperties().setProperty(
    'prevListId', JSON.stringify(prevListIds));
  processDescendents(e.getDescendants(), e);
}

function getLastListItemById(id) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  for (var i = doc.getNumChildren() - 1; i >= 0; i--) {
    var child = doc.getChild(i);
    if (child.getType() == DocumentApp.ElementType.LIST_ITEM &&
        child.getListId() == id) {
      return child;
    }
  }
  return null;
}

function getLastParagraph(element) {
  var doc = DocumentApp.getActiveDocument();
  if (element != null) {
    doc = element;
  }
  for (var i = doc.getNumChildren() - 1; i >= 0; i--) {
    var child = doc.getChild(i);
    if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
      return child;
    }
  }
  return null;
}

function getLastTableCell() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  for (var i = doc.getNumChildren() - 1; i >= 0; i--) {
    var child = doc.getChild(i);
    if (child.getType() == DocumentApp.ElementType.TABLE) {
      return child.getRow(0).getCell(0);
    }
  }
  return null;
}

function getLastItemToAppend() {
  var level = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('listNestingLevel'));
  var prevListId = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('prevListId'))[level];
  var blockquote = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('inBlockquote'));
  if (level >= 0) {
    return getLastListItemById(prevListId);
  } else if (blockquote) {
    var cell = getLastTableCell();
    return getLastParagraph(cell);
  } else {
    return getLastParagraph(null);
  }
}

function xmlGetPath(element) {
  var path = "";
  while (!element.isRootElement()) {
    path = element.getQualifiedName() + "/" + path;
    element = element.getParentElement();
  }
  return path;
}

function onInstall(e) {
	onOpen(e);
}


function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Match footnote style to selection', 'matchStyle')
    .addItem('Update footnotes', 'updateFootnotes')
    .addSeparator()
    .addItem('Use as default', 'saveDefault')
    .addItem('Restore default', 'restoreDefault')
    .addToUi();
}


function matchStyle() {
  var docProps = PropertiesService.getDocumentProperties();
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  var selection = doc.getSelection();
  var footnotes = doc.getFootnotes();
  
  var para;
  var text;
  var index;
  var para_attributes;
  var text_attributes;
  
  var props = {};
  var final_props = {};
  var text_props = [
    'BOLD',
    'ITALIC',
    'LINK_URL',
    'UNDERLINE'
  ];
  
  // Determine the paragraph, text and text position to clone attributes
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i in elements) {
      var element = elements[i].getElement();
      if (elements[i].isPartial()) {
        para = element.getParent();
        text = element.asText();
        index = elements[i].getStartOffset();  
      } else {
        para = element;
        text = element.asParagraph().getChild(0);
        index = 0;
      };
      break
    };
  } else {
    para = cursor.getElement().getParent();
    text = cursor.getElement().asText();
    index = cursor.getOffset();
  };
  
  para_attributes = para.getAttributes();
  text_attributes = text.getAttributes(index);
  
  // Create a new object for the attributes to be applied
  
  for (var para_att in para_attributes) {
    if (para_att === 'HEADING') continue;
    props[para_att] = para_attributes[para_att];
  };
  
  for (var text_att in text_attributes) {
    props[text_att] = text_attributes[text_att];
  };
  
  for (var i in text_props) {
    delete props[text_props[i]];
  };
  
  
  // Create final object with correct formatting of keys/values
  for (var prop in props) {
    var doc_keyname = prop + '_fs';
    if (props[prop] == null) {
      final_props[doc_keyname] = null;
    } else {
      final_props[doc_keyname] = props[prop];
    };
  };
  
  // Set props and update
  docProps.setProperties(final_props);
  updateFootnotes();
  
}


function updateFootnotes() {
	var userProps = PropertiesService.getUserProperties();
	var docProps = PropertiesService.getDocumentProperties();
  var user_props = userProps.getProperties();
  var doc_props = docProps.getProperties();
    
	var doc = DocumentApp.getActiveDocument();
	var footnotes = doc.getFootnotes();
    
	var final_props = {};
  var h_align;
  
	if (!('FONT_SIZE_fs' in doc_props)) {
		if (!('FONT_SIZE_fs' in user_props)) {
			//DocumentApp.getUi().alert('No footnote style found.\n\nSet one by selecting some text and clicking\n\'Match footnote style to selection\'.')
		} else {
			docProps.setProperties(user_props);
		};
  };

	for (var i in doc_props) {
		var orig_key = i.substring(0, i.length - 3);
		if (doc_props[i] == '') {
			final_props[orig_key] = null;
		} else {
			final_props[orig_key] = doc_props[i];
		};
	};
    
  h_align = final_props['HORIZONTAL_ALIGNMENT'];
  delete final_props['HORIZONTAL_ALIGNMENT'];
  delete final_props['HEADING'];

  if (h_align) {
    final_props[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
      eval('DocumentApp.HorizontalAlignment.' + h_align.toUpperCase());
  };

	for (var j in footnotes) {
		var paras = footnotes[j].getFootnoteContents().getParagraphs();
		for (var k in paras) {
			paras[k].setAttributes(final_props);
		};
	};
}


function saveDefault() {
	var userProps = PropertiesService.getUserProperties();
	var doc_props = PropertiesService.getDocumentProperties().getProperties();
    
	userProps.setProperties(doc_props);
}


function restoreDefault() {
	var docProps = PropertiesService.getDocumentProperties();
	var user_props = PropertiesService.getUserProperties().getProperties();
    
	docProps.setProperties(user_props);
	updateFootnotes();
}


function resetStyle() {
	var docProps = PropertiesService.getDocumentProperties();
	var props = {
    FONT_SIZE_fs: 10,
    LINE_SPACING_fs: 0,
    SPACING_BEFORE_fs: 0,
    SPACING_AFTER_fs: 0,
    HORIZONTAL_ALIGNMENT_fs: 'Left',
    STRIKETHROUGH_fs: null,
    FOREGROUND_COLOR_fs: null,
    FONT_FAMILY_fs: null,
    BACKGROUND_COLOR_fs: null,
    INDENT_END_fs: null,
    INDENT_START_fs: null,
    INDENT_FIRST_LINE_fs: null
	};
    
    docProps.deleteAllProperties().setProperties(props);
    updateFootnotes();
}


/**
* Testing only
*/

function clearProps() {
  var userProps = PropertiesService.getUserProperties();
	var docProps = PropertiesService.getDocumentProperties();
	userProps.deleteAllProperties();
  docProps.deleteAllProperties();
}

function logProps() {
  //Logger.log(PropertiesService.getUserProperties().getProperties());
  Logger.log(PropertiesService.getDocumentProperties().getProperties());
}

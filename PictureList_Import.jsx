/*****************************************************************
 *
 * Copyright (c) 2021 Dustin Liaw
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is furnished
 * to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 *****************************************************************/

#include "xlsx.extendscript.js"

const ADOBE_PHOTOSHOP = "Adobe Photoshop";
const ADOBE_ILLUSTRATOR = "Adobe Illustrator";
const ADOBE_INDESIGN = "Adobe InDesign";

const appFiletypes = {};
appFiletypes[ADOBE_PHOTOSHOP] = [
  "Photoshop:*.psd;*.pdd;*.psdt",
  "Large Document Format:*.psb"
];
appFiletypes[ADOBE_ILLUSTRATOR] = [
  "Illustrator:*.ai",
];
appFiletypes[ADOBE_INDESIGN] = [
  "InDesign:*.indd"
];

const WHITESPACE_REGEX = /\s/g;

// Localization
$.localize = true;
const UI_MESSAGES = {
  OK: { en: "OK" },
  CANCEL: { en: "Cancel" },
  UNSUPPORTED_APPLICATION: { en: "This application is not supported" },
  SELECT_CC_FILE: { en: "Select Creative Cloud File" },
  SELECT_PICTURE_LIST: { en: "Select Picture List File" },
  FILE_LOAD_CONFIRM: {
    en: app.name + " source file: %1\n\nPicture list: %2\n\nLoad files? This operation may take a while." },
  PICTURE_LIST_LOAD_ERROR: { en: "Failed to open picture list: %1" },
  SELECT_WORKSHEET_TITLE: { en: "Select Worksheet"},
  NO_TARGET_LANGS_FOUND: { en: "No target languages found" },
  NO_TARGET_LANGS_SELECTED: { en: "No target languages selected" },
  SELECT_LANGUAGES_TITLE: { en: "Select Source/Target Languages" },
  SOURCE_LANG_GROUP_TITLE: { en: "Source language" },
  TARGET_LANG_GROUP_TITLE: { en: "Target languages" },
  MISSING_SHEET_ROWS: { en: "Picture list entries not found in file:\n\n%1\n\nProceed?" },
  FILE_LOCALIZE_SUCCESS: { en: "Successfully generated %1 localized file(s)" },
  FILE_LOCALIZE_ERROR: { en: "Failed to generate localized files: %1" }
};


/**
 * Opens Adobe CC file selection dialog
 * @returns File
 */
function getCCFile() {
  if (app.documents.length === 0) {
    var filetypes = appFiletypes[app.name];
    if (filetypes === undefined) {
      alert(UI_MESSAGES.UNSUPPORTED_APPLICATION);
      return null;
    }
    filetypes.push("All Formats:*.*");
    return File.openDialog(UI_MESSAGES.SELECT_CC_FILE, filetypes.join(","));
  } else {
    return app.activeDocument.fullName;
  }
}

/**
 * Opens picture list file selection dialog
 * @returns File
 */
function getPictureListFile() {
  return File.openDialog(UI_MESSAGES.SELECT_PICTURE_LIST, "Excel Files:*.xlsx;*.xls,All Formats:*.*");
}

/**
 * Add target language to source filename
 * @param {String} sourceName 
 * @param {String} targetLang 
 * @returns String
 */
function generateTargetFilename(sourceName, targetLang) {
  var nameComponents = sourceName.split(".");
  nameComponents[Math.max(0, nameComponents.length - 2)] += "_" + targetLang;
  return nameComponents.join(".");
}

/**
 * Open source/target language selection dialog
 * @returns Object containing selected source (string) and target languages (string array)
 */
function userSelectLanguages(languages) {
  var langConfigWindow = new Window("dialog", UI_MESSAGES.SELECT_LANGUAGES_TITLE);
  langConfigWindow.alignChildren = "left";
  var srcLangGroup = langConfigWindow.add("group")
  srcLangGroup.add("statictext", undefined, UI_MESSAGES.SOURCE_LANG_GROUP_TITLE);
  var srcLangDropdown = srcLangGroup.add("dropdownlist", undefined, languages);
  srcLangDropdown.selection = 0;
  srcLangDropdown.currentLang = languages[0];

  var targetLangGroup = langConfigWindow.add("group");
  targetLangGroup.add("statictext", undefined, UI_MESSAGES.TARGET_LANG_GROUP_TITLE);
  targetLangGroup.alignChildren = "left";
  var targetLangCheckboxes = {};
  for (var i = 0; i < languages.length; ++i) {
    var checkbox = targetLangGroup.add("checkbox", undefined, languages[i]);
    targetLangCheckboxes[languages[i]] = checkbox;
    if (i === 0) {
      checkbox.enabled = false;
    } else {
      checkbox.value = true;
    }
  }

  // Disable selected source language's target checkbox
  srcLangDropdown.onChange = function() {
    if (srcLangDropdown.selection.text !== srcLangDropdown.currentLang) {
      targetLangCheckboxes[srcLangDropdown.currentLang].enabled = true;
      targetLangCheckboxes[srcLangDropdown.selection.text].enabled = false;
      targetLangCheckboxes[srcLangDropdown.selection.text].value = false;
      srcLangDropdown.currentLang = srcLangDropdown.selection.text;
    }
  }

  var buttonGroup = langConfigWindow.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.add("button", undefined, UI_MESSAGES.OK);
  buttonGroup.add("button", undefined, UI_MESSAGES.CANCEL);
  if (langConfigWindow.show() === 2) {
    return {
      source: null
    };
  }

  var selectedTargetLangs = [];
  for (var i = 0; i < languages.length; ++i) {
    if (targetLangCheckboxes[languages[i]].value) {
      selectedTargetLangs.push(languages[i]);
    }
  }

  var selectedLanguages = {
    source: srcLangDropdown.selection.text,
    target: selectedTargetLangs,
  };
  return selectedLanguages;
}

/**
 * Get document text layers/frames
 * @param {Document} document
 * @returns Array of Layers or TextFrameItems
 */
function getDocumentTextLayers(document) {
  var textLayers;

  switch (app.name) {
    case ADOBE_ILLUSTRATOR:
    case ADOBE_INDESIGN:
      textLayers = document.textFrames;
      break;
    case ADOBE_PHOTOSHOP:
    default:
      textLayers = [];
      for (var i = 0; i < document.layers.length; ++i) {
        if (document.layers[i].kind === LayerKind.TEXT) {
          textLayers.push(document.layers[i]);
        }
      }
      break;
  }

  return textLayers;
}

/**
 * Get text content of layer/frame
 * @param {Layer or TextFrame} layer 
 * @returns String
 */
function getLayerText(layer) {
  var layerText;

  switch (app.name) {
    case ADOBE_ILLUSTRATOR:
    case ADOBE_INDESIGN:
      layerText = layer.contents;
      break;
    case ADOBE_PHOTOSHOP:
    default:
      layerText = layer.textItem.contents;
  }

  return layerText;
}

/**
 * Set text content of layer/frame
 * @param {Layer or TextFrameItem} layer 
 * @param {String} newText 
 */
function setLayerText(layer, newText) {
  switch (app.name) {
    case ADOBE_ILLUSTRATOR:
    case ADOBE_INDESIGN:
      layer.contents = newText;
      break;
    case ADOBE_PHOTOSHOP:
    default:
      layer.textItem.contents = newText;
      break;
  }
}

/**
 * Opens worksheet selection dialog and loads selected worksheet JSON
 * @param {Workbook} workbook 
 * @returns Worksheet object
 */
function loadWorksheetData(workbook) {
  // Worksheet selection
  var sheetConfigWindow = new Window("dialog", UI_MESSAGES.SELECT_WORKSHEET_TITLE);
  var sheetGroup = sheetConfigWindow.add("group");
  var sheetDropdown = sheetGroup.add("dropdownlist", undefined, workbook.SheetNames);
  sheetDropdown.selection = 0;
  
  var buttonGroup = sheetConfigWindow.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.add("button", undefined, UI_MESSAGES.OK);
  buttonGroup.add("button", undefined, UI_MESSAGES.CANCEL);
  if (sheetConfigWindow.show() === 2) {
    return null;
  }

  // Convert sheet into JSON
  var worksheet = workbook.Sheets[sheetDropdown.selection];
  return XLSX.utils.sheet_to_json(worksheet, {raw: false});
}

/**
 * Find layer/frame with matching text content
 * @param {[Layer or TextFrameItem]} textLayers 
 * @param {[int]} layerIndices
 * @param {String} searchText 
 * @returns Layer/TextFrameItem
 */
 function findMatchingTextLayer(layerTextStrings, layerIndices, searchText) {  //TODO: Sort/otherwise optimize search
  if (layerIndices.length > 0) {
    for (var i = 0; i < layerIndices.length; ++i) {
      if (layerTextStrings[i] === searchText) {
        var matchIndex = layerIndices[i];
        layerTextStrings.splice(i, 1);
        layerIndices.splice(i, 1);
        return matchIndex;
      }
    }
  }

  return -1;
}

/**
 * Map worksheet rows to layer indices
 * @param {File} ccFile 
 * @param {Worksheet object} worksheetData 
 * @param {String} sourceLang 
 * @returns Array of integers
 */
function mapTextLayers(ccFile, worksheetData, sourceLang) {
  // var hasActiveDocument = app.activeDocument === null;
  var sourceDocument = app.open(ccFile);
  var textLayers = getDocumentTextLayers(sourceDocument);
  
  var layerIndices = [];
  var layerTextStrings = [];
  for (var i = 0; i < textLayers.length; ++i) {
    layerIndices.push(i);
    layerTextStrings.push(getLayerText(textLayers[i]).replace(WHITESPACE_REGEX, ""));
  }
  sourceDocument.close();

  var textLayerMap = [];
  var missingText = [];
  for (var dataRow = 0; dataRow < worksheetData.length; ++dataRow) {
    var trimmedSourceText = worksheetData[dataRow][sourceLang].replace(WHITESPACE_REGEX, "");
    var layerMatchIndex = findMatchingTextLayer(layerTextStrings, layerIndices, trimmedSourceText)
    textLayerMap.push(layerMatchIndex);
    if (layerMatchIndex === -1) {
      missingText.push(worksheetData[dataRow][sourceLang]);
    }
  }

  if (missingText.length > 0) {
    if (!confirm(localize(UI_MESSAGES.MISSING_SHEET_ROWS, "-" + missingText.join("\n-")))) {
      return null;
    }
  }

  return textLayerMap;

}

/**
 * Copy source CC file and replace text for all selected target languages
 * @param {File} ccFile 
 * @param {Object} selectedLanguages 
 * @param {Array} worksheetData 
 */
function generateLocalizedFiles(ccFile, selectedLanguages, worksheetData, textLayerMap) {
  // Generated localized files
  for (var langIndex = 0; langIndex < selectedLanguages.target.length; ++langIndex) {
    var currentLang = selectedLanguages.target[langIndex];
    var targetFilename = generateTargetFilename(ccFile.name, currentLang);
    var targetFile = new File(ccFile.path + "/" + targetFilename);
    ccFile.copy(targetFile);
    var targetDoc = app.open(targetFile);

    var textLayers = getDocumentTextLayers(targetDoc);
    for (var dataRow = 0; dataRow < worksheetData.length; ++dataRow) {
      var layerIndex = textLayerMap[dataRow];
      if (layerIndex !== -1) {
        setLayerText(textLayers[layerIndex], worksheetData[dataRow][currentLang]);
      }
    }

    targetDoc.save();
    targetDoc.close();
  }
}

/**
 * main function
 * @returns Nothing
 */
function main() {
  const ccFile = getCCFile();
  if (ccFile === null) {
    return;
  }

  const pictureListFile = getPictureListFile();
  if (pictureListFile === null) {
    return;
  }

  if (!confirm(localize(UI_MESSAGES.FILE_LOAD_CONFIRM, ccFile.name, pictureListFile.name))) {
    return;
  }

  var workbook;
  try {
    workbook = XLSX.readFile(pictureListFile);
  } catch (err) {
    alert(localize(UI_MESSAGES.PICTURE_LIST_LOAD_ERROR, err.message));
    return;
  }

  var worksheetData = loadWorksheetData(workbook);
  if (worksheetData === null) {
    return;
  }

  // Get list of languages
  var languages = Object.keys(worksheetData[0]);
  var rowNumPropIndex = languages.indexOf("__rowNum__")
  if (rowNumPropIndex !== -1) {
    languages.splice(rowNumPropIndex, 1);
  }

  if (languages.length < 2) {
    alert(UI_MESSAGES.NO_TARGET_LANGS_FOUND);
    return;
  }

  var selectedLanguages = userSelectLanguages(languages);
  if (selectedLanguages.source === null) {
    return;
  }
  if (selectedLanguages.target.length === 0) {
    alert(UI_MESSAGES.NO_TARGET_LANGS_SELECTED);
    return;
  }

  var textLayerMap = mapTextLayers(ccFile, worksheetData, selectedLanguages.source);
  if (textLayerMap === null) {
    return;
  }
  
  try {
    generateLocalizedFiles(ccFile, selectedLanguages, worksheetData, textLayerMap);
    alert(localize(UI_MESSAGES.FILE_LOCALIZE_SUCCESS, selectedLanguages.target.length));
  } catch (err) {
    alert(localize(UI_MESSAGES.FILE_LOCALIZE_ERROR, err.description));
  }
}

main();

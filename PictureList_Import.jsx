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


/**
 * Opens Adobe CC file selection dialog
 * @returns File
 */
function getCCFile() {
  if (app.documents.length === 0) {
    var filetypes = appFiletypes[app.name];
    if (filetypes === undefined) {
      alert("Unsupported Creative Cloud app");
      return null;
    }
    filetypes.push("All Formats:*.*");
    return File.openDialog("Select Creative Cloud File", filetypes.join(","));
  } else {
    return app.activeDocument.fullName;
  }
}

/**
 * Opens picture list file selection dialog
 * @returns File
 */
function getPictureListFile() {
  return File.openDialog("Select Picture List File", "Excel Files:*.xlsx;*.xls,All Formats:*.*");
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
  var langConfigWindow = new Window("dialog", "Select source/target languages");
  langConfigWindow.alignChildren = "left";
  var srcLangGroup = langConfigWindow.add("group {_: StaticText {text: \"Source language:\"}}");
  var srcLangDropdown = srcLangGroup.add("dropdownlist", undefined, languages);
  srcLangDropdown.selection = 0;
  srcLangDropdown.currentLang = languages[0];

  var targetLangGroup = langConfigWindow.add("group {_: StaticText {text: \"Target languages:\"}}");
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
  buttonGroup.add("button", undefined, "OK");
  buttonGroup.add("button", undefined, "Cancel");
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
      //alert(document.layers.length + " document layers")
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
 * Find layer/frame with matching text content
 * @param {[Layer or TextFrameItem]} textLayers 
 * @param {String} searchText 
 * @returns Layer/TextFrameItem
 */
function findMatchingTextLayer(textLayers, searchText) {  //TODO: Sort/otherwise optimize search
  var whitespaceRegex = /\s/g;
  for (var i = 0; i < textLayers.length; ++i) {
    var layerText = getLayerText(textLayers[i]).replace(whitespaceRegex, "");
    if (getLayerText(textLayers[i]).replace(whitespaceRegex, "") === searchText.replace(whitespaceRegex, "")) {
      return i;
    }
  }

  return -1;
}

/**
 * Copy source CC file and replace text for all selected target languages
 * @param {File} ccFile 
 * @param {Object} selectedLanguages 
 * @param {Array} worksheetData 
 */
function generateLocalizedFiles(ccFile, selectedLanguages, worksheetData) {
  // Generated localized files
  var dataLayerIndices = [];

  for (var langIndex = 0; langIndex < selectedLanguages.target.length; ++langIndex) {
    var currentLang = selectedLanguages.target[langIndex];
    var targetFilename = generateTargetFilename(ccFile.name, currentLang);
    var targetFile = new File(ccFile.path + "/" + targetFilename);
    ccFile.copy(targetFile);
    var targetDoc = app.open(targetFile);

    var textLayers = getDocumentTextLayers(targetDoc);
    for (var dataRow = 0; dataRow < worksheetData.length; ++dataRow) {
      // Build index for subsequent target documents
      if (dataLayerIndices.length < dataRow + 1) {
        var matchingLayerIndex = findMatchingTextLayer(textLayers, worksheetData[dataRow][selectedLanguages.source]);
        dataLayerIndices.push(matchingLayerIndex);

        if (matchingLayerIndex === -1) {
          // TODO: Log error for missing text layer
        }
      }

      var layerIndex = dataLayerIndices[dataRow];
      if (layerIndex !== -1) {
        setLayerText(textLayers[layerIndex], worksheetData[dataRow][currentLang]);
      }
    }

    targetDoc.save();
    targetDoc.close();
  }
}

/**
 * Opens worksheet selection dialog and loads selected worksheet JSON
 * @param {Workbook object} workbook 
 * @returns 
 */
function loadWorksheetData(workbook) {
  // Worksheet selection
  var sheetConfigWindow = new Window("dialog", "Select worksheet");
  var sheetGroup = sheetConfigWindow.add("group {_: StaticText {text: \"Worksheet:\"}}");
  var sheetDropdown = sheetGroup.add("dropdownlist", undefined, workbook.SheetNames);
  sheetDropdown.selection = 0;
  
  var buttonGroup = sheetConfigWindow.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.add("button", undefined, "OK");
  buttonGroup.add("button", undefined, "Cancel");
  if (sheetConfigWindow.show() === 2) {
    return null;
  }

  // Convert sheet into JSON
  var worksheet = workbook.Sheets[sheetDropdown.selection];
  return XLSX.utils.sheet_to_json(worksheet, {raw: false});
}


function main() {
  const ccFile = getCCFile();
  if (ccFile === null) {
    return;
  }

  const pictureListFile = getPictureListFile();
  if (pictureListFile === null) {
    return;
  }

  if (!confirm("Creative Cloud file: " + ccFile.name +
      "\n\nPicture list: " + pictureListFile.name +
      "\n\nLoad files? This operation may take a while.")) {
    return;
  }

  var workbook;
  try {
    workbook = XLSX.readFile(pictureListFile);
  } catch (err) {
    alert("Failed to open picture list: " + err.message);
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
    alert("No target languages found");
    return;
  }

  var selectedLanguages = userSelectLanguages(languages);
  if (selectedLanguages.source === null) {
    return;
  }
  if (selectedLanguages.target.length === 0) {
    alert("No target languages selected");
    return;
  }

  generateLocalizedFiles(ccFile, selectedLanguages, worksheetData);
}

main();

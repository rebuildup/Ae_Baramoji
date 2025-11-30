/*
The MIT License

Copyright (c) 2024 Nisai(original)
Copyright (c) 2025 Modified version by 361do_sleep

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。
ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
*/
(function () {
  var progressWindow = null;
  var progressBar = null;
  var statusText = null;

  function showProgressWindow() {
    if (progressWindow == null || !progressWindow.visible) {
      progressWindow = new Window(
        "palette",
        "Text decomposition in progress",
        undefined
      );
      progressWindow.orientation = "column";
      progressWindow.alignChildren = ["fill", "center"];

      progressBar = progressWindow.add("progressbar", undefined, 0, 100);
      progressBar.preferredSize.width = 300;

      statusText = progressWindow.add(
        "statictext",
        [0, 0, 300, 30],
        "Preparing..."
      );
      statusText.alignment = "center";

      progressWindow.updateProgress = function (targetValue, text) {
        if (!progressBar) return;
        var currentValue = Math.round(progressBar.value) || 0;
        var clampedTarget = Math.max(
          0,
          Math.min(100, Math.round(targetValue || 0))
        );
        var nextValue = Math.max(currentValue, clampedTarget);
        progressBar.value = nextValue;
        statusText.text = text || String(nextValue) + "% complete";
        try {
          progressWindow.update();
        } catch (e) { }
      };

      progressWindow.center();
      progressWindow.show();
    }
  }

  function updateProgress(value, text) {
    if (progressWindow && progressWindow.updateProgress) {
      progressWindow.updateProgress(value, text);
    }
  }

  function closeProgressWindow() {
    if (progressWindow) {
      try {
        progressWindow.close();
      } catch (e) { }
      progressWindow = null;
      progressBar = null;
      statusText = null;
    }
  }

  try {
    app.beginUndoGroup("DecomposeTextLayers");

    showProgressWindow();
    updateProgress(0, "Initializing...");

    var comp = app.project.activeItem;
    if (!comp || !(comp instanceof CompItem)) {
      alert(
        "No composition is active. Please open a composition and select text layers."
      );
      closeProgressWindow();
      app.endUndoGroup();
      return;
    }

    var curTime = comp.time;
    var selLayers = comp.selectedLayers;
    if (!selLayers || selLayers.length === 0) {
      alert("No layers selected. Please select one or more text layers.");
      closeProgressWindow();
      app.endUndoGroup();
      return;
    }

    updateProgress(2, "Inspecting layers...");

    function degToRad(d) {
      return d * (Math.PI / 180);
    }

    function rotationToMatrix(rot) {
      var rx = degToRad(rot[0]);
      var ry = degToRad(rot[1]);
      var rz = degToRad(rot[2]);

      var cosX = Math.cos(rx),
        sinX = Math.sin(rx);
      var cosY = Math.cos(ry),
        sinY = Math.sin(ry);
      var cosZ = Math.cos(rz),
        sinZ = Math.sin(rz);

      return [
        [cosY * cosZ, -cosY * sinZ, sinY],
        [
          sinX * sinY * cosZ + cosX * sinZ,
          -sinX * sinY * sinZ + cosX * cosZ,
          -sinX * cosY,
        ],
        [
          -cosX * sinY * cosZ + sinX * sinZ,
          cosX * sinY * sinZ + sinX * cosZ,
          cosX * cosY,
        ],
      ];
    }

    function multiplyMatrix3x3(a, b) {
      var r = [];
      for (var i = 0; i < 3; i++) {
        r[i] = [];
        for (var j = 0; j < 3; j++) {
          r[i][j] = a[i][0] * b[0][j] + a[i][1] * b[1][j] + a[i][2] * b[2][j];
        }
      }
      return r;
    }

    function applyMatrixToOffset(matrix, xOffset, yOffset, zOffset) {
      return [
        xOffset * matrix[0][0] +
        yOffset * matrix[0][1] +
        zOffset * matrix[0][2],
        xOffset * matrix[1][0] +
        yOffset * matrix[1][1] +
        zOffset * matrix[1][2],
        xOffset * matrix[2][0] +
        yOffset * matrix[2][1] +
        zOffset * matrix[2][2],
      ];
    }

    function setPosition(transformProperty, newPosition) {
      if (!transformProperty) return;
      if (transformProperty.numKeys === 0) {
        transformProperty.setValue(newPosition);
      } else {
        try {
          transformProperty.setValue(newPosition);
        } catch (e) { }
      }
    }

    function getPropertyArray(property, layer) {
      var txtSave = layer.property("Source Text").value;
      var expression =
        "var output_txt=[]; for(var i=0;i<text.sourceText.length;i++){output_txt.push(text.sourceText.getStyleAt(i)." +
        property +
        ");}output_txt;";
      layer.property("Source Text").expression = expression;
      var output = layer.property("Source Text").value;
      layer.property("Source Text").expression = "";
      layer.text.sourceText.setValue(txtSave);
      return String(output);
    }

    function getLayerSize(layer) {
      var bounds = layer.sourceRectAtTime(app.project.activeItem.time, false);
      return [bounds.width, bounds.height];
    }

    function hasDeepGlow(layer) {
      try {
        var ef = layer.property("ADBE Effect Parade");
        if (!ef) return false;
        if (ef.property("PEDG") || ef.property("PEDG2")) return true;
      } catch (e) { }
      return false;
    }

    function adjustLayerPosition(layer, posX, posY) {
      try {
        var layerBounds = layer.sourceRectAtTime(curTime, false);
        var left = layerBounds.left,
          top = layerBounds.top,
          width = layerBounds.width,
          height = layerBounds.height;
        layer.transform.anchorPoint.setValue([
          width / 2 + left,
          height / 2 + top,
        ]);
      } catch (e) { }
      try {
        if (!layer.transform.position.dimensionsSeparated) {
          setPosition(layer.transform("ADBE Position"), [posX, posY]);
        } else {
          setPosition(layer.transform("ADBE Position_0"), posX);
          setPosition(layer.transform("ADBE Position_1"), posY);
        }
      } catch (e) {
        try {
          layer.transform.position.setValue([posX, posY]);
        } catch (e2) { }
      }
    }

    for (var layerIdx = 0; layerIdx < selLayers.length; layerIdx++) {
      updateProgress(
        Math.round(2 + (layerIdx / selLayers.length) * 3),
        "Processing layer " + (layerIdx + 1) + "/" + selLayers.length + "..."
      );

      var textLayer = selLayers[layerIdx];
      if (!(textLayer instanceof TextLayer)) continue;

      if (hasDeepGlow(textLayer)) {
        alert(
          "Please temporarily remove the Deep Glow effect from the layer: " +
          textLayer.name
        );
        continue;
      }

      updateProgress(6, "Loading layer info...");

      var originalScale = [100, 100];
      try {
        var s = textLayer.transform.scale.value;
        if (s.length === 2) originalScale = [s[0], s[1]];
        else if (s.length === 3) originalScale = [s[0], s[1]];
      } catch (e) { }

      var originalSolo = false;
      try {
        originalSolo = textLayer.solo === true;
      } catch (e) { }

      var layerInPoint = textLayer.inPoint;
      var layerOutPoint = textLayer.outPoint;

      var originalText = textLayer.text.sourceText.value;
      var textContent = String(originalText);

      // keep newline characters so style arrays stay aligned with character indices
      var cleanedForChars = textContent.replace(/\u0003/g, "");
      updateProgress(8, "Retrieving text styles...");
      var textFontSize = getPropertyArray("fontSize", textLayer).split(",");
      var textFont = getPropertyArray("font", textLayer).split(",");
      var textApplyFill = getPropertyArray("applyFill", textLayer).split(",");
      var textFillColor_txt = getPropertyArray("fillColor", textLayer).split(
        ","
      );
      var textApplyStroke = getPropertyArray("applyStroke", textLayer).split(
        ","
      );
      var textStrokeColor_txt = getPropertyArray(
        "strokeColor",
        textLayer
      ).split(",");
      var textStrokeWidth = getPropertyArray("strokeWidth", textLayer).split(
        ","
      );

      var textTracking = getPropertyArray("tracking", textLayer).split(",");
      var textBaselineShift = getPropertyArray(
        "baselineShift",
        textLayer
      ).split(",");
      //var textAllCaps = getPropertyArray("allCaps", textLayer).split(",");
      //var textSmallCaps = getPropertyArray("smallCaps", textLayer).split(",");
      var textStrokeOverFill = getPropertyArray(
        "strokeOverFill",
        textLayer
      ).split(",");
      var textFauxBold = getPropertyArray("fauxBold", textLayer).split(",");
      var textFauxItalic = getPropertyArray("fauxItalic", textLayer).split(",");
      var textTsume = getPropertyArray("tsume", textLayer).split(",");

      updateProgress(12, "Formatting color info...");
      var textFillColor = [];
      for (var i = 0; i < textFillColor_txt.length; i += 3) {
        textFillColor.push(textFillColor_txt.slice(i, i + 3));
      }
      var textStrokeColor = [];
      for (var i = 0; i < textStrokeColor_txt.length; i += 3) {
        textStrokeColor.push(textStrokeColor_txt.slice(i, i + 3));
      }

      updateProgress(16, "Handling whitespace...");
      var specialCharRegex = /[\s\n\r]/g;
      var specialCharIndices = [];
      var text = cleanedForChars;
      for (var ci = 0; ci < text.length; ci++) {
        if (specialCharRegex.test(text[ci])) specialCharIndices.push(ci);
      }
      var cleanText = text.replace(/\s+/g, "");
      for (var j = specialCharIndices.length - 1; j >= 0; j--) {
        var idx = specialCharIndices[j];
        if (textFontSize.length > idx) textFontSize.splice(idx, 1);
        if (textFont.length > idx) textFont.splice(idx, 1);
        if (textApplyFill.length > idx) textApplyFill.splice(idx, 1);
        if (textFillColor.length > idx) textFillColor.splice(idx, 1);
        if (textApplyStroke.length > idx) textApplyStroke.splice(idx, 1);
        if (textStrokeColor.length > idx) textStrokeColor.splice(idx, 1);
        if (textStrokeWidth.length > idx) textStrokeWidth.splice(idx, 1);

        if (textTracking.length > idx) textTracking.splice(idx, 1);
        if (textBaselineShift.length > idx) textBaselineShift.splice(idx, 1);
        //if (textAllCaps.length > idx) textAllCaps.splice(idx, 1);
        //if (textSmallCaps.length > idx) textSmallCaps.splice(idx, 1);
        if (textStrokeOverFill.length > idx) textStrokeOverFill.splice(idx, 1);
        if (textFauxBold.length > idx) textFauxBold.splice(idx, 1);
        if (textFauxItalic.length > idx) textFauxItalic.splice(idx, 1);
        if (textTsume.length > idx) textTsume.splice(idx, 1);
      }

      updateProgress(22, "Converting text to shapes...");
      for (var sdel = 0; sdel < comp.selectedLayers.length; sdel++)
        comp.selectedLayers[sdel].selected = false;
      textLayer.selected = true;
      app.executeCommand(3781);

      var shapeLayer = comp.selectedLayers[0];
      if (!shapeLayer) {
        alert("Failed to create shapes from text for layer: " + textLayer.name);
        continue;
      }

      var shapeContents = shapeLayer.property("Contents");
      updateProgress(30, "Duplicating shape layers...");
      for (var p = shapeContents.numProperties - 1; p > 0; p--) {
        var duplicatedShape = comp.selectedLayers[0].duplicate();
        duplicatedShape.selected = true;
      }
      var allShapes = comp.selectedLayers;

      updateProgress(36, "Refining shapes...");
      for (var si = 0; si < allShapes.length; si++) {
        var curShape = allShapes[si];
        curShape.enabled = false;
        var curContents = curShape.property("Contents");
        for (var pi = curContents.numProperties; pi > 0; pi--) {
          if (pi !== si + 1) {
            try {
              curContents.property(pi).remove();
            } catch (e) { }
          }
        }
      }

      var shapeAnchorX = [],
        shapeAnchorY = [],
        shapePositionX = [],
        shapePositionY = [];
      for (var s = 0; s < allShapes.length; s++) {
        try {
          var cur = allShapes[s];
          var shapeRot = cur.transform.rotation
            ? cur.transform.rotation.value
            : 0;
          var shapeAnchor = cur.transform.anchorPoint.value;
          var shapeBounds = cur.sourceRectAtTime(curTime, false);
          var centerX = shapeBounds.width / 2 + shapeBounds.left;
          var centerY = shapeBounds.height / 2 + shapeBounds.top;

          var offsetX = (centerX - shapeAnchor[0]) * (originalScale[0] / 100);
          var offsetY = (centerY - shapeAnchor[1]) * (originalScale[1] / 100);
          var cosR = Math.cos(shapeRot * (Math.PI / 180));
          var sinR = Math.sin(shapeRot * (Math.PI / 180));
          var shpPos = cur.transform.position.value;

          shapeAnchorX.push(centerX);
          shapeAnchorY.push(centerY);
          shapePositionX.push(shpPos[0] + offsetX * cosR - offsetY * sinR);
          shapePositionY.push(shpPos[1] + offsetX * sinR + offsetY * cosR);
        } catch (e) {
          shapeAnchorX.push(0);
          shapeAnchorY.push(0);
          var baseP = textLayer.transform.position.value;
          shapePositionX.push(baseP[0]);
          shapePositionY.push(baseP[1]);
        }
        updateProgress(
          36 + Math.round((s / Math.max(1, allShapes.length)) * 10),
          "Analyzing shape " + (s + 1) + "/" + allShapes.length + "..."
        );
      }

      updateProgress(48, "Creating character layers...");
      var resultLayers = [];
      for (var ci = cleanText.length - 1; ci >= 0; ci--) {
        var characterLayer = textLayer.duplicate();
        characterLayer.enabled = true;
        characterLayer.name = cleanText[ci];
        try {
          characterLayer.solo = originalSolo;
        } catch (e) { }
        resultLayers.unshift(characterLayer);
      }

      for (var charIndex = 0; charIndex < cleanText.length; charIndex++) {
        updateProgress(
          50 + Math.round((charIndex / Math.max(1, cleanText.length)) * 40),
          "Processing character " +
          (charIndex + 1) +
          "/" +
          cleanText.length +
          "..."
        );

        var characterLayer = resultLayers[charIndex];

        var charTextDocument = textLayer.text.sourceText.valueAtTime(
          curTime,
          true
        );
        charTextDocument.text = cleanText[charIndex];

        try {
          if (charIndex < textFontSize.length) {
            if (textFontSize[charIndex] !== "")
              charTextDocument.fontSize = Number(textFontSize[charIndex]);
            if (textFont[charIndex] !== "")
              charTextDocument.font = textFont[charIndex];

            if (textApplyFill[charIndex] === "true") {
              charTextDocument.applyFill = true;
              if (textFillColor[charIndex])
                charTextDocument.fillColor = textFillColor[charIndex];
            } else {
              charTextDocument.applyFill = false;
            }

            if (textApplyStroke[charIndex] === "true") {
              charTextDocument.applyStroke = true;
              if (textStrokeColor[charIndex])
                charTextDocument.strokeColor = textStrokeColor[charIndex];
              if (textStrokeWidth[charIndex] !== "")
                charTextDocument.strokeWidth = Number(
                  textStrokeWidth[charIndex]
                );
            } else {
              charTextDocument.applyStroke = false;
            }

            if (
              textTracking[charIndex] !== undefined &&
              textTracking[charIndex] !== ""
            )
              charTextDocument.tracking = Number(textTracking[charIndex]);
            if (
              textBaselineShift[charIndex] !== undefined &&
              textBaselineShift[charIndex] !== ""
            )
              charTextDocument.baselineShift = Number(
                textBaselineShift[charIndex]
              );
            /*
            if (textAllCaps[charIndex] !== undefined)
              charTextDocument.allCaps = textAllCaps[charIndex] === "true";
            if (textSmallCaps[charIndex] !== undefined)
              charTextDocument.smallCaps = textSmallCaps[charIndex] === "true";
            */
            if (textStrokeOverFill[charIndex] !== undefined)
              charTextDocument.strokeOverFill =
                textStrokeOverFill[charIndex] === "true";
            if (textFauxBold[charIndex] !== undefined)
              charTextDocument.fauxBold = textFauxBold[charIndex] === "true";
            if (textFauxItalic[charIndex] !== undefined)
              charTextDocument.fauxItalic =
                textFauxItalic[charIndex] === "true";
            if (
              textTsume[charIndex] !== undefined &&
              textTsume[charIndex] !== ""
            )
              charTextDocument.tsume = Number(textTsume[charIndex]);
          }
        } catch (styleError) { }

        characterLayer.text.sourceText.setValue(charTextDocument);

        try {
          var tSize = getLayerSize(characterLayer);
          var sSize = getLayerSize(allShapes[charIndex]);

          var originalHScale = charTextDocument.horizontalScale || 100;
          var originalVScale = charTextDocument.verticalScale || 100;

          var newHScale = (sSize[0] / (tSize[0] || 1)) * originalHScale;
          var newVScale = (sSize[1] / (tSize[1] || 1)) * originalVScale;

          charTextDocument.horizontalScale = Math.min(
            Math.max(newHScale, 0),
            1000
          );
          charTextDocument.verticalScale = Math.min(
            Math.max(newVScale, 0),
            1000
          );

          characterLayer.text.sourceText.setValue(charTextDocument);
        } catch (e) { }

        try {
          var lb = characterLayer.sourceRectAtTime(curTime, false);
          var charAnchorLocal = [
            lb.width / 2 + lb.left,
            lb.height / 2 + lb.top,
          ];

          try {
            characterLayer.transform.anchorPoint.setValue(charAnchorLocal);
          } catch (e) { }

          var textAP = textLayer.transform.anchorPoint.value;

          var sx = shapeAnchorX[charIndex];
          var sy = shapeAnchorY[charIndex];

          var localOffsetX = (sx - textAP[0]) * (originalScale[0] / 100);
          var localOffsetY = (sy - textAP[1]) * (originalScale[1] / 100);
          var localOffsetZ = 0;

          if (!textLayer.threeDLayer) {
            var rotZ = textLayer.transform.rotation
              ? textLayer.transform.rotation.value
              : 0;
            var cosR = Math.cos(degToRad(rotZ));
            var sinR = Math.sin(degToRad(rotZ));
            var rotX = localOffsetX * cosR - localOffsetY * sinR;
            var rotY = localOffsetX * sinR + localOffsetY * cosR;
            var basePos = textLayer.transform.position.value;
            var finalX = basePos[0] + rotX;
            var finalY = basePos[1] + rotY;
            try {
              if (!characterLayer.transform.position.dimensionsSeparated) {
                setPosition(characterLayer.transform("ADBE Position"), [
                  finalX,
                  finalY,
                ]);
              } else {
                setPosition(
                  characterLayer.transform("ADBE Position_0"),
                  finalX
                );
                setPosition(
                  characterLayer.transform("ADBE Position_1"),
                  finalY
                );
              }
            } catch (e) {
              characterLayer.transform.position.setValue([finalX, finalY]);
            }
          } else {
            var ori = [0, 0, 0];
            try {
              ori = textLayer.transform.orientation.value;
            } catch (e) { }
            var xRot = textLayer.transform.xRotation
              ? textLayer.transform.xRotation.value
              : 0;
            var yRot = textLayer.transform.yRotation
              ? textLayer.transform.yRotation.value
              : 0;
            var zRot = textLayer.transform.zRotation
              ? textLayer.transform.zRotation.value
              : 0;
            var rotationVals = [xRot, yRot, zRot];

            var oriMat = rotationToMatrix(ori);
            var rotMat = rotationToMatrix(rotationVals);
            var finalMat = multiplyMatrix3x3(oriMat, rotMat);

            var applied = applyMatrixToOffset(
              finalMat,
              localOffsetX,
              localOffsetY,
              localOffsetZ
            );

            var basePos = textLayer.transform.position.value;
            var finalPos3 = [
              basePos[0] + applied[0],
              basePos[1] + applied[1],
              (basePos.length > 2 ? basePos[2] : 0) + (applied[2] || 0),
            ];

            try {
              if (!characterLayer.transform.position.dimensionsSeparated) {
                setPosition(
                  characterLayer.transform("ADBE Position"),
                  finalPos3
                );
              } else {
                setPosition(
                  characterLayer.transform("ADBE Position_0"),
                  finalPos3[0]
                );
                setPosition(
                  characterLayer.transform("ADBE Position_1"),
                  finalPos3[1]
                );
                if (characterLayer.transform("ADBE Position_2"))
                  setPosition(
                    characterLayer.transform("ADBE Position_2"),
                    finalPos3[2]
                  );
              }
            } catch (e) {
              try {
                characterLayer.transform.position.setValue(finalPos3);
              } catch (e2) { }
            }
          }
        } catch (posError) {
          try {
            characterLayer.transform.position.setValue(
              textLayer.transform.position.value
            );
          } catch (e2) { }
        }

        characterLayer.inPoint = layerInPoint;
        characterLayer.outPoint = layerOutPoint;
      }

      try {
        textLayer.enabled = false;
      } catch (e) { }

      for (var rem = 0; rem < allShapes.length; rem++) {
        try {
          allShapes[rem].remove();
        } catch (e) { }
      }

      try {
        for (var ss = 0; ss < comp.selectedLayers.length; ss++)
          comp.selectedLayers[ss].selected = false;
      } catch (e) { }
      for (var rr = 0; rr < resultLayers.length; rr++) {
        try {
          resultLayers[rr].selected = true;
        } catch (e) { }
      }

      updateProgress(
        Math.round(50 + ((layerIdx + 1) / selLayers.length) * 40),
        "Layer " + (layerIdx + 1) + "/" + selLayers.length + " completed"
      );
    }

    updateProgress(95, "Finalizing...");
    $.sleep(200);
    updateProgress(100, "Completed!");
    $.sleep(300);
    closeProgressWindow();

    app.endUndoGroup();
  } catch (err) {
    try {
      app.endUndoGroup();
    } catch (e) { }
    closeProgressWindow();
    alert("Error: " + (err && err.toString ? err.toString() : err));
  }
})();

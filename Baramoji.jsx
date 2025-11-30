/*
The MIT License

Copyright (c) 2025 361do_sleep

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。
ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
*/

(function BaramojiUnifiedEntry(thisObj) {
  function ProgressDialog(title) {
    this.win = new Window("palette", title || "Processing...", undefined, {
      resizeable: false,
    });
    this.win.orientation = "column";
    this.win.alignChildren = ["fill", "top"];
    this.msg = this.win.add("statictext", undefined, "Ready");
    this.msg.alignment = ["fill", "top"];
    this.bar = this.win.add("progressbar", undefined, 0, 100);
    this.bar.preferredSize = [280, 16];
    this.setMax = function (max) {
      try {
        this.bar.maxvalue = Math.max(1, max);
        this.bar.value = 0;
      } catch (e) { }
    };
    this.step = function (text) {
      try {
        if (text) this.msg.text = text;
        this.bar.value = Math.min(this.bar.value + 1, this.bar.maxvalue);
        this.win.update();
      } catch (e) { }
    };
    this.show = function () {
      try {
        this.win.show();
        this.win.update();
      } catch (e) { }
    };
    this.close = function () {
      try {
        this.win.close();
      } catch (e) { }
    };
  }

  function runDecomposeTextToTextLayers(progress) {
    try {
      app.beginUndoGroup("DecomposeTextLayers");

      var comp = app.project.activeItem;
      if (!comp || !(comp instanceof CompItem)) {
        alert(
          "No composition is active. Please open a composition and select text layers."
        );
        app.endUndoGroup();
        return;
      }

      var curTime = comp.time;
      var selLayers = comp.selectedLayers;
      if (!selLayers || selLayers.length === 0) {
        alert("No layers selected. Please select one or more text layers.");
        app.endUndoGroup();
        return;
      }

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

      var totalCharsEstimate = 0;
      for (var li = 0; li < selLayers.length; li++) {
        var tl = selLayers[li];
        if (tl instanceof TextLayer) {
          var txt = String(tl.text.sourceText.value)
            .replace(/\r|\n|\u0003/g, "")
            .replace(/\s+/g, "");
          totalCharsEstimate += Math.max(1, txt.length);
        }
      }
      var overheadPerLayer = 6;
      var overheadTotal = Math.max(0, selLayers.length * overheadPerLayer);
      var prog = progress || new ProgressDialog("Decompose: Text → Texts");
      prog.setMax(Math.max(1, totalCharsEstimate + overheadTotal));
      if (!progress) prog.show();

      for (var layerIdx = 0; layerIdx < selLayers.length; layerIdx++) {
        var textLayer = selLayers[layerIdx];
        if (!(textLayer instanceof TextLayer)) continue;

        if (hasDeepGlow(textLayer)) {
          alert(
            "Please temporarily remove the Deep Glow effect from the layer: " +
            textLayer.name
          );
          continue;
        }

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
        var textStrokeOverFill = getPropertyArray(
          "strokeOverFill",
          textLayer
        ).split(",");
        var textFauxBold = getPropertyArray("fauxBold", textLayer).split(",");
        var textFauxItalic = getPropertyArray("fauxItalic", textLayer).split(
          ","
        );
        var textTsume = getPropertyArray("tsume", textLayer).split(",");

        var textFillColor = [];
        for (var i = 0; i < textFillColor_txt.length; i += 3)
          textFillColor.push(textFillColor_txt.slice(i, i + 3));
        var textStrokeColor = [];
        for (var j = 0; j < textStrokeColor_txt.length; j += 3)
          textStrokeColor.push(textStrokeColor_txt.slice(j, j + 3));

        var specialCharRegex = /[\s\n\r]/g;
        var specialCharIndices = [];
        var text = cleanedForChars;
        for (var ci = 0; ci < text.length; ci++)
          if (specialCharRegex.test(text[ci])) specialCharIndices.push(ci);
        var cleanText = text.replace(/\s+/g, "");
        for (var k = specialCharIndices.length - 1; k >= 0; k--) {
          var idx = specialCharIndices[k];
          if (textFontSize.length > idx) textFontSize.splice(idx, 1);
          if (textFont.length > idx) textFont.splice(idx, 1);
          if (textApplyFill.length > idx) textApplyFill.splice(idx, 1);
          if (textFillColor.length > idx) textFillColor.splice(idx, 1);
          if (textApplyStroke.length > idx) textApplyStroke.splice(idx, 1);
          if (textStrokeColor.length > idx) textStrokeColor.splice(idx, 1);
          if (textStrokeWidth.length > idx) textStrokeWidth.splice(idx, 1);
          if (textTracking.length > idx) textTracking.splice(idx, 1);
          if (textBaselineShift.length > idx) textBaselineShift.splice(idx, 1);
          if (textStrokeOverFill.length > idx)
            textStrokeOverFill.splice(idx, 1);
          if (textFauxBold.length > idx) textFauxBold.splice(idx, 1);
          if (textFauxItalic.length > idx) textFauxItalic.splice(idx, 1);
          if (textTsume.length > idx) textTsume.splice(idx, 1);
        }

        for (var sdel = 0; sdel < comp.selectedLayers.length; sdel++)
          comp.selectedLayers[sdel].selected = false;
        textLayer.selected = true;
        prog.step("Converting text to shapes...");
        app.executeCommand(3781);

        var shapeLayer = comp.selectedLayers[0];
        if (!shapeLayer) {
          alert(
            "Failed to create shapes from text for layer: " + textLayer.name
          );
          continue;
        }

        var shapeContents = shapeLayer.property("Contents");
        for (var p = shapeContents.numProperties - 1; p > 0; p--) {
          var duplicatedShape = comp.selectedLayers[0].duplicate();
          duplicatedShape.selected = true;
        }
        prog.step("Duplicated shape layers");
        var allShapes = comp.selectedLayers;

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
        prog.step("Isolated shape groups");

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
        }

        var resultLayers = [];
        for (var ci2 = cleanText.length - 1; ci2 >= 0; ci2--) {
          var characterLayer = textLayer.duplicate();
          characterLayer.enabled = true;
          characterLayer.name = cleanText[ci2];
          try {
            characterLayer.solo = originalSolo;
          } catch (e) { }
          resultLayers.unshift(characterLayer);
        }
        prog.step("Created character layers");

        for (var charIndex = 0; charIndex < cleanText.length; charIndex++) {
          var characterLayer2 = resultLayers[charIndex];
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

          characterLayer2.text.sourceText.setValue(charTextDocument);

          try {
            var tSize = getLayerSize(characterLayer2);
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
            characterLayer2.text.sourceText.setValue(charTextDocument);
          } catch (e) { }

          try {
            var lb = characterLayer2.sourceRectAtTime(curTime, false);
            var charAnchorLocal = [
              lb.width / 2 + lb.left,
              lb.height / 2 + lb.top,
            ];
            try {
              characterLayer2.transform.anchorPoint.setValue(charAnchorLocal);
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
                if (!characterLayer2.transform.position.dimensionsSeparated) {
                  setPosition(characterLayer2.transform("ADBE Position"), [
                    finalX,
                    finalY,
                  ]);
                } else {
                  setPosition(
                    characterLayer2.transform("ADBE Position_0"),
                    finalX
                  );
                  setPosition(
                    characterLayer2.transform("ADBE Position_1"),
                    finalY
                  );
                }
              } catch (e) {
                characterLayer2.transform.position.setValue([finalX, finalY]);
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
              var basePos3 = textLayer.transform.position.value;
              var finalPos3 = [
                basePos3[0] + applied[0],
                basePos3[1] + applied[1],
                (basePos3.length > 2 ? basePos3[2] : 0) + (applied[2] || 0),
              ];
              try {
                if (!characterLayer2.transform.position.dimensionsSeparated) {
                  setPosition(
                    characterLayer2.transform("ADBE Position"),
                    finalPos3
                  );
                } else {
                  setPosition(
                    characterLayer2.transform("ADBE Position_0"),
                    finalPos3[0]
                  );
                  setPosition(
                    characterLayer2.transform("ADBE Position_1"),
                    finalPos3[1]
                  );
                  if (characterLayer2.transform("ADBE Position_2"))
                    setPosition(
                      characterLayer2.transform("ADBE Position_2"),
                      finalPos3[2]
                    );
                }
              } catch (e) {
                try {
                  characterLayer2.transform.position.setValue(finalPos3);
                } catch (e2) { }
              }
            }
          } catch (posError) {
            try {
              characterLayer2.transform.position.setValue(
                textLayer.transform.position.value
              );
            } catch (e2) { }
          }

          characterLayer2.inPoint = layerInPoint;
          characterLayer2.outPoint = layerOutPoint;
          try {
            characterLayer2.solo = originalSolo;
          } catch (e) { }

          prog.step("Characters processed: " + (charIndex + 1));
        }

        try {
          textLayer.enabled = false;
        } catch (e) { }
        for (var rem = 0; rem < allShapes.length; rem++) {
          try {
            allShapes[rem].remove();
          } catch (e) { }
        }
        prog.step("Cleaned up temp shapes");
        try {
          for (var ss = 0; ss < comp.selectedLayers.length; ss++)
            comp.selectedLayers[ss].selected = false;
        } catch (e) { }
        for (var rr = 0; rr < resultLayers.length; rr++) {
          try {
            resultLayers[rr].selected = true;
          } catch (e) { }
        }
        prog.step("Selected result layers");
      }

      app.endUndoGroup();
      if (!progress) prog.close();
    } catch (err) {
      try {
        app.endUndoGroup();
      } catch (e) { }
      alert("Error: " + (err && err.toString ? err.toString() : err));
    }
  }

  function runDecomposeTextToShapeLayers(progress) {
    try {
      app.beginUndoGroup("DecomposeTextToShapeLayers");

      var comp = app.project.activeItem;
      if (!comp || !(comp instanceof CompItem)) {
        alert(
          "No composition is active. Please open a composition and select text layers."
        );
        app.endUndoGroup();
        return;
      }

      var selLayers = comp.selectedLayers;
      if (!selLayers || selLayers.length === 0) {
        alert("No layers selected. Please select one or more text layers.");
        app.endUndoGroup();
        return;
      }

      var totalLayers = 0;
      for (var i1 = 0; i1 < selLayers.length; i1++)
        if (selLayers[i1] instanceof TextLayer) totalLayers++;
      var prog = progress || new ProgressDialog("Decompose: Text → Shapes");
      prog.setMax(Math.max(1, totalLayers));
      if (!progress) prog.show();

      for (var layerIdx = 0; layerIdx < selLayers.length; layerIdx++) {
        var textLayer = selLayers[layerIdx];
        if (!(textLayer instanceof TextLayer)) continue;

        var layerInPoint = textLayer.inPoint;
        var layerOutPoint = textLayer.outPoint;

        var textContent = String(textLayer.text.sourceText.value);
        // keep newline characters so style arrays stay aligned with character indices
        var cleanedForChars = textContent.replace(/\u0003/g, "");
        var cleanText = cleanedForChars.replace(/\s+/g, "");

        for (var sdel = 0; sdel < comp.selectedLayers.length; sdel++)
          comp.selectedLayers[sdel].selected = false;
        textLayer.selected = true;
        app.executeCommand(3781);

        var baseShapeLayer = comp.selectedLayers[0];
        if (!baseShapeLayer) {
          alert(
            "Failed to create shapes from text for layer: " + textLayer.name
          );
          continue;
        }

        var shapeContents = baseShapeLayer.property("Contents");
        var totalShapes = shapeContents.numProperties;

        var resultLayers = [];
        for (var i = totalShapes - 1; i >= 0; i--) {
          var dup = baseShapeLayer.duplicate();
          var dupContents = dup.property("Contents");
          for (var j = dupContents.numProperties; j > 0; j--) {
            if (j !== i + 1) {
              try {
                dupContents.property(j).remove();
              } catch (e) { }
            }
          }
          var charName = cleanText[i] ? cleanText[i] : i + 1;
          dup.name = "char_" + charName;
          dup.inPoint = layerInPoint;
          dup.outPoint = layerOutPoint;

          try {
            var laytrans = dup.property("ADBE Transform Group");

            if (
              laytrans.property("ADBE Position").numKeys != 0 ||
              laytrans.property("ADBE Position_0").numKeys != 0 ||
              laytrans.property("ADBE Position_1").numKeys != 0 ||
              laytrans.property("ADBE Anchor Point").numKeys != 0
            ) {
            } else {
              var RZ = laytrans.property("ADBE Rotate Z").value;
              var sourceRect = dup.sourceRectAtTime(0, true);
              var CX = sourceRect.width * 0.5 + sourceRect.left;
              var CY = sourceRect.height * 0.5 + sourceRect.top;
              var scale = laytrans.property("ADBE Scale").value;
              var SX = scale[0];
              var SY = scale[1];
              var anchor = laytrans.property("ADBE Anchor Point").value;
              var APX = anchor[0];
              var APY = anchor[1];

              var PX, PY;
              if (!laytrans.property("ADBE Position").dimensionsSeparated) {
                var pos = laytrans.property("ADBE Position").value;
                PX = pos[0];
                PY = pos[1];
              } else {
                PX = laytrans.property("ADBE Position_0").value;
                PY = laytrans.property("ADBE Position_1").value;
              }

              laytrans.property("ADBE Anchor Point").setValue([CX, CY, 0]);

              var DX = (CX - APX) * 0.01 * SX;
              var DY = (CY - APY) * 0.01 * SY;
              var rotRad = RZ * (Math.PI / 180);

              var newX = PX + (DX * Math.cos(rotRad) - DY * Math.sin(rotRad));
              var newY = PY + (DX * Math.sin(rotRad) + DY * Math.cos(rotRad));

              if (!laytrans.property("ADBE Position").dimensionsSeparated) {
                laytrans.property("ADBE Position").setValue([newX, newY, 0]);
              } else {
                laytrans.property("ADBE Position_0").setValue(newX);
                laytrans.property("ADBE Position_1").setValue(newY);
              }
            }
          } catch (e) { }

          resultLayers.push(dup);
        }

        try {
          textLayer.enabled = false;
        } catch (e) { }
        try {
          baseShapeLayer.remove();
        } catch (e) { }
        for (var rr = 0; rr < resultLayers.length; rr++) {
          try {
            resultLayers[rr].selected = true;
          } catch (e) { }
        }
        prog.step("Layers processed: " + (layerIdx + 1));
      }

      app.endUndoGroup();
      if (!progress) prog.close();
    } catch (err) {
      try {
        app.endUndoGroup();
      } catch (e) { }
      alert("Error: " + (err && err.toString ? err.toString() : err));
    }
  }

  function runDecomposeTextToShapeParts(progress) {
    (function () {
      function executePartsDecompose() {
        var comp = app.project.activeItem;
        if (!comp || !(comp instanceof CompItem)) {
          alert(
            "No composition is active. Please open a composition and select text layers or shape layers."
          );
          return;
        }
        var selectedLayers = comp.selectedLayers;
        if (selectedLayers.length === 0) {
          alert(
            "No layers selected. Please select one or more text layers or shape layers."
          );
          return;
        }
        app.beginUndoGroup("Text to Parts Decompose");
        try {
          var textLayerIndices = [];
          var layerProperties = [];
          for (var i = 0; i < selectedLayers.length; i++) {
            var layer = selectedLayers[i];
            var isTextLayer = layer instanceof TextLayer;
            var isShapeLayer =
              layer instanceof AVLayer ||
              layer.constructor.name === "ShapeLayer";
            var hasVectorGroup = false;

            try {
              hasVectorGroup =
                layer.property("ADBE Root Vectors Group") !== null;
            } catch (e) {
              hasVectorGroup = false;
            }

            if (isTextLayer || (isShapeLayer && hasVectorGroup)) {
              textLayerIndices.push(layer.index);
              layerProperties.push(captureBasicProperties(layer));
            }

            layer.selected = false;
          }
          textLayerIndices.sort(function (a, b) {
            return a - b;
          });

          if (textLayerIndices.length === 0) {
            alert(
              "No valid text layers or shape layers found in selection. Please select layers with vector content."
            );
            app.endUndoGroup();
            return;
          }

          var prog = progress || new ProgressDialog("Decompose: Text → Parts");
          prog.setMax(Math.max(1, textLayerIndices.length));
          if (!progress) prog.show();

          for (var ii = 0; ii < textLayerIndices.length; ii++) {
            var layerIndex = textLayerIndices[ii];
            var originalProps = layerProperties[ii];

            var currentLayer = comp.layers[layerIndex];
            currentLayer.selected = true;

            var baseShapeLayer;
            if (currentLayer instanceof TextLayer) {
              app.executeCommand(3781);
              baseShapeLayer =
                comp.selectedLayers && comp.selectedLayers.length > 0
                  ? comp.selectedLayers[0]
                  : null;
              if (!baseShapeLayer) {
                alert("Failed to create shapes from text for a layer.");
                currentLayer.selected = false;
                continue;
              }
            } else {
              try {
                if (!currentLayer.property("ADBE Root Vectors Group")) {
                  alert(
                    "Selected layer does not contain vector content: " +
                    currentLayer.name
                  );
                  currentLayer.selected = false;
                  continue;
                }
                baseShapeLayer = currentLayer;
              } catch (e) {
                alert("Error processing shape layer: " + e.toString());
                currentLayer.selected = false;
                continue;
              }
            }

            var shapeLabel = undefined;
            try {
              shapeLabel = baseShapeLayer.label;
            } catch (e) { }

            processPartsMerge(baseShapeLayer);

            var keepOriginal = currentLayer.constructor.name === "ShapeLayer";
            var resultLayers = processPartsDecompose(
              baseShapeLayer,
              originalProps,
              shapeLabel,
              keepOriginal
            );

            if (resultLayers && resultLayers.length > 0) {
              try {
                for (var s = 0; s < comp.selectedLayers.length; s++) {
                  comp.selectedLayers[s].selected = false;
                }

                for (var r = 0; r < resultLayers.length; r++) {
                  resultLayers[r].selected = true;
                }
              } catch (e) { }
            }

            prog.step("Layers processed: " + (ii + 1));
          }
        } catch (error) {
          alert("Error occurred: " + error.toString());
        }
        app.endUndoGroup();
      }

      function captureBasicProperties(layer) {
        var props = {};
        try {
          props.name = layer.name;
          props.inPoint = layer.inPoint;
          props.outPoint = layer.outPoint;
          props.enabled = layer.enabled;
          props.solo = layer.solo;
          props.shy = layer.shy;
          props.locked = layer.locked;
          props.label = layer.label;
          props.comment = layer.comment;
          props.threeDLayer = layer.threeDLayer;
          props.parent = layer.parent;
          props.blendingMode = layer.blendingMode;
        } catch (e) { }
        return props;
      }

      function applyBasicProperties(layer, props) {
        try {
          layer.inPoint = props.inPoint;
          layer.outPoint = props.outPoint;
          layer.enabled = props.enabled;
          layer.solo = props.solo;
          layer.shy = props.shy;
          layer.locked = props.locked;
          layer.comment = props.comment;
          layer.threeDLayer = props.threeDLayer;
          layer.blendingMode = props.blendingMode;
          if (props.parent) {
            try {
              layer.parent = props.parent;
            } catch (e) { }
          }
        } catch (e) { }
      }

      function processPartsMerge(layer) {
        var vectorGroup = layer.property("ADBE Root Vectors Group");
        var prowloop = 1;
        var guardMerge = 0;
        while (prowloop <= vectorGroup.numProperties) {
          if (++guardMerge > 20000) throw new Error("processPartsMerge: loop guard hit (prowloop)");
          var pathcount = 1;
          var pathnum = 3;
          var nowtext = vectorGroup.property(prowloop).property(2);
          if (nowtext.numProperties == 3) {
            pathnum--;
          }
          var area = [];
          var maxarea = 0;
          var mavec = true;
          var pathveccheckloop = 1;
          var areavecchash = [];
          var pathnowloop = 0;
          var guardVecCheck = 0;
          while (pathveccheckloop <= nowtext.numProperties - pathnum) {
            if (++guardVecCheck > 20000) throw new Error("processPartsMerge: loop guard hit (pathveccheckloop)");
            var nexdex = 0;
            var areachash = 0;
            var nowpath = nowtext.property(pathveccheckloop).property(2);
            var nowpathpoint = nowpath.value.vertices;
            for (var i = 0; i < nowpathpoint.length; i++) {
              nexdex = (i + 1) % nowpathpoint.length;
              areachash += nowpathpoint[i][0] * nowpathpoint[nexdex][1];
              areachash -= nowpathpoint[nexdex][0] * nowpathpoint[i][1];
            }
            if (Math.abs(areachash) > Math.abs(maxarea)) {
              maxarea = areachash;
            }
            areavecchash[pathveccheckloop - 1] = areachash;
            pathveccheckloop++;
          }
          if (maxarea > 0) {
            mavec = false;
          }
          var guardPathCount = 0;
          while (pathcount <= nowtext.numProperties - pathnum) {
            if (++guardPathCount > 20000) throw new Error("processPartsMerge: loop guard hit (pathcount)");
            area[pathcount - 1] = areavecchash[pathnowloop];
            if (area[pathcount - 1] < 0 == mavec) {
              var countloop = 1;
              var areamove = false;
              nowtext.addProperty("ADBE Vector Group").moveTo(pathcount);
              nowtext
                .property(pathcount)
                .property("ADBE Vector Materials Group")
                .remove();
              nowtext.property(pathcount).name = nowtext.property(
                pathcount + 1
              ).name;
              nowtext
                .property(pathcount)
                .property(2)
                .addProperty("ADBE Vector Shape - Group");
              nowtext
                .property(pathcount)
                .property(2)
                .property(1)
                .property(2)
                .setValue(nowtext.property(pathcount + 1).property(2).value);
              nowtext.property(pathcount).property(2).property(1).name =
                nowtext.property(pathcount + 1).name;
              nowtext.property(pathcount + 1).remove();
              var guardCountLoop = 0;
              while (countloop < pathcount) {
                if (++guardCountLoop > 20000) throw new Error("processPartsMerge: loop guard hit (countloop)");
                if (area[countloop - 1] < area[pathcount - 1]) {
                  nowtext.property(pathcount).moveTo(countloop);
                  area.splice(countloop - 1, 0, area[pathcount - 1]);
                  areamove = true;
                }
                countloop++;
              }
              if (!areamove) {
                nowtext.property(pathcount).moveTo(countloop);
              }
              pathcount++;
            } else {
              nowtext.property(pathcount).moveTo(nowtext.numProperties - 3);
              pathnum++;
            }
            pathnowloop++;
          }
          var countloop2 = 1;
          var guardCountLoop2 = 0;
          while (countloop2 < pathcount) {
            if (++guardCountLoop2 > 20000) throw new Error("processPartsMerge: loop guard hit (countloop2)");
            var contentsloop = 0;
            var Mflag = false;
            var guardContentsLoop = 0;
            while (contentsloop < nowtext.numProperties - (pathcount + 2)) {
              if (++guardContentsLoop > 20000) throw new Error("processPartsMerge: loop guard hit (contentsloop)");
              var ppflag = false;
              var checkpoint = [
                nowtext.property(pathcount + contentsloop).property(2).value
                  .vertices[0][0],
                nowtext.property(pathcount + contentsloop).property(2).value
                  .vertices[0][1],
              ];
              var cn = 0;
              var nowtexpath = nowtext
                .property(countloop2)
                .property(2)
                .property(1)
                .property(2);
              var nowpathpoint = nowtexpath.value.vertices;
              var nowpathoutT = nowtexpath.value.outTangents;
              var nowpathinT = nowtexpath.value.inTangents;
              var bezposX = [];
              var bezposY = [];
              for (var k = 0; k < nowpathpoint.length; k++) {
                var beznexdex = (k + 1) % nowpathpoint.length;
                var bez3 = 0.125;
                var p0 = [nowpathpoint[k][0], nowpathpoint[k][1]];
                var p1 = [
                  nowpathpoint[k][0] + nowpathoutT[k][0],
                  nowpathpoint[k][1] + nowpathoutT[k][1],
                ];
                var p2 = [
                  nowpathpoint[beznexdex][0] + nowpathinT[beznexdex][0],
                  nowpathpoint[beznexdex][1] + nowpathinT[beznexdex][1],
                ];
                var p3 = [
                  nowpathpoint[beznexdex][0],
                  nowpathpoint[beznexdex][1],
                ];
                bezposX[k] =
                  bez3 * p0[0] +
                  3 * bez3 * p1[0] +
                  3 * bez3 * p2[0] +
                  bez3 * p3[0];
                bezposY[k] =
                  bez3 * p0[1] +
                  3 * bez3 * p1[1] +
                  3 * bez3 * p2[1] +
                  bez3 * p3[1];
              }
              var u = 0;
              var bppointpos = [];
              for (var j2 = 0; j2 < nowpathpoint.length * 2; j2++) {
                if (j2 % 2 == 0) {
                  bppointpos[j2] = nowpathpoint[u];
                } else {
                  bppointpos[j2] = [bezposX[u], bezposY[u]];
                  u++;
                }
              }
              for (var i3 = 0; i3 < bppointpos.length; i3++) {
                var nexdex2 = (i3 + 1) % bppointpos.length;
                if (
                  (bppointpos[i3][1] <= checkpoint[1] &&
                    bppointpos[nexdex2][1] > checkpoint[1]) ||
                  (bppointpos[i3][1] > checkpoint[1] &&
                    bppointpos[nexdex2][1] <= checkpoint[1])
                ) {
                  var vt =
                    (checkpoint[1] - bppointpos[i3][1]) /
                    (bppointpos[nexdex2][1] - bppointpos[i3][1]);
                  if (
                    checkpoint[0] <
                    bppointpos[i3][0] +
                    vt * (bppointpos[nexdex2][0] - bppointpos[i3][0])
                  ) {
                    cn++;
                  }
                }
              }
              if (cn % 2 == 0) {
                ppflag = false;
              } else {
                ppflag = true;
              }
              if (ppflag == true) {
                nowtext
                  .property(countloop2)
                  .property(2)
                  .addProperty("ADBE Vector Shape - Group")
                  .moveTo(2);
                nowtext
                  .property(countloop2)
                  .property(2)
                  .property(2)
                  .property(2)
                  .setValue(
                    nowtext.property(pathcount + contentsloop).property(2).value
                  );
                nowtext.property(countloop2).property(2).property(2).name =
                  nowtext.property(pathcount + contentsloop).name;
                nowtext.property(pathcount + contentsloop).remove();
                contentsloop--;
                if (Mflag == false) {
                  nowtext
                    .property(countloop2)
                    .property(2)
                    .addProperty("ADBE Vector Filter - Merge");
                  Mflag = true;
                }
              }
              contentsloop++;
            }
            countloop2++;
          }
          prowloop++;
        }
      }

      function processPartsDecompose(
        layer,
        originalProps,
        targetLabel,
        keepOriginal
      ) {
        var vectorGroup = layer.property("ADBE Root Vectors Group");
        var proloop = 0;
        var texnum = vectorGroup.numProperties;
        var resultLayers = [];
        var guardProloop = 0;
        while (proloop < texnum) {
          if (++guardProloop > 20000) throw new Error("processPartsDecompose: loop guard hit (proloop)");
          var character = vectorGroup.property(1);
          var contents = character.property(2);
          var pronum = contents.numProperties - 3;
          var prowloop = 1;
          var guardProwloop = 0;
          while (prowloop < pronum) {
            if (++guardProwloop > 20000) throw new Error("processPartsDecompose: loop guard hit (prowloop)");
            var duplicatedLayer = layer.duplicate();
            duplicatedLayer.name = character.name + " Outline ";
            var dupContents = duplicatedLayer
              .property("ADBE Root Vectors Group")
              .property(1)
              .property(2);
            var guardDupContents = 0;
            while (dupContents.numProperties > 4) {
              if (++guardDupContents > 20000) throw new Error("processPartsDecompose: loop guard hit (dupContents)");
              dupContents.property(2).remove();
            }
            if (
              dupContents.property(2).matchName == "ADBE Vector Filter - Merge"
            ) {
              dupContents.property(2).remove();
            }
            contents.property(1).remove();
            var guardDupRemove = 0;
            while (
              duplicatedLayer.property("ADBE Root Vectors Group")
                .numProperties > 1
            ) {
              if (++guardDupRemove > 20000) throw new Error("processPartsDecompose: loop guard hit (dupRemove)");
              duplicatedLayer
                .property("ADBE Root Vectors Group")
                .property(2)
                .remove();
            }
            adjustAnchorPoint(duplicatedLayer, 2);
            applyBasicProperties(duplicatedLayer, originalProps);
            if (typeof targetLabel !== "undefined") {
              try {
                duplicatedLayer.label = targetLabel;
              } catch (e) { }
            }
            try {
              // 最初のパーツは元のレイヤーの前に、それ以降は最後に作成されたレイヤーの前に挿入
              if (resultLayers.length === 0) {
                duplicatedLayer.moveBefore(layer);
              } else {
                // 最後に作成されたレイヤーの前に挿入
                var lastLayer = resultLayers[resultLayers.length - 1];
                duplicatedLayer.moveBefore(lastLayer);
              }
            } catch (e) { }
            resultLayers.push(duplicatedLayer);
            prowloop++;
          }
          var finalLayer = layer.duplicate();
          finalLayer.name = character.name + " Outline ";
          var finalContents = finalLayer
            .property("ADBE Root Vectors Group")
            .property(1)
            .property(2);
          if (
            finalContents.property(2).matchName == "ADBE Vector Filter - Merge"
          ) {
            finalContents.property(2).remove();
          }
          var guardFinalRemove = 0;
          while (
            finalLayer.property("ADBE Root Vectors Group").numProperties > 1
          ) {
            if (++guardFinalRemove > 20000) throw new Error("processPartsDecompose: loop guard hit (finalRemove)");
            finalLayer.property("ADBE Root Vectors Group").property(2).remove();
          }
          adjustAnchorPoint(finalLayer, 2);
          applyBasicProperties(finalLayer, originalProps);
          if (typeof targetLabel !== "undefined") {
            try {
              finalLayer.label = targetLabel;
            } catch (e) { }
          }
          try {
            // 最終レイヤーは最後に作成されたレイヤーの前に挿入
            if (resultLayers.length === 0) {
              finalLayer.moveBefore(layer);
            } else {
              var lastLayer = resultLayers[resultLayers.length - 1];
              finalLayer.moveBefore(lastLayer);
            }
          } catch (e) { }
          resultLayers.push(finalLayer);
          vectorGroup.property(1).remove();
          proloop++;
        }
        if (keepOriginal) {
          try {
            layer.enabled = false;
          } catch (e) { }
        } else {
          try {
            layer.remove();
          } catch (e) { }
        }
        return resultLayers;
      }

      function adjustAnchorPoint(layer, pet) {
        var laytrans = layer.property("ADBE Transform Group");
        if (
          laytrans.property("ADBE Position").numKeys != 0 ||
          laytrans.property("ADBE Position_0").numKeys != 0 ||
          laytrans.property("ADBE Position_1").numKeys != 0 ||
          laytrans.property("ADBE Anchor Point").numKeys != 0
        ) {
          return;
        }
        try {
          var RZ = laytrans.property("ADBE Rotate Z").value;
          var sourceRect = layer.sourceRectAtTime(0, true);
          var CX = sourceRect.width * 0.5 + sourceRect.left;
          var CY = sourceRect.height * 0.5 + sourceRect.top;
          var scale = laytrans.property("ADBE Scale").value;
          var SX = scale[0];
          var SY = scale[1];
          var anchor = laytrans.property("ADBE Anchor Point").value;
          var APX = anchor[0];
          var APY = anchor[1];
          var PX, PY;
          if (!laytrans.property("ADBE Position").dimensionsSeparated) {
            var pos = laytrans.property("ADBE Position").value;
            PX = pos[0];
            PY = pos[1];
          } else {
            PX = laytrans.property("ADBE Position_0").value;
            PY = laytrans.property("ADBE Position_1").value;
          }
          laytrans.property("ADBE Anchor Point").setValue([0, 0, 0]);
          if (pet == 1) {
            layer
              .property("ADBE Root Vectors Group")
              .property(1)
              .property(3)
              .property("ADBE Vector Anchor")
              .setValue([CX, CY]);
          } else if (pet == 2) {
            layer
              .property("ADBE Root Vectors Group")
              .property(1)
              .property(2)
              .property(1)
              .property(3)
              .property("ADBE Vector Anchor")
              .setValue([CX, CY]);
          } else {
            laytrans.property("ADBE Anchor Point").setValue([CX, CY, 0]);
          }
          var DX = (CX - APX) * 0.01 * SX;
          var DY = (CY - APY) * 0.01 * SY;
          var rotRad = RZ * (Math.PI / 180);
          var newX = PX + (DX * Math.cos(rotRad) - DY * Math.sin(rotRad));
          var newY = PY + (DX * Math.sin(rotRad) + DY * Math.cos(rotRad));
          if (!laytrans.property("ADBE Position").dimensionsSeparated) {
            laytrans.property("ADBE Position").setValue([newX, newY, 0]);
          } else {
            laytrans.property("ADBE Position_0").setValue(newX);
            laytrans.property("ADBE Position_1").setValue(newY);
          }
        } catch (e) { }
      }

      executePartsDecompose();
    })();
  }

  function buildUI(thisObj) {
    var pal =
      thisObj instanceof Panel
        ? thisObj
        : new Window("palette", "Baramoji", undefined, { resizeable: true });
    if (!pal) return pal;

    pal.orientation = "column";
    pal.alignChildren = ["fill", "fill"];
    var grp = pal.add("group");
    grp.orientation = "column";
    grp.alignChildren = ["fill", "fill"];
    grp.alignment = ["fill", "fill"];

    var btn1 = grp.add("button", undefined, "Texts");
    var btn2 = grp.add("button", undefined, "Shapes");
    var btn3 = grp.add("button", undefined, "Parts");
    btn1.alignment = ["fill", "fill"];
    btn2.alignment = ["fill", "fill"];
    btn3.alignment = ["fill", "fill"];

    btn1.onClick = function () {
      runDecomposeTextToTextLayers();
    };
    btn2.onClick = function () {
      runDecomposeTextToShapeLayers();
    };
    btn3.onClick = function () {
      runDecomposeTextToShapeParts();
    };

    pal.onResizing = pal.onResize = function () {
      this.layout.resize();
      try {
        var topBottom = this.margins
          ? this.margins.top + this.margins.bottom
          : 0;
        var availableH = Math.max(0, this.size.height - topBottom);
        var spacing = grp.spacing || 0;
        var count = 3;
        var perH = Math.max(
          24,
          Math.floor((availableH - spacing * (count - 1)) / count)
        );
        btn1.preferredSize = [0, perH];
        btn2.preferredSize = [0, perH];
        btn3.preferredSize = [0, perH];
        grp.layout.layout(true);
      } catch (e) { }
    };
    return pal;
  }

  var ui = buildUI(thisObj);
  if (ui instanceof Window) {
    try {
      ui.center();
      ui.show();
    } catch (e) { }
  } else if (ui) {
    try {
      ui.layout.layout(true);
      ui.layout.resize();
      ui.visible = true;
    } catch (e) { }
  }
})(this);

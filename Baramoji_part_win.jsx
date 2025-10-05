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

(function () {
  var progressWindow = null;
  var progressBar = null;
  var statusText = null;

  function showProgressWindow() {
    if (progressWindow == null || !progressWindow.visible) {
      progressWindow = new Window(
        "palette",
        "Text to Parts Decompose",
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
        } catch (e) {}
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
      } catch (e) {}
      progressWindow = null;
      progressBar = null;
      statusText = null;
    }
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
    } catch (e) {}
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
        } catch (e) {}
      }
    } catch (e) {}
  }

  function processPartsMerge(layer) {
    var vectorGroup = layer.property("ADBE Root Vectors Group");
    var prowloop = 1;
    while (prowloop <= vectorGroup.numProperties) {
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
      while (pathveccheckloop <= nowtext.numProperties - pathnum) {
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
      while (pathcount <= nowtext.numProperties - pathnum) {
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
          while (countloop < pathcount) {
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
      while (countloop2 < pathcount) {
        var contentsloop = 0;
        var Mflag = false;
        while (contentsloop < nowtext.numProperties - (pathcount + 2)) {
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
            var p3 = [nowpathpoint[beznexdex][0], nowpathpoint[beznexdex][1]];
            bezposX[k] =
              bez3 * p0[0] + 3 * bez3 * p1[0] + 3 * bez3 * p2[0] + bez3 * p3[0];
            bezposY[k] =
              bez3 * p0[1] + 3 * bez3 * p1[1] + 3 * bez3 * p2[1] + bez3 * p3[1];
          }
          var u = 0;
          var bppointpos = [];
          for (var j = 0; j < nowpathpoint.length * 2; j++) {
            if (j % 2 == 0) {
              bppointpos[j] = nowpathpoint[u];
            } else {
              bppointpos[j] = [bezposX[u], bezposY[u]];
              u++;
            }
          }
          for (var i2 = 0; i2 < bppointpos.length; i2++) {
            var nexdex2 = (i2 + 1) % bppointpos.length;
            if (
              (bppointpos[i2][1] <= checkpoint[1] &&
                bppointpos[nexdex2][1] > checkpoint[1]) ||
              (bppointpos[i2][1] > checkpoint[1] &&
                bppointpos[nexdex2][1] <= checkpoint[1])
            ) {
              var vt =
                (checkpoint[1] - bppointpos[i2][1]) /
                (bppointpos[nexdex2][1] - bppointpos[i2][1]);
              if (
                checkpoint[0] <
                bppointpos[i2][0] +
                  vt * (bppointpos[nexdex2][0] - bppointpos[i2][0])
              ) {
                cn++;
              }
            }
          }
          ppflag = cn % 2 != 0;
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
    } catch (e) {}
  }

  function processPartsDecompose(layer, originalProps, targetLabel) {
    var vectorGroup = layer.property("ADBE Root Vectors Group");
    var proloop = 0;
    var texnum = vectorGroup.numProperties;
    var resultLayers = [];
    while (proloop < texnum) {
      var character = vectorGroup.property(1);
      var contents = character.property(2);
      var pronum = contents.numProperties - 3;
      var prowloop = 1;
      while (prowloop < pronum) {
        var duplicatedLayer = layer.duplicate();
        duplicatedLayer.name = character.name + " Outline ";
        var dupContents = duplicatedLayer
          .property("ADBE Root Vectors Group")
          .property(1)
          .property(2);
        while (dupContents.numProperties > 4) {
          dupContents.property(2).remove();
        }
        if (dupContents.property(2).matchName == "ADBE Vector Filter - Merge") {
          dupContents.property(2).remove();
        }
        contents.property(1).remove();
        while (
          duplicatedLayer.property("ADBE Root Vectors Group").numProperties > 1
        ) {
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
          } catch (e) {}
        }
        try {
          duplicatedLayer.moveBefore(layer);
        } catch (e) {}
        resultLayers.push(duplicatedLayer);
        prowloop++;
        updateProgress(
          30 + Math.min(resultLayers.length, 40),
          "Extracting outlines..."
        );
      }
      var finalLayer = layer.duplicate();
      finalLayer.name = character.name + " Outline ";
      var finalContents = finalLayer
        .property("ADBE Root Vectors Group")
        .property(1)
        .property(2);
      if (finalContents.property(2).matchName == "ADBE Vector Filter - Merge") {
        finalContents.property(2).remove();
      }
      while (finalLayer.property("ADBE Root Vectors Group").numProperties > 1) {
        finalLayer.property("ADBE Root Vectors Group").property(2).remove();
      }
      adjustAnchorPoint(finalLayer, 2);
      applyBasicProperties(finalLayer, originalProps);
      if (typeof targetLabel !== "undefined") {
        try {
          finalLayer.label = targetLabel;
        } catch (e) {}
      }
      try {
        finalLayer.moveBefore(layer);
      } catch (e) {}
      resultLayers.push(finalLayer);
      vectorGroup.property(1).remove();
      proloop++;
      updateProgress(75, "Finalizing character...");
    }
    try {
      layer.remove();
    } catch (e) {}
    if (resultLayers.length > 0) {
      resultLayers.reverse();
      try {
        resultLayers[0].selected = true;
      } catch (e) {}
      for (var i = 1; i < resultLayers.length; i++) {
        try {
          resultLayers[i].moveAfter(resultLayers[i - 1]);
          resultLayers[i].selected = true;
        } catch (e) {}
      }
    }
  }

  try {
    app.beginUndoGroup("Text to Parts Decompose");
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

    var selectedLayers = comp.selectedLayers;
    if (!selectedLayers || selectedLayers.length === 0) {
      alert("No layers selected. Please select one or more text layers.");
      closeProgressWindow();
      app.endUndoGroup();
      return;
    }

    updateProgress(5, "Scanning selection...");

    try {
      var textLayerIndices = [];
      var layerProperties = [];
      for (var iSel = 0; iSel < selectedLayers.length; iSel++) {
        if (selectedLayers[iSel] instanceof TextLayer) {
          textLayerIndices.push(selectedLayers[iSel].index);
          layerProperties.push(captureBasicProperties(selectedLayers[iSel]));
        }
        selectedLayers[iSel].selected = false;
      }
      textLayerIndices.sort(function (a, b) {
        return a - b;
      });

      for (var i = textLayerIndices.length - 1; i >= 0; i--) {
        var layerIndex = textLayerIndices[i];
        var originalProps = layerProperties[i];

        comp.layers[layerIndex].selected = true;
        updateProgress(12, "Creating shapes from text...");
        app.executeCommand(3781);

        var baseShapeLayer =
          comp.selectedLayers && comp.selectedLayers.length > 0
            ? comp.selectedLayers[0]
            : null;
        if (!baseShapeLayer) {
          alert("Failed to create shapes from text for a layer.");
          comp.layers[layerIndex].selected = false;
          continue;
        }

        var shapeLabel = undefined;
        try {
          shapeLabel = baseShapeLayer.label;
        } catch (e) {}

        updateProgress(28, "Merging parts...");
        processPartsMerge(baseShapeLayer);

        updateProgress(42, "Isolating outlines...");
        updateProgress(55, "Decomposing to parts...");
        processPartsDecompose(baseShapeLayer, originalProps, shapeLabel);

        var layerProgress =
          60 +
          Math.round(
            ((textLayerIndices.length - 1 - i) /
              Math.max(1, textLayerIndices.length)) *
              35
          );
        updateProgress(
          Math.min(layerProgress, 95),
          "Processed layer " +
            (textLayerIndices.length - i) +
            "/" +
            textLayerIndices.length
        );
      }
    } catch (error) {
      alert("Error occurred: " + error.toString());
    }

    updateProgress(96, "Finalizing...");
    $.sleep(150);
    updateProgress(100, "Completed!");
    $.sleep(250);
    closeProgressWindow();
    app.endUndoGroup();
  } catch (err) {
    try {
      app.endUndoGroup();
    } catch (e) {}
    closeProgressWindow();
    alert("Error: " + (err && err.toString ? err.toString() : err));
  }
})();

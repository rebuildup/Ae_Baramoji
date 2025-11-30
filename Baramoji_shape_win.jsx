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
        "Decompose Text To Shapes",
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
    app.beginUndoGroup("DecomposeTextToShapeLayers");

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

    updateProgress(4, "Inspecting layers...");

    for (var layerIdx = 0; layerIdx < selLayers.length; layerIdx++) {
      updateProgress(
        Math.min(
          8 + Math.round((layerIdx / Math.max(1, selLayers.length)) * 10),
          18
        ),
        "Processing layer " + (layerIdx + 1) + "/" + selLayers.length + "..."
      );

      var textLayer = selLayers[layerIdx];
      if (!(textLayer instanceof TextLayer)) {
        continue;
      }

      var layerInPoint = textLayer.inPoint;
      var layerOutPoint = textLayer.outPoint;

      var textContent = String(textLayer.text.sourceText.value);
      var cleanedForChars = textContent.replace(/\r|\n|\u0003/g, "");
      var cleanText = cleanedForChars.replace(/\s+/g, "");

      for (var sdel = 0; sdel < comp.selectedLayers.length; sdel++)
        comp.selectedLayers[sdel].selected = false;
      textLayer.selected = true;
      updateProgress(20, "Converting text to shapes...");
      app.executeCommand(3781);

      var baseShapeLayer = comp.selectedLayers[0];
      if (!baseShapeLayer) {
        alert("Failed to create shapes from text for layer: " + textLayer.name);
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

        try {
          dup.moveBefore(textLayer);
        } catch (e) { }

        resultLayers.push(dup);

        var progressBase = 22;
        var progressSpan = 60;
        var step =
          progressBase +
          Math.round(
            ((totalShapes - 1 - i) / Math.max(1, totalShapes)) * progressSpan
          );
        updateProgress(
          step,
          "Isolating shape " + (totalShapes - i) + "/" + totalShapes + "..."
        );
      }

      updateProgress(86, "Cleaning up...");

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

      updateProgress(
        Math.min(90 + Math.round(((layerIdx + 1) / selLayers.length) * 8), 98),
        "Layer " + (layerIdx + 1) + "/" + selLayers.length + " completed"
      );
    }

    updateProgress(99, "Finalizing...");
    $.sleep(150);
    updateProgress(100, "Completed!");
    $.sleep(250);
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

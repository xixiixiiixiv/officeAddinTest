(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
    Office.initialize = function (reason) {
        $(document).ready(readyDocument());
    };

    function readyDocument() {
        // FabricUI 通知メカニズムを初期化して、非表示にします
        var element = document.querySelector('.ms-MessageBanner');
        messageBanner = new fabric.MessageBanner(element);
        messageBanner.hideBanner();

        // https://msdn.microsoft.com/ja-jp/library/office/fp161062.aspx
        // ドキュメント内で選択が変更されるときに発生
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, displaySelectedCells);

        // Excel 2016 を使用していない場合は、フォールバック ロジックを使用してください。
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            // $("#template-description").text("このサンプルでは、スプレッドシートで選ばれたセルの値が表示されます。(ExcelApi 1.1 Not Supported)");
            $('#button-text').text("表示!");
            $('#button-desc').text("選択範囲が表示されます");

            $('#highlight-button').click(displaySelectedCells);
            //return;
        } else {
            //$("#template-description").text("このサンプルでは、スプレッドシートで選択したセルから最も高い値が強調表示されます。");
            $('#button-text').text("強調表示!");
            $('#button-desc').text("最大数が強調表示されます。");

            loadSampleData();

            // 強調表示ボタンのクリック イベント ハンドラーを追加します。
            $('#highlight-button').click(hightlightHighestValue);
        }

        // 画像挿入
        if (Office.context.requirements.isSetSupported("ImageCoercion")) {
            //https://msdn.microsoft.com/ja-jp/library/office/fp142145.aspx
            $("#template-description").text("ImageCoercion is supported.");
            // Excelでの画像挿入
            insertViaImageCoercion();
        } else {
            $("#template-description").text("ImageCoercion is NOT supported.");
            // Wordでの画像挿入
            insertViaOoxml();
        }
    };

    // https://qr4office.azurewebsites.net/App/Home/Home.js
    function insertViaImageCoercion() {
        // Document.setSelectedDataAsync メソッド 
        // ドキュメント内の現在の選択範囲にデータを書き込みます。
        // https://msdn.microsoft.com/ja-jp/library/office/fp142145.aspx
        //Office.context.document.setSelectedDataAsync(currentBinaryImage,
        //    { coercionType: Office.CoercionType.Image },
        //    function (result) {
        //        if (result.status !== Office.AsyncResultStatus.Succeeded) {
        //            app.showNotification(
        //                'Error inserting into selection:',
        //                "An unexpected error occured.  " +
        //                GeneratePleaseCopyPasteText() + "  Error message: " + result.error.message);
        //        }
        //    }
        //);
    }

    // https://qr4office.azurewebsites.net/App/Home/Home.js
    function insertViaOoxml() {
        //var dataToPassToService = {
        //    'width': fullWidth,
        //    'height': fullHeight,
        //    'count': countBeforeInserting,
        //    'token': token,
        //    'version': LATEST_VERSION
        //};
        //$.ajax({
        //    url: '../../api/XML',
        //    type: 'POST',
        //    data: JSON.stringify(dataToPassToService),
        //    contentType: "application/json;charset=utf-8",
        //    cache: false
        //}).done(function (data) {
        //    replaceXMLAndInsert(data);
        //    countBeforeInserting = 0;
        //}).fail(function () {
        //    app.showNotification('Error:', 'Error inserting image, you may ' +
        //        'try copy-pasting the image instead.');
        //});

        //function replaceXMLAndInsert(xml) {
        //    xml = xml.replace('{{{BINARY_IMAGE_DATA}}}', currentBinaryImage);

        //    // Only way to get here is if button is visible.  In which case there's
        //    //     also a preview image, and a current OOXML loaded.
        //    Office.context.document.setSelectedDataAsync(xml,
        //        { coercionType: Office.CoercionType.Ooxml },
        //        function (result) {
        //            if (result.status !== Office.AsyncResultStatus.Succeeded) {
        //                var message = result.error.message;
        //                // 1000 code is "Invalid Coercion Type", but unlike the name, doesn't get localized cross languages
        //                // 5007 code is "unsupported enumeration", but unlike the name, doesn't get localized cross languages
        //                if (result.error.code === 1000 || result.error.code == 5007) {
        //                    message = 'Unfortunately, the programmatic insertion of images is ' +
        //                        'not supported in this Office application yet.  ' +
        //                        GeneratePleaseCopyPasteText();
        //                }
        //                app.showNotification(
        //                    'Error inserting into selection:', message);
        //            }
        //        }
        //    );
        //}
    }

    // サンプルデータ埋め込み。Office2016以降で動作。
    function loadSampleData() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            $("#template-description").text("ExcelApi 1.1 Not Supported");
            return;
        }

        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Excel オブジェクト モデルに対してバッチ操作を実行します
        Excel.run(function (ctx) {
            // 作業中のシートに対するプロキシ オブジェクトを作成します
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // ワークシートにサンプル データを書き込むコマンドをキューに入れます
            sheet.getRange("B3:D5").values = values;

            // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    // Excel2016 以降のみで動作
    function hightlightHighestValue() {
        // Excel オブジェクト モデルに対してバッチ操作を実行します
        Excel.run(function (ctx) {
            // 選択された範囲に対するプロキシ オブジェクトを作成し、そのプロパティを読み込みます
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // セルを検索して強調表示します
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // セルを強調表示
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    // 現在選択されているセルの内容を表示
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('選択されたテキスト:', '"' + result.value + '"');
                } else {
                    showNotification('エラー', result.error.message);
                }
            });
    }

    // エラーを処理するためのヘルパー関数
    function errorHandler(error) {
        // Excel.run の実行から浮かび上がってくるすべての累積エラーをキャッチする必要があります
        showNotification("エラー", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 通知を表示するヘルパー関数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

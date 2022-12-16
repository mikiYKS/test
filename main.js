$(document).ready(function () {
  $("#run").click(() => tryCatch(getKakuin));
});

function getKakuin() {
  var authenticator;
  var client_id = "d81628a2-bd53-4116-a8a6-c57377eececd";
  var redirect_url = "https://mikiyks.github.io/test/";
  var scope = "https://graph.microsoft.com/Files.Read.All";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function (token) {
      access_token = token.access_token;
      //API呼び出し
      $(function () {
        $.ajax({
          url:
            "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/drive/items/01SG44IHMJY6HM4OB2XJGZ34EYB77ZANB2",
          type: "GET",
          beforeSend: function (xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          }
        }).then(
          async function (data) {
            const obj = data["@microsoft.graph.downloadUrl"];
            var kakuinbase64 = await getImageBase64(obj);
            //ここからkakuinbase64を張り付ける処理
            inkanpaste(kakuinbase64);

            //ログ出力
            var fileName = Office.context.document.url.match(".+/(.+?)([\?#;].*)?$")[1];
            inkanLog('角印', fileName);

          },
          function (data) {
            console.log(data);
          }
        );
      });
    })
    .catch(OfficeHelpers.Utilities.log);
}

// バイナリ画像をbase64で返す
async function getImageBase64(url) {
  const response = await fetch(url);
  const contentType = response.headers.get("content-type");
  const arrayBuffer = await response.arrayBuffer();
  let base64String = btoa(String.fromCharCode.apply(null, new Uint8Array(arrayBuffer)));
  //return `data:${contentType};base64,${base64String}`;
  return base64String;
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

Office.initialize = function (reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

async function onWorkSheetSingleClick(x, y, pic) {
  await Excel.run(async (context) => {

    const shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
    const shpStampImage = shapes.addImage(pic);
    shpStampImage.name = "印鑑";
    shpStampImage.left = x;
    shpStampImage.top = y;
    await context.sync();
  });
}

async function inkanpaste(pic) {
  await Excel.run(async (context) => {
    //アクティブセルの位置取得
    const cell = context.workbook.getActiveCell();
    cell.load("left").load("top");
    await context.sync();
    //印鑑生成実行
    onWorkSheetSingleClick(cell.left, cell.top, pic);
  });
}

//SharePointListにログ出力
function inkanLog(inkanName, inkanFile) {

  var authenticator;
  var client_id = "d81628a2-bd53-4116-a8a6-c57377eececd";
  var redirect_url = "https://mikiyks.github.io/test/";
  var scope = "https://graph.microsoft.com/Sites.ReadWrite.All";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function (token) {
      access_token = token.access_token;

      $(function () {
        $.ajax({
          url:
            "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/lists/6aac0560-622e-4ee1-ba8f-73b32d8e9f05/items",
          type: "POST",
          data: JSON.stringify({
            fields: {
              Title: inkanName,
              FileName: inkanFile
            }
          }),
          contentType: 'application/json',
          beforeSend: function (xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          }
        }).then(
          async function (data) {
          },
          function (data) {
            console.log(data);
          }
        );
      });
    });
}

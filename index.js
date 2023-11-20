// load the module
let GroupDocs = require("groupdocs-conversion-cloud");
let fs = require("fs").promises;
let path = require("path");
const util = require("util");

// get your ClientId and ClientSecret at https://dashboard.groupdocs.cloud (free registration is required).
let clientId = "c8239bbb-2fbb-4ec6-ae9b-1aa12e2c4178";
let clientSecret = "ad6923c1d1acc12576cd59c19a977957";
let gdClientStorage = "ppts";

async function convertPPTtoHTML() {
  let fontPath = "Fonts/calibri.ttf";

  // configure
  let config = new GroupDocs.Configuration(clientId, clientSecret);
  config.apiBaseUrl = "https://api.groupdocs.cloud";

  // file api
  let fileApi = GroupDocs.FileApi.fromConfig(config);

  // storage api
  let storageApi = GroupDocs.StorageApi.fromConfig(config);

  // construct Api
  let convertApi = GroupDocs.ConvertApi.fromConfig(config);

  try {
    // Font Checking
    let existResponse = await storageApi.objectExists(
      new GroupDocs.ObjectExistsRequest(fontPath, gdClientStorage),
    );

    let filePath;
    console.log("Font checking: ", existResponse.exists);
    if (existResponse.exists === false) {
      let file = await fs.readFile(path.join(__dirname, fontPath));
      let uploadRequest = new GroupDocs.UploadFileRequest(
        "/" + fontPath,
        file,
        gdClientStorage,
      );

      filePath = await fileApi.uploadFile(uploadRequest);
    }

    console.log("filePath: ", filePath);

    let fileName = "test-ppt.pptx";
    let htmlOutputPath = "test-ppt.html";
    let pptxInputPath = path.join(__dirname, "pptx", fileName);
    let fileNameWithoutExtension = fileName.replace(/\.[^/.]+$/, "");

    // PPTx file checking
    console.log("start file reading: ", fileName);
    const fileContent = await fs.readFile(pptxInputPath);

    let fileUploadRequest = new GroupDocs.UploadFileRequest(
      fileName,
      fileContent,
      gdClientStorage,
    );
    let fileUploadResponse = await fileApi.uploadFile(fileUploadRequest);
    const fileUploadedPath = fileUploadResponse.uploaded[0];
    console.log(
      "\x1b[32m%s\x1b[0m",
      `Successfully uploaded file ${fileNameWithoutExtension} to groupdocs cloud`,
    );
    console.log(
      "\x1b[36m%s\x1b[0m",
      util.inspect(fileUploadResponse, {
        showHidden: false,
        depth: null,
        colors: true,
      }),
    );

    // Html Convert Option
    let convertOptions = new GroupDocs.WebConvertOptions();
    convertOptions.fixedLayout = true;
    convertOptions.fixedLayoutShowBorders = false;

    // Presentation Load Options
    let loadOptions = new GroupDocs.PresentationLoadOptions();
    loadOptions.defaultFont = "Calibri";
    loadOptions.hideComments = true;

    const conversionSettings = new GroupDocs.ConvertSettings();
    conversionSettings.storageName = gdClientStorage;
    conversionSettings.filePath = fileUploadedPath;
    conversionSettings.format = "html";
    conversionSettings.loadOptions = loadOptions;
    conversionSettings.convertOptions = convertOptions;
    conversionSettings.outputPath = htmlOutputPath;
    conversionSettings.fontsPath = "Fonts";

    console.log("conversionSettings: ", conversionSettings);
    let conversionRequest = new GroupDocs.ConvertDocumentRequest(
      conversionSettings,
    );
    let conversionResponse =
      await convertApi.convertDocument(conversionRequest);
    console.log("\x1b[32m%s\x1b[0m", "HTML Converted Successfully");
    console.log("\x1b[36m%s\x1b[0m", conversionResponse);
    console.log("-----------------------------------------");
    return JSON.stringify(conversionResponse, null, 2);
  } catch (err) {
    return err;
  }
}

(async () => {
  await convertPPTtoHTML()
    .then((res) => {
      console.log("Document converted successfully");
      console.log("-----------------------------------------");
      console.log(res);
    })
    .catch((err) => {
      console.log("Error occurred while converting the document:", err);
    });
})();

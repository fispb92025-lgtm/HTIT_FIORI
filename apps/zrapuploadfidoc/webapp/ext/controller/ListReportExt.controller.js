sap.ui.define([
  "sap/ui/core/mvc/ControllerExtension",
  "sap/m/MessageToast",
  "sap/m/MessageBox",
  "sap/ui/core/Fragment"
], function (
  ControllerExtension,
  MessageToast,
  MessageBox,
  Fragment
) {
  "use strict";

  return ControllerExtension.extend("zrapuploadfidoc.ext.controller.ListReportExt", {
    metadata: {
      methods: {
        UploadFile: { public: true },
        onUploadPress: { public: true },
        onTempDownload: { public: true },
        onCancelPress: { public: true },
        onFileChange: { public: true },
        Post_Document: { public: true },
      }
    },

    _NS: "com.sap.gateway.srvd.zrap_ui_data_fidoc.v0001.",
    _FRAGMENT_NAME: "zrapuploadfidoc.ext.fragment.uploadFileDialog",
    _FRAGMENT_ID: "uploadFileDialog",

    onInit: function () {
      this._oDialog = null;
      this._oSelectedFile = null;
    },

    _api: function () {
      return this.base.getExtensionAPI();
    },

    _model: function () {
      return this._api().getModel();
    },

    _view: function () {
      return this.base.getView();
    },

    _getFragmentId: function () {
      return this._view().createId(this._FRAGMENT_ID);
    },

    _byFragmentId: function (sId) {
      return Fragment.byId(this._getFragmentId(), sId);
    },

    UploadFile: async function () {
      if (!this._oDialog) {
        this._oDialog = await Fragment.load({
          id: this._getFragmentId(),
          name: this._FRAGMENT_NAME,
          controller: this
        });

        this._view().addDependent(this._oDialog);
      }

      this._oSelectedFile = null;

      var oUploader = this._byFragmentId("idFileUpload");
      if (oUploader) {
        oUploader.clear();
        oUploader.setValueState("None");
      }

      this._oDialog.open();
    },

    onFileChange: function (oEvent) {
      var aFiles = oEvent.getParameter("files");
      this._oSelectedFile = aFiles && aFiles.length ? aFiles[0] : null;
    },

    onCancelPress: function () {
      if (this._oDialog) {
        this._oDialog.close();
      }
    },

    onUploadPress: async function () {
      try {
        var oUploader = this._byFragmentId("idFileUpload");
        var oFile = this._oSelectedFile;

        if (!oFile) {
          oUploader.setValueState("Error");
          oUploader.setValueStateText("Please choose a file");
          MessageToast.show("Please choose a file");
          return;
        }

        var sFileName = oFile.name;
        var sMimeType = oFile.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        var sExtension = this._getFileExtension(sFileName);
        var sBase64 = await this._readFileAsBase64(oFile);

        var oResult = await this._callCollectionAction("fileUpload", {
          mimeType: sMimeType,
          fileName: sFileName,
          fileContent: sBase64,
          fileExtension: sExtension
        });

        var sUuidFile = oResult && oResult.UuidFile;

        if (this._oDialog) {
          this._oDialog.close();
        }

        MessageToast.show((oResult && oResult.Message) || "Upload successfully");

        if (sUuidFile) {
          await this._setUuidFileFilter(sUuidFile);
        }

        await this._refreshListReport();

        this._oSelectedFile = null;

        if (oUploader) {
          oUploader.clear();
          oUploader.setValueState("None");
        }
      } catch (oError) {
        MessageBox.error(this._getErrorText(oError));
      }
    },

    onTempDownload: async function () {
      try {
        var oResult = await this._callCollectionAction("downloadFile", {});

        if (!oResult || !oResult.fileContent) {
          MessageToast.show("No file content returned");
          return;
        }

        this._downloadBase64File(
          oResult.fileContent,
          oResult.fileName || "Template",
          oResult.fileExtension || "xlsx",
          oResult.mimeType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
      } catch (oError) {
        MessageBox.error(this._getErrorText(oError));
      }
    },

    _callCollectionAction: async function (sActionName, mParameters) {
      var oModel = this._model();
      var sPath = "/DataDocument/" + this._NS + sActionName + "(...)";
      var oOperation = oModel.bindContext(sPath);

      Object.keys(mParameters || {}).forEach(function (sKey) {
        oOperation.setParameter(sKey, mParameters[sKey]);
      });

      await oOperation.execute();

      var oBoundContext = oOperation.getBoundContext();
      return oBoundContext ? oBoundContext.getObject() : null;
    },

    _readFileAsBase64: function (oFile) {
      return new Promise(function (resolve, reject) {
        var oReader = new FileReader();

        oReader.onload = function (oEvent) {
          var sResult = oEvent.target.result || "";
          var sBase64 = sResult.split(",")[1] || "";
          resolve(sBase64);
        };

        oReader.onerror = function () {
          reject(new Error("Cannot read file"));
        };

        oReader.readAsDataURL(oFile);
      });
    },

    _downloadBase64File: function (sBase64, sFileName, sExtension, sMimeType) {
      var sCleanBase64 = sBase64.indexOf(",") >= 0 ? sBase64.split(",")[1] : sBase64;
      var sBinary = window.atob(sCleanBase64);
      var aBytes = new Uint8Array(sBinary.length);

      for (var i = 0; i < sBinary.length; i++) {
        aBytes[i] = sBinary.charCodeAt(i);
      }

      var oBlob = new Blob([aBytes], { type: sMimeType });
      var sUrl = URL.createObjectURL(oBlob);
      var oLink = document.createElement("a");

      sFileName = sFileName || "download";
      sExtension = sExtension || "xlsx";

      if (sFileName.toLowerCase().indexOf("." + sExtension.toLowerCase()) < 0) {
        sFileName = sFileName + "." + sExtension;
      }

      oLink.href = sUrl;
      oLink.download = sFileName;
      document.body.appendChild(oLink);
      oLink.click();
      document.body.removeChild(oLink);

      URL.revokeObjectURL(sUrl);
    },

    _getFileExtension: function (sFileName) {
      var iIndex = sFileName.lastIndexOf(".");
      return iIndex >= 0 ? sFileName.substring(iIndex + 1).toLowerCase() : "";
    },

    _setUuidFileFilter: async function (sUuidFile) {
      try {
        if (this._api().setFilterValues) {
          await this._api().setFilterValues("UuidFile", "EQ", sUuidFile);
        }
      } catch (e) {
        // Không chặn upload nếu hệ thống không cho set filter bằng ExtensionAPI
      }
    },

    _refresh: function () {
      try {
        this._model().refresh();
      } catch (e) {
        // ignore
      }
    },

    _getErrorText: function (oError) {
      if (!oError) {
        return "Unknown error";
      }

      if (oError.message) {
        return oError.message;
      }

      if (oError.responseText) {
        try {
          var oResponse = JSON.parse(oError.responseText);
          return oResponse.error && oResponse.error.message
            ? oResponse.error.message
            : oError.responseText;
        } catch (e) {
          return oError.responseText;
        }
      }

      return String(oError);
    },

    Post_Document: async function () {
      try {
        var aContexts = this._api().getSelectedContexts();

        if (!aContexts || aContexts.length === 0) {
          MessageToast.show("Please select at least one row");
          return;
        }

        var mUnique = {};

        aContexts.forEach(function (oContext) {
          var oRow = oContext.getObject();

          var sUuidFile = oRow.UuidFile;
          var sDocumentSequenceNo = oRow.DocumentSequenceNo;
          var sCompanyCode = oRow.CompanyCode;

          var sKey = [
            sUuidFile,
            sDocumentSequenceNo,
            sCompanyCode
          ].join("|");

          if (!mUnique[sKey]) {
            mUnique[sKey] = {
              UuidFile: sUuidFile,
              DocumentSequenceNo: sDocumentSequenceNo,
              CompanyCode: sCompanyCode
            };
          }
        });

        var aPayload = Object.keys(mUnique).map(function (sKey) {
          return mUnique[sKey];
        });

        var sJson = JSON.stringify(aPayload);
        var sBase64 = this._stringToBase64(sJson);

        var oResult = await this._callCollectionAction("Post_Document", {
          mimeType: "application/json",
          fileName: "Post_Document.json",
          fileContent: sBase64,
          fileExtension: "json"
        });

        MessageToast.show("Post Document successfully");

        this._model().refresh();

        return oResult;
      } catch (oError) {
        MessageBox.error(this._getErrorText(oError));
      }
    },

    _stringToBase64: function (sValue) {
      return btoa(unescape(encodeURIComponent(sValue)));
    },

    _refreshListReport: async function () {
      try {
        var oExtensionAPI = this._api();

        /*
         * Cách chuẩn cho Fiori Elements List Report.
         * refresh() của ExtensionAPI sẽ refresh lại binding hiện tại của List Report.
         */
        if (oExtensionAPI && typeof oExtensionAPI.refresh === "function") {
          await oExtensionAPI.refresh();
          return;
        }
      } catch (e1) {
        // fallback bên dưới
      }

      try {
        /*
         * Fallback cho OData V4 model.
         */
        var oModel = this._model();

        if (oModel && typeof oModel.refresh === "function") {
          oModel.refresh();
        }
      } catch (e2) {
        // ignore
      }
    },

  });
});
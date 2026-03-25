sap.ui.define([
  "sap/ui/core/mvc/ControllerExtension",
  "sap/fe/templates/ListReport/ExtensionAPI",
  "sap/m/MessageToast",
  "sap/ui/core/Messaging",
  "sap/ui/core/message/Message",
  "sap/ui/core/message/MessageType",
  "sap/ui/core/Fragment",
  "sap/ui/core/util/File"
], function (
  ControllerExtension,
  ExtensionAPI,
  MessageToast,
  Messaging,
  Message,
  MessageType,
  Fragment,
  FileUtil
) {
  "use strict";

  /**
   * Clean, compact controller extension for FE List Report
   * - Uses async/await + one unified action invoker
   * - Consistent message handling (FE Messaging → fallback to MessageManager)
   * - Small helpers for i18n, file, Base64
   */
  return ControllerExtension.extend("zuploadfidoc.ext.controller.ListReportExt", {
    metadata: { methods: { UploadFile: { public: true }, PostDocument: { public: true } } },

    // ===== Constants ==========================================================
    _NS: "com.sap.gateway.srvd.zsd_upload_fidoc.v0001.", // OData namespace prefix
    _DIALOG_ID: "uploadFileDialog",
    _FRAGMENT: "zuploadfidoc.ext.fragment.uploadFileDialog",

    // ===== Shortcuts ==========================================================
    _api() { return this.base.getExtensionAPI(); },
    _model() { return this._api().getModel(); },
    _i18n() { return this._api().getModel("i18n"); },
    _t(key, def) {
      const b = this._i18n()?.getResourceBundle?.();
      try { return b?.getText?.(key) ?? def ?? key; } catch (e) { return def ?? key; }
    },

    // ===== when a row is selected, gray out "Post Document" if any row is already posted
    onInit() {
      // đợi view dựng xong 1 nhịp, rồi wire theo đúng vòng đời mdc.Table
      setTimeout(() => this._wireSelectionEvents(), 0);
    },

    // ===== Lifecycle ==========================================================
    onExit() { if (this._dlg) { this._dlg.destroy(); this._dlg = null; } },

    // ===== Public: Open Upload Dialog ========================================
    async UploadFile() {
      if (!this._dlg) {
        this._dlg = await this._api().loadFragment({ id: this._DIALOG_ID, name: this._FRAGMENT, controller: this });
      }
      this._dlg.open();
    },

    // ===== Public: Post Document (bound-action on collection) =================
    async PostDocument() {
      await this._secured(async () => {

        const csv = this._collectSelectedPairs();
        if (!csv) {
          sap.m.MessageToast.show(this._t("selectRowFirst", "Vui lòng chọn ít nhất 1 dòng."));
          return;
        }

        // Approach A (mass): BE action nhận collection Items
        await this._invokeAction("postFI", {
          companycode: csv.companycodeCsv,           // <-- CHỮ THƯỜNG & là CHUỖI CSV
          documentsequenceno: csv.documentsequencenoCsv
        });

        await this._refreshListReport();
        MessageToast.show(this._t("Post executed, please check the Message.", "Đã thực hiện Post, vui lòng kiểm tra Message."));
      });
    },

    // ===== Upload Handlers ====================================================
    async onFileChange(oEvent) {
      const f = (oEvent.getParameter("files") || [])[0];
      if (!f) return;

      this._file = {
        type: f.type || "",
        name: f.name || "",
        ext: (f.name || "").split(".").pop() || ""
      };

      await this._secured(() => this._readAsDataUrl(f).then((url) => {
        const m = String(url).match(/,(.*)$/); // take Base64 part only
        this._file.content = m && m[1] ? m[1] : "";
      }));
    },

    async onUploadPress() {
      if (!this._file?.content) {
        MessageToast.show(this._t("uploadFileErrMeg", "Vui lòng chọn tệp."));
        return;
      }

      await this._secured(async () => {
        await this._invokeAction("fileUpload", {
          mimeType: this._file.type,
          fileName: this._file.name,
          fileContent: this._file.content,
          fileExtension: this._file.ext
        });

        await this._refreshListReport();
        MessageToast.show(this._t("uploadFileSuccMsg", "Tải lên thành công."));
        this._resetDialog();
      });
    },

    onCancelPress() { this._resetDialog(); },

    // ===== Download Template (generic file downloader) ========================
    async onTempDownload() {
      var oModel = this.base.getExtensionAPI().getModel();
      var oI18n = this.base.getExtensionAPI().getModel("i18n");
      var oBundle = oI18n && oI18n.getResourceBundle && oI18n.getResourceBundle();

      var sActionPath = "/DataFIDoc/" + this._NS + "downloadFile(...)";
      var oOperation = oModel && oModel.bindContext(sActionPath);

      var that = this;

      function onSuccess() {
        var oResult = oOperation.getBoundContext().getObject() || {};

        // Chuẩn hoá Base64 (bỏ data:, strip whitespace, URL-safe -> standard, pad '=')
        var sB64 = that._convertBase64(oResult.fileContent);

        // Mặc định/chuẩn hoá MIME cho Excel nếu BE trả thiếu/khác chuẩn
        var sExt = (oResult.fileExtension || "").toLowerCase();
        var sMime = oResult.mimeType || "";
        if (sExt === "xlsx" && !/officedocument\.spreadsheetml\.sheet$/i.test(sMime)) {
          sMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        }
        if (!sMime) sMime = "application/octet-stream";

        // Giải mã Base64 -> Blob (nhị phân)
        var bin = atob(sB64);
        var aBin = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) aBin[i] = bin.charCodeAt(i);
        var oBlob = new Blob([aBin], { type: sMime });

        // Lưu file: KHÔNG truyền charset cho dữ liệu nhị phân
        FileUtil.save(oBlob, oResult.fileName || "result", sExt || "bin", sMime);

        MessageToast.show(oBundle ? oBundle.getText("downloadTempSuccMsg") : "Đã tải file.");
      }

      function onError(oError) {
        var thatAPI = that.base.getExtensionAPI();
        thatAPI.getEditFlow().securedExecution(function () {
          Messaging.addMessages(new Message({
            message: oError.message,
            target: "",
            persistent: true,
            type: MessageType.Error,
            code: oError.error && oError.error.code
          }));
          var aDetails = (oError.error && oError.error.details) || [];
          aDetails.forEach(function (e) {
            Messaging.addMessages(new Message({
              message: e.message,
              target: "",
              persistent: true,
              type: MessageType.Error,
              code: e.error && e.error.code
            }));
          });
        });
      }

      oOperation.invoke().then(onSuccess, onError);
    },

    // ===== Download XML (BE may return raw XML or Base64) =====================
    async GetXML() {
      await this._secured(async () => {

        const csv = this._collectSelectedPairs();
        if (!csv) {
          sap.m.MessageToast.show(this._t("selectRowFirst", "Vui lòng chọn ít nhất 1 dòng."));
          return;
        }

        const res = await this._invokeAction("downloadXML", {
          companycode: csv.companycodeCsv,           // <-- CHỮ THƯỜNG & là CHUỖI CSV
          documentsequenceno: csv.documentsequencenoCsv
        });
        this._saveFileResult(res, "xml", "application/xml");
        MessageToast.show(this._t("downloadTempSuccMsg", "Đã tải XML."));
      });
    },

    // ========================================================================
    // Action invoker (OData V4 bound action on collection)
    // ========================================================================
    async _invokeAction(actionName, params) {
      const path = `/DataFIDoc/${this._NS}${actionName}(...)`;
      const op = this._model().bindContext(path); // ODataContextBinding

      // set ALL provided parameters (đúng key đúng chữ thường như BE định nghĩa)
      if (params && typeof op.setParameter === "function") {
        Object.entries(params).forEach(([k, v]) => {
          if (v !== undefined && v !== null && v !== "") {
            op.setParameter(k, String(v));
          }
        });
      }

      try {
        await op.invoke();
      } catch (e) {
        this._pushODataErrors(e);
        this._openFEMessages();
        throw e;
      }

      const ctx = op.getBoundContext?.();
      return ctx?.getObject?.() || {};
    },

    // === Helpers: Status button Post =========================

    // --- 1) Wire all selection events robustly ---
    _wireSelectionEvents() {
      const view = this.base.getView();

      // 1) lấy mdc.Table theo stable-id; fallback tìm theo type
      let mdcTable =
        view.byId("zuploadfidoc::DataFIDocList--fe::table::DataFIDoc::LineItem::Table") ||
        view.findAggregatedObjects(true, c => c.isA?.("sap.ui.mdc.Table"))[0];
      if (!mdcTable) return;

      // 2) CHỜ mdc.Table initialized trước khi gắn listener
      const afterInit = async () => {
        try {
          // a) sự kiện chung của mdc.Table (Responsive & Grid)
          mdcTable.attachSelectionChange(this._updatePostButtonState, this);

          // b) GridTable: inner sap.ui.table.Table
          const inner = mdcTable.getTable?.(); // sap.ui.table.Table khi type = GridTable
          inner?.attachRowSelectionChange?.(this._updatePostButtonState, this);
          inner?.attachRowsUpdated?.(this._updatePostButtonState, this);

          // c) ResponsiveTable: SelectionPlugin
          const plugins = mdcTable.getPlugins?.() || [];
          const selPlugin = plugins.find(p =>
            p.isA?.("sap.m.plugins.SelectionPlugin") ||
            p.isA?.("sap.ui.table.plugins.MultiSelectionPlugin")
          );
          selPlugin?.attachSelectionChange?.(this._updatePostButtonState, this);

          // d) sau render cũng cập nhật lại (phòng khi FE re-render)
          mdcTable.addEventDelegate({ onAfterRendering: () => this._updatePostButtonState() }, this);

          // e) chạy lần đầu
          this._updatePostButtonState();
        } catch (e) {
          // no-op
        }
      };

      // API mdc.Table có promise initialized(); nếu chưa có thì bắt 1 lần
      const initPromise = typeof mdcTable.initialized === "function"
        ? mdcTable.initialized()
        : Promise.resolve();
      initPromise.then(afterInit);
    },

    // --- 2) Enable/disable action theo selection ---
    _updatePostButtonState() {
      const sel = this._api().getSelectedContexts?.() || [];

      // điều kiện “đã post”: có số chứng từ
      const hasPosted = sel.some(c => {
        const o = c.getObject?.() || {};
        return !!o.Accountingdocument;  // ĐÚNG chính tả field của bạn
      });

      const canPost = sel.length > 0 && !hasPosted;

      // Cách A: FE API (nếu có)
      try {
        if (typeof this._api().setActionEnabled === "function") {
          this._api().setActionEnabled("PostDocument", canPost);
          return;
        }
      } catch (e) { /* ignore */ }

      // Cách B: set trực tiếp vào CustomAction
      const btn = this.base.getView()
        .byId("zuploadfidoc::DataFIDocList--fe::table::DataFIDoc::LineItem::CustomAction::PostDocument");
      btn?.setEnabled?.(canPost);
    },

    // === Helpers: collect selected keys (unique) =========================
    _collectSelectedPairs() {
      const ctxs = this._api().getSelectedContexts?.() || [];
      if (!ctxs.length) return null;

      const ccSet = new Set();     // Companycode duy nhất
      const dsSet = new Set();     // Documentsequenceno duy nhất (không trùng)

      ctxs.forEach((c) => {
        const o = c.getObject?.() || {};
        const cc = (o.Companycode ?? "").toString().trim();
        const ds = (o.Documentsequenceno ?? "").toString().trim();
        if (cc) ccSet.add(cc);
        if (ds) dsSet.add(ds);
      });

      if (!ccSet.size || !dsSet.size) return null;

      // (tùy chọn) sắp xếp số tăng dần nếu toàn là số
      const dsArr = Array.from(dsSet);
      const numeric = dsArr.every(x => /^\d+$/.test(x));
      if (numeric) dsArr.sort((a, b) => Number(a) - Number(b)); else dsArr.sort();

      return {
        companycodeCsv: Array.from(ccSet).join(","),   // ví dụ: "6710"
        documentsequencenoCsv: dsArr.join(",")         // ví dụ: "1,6,9"
      };
    },

    // ========================================================================
    // Helpers: securedExecution, i18n, messages, Base64, file saving
    // ========================================================================
    _secured(fn) { return this._api().getEditFlow().securedExecution(fn, { busy: { set: true } }); },
    async _safeRefresh() { if (this._model()?.refresh) { await this._model().refresh(); } },

    async _refreshListReport() {
      // 1) FE-native refresh (ưu tiên)
      const api = this._api();
      if (typeof api.refresh === "function") {
        await api.refresh();              // FE sẽ rebind đúng cách
        return;
      }
      // 2) Fallback: refresh model
      if (this._model()?.refresh) {
        await this._model().refresh();
      }
      // 3) Tùy chọn: kích hoạt FilterBar search (nếu có stable ID)
      try {
        const fb = this.byId?.("zuploadfidoc::DataFIDocList--fe::FilterBar::DataFIDoc"); // đổi ID cho đúng app bạn
        fb?.triggerSearch?.();
      } catch (e) { /* no-op */ }

      this._updatePostButtonState();
    },

    _resetDialog() {
      try {
        const fu = Fragment.byId(this._DIALOG_ID, "idFileUpload");
        fu?.clear?.();
      } catch (e) { /* no-op */ }
      this._file = null;
      if (this._dlg) { this._dlg.close?.(); this._dlg.destroy?.(); this._dlg = null; }
    },

    _mm() { return sap.ui.getCore().getMessageManager?.(); },
    _addMessages(arr) {
      if (Messaging?.addMessages) { Messaging.addMessages(arr); return; }
      this._mm()?.addMessages?.(arr);
    },
    _openFEMessages() {
      const h = this._api().getEditFlow?.().getMessageHandler?.();
      h?.showMessages?.();
    },

    _pushODataErrors(err) {
      const root = err?.error || err?.cause?.error || {};
      const bag = [];
      const rootMsg = root?.message || err?.message;
      if (typeof rootMsg === "string" && rootMsg.trim()) {
        bag.push(new Message({ message: rootMsg, type: MessageType.Error, persistent: true, code: root?.code }));
      }
      if (Array.isArray(root?.details)) {
        root.details.forEach((d) => {
          if (d?.message) {
            bag.push(new Message({ message: d.message, type: MessageType.Error, persistent: true, code: d.code, target: d.target || "" }));
          }
        });
      }
      if (bag.length) this._addMessages(bag);
    },

    _normalizeB64(s) {
      let x = String(s || "").replace(/_/g, "/").replace(/-/g, "+");
      const m = x.length % 4; if (m) x += "=".repeat(4 - m);
      return x;
    },
    _looksLikeB64(s) {
      if (!s) return false;
      if (String(s).startsWith("<") || String(s).includes("<?xml")) return false;
      return /^[A-Za-z0-9+/=]+$/.test(s) && (s.length % 4 === 0);
    },

    _convertBase64(sUrlSafeBase64) {
        // Chuyển URL-safe Base64 -> Base64 chuẩn
        return String(sUrlSafeBase64 || "").replace(/_/g, "/").replace(/-/g, "+");
        // Nếu cần padding:
        // var m = standard.length % 4;
        // if (m > 0) { standard += "=".repeat(4 - m); }
    },

    _toBlobFromResult(res, defaultExt, defaultMime) {
      const mime = res?.mimeType || defaultMime || "application/octet-stream";
      const name = res?.fileName || "result";
      const ext = res?.fileExtension || defaultExt || "bin";
      const payload = res?.fileContent || "";

      if (this._looksLikeB64(payload)) {
        const b64 = this._normalizeB64(payload);
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        return { blob: new Blob([bytes], { type: mime }), name, ext, mime };
      }
      // treat as text
      const textBytes = new TextEncoder().encode(String(payload));
      return { blob: new Blob([textBytes], { type: `${mime};charset=utf-8` }), name, ext, mime };
    },

    _saveFileResult(res, defaultExt, defaultMime) {
      const { blob, name, ext, mime } = this._toBlobFromResult(res, defaultExt, defaultMime);
      FileUtil.save(blob, name, ext, mime, "utf-8");
    },

    _readAsDataUrl(file) {
      return new Promise((resolve, reject) => {
        try {
          const r = new FileReader();
          r.onload = (e) => resolve(e?.target?.result || "");
          r.onerror = reject;
          r.readAsDataURL(file);
        } catch (e) { reject(e); }
      });
    }
  });
});

sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"zrapuploadfidoc/test/integration/pages/DataDocumentList",
	"zrapuploadfidoc/test/integration/pages/DataDocumentObjectPage"
], function (JourneyRunner, DataDocumentList, DataDocumentObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('zrapuploadfidoc') + '/test/flp.html#app-preview',
        pages: {
			onTheDataDocumentList: DataDocumentList,
			onTheDataDocumentObjectPage: DataDocumentObjectPage
        },
        async: true
    });

    return runner;
});


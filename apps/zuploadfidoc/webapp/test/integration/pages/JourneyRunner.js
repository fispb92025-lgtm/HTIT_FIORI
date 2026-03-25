sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"zuploadfidoc/test/integration/pages/DataFIDocList",
	"zuploadfidoc/test/integration/pages/DataFIDocObjectPage"
], function (JourneyRunner, DataFIDocList, DataFIDocObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('zuploadfidoc') + '/test/flp.html#app-preview',
        pages: {
			onTheDataFIDocList: DataFIDocList,
			onTheDataFIDocObjectPage: DataFIDocObjectPage
        },
        async: true
    });

    return runner;
});


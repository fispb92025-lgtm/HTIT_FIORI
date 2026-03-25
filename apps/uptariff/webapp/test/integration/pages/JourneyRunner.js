sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"uptariff/test/integration/pages/XLHeadList",
	"uptariff/test/integration/pages/XLHeadObjectPage"
], function (JourneyRunner, XLHeadList, XLHeadObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('uptariff') + '/test/flp.html#app-preview',
        pages: {
			onTheXLHeadList: XLHeadList,
			onTheXLHeadObjectPage: XLHeadObjectPage
        },
        async: true
    });

    return runner;
});


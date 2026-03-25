sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"htitlog/test/integration/pages/ZC_TBCAT_REV_LOG003List",
	"htitlog/test/integration/pages/ZC_TBCAT_REV_LOG003ObjectPage"
], function (JourneyRunner, ZC_TBCAT_REV_LOG003List, ZC_TBCAT_REV_LOG003ObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('htitlog') + '/test/flp.html#app-preview',
        pages: {
			onTheZC_TBCAT_REV_LOG003List: ZC_TBCAT_REV_LOG003List,
			onTheZC_TBCAT_REV_LOG003ObjectPage: ZC_TBCAT_REV_LOG003ObjectPage
        },
        async: true
    });

    return runner;
});


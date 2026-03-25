sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"headercf/test/integration/pages/ZC_TBCAT_REV_CFH000List",
	"headercf/test/integration/pages/ZC_TBCAT_REV_CFH000ObjectPage"
], function (JourneyRunner, ZC_TBCAT_REV_CFH000List, ZC_TBCAT_REV_CFH000ObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('headercf') + '/test/flp.html#app-preview',
        pages: {
			onTheZC_TBCAT_REV_CFH000List: ZC_TBCAT_REV_CFH000List,
			onTheZC_TBCAT_REV_CFH000ObjectPage: ZC_TBCAT_REV_CFH000ObjectPage
        },
        async: true
    });

    return runner;
});


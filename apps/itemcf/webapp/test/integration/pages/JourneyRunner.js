sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"itemcf/test/integration/pages/ZC_TBCAT_REV_CFI000List",
	"itemcf/test/integration/pages/ZC_TBCAT_REV_CFI000ObjectPage"
], function (JourneyRunner, ZC_TBCAT_REV_CFI000List, ZC_TBCAT_REV_CFI000ObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('itemcf') + '/test/flp.html#app-preview',
        pages: {
			onTheZC_TBCAT_REV_CFI000List: ZC_TBCAT_REV_CFI000List,
			onTheZC_TBCAT_REV_CFI000ObjectPage: ZC_TBCAT_REV_CFI000ObjectPage
        },
        async: true
    });

    return runner;
});


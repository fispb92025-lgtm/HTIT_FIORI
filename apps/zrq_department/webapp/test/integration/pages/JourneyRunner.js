sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"zrqdepartment/test/integration/pages/ZC_TBRQ_DEPARTList",
	"zrqdepartment/test/integration/pages/ZC_TBRQ_DEPARTObjectPage"
], function (JourneyRunner, ZC_TBRQ_DEPARTList, ZC_TBRQ_DEPARTObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('zrqdepartment') + '/test/flp.html#app-preview',
        pages: {
			onTheZC_TBRQ_DEPARTList: ZC_TBRQ_DEPARTList,
			onTheZC_TBRQ_DEPARTObjectPage: ZC_TBRQ_DEPARTObjectPage
        },
        async: true
    });

    return runner;
});


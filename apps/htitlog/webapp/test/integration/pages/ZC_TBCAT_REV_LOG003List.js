sap.ui.define(['sap/fe/test/ListReport'], function(ListReport) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ListReport(
        {
            appId: 'htitlog',
            componentId: 'ZC_TBCAT_REV_LOG003List',
            contextPath: '/ZC_TBCAT_REV_LOG003'
        },
        CustomPageDefinitions
    );
});
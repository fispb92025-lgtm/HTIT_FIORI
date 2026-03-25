sap.ui.define(['sap/fe/test/ListReport'], function(ListReport) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ListReport(
        {
            appId: 'headercf',
            componentId: 'ZC_TBCAT_REV_CFH000List',
            contextPath: '/ZC_TBCAT_REV_CFH000'
        },
        CustomPageDefinitions
    );
});
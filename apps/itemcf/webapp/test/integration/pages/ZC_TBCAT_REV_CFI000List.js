sap.ui.define(['sap/fe/test/ListReport'], function(ListReport) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ListReport(
        {
            appId: 'itemcf',
            componentId: 'ZC_TBCAT_REV_CFI000List',
            contextPath: '/ZC_TBCAT_REV_CFI000'
        },
        CustomPageDefinitions
    );
});
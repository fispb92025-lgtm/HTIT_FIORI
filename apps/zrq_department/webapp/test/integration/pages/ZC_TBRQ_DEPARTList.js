sap.ui.define(['sap/fe/test/ListReport'], function(ListReport) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ListReport(
        {
            appId: 'zrqdepartment',
            componentId: 'ZC_TBRQ_DEPARTList',
            contextPath: '/ZC_TBRQ_DEPART'
        },
        CustomPageDefinitions
    );
});
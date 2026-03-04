sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'zrqdepartment',
            componentId: 'ZC_TBRQ_DEPARTObjectPage',
            contextPath: '/ZC_TBRQ_DEPART'
        },
        CustomPageDefinitions
    );
});
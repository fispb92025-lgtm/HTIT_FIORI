sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'htitlog',
            componentId: 'ZC_TBCAT_REV_LOG003ObjectPage',
            contextPath: '/ZC_TBCAT_REV_LOG003'
        },
        CustomPageDefinitions
    );
});
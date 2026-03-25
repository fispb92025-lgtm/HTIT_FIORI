sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'headercf',
            componentId: 'ZC_TBCAT_REV_CFH000ObjectPage',
            contextPath: '/ZC_TBCAT_REV_CFH000'
        },
        CustomPageDefinitions
    );
});
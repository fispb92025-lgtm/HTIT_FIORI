sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'itemcf',
            componentId: 'ZC_TBCAT_REV_CFI000ObjectPage',
            contextPath: '/ZC_TBCAT_REV_CFI000'
        },
        CustomPageDefinitions
    );
});
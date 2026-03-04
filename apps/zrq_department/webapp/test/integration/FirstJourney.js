sap.ui.define([
    "sap/ui/test/opaQunit",
    "./pages/JourneyRunner"
], function (opaTest, runner) {
    "use strict";

    function journey() {
        QUnit.module("First journey");

        opaTest("Start application", function (Given, When, Then) {
            Given.iStartMyApp();

            Then.onTheZC_TBRQ_DEPARTList.iSeeThisPage();

        });


        opaTest("Navigate to ObjectPage", function (Given, When, Then) {
            // Note: this test will fail if the ListReport page doesn't show any data
            
            When.onTheZC_TBRQ_DEPARTList.onFilterBar().iExecuteSearch();
            
            Then.onTheZC_TBRQ_DEPARTList.onTable().iCheckRows();

            When.onTheZC_TBRQ_DEPARTList.onTable().iPressRow(0);
            Then.onTheZC_TBRQ_DEPARTObjectPage.iSeeThisPage();

        });

        opaTest("Teardown", function (Given, When, Then) { 
            // Cleanup
            Given.iTearDownMyApp();
        });
    }

    runner.run([journey]);
});
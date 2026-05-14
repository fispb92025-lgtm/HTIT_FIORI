sap.ui.define([
    "sap/ui/test/opaQunit",
    "./pages/JourneyRunner"
], function (opaTest, runner) {
    "use strict";

    function journey() {
        QUnit.module("First journey");

        opaTest("Start application", function (Given, When, Then) {
            Given.iStartMyApp();

            Then.onTheDataDocumentList.iSeeThisPage();
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Uuidfile");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("UUID");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Messagetxt");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Messagetype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Accountingdocument");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Documentsequenceno");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Companycode");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Ytransaction");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Customer");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Vendor");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Reference");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Documentdate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Postingdate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Documenttype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Documentheadertext");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Transactioncurrency");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Grossinvoiceamount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cccurrencyamount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Businessplace");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Sectioncode");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Branchnumber");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Paymentblock");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Baselinedate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashdiscountamount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashdiscountbase");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Paymentmethod");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Pmtmethsupplement");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Paymentreference");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Invoicereference");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Invoicereferencefiscalyear");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Invoiceitemreference");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Partnerbanktype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Housebank");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Accountid");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Instruction1");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Instruction2");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Instruction3");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Instruction4");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Reasoncode");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Termsofpayment");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashdiscountdays1");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cdpercentage1");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashdiscountdays2");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cdpercentage2");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Netpmttermsperiod");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Fixed");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Reconaccount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Assignmentheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Textheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Businessareaheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Partnerbusinessareaheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Contractnumber");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Contracttype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Referencekey1");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Referencekey2");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Referencekey3");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Scbindicator");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Supplyctryreg");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Serviceindicator");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Flowtypeheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Creditcontrolarea");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Reportingctryreg");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Eutriangulardeal");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Planninglevel");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Planningdate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Negativepostingheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxcodeheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxreportingdate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxfulfilldate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxdate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxctryreg");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Exchangerate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Translationdate");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashflowheader");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashflowheaderdesc");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Title");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Name");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Name2");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Name3");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Name4");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Street");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("City");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Postalcode");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Pobox");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Poboxwithoutno");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Poboxpostalcode");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Countryregion");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Region");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Emailaddress");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Bankcountryregion");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Bankkey");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Bankaccount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Swiftbic");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Bankreference");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Bankcontrolkey");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Iban");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Paymentsystem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Aliastype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Bankaccountalias");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Liableforvat");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxtype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxnumbertype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxnumber1");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxnumber2");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxnumber3");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxnumber4");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxnumber5");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Vatregistrationno");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Naturalperson");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Salesequalizationtax");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Typeofbusiness");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Typeofindustry");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Repsname");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Instructionkey");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Dmeindicator");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Companycodeagain");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Glaccount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Itemtext");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Debitcredit");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Grossitemamount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Ccitemamount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxcodeitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxjurisdiction");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Assignmentitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Costcenter");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Profitcenter");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Partnerprofitcenter");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Orderno");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Reportingsegment");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Negativepostingitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Wbselement");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Businessareaitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Businessprocess");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Controllingarea");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Activitytype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Costobject");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Functionalarea");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Withoutcashdiscount");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Personnelnumber");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Salesdocument");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Salesdocumentitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Schedulelinenumber");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Plant");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Material");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Network");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Operationactivity");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Workitemid");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Commitmentitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Fundscenter");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxamountintrcurrency");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxamountincccurrency");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxbaseamtintransactioncurr");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Taxbaseamtincccurrency");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Yfund");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Ygrant");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Baseunitofmeasure");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Quantity");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Partnerbusinessareaitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Servicedocumenttype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Servicedocument");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Servicedocumentitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Financialtransactiontype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Jointventure");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Equitygroup");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Recoveryindicator");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Tradingpartnerno");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashflowitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Cashflowitemdesc");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Providercontract");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Contractitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Customeritem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Customergroup");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Industry");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Countryregionkey");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Salesorganization");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Distributionchannel");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Division");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Billingtype");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Longtext");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Housebankitem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Accountiditem");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Created By");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Created On");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Changed By");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Changed On");
            Then.onTheDataDocumentList.onFilterBar().iCheckFilterField("Changed On");
            Then.onTheDataDocumentList.onTable().iCheckColumns(186, {"UuidFile":{"header":"Uuidfile"},"UUID":{"header":"UUID"},"MessageTxt":{"header":"Messagetxt"},"MessageType":{"header":"Messagetype"},"AccountingDocument":{"header":"Accountingdocument"},"DocumentSequenceNo":{"header":"Documentsequenceno"},"CompanyCode":{"header":"Companycode"},"YTransaction":{"header":"Ytransaction"},"Customer":{"header":"Customer"},"Vendor":{"header":"Vendor"},"Reference":{"header":"Reference"},"DocumentDate":{"header":"Documentdate"},"PostingDate":{"header":"Postingdate"},"DocumentType":{"header":"Documenttype"},"DocumentHeaderText":{"header":"Documentheadertext"},"TransactionCurrency":{"header":"Transactioncurrency"},"GrossInvoiceAmount":{"header":"Grossinvoiceamount"},"CcCurrencyAmount":{"header":"Cccurrencyamount"},"BusinessPlace":{"header":"Businessplace"},"SectionCode":{"header":"Sectioncode"},"BranchNumber":{"header":"Branchnumber"},"PaymentBlock":{"header":"Paymentblock"},"BaselineDate":{"header":"Baselinedate"},"CashDiscountAmount":{"header":"Cashdiscountamount"},"CashDiscountBase":{"header":"Cashdiscountbase"},"PaymentMethod":{"header":"Paymentmethod"},"PmtmethSupplement":{"header":"Pmtmethsupplement"},"PaymentReference":{"header":"Paymentreference"},"InvoiceReference":{"header":"Invoicereference"},"InvoiceReferenceFiscalYear":{"header":"Invoicereferencefiscalyear"},"InvoiceItemReference":{"header":"Invoiceitemreference"},"PartnerBankType":{"header":"Partnerbanktype"},"HouseBank":{"header":"Housebank"},"AccountID":{"header":"Accountid"},"Instruction1":{"header":"Instruction1"},"Instruction2":{"header":"Instruction2"},"Instruction3":{"header":"Instruction3"},"Instruction4":{"header":"Instruction4"},"ReasonCode":{"header":"Reasoncode"},"TermsOfPayment":{"header":"Termsofpayment"},"CashDiscountDays1":{"header":"Cashdiscountdays1"},"CdPercentage1":{"header":"Cdpercentage1"},"CashDiscountDays2":{"header":"Cashdiscountdays2"},"CdPercentage2":{"header":"Cdpercentage2"},"NetpmtTermsPeriod":{"header":"Netpmttermsperiod"},"Fixed":{"header":"Fixed"},"ReconAccount":{"header":"Reconaccount"},"AssignmentHeader":{"header":"Assignmentheader"},"TextHeader":{"header":"Textheader"},"BusinessAreaHeader":{"header":"Businessareaheader"},"PartnerBusinessAreaHeader":{"header":"Partnerbusinessareaheader"},"ContractNumber":{"header":"Contractnumber"},"ContractType":{"header":"Contracttype"},"Referencekey1":{"header":"Referencekey1"},"Referencekey2":{"header":"Referencekey2"},"Referencekey3":{"header":"Referencekey3"},"ScbIndicator":{"header":"Scbindicator"},"Supplyctryreg":{"header":"Supplyctryreg"},"ServiceIndicator":{"header":"Serviceindicator"},"FlowTypeHeader":{"header":"Flowtypeheader"},"CreditControlArea":{"header":"Creditcontrolarea"},"Reportingctryreg":{"header":"Reportingctryreg"},"Eutriangulardeal":{"header":"Eutriangulardeal"},"PlanningLevel":{"header":"Planninglevel"},"PlanningDate":{"header":"Planningdate"},"NegativePostingHeader":{"header":"Negativepostingheader"},"TaxCodeHeader":{"header":"Taxcodeheader"},"TaxReportingDate":{"header":"Taxreportingdate"},"TaxFulFillDate":{"header":"Taxfulfilldate"},"TaxDate":{"header":"Taxdate"},"Taxctryreg":{"header":"Taxctryreg"},"ExchangeRate":{"header":"Exchangerate"},"TranslationDate":{"header":"Translationdate"},"CashFlowHeader":{"header":"Cashflowheader"},"CashFlowHeaderDesc":{"header":"Cashflowheaderdesc"},"Title":{"header":"Title"},"Name":{"header":"Name"},"Name2":{"header":"Name2"},"Name3":{"header":"Name3"},"Name4":{"header":"Name4"},"Street":{"header":"Street"},"City":{"header":"City"},"PostalCode":{"header":"Postalcode"},"Pobox":{"header":"Pobox"},"PoboxWithoutNo":{"header":"Poboxwithoutno"},"PoboxPostalCode":{"header":"Poboxpostalcode"},"CountryRegion":{"header":"Countryregion"},"Region":{"header":"Region"},"EmailAddress":{"header":"Emailaddress"},"BankCountryRegion":{"header":"Bankcountryregion"},"BankKey":{"header":"Bankkey"},"BankAccount":{"header":"Bankaccount"},"Swiftbic":{"header":"Swiftbic"},"BankReference":{"header":"Bankreference"},"BankControlKey":{"header":"Bankcontrolkey"},"Iban":{"header":"Iban"},"PaymentSystem":{"header":"Paymentsystem"},"AliasType":{"header":"Aliastype"},"BankAccountAlias":{"header":"Bankaccountalias"},"Liableforvat":{"header":"Liableforvat"},"TaxType":{"header":"Taxtype"},"TaxNumberType":{"header":"Taxnumbertype"},"TaxNumber1":{"header":"Taxnumber1"},"TaxNumber2":{"header":"Taxnumber2"},"TaxNumber3":{"header":"Taxnumber3"},"TaxNumber4":{"header":"Taxnumber4"},"TaxNumber5":{"header":"Taxnumber5"},"VatRegisTrationNo":{"header":"Vatregistrationno"},"NaturalPerson":{"header":"Naturalperson"},"SaleSequalizationTax":{"header":"Salesequalizationtax"},"TypeofBusiness":{"header":"Typeofbusiness"},"TypeofIndustry":{"header":"Typeofindustry"},"RepsName":{"header":"Repsname"},"InstructionKey":{"header":"Instructionkey"},"DmeIndicator":{"header":"Dmeindicator"},"CompanyCodeAgain":{"header":"Companycodeagain"},"GLAccount":{"header":"Glaccount"},"ItemText":{"header":"Itemtext"},"DebitCredit":{"header":"Debitcredit"},"GrossItemAmount":{"header":"Grossitemamount"},"CcItemAmount":{"header":"Ccitemamount"},"TaxCodeItem":{"header":"Taxcodeitem"},"Taxjurisdiction":{"header":"Taxjurisdiction"},"AssignmentItem":{"header":"Assignmentitem"},"CostCenter":{"header":"Costcenter"},"ProfitCenter":{"header":"Profitcenter"},"PartnerProfitCenter":{"header":"Partnerprofitcenter"},"OrderNo":{"header":"Orderno"},"ReportingSegment":{"header":"Reportingsegment"},"NegativePostingItem":{"header":"Negativepostingitem"},"WbsElement":{"header":"Wbselement"},"BusinessAreaItem":{"header":"Businessareaitem"},"BusinessProcess":{"header":"Businessprocess"},"ControllingArea":{"header":"Controllingarea"},"ActivityType":{"header":"Activitytype"},"CostObject":{"header":"Costobject"},"FunctionalArea":{"header":"Functionalarea"},"WithoutCashDiscount":{"header":"Withoutcashdiscount"},"PersonnelNumber":{"header":"Personnelnumber"},"SalesDocument":{"header":"Salesdocument"},"SalesDocumentItem":{"header":"Salesdocumentitem"},"SchedulelineNumber":{"header":"Schedulelinenumber"},"Plant":{"header":"Plant"},"Material":{"header":"Material"},"Network":{"header":"Network"},"OperationActivity":{"header":"Operationactivity"},"WorkItemId":{"header":"Workitemid"},"CommitmentItem":{"header":"Commitmentitem"},"FundsCenter":{"header":"Fundscenter"},"TaxAmountIntrcurrency":{"header":"Taxamountintrcurrency"},"TaxAmountIncccurrency":{"header":"Taxamountincccurrency"},"TaxbaseAmtIntransactioncurr":{"header":"Taxbaseamtintransactioncurr"},"TaxbaseAmtIncccurrency":{"header":"Taxbaseamtincccurrency"},"Yfund":{"header":"Yfund"},"Ygrant":{"header":"Ygrant"},"BaseUnitofMeasure":{"header":"Baseunitofmeasure"},"Quantity":{"header":"Quantity"},"PartnerBusinessAreaItem":{"header":"Partnerbusinessareaitem"},"ServiceDocumentType":{"header":"Servicedocumenttype"},"ServiceDocument":{"header":"Servicedocument"},"ServiceDocumentItem":{"header":"Servicedocumentitem"},"FinancialTransactionType":{"header":"Financialtransactiontype"},"JointVenture":{"header":"Jointventure"},"EquityGroup":{"header":"Equitygroup"},"RecoveryIndicator":{"header":"Recoveryindicator"},"TradingPartnerNo":{"header":"Tradingpartnerno"},"CashFlowItem":{"header":"Cashflowitem"},"CashFlowItemdesc":{"header":"Cashflowitemdesc"},"ProviderContract":{"header":"Providercontract"},"ContractItem":{"header":"Contractitem"},"CustomerItem":{"header":"Customeritem"},"CustomerGroup":{"header":"Customergroup"},"Industry":{"header":"Industry"},"CountryRegionKey":{"header":"Countryregionkey"},"SalesOrganization":{"header":"Salesorganization"},"DistributionChannel":{"header":"Distributionchannel"},"Division":{"header":"Division"},"BillingType":{"header":"Billingtype"},"LongText":{"header":"Longtext"},"HouseBankItem":{"header":"Housebankitem"},"AccountidItem":{"header":"Accountiditem"},"CreatedBy":{"header":"Created By"},"CreatedAt":{"header":"Created On"},"LocalLastChangedBy":{"header":"Changed By"},"LocalLastChangedAt":{"header":"Changed On"},"LastChangedAt":{"header":"Changed On"}});

        });


        opaTest("Navigate to ObjectPage", function (Given, When, Then) {
            // Note: this test will fail if the ListReport page doesn't show any data
            
            When.onTheDataDocumentList.onFilterBar().iExecuteSearch();
            
            Then.onTheDataDocumentList.onTable().iCheckRows();

            When.onTheDataDocumentList.onTable().iPressRow(0);
            Then.onTheDataDocumentObjectPage.iSeeThisPage();

        });

        opaTest("Teardown", function (Given, When, Then) { 
            // Cleanup
            Given.iTearDownMyApp();
        });
    }

    runner.run([journey]);
});
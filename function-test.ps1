function Convert-LicenseString {
    # Remove first set of characters matching "$STRING:"
    $String = $args[0] -Replace '^[^:]+:\s*', ''
    # Rename licenses to human-readable format
    
    switch ($String) {
        'O365_BUSINESS_ESSENTIALS'              {'Office365BusinessEssentials'}
        'O365_BUSINESS_PREMIUM'                 {'Office365BusinessPremium'}
        'DESKLESSPACK'                          {'Office365(PlanK1)'}
        'DESKLESSWOFFPACK'				        {'Office365(PlanK2)'}
        'LITEPACK'						        {'Office365(PlanP1)'}
        'EXCHANGESTANDARD'				        {'Office365ExchangeOnlineOnly'}
        'STANDARDPACK'					        {'EnterprisePlanE1'}
        'STANDARDWOFFPACK'				        {'Office365(PlanE2)'}
        'ENTERPRISEPACK'					    {'EnterprisePlanE3'}
        'ENTERPRISEPACKLRG'				        {'EnterprisePlanE3'}
        'ENTERPRISEWITHSCAL'				    {'EnterprisePlanE4'}
        'STANDARDPACK_STUDENT'			        {'Office365(PlanA1)forStudents'}
        'STANDARDWOFFPACKPACK_STUDENT'	        {'Office365(PlanA2)forStudents'}
        'ENTERPRISEPACK_STUDENT'			    {'Office365(PlanA3)forStudents'}
        'ENTERPRISEWITHSCAL_STUDENT'		    {'Office365(PlanA4)forStudents'}
        'STANDARDPACK_FACULTY'			        {'Office365(PlanA1)forFaculty'}
        'STANDARDWOFFPACKPACK_FACULTY'	        {'Office365(PlanA2)forFaculty'}
        'ENTERPRISEPACK_FACULTY'			    {'Office365(PlanA3)forFaculty'}
        'ENTERPRISEWITHSCAL_FACULTY'		    {'Office365(PlanA4)forFaculty'}
        'ENTERPRISEPACK_B_PILOT'			    {'Office365(EnterprisePreview)'}
        'STANDARD_B_PILOT'				        {'Office365(SmallBusinessPreview)'}
        'VISIOCLIENT'					        {'VisioProOnline'}
        'POWER_BI_ADDON'					    {'Office365PowerBIAddon'}
        'POWER_BI_INDIVIDUAL_USE'		        {'PowerBIIndividualUser'}
        'POWER_BI_STANDALONE'			        {'PowerBIStandAlone'}
        'POWER_BI_STANDARD'				        {'Power-BIStandard'}
        'PROJECTESSENTIALS'				        {'ProjectLite'}
        'PROJECTCLIENT'					        {'ProjectProfessional'}
        'PROJECTONLINE_PLAN_1'			        {'ProjectOnline'}
        'PROJECTONLINE_PLAN_2'			        {'ProjectOnlineandPRO'}
        'ProjectPremium'					    {'ProjectOnlinePremium'}
        'ECAL_SERVICES'					        {'ECAL'}
        'EMS'							        {'EnterpriseMobilitySuite'}
        'RIGHTSMANAGEMENT_ADHOC'			    {'WindowsAzureRightsManagement'}
        'MCOMEETADV'						    {'PSTNconferencing'}
        'SHAREPOINTSTORAGE'				        {'SharePointstorage'}
        'PLANNERSTANDALONE'				        {'PlannerStandalone'}
        'CRMIUR'							    {'CMRIUR'}
        'BI_AZURE_P1'					        {'PowerBIReportingandAnalytics'}
        'INTUNE_A'						        {'WindowsIntunePlanA'}
        'PROJECTWORKMANAGEMENT'			        {'Office365PlannerPreview'}
        'ATP_ENTERPRISE'					    {'ExchangeOnlineAdvancedThreatProtection'}
        'EQUIVIO_ANALYTICS'				        {'Office365AdvancedeDiscovery'}
        'AAD_BASIC'						        {'AzureActiveDirectoryBasic'}
        'RMS_S_ENTERPRISE'				        {'AzureActiveDirectoryRightsManagement'}
        'AAD_PREMIUM'					        {'AzureActiveDirectoryPremium'}
        'MFA_PREMIUM'					        {'AzureMulti-FactorAuthentication'}
        'STANDARDPACK_GOV'				        {'MicrosoftOffice365(PlanG1)forGovernment'}
        'STANDARDWOFFPACK_GOV'			        {'MicrosoftOffice365(PlanG2)forGovernment'}
        'ENTERPRISEPACK_GOV'				    {'MicrosoftOffice365(PlanG3)forGovernment'}
        'ENTERPRISEWITHSCAL_GOV'			    {'MicrosoftOffice365(PlanG4)forGovernment'}
        'DESKLESSPACK_GOV'				        {'MicrosoftOffice365(PlanK1)forGovernment'}
        'ESKLESSWOFFPACK_GOV'			        {'MicrosoftOffice365(PlanK2)forGovernment'}
        'EXCHANGESTANDARD_GOV'			        {'MicrosoftOffice365ExchangeOnline(Plan1)onlyforGovernment'}
        'EXCHANGEENTERPRISE_GOV'			    {'MicrosoftOffice365ExchangeOnline(Plan2)onlyforGovernment'}
        'SHAREPOINTDESKLESS_GOV'			    {'SharePointOnlineKiosk'}
        'EXCHANGE_S_DESKLESS_GOV'		        {'ExchangeKiosk'}
        'RMS_S_ENTERPRISE_GOV'			        {'WindowsAzureActiveDirectoryRightsManagement'}
        'OFFICESUBSCRIPTION_GOV'			    {'OfficeProPlus'}
        'MCOSTANDARD_GOV'				        {'LyncPlan2G'}
        'SHAREPOINTWAC_GOV'				        {'OfficeOnlineforGovernment'}
        'SHAREPOINTENTERPRISE_GOV'		        {'SharePointPlan2G'}
        'EXCHANGE_S_ENTERPRISE_GOV'		        {'ExchangePlan2G'}
        'EXCHANGE_S_ARCHIVE_ADDON_GOV'	        {'ExchangeOnlineArchiving'}
        'EXCHANGE_S_DESKLESS'			        {'ExchangeOnlineKiosk'}
        'SHAREPOINTDESKLESS'				    {'SharePointOnlineKiosk'}
        'SHAREPOINTWAC'					        {'OfficeOnline'}
        'YAMMER_ENTERPRISE'				        {'YammerfortheStarshipEnterprise'}
        'EXCHANGE_L_STANDARD'			        {'ExchangeOnline(Plan1)'}
        'MCOLITE'						        {'LyncOnline(Plan1)'}
        'SHAREPOINTLITE'					    {'SharePointOnline(Plan1)'}
        'OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ'    {'OfficeProPlus'}
        'EXCHANGE_S_STANDARD_MIDMARKET'	        {'ExchangeOnline(Plan1)'}
        'MCOSTANDARD_MIDMARKET'			        {'LyncOnline(Plan1)'}
        'SHAREPOINTENTERPRISE_MIDMARKET'	    {'SharePointOnline(Plan1)'}
        'OFFICESUBSCRIPTION'				    {'OfficeProPlus'}
        'YAMMER_MIDSIZE'					    {'Yammer'}
        'DYN365_ENTERPRISE_PLAN1'		        {'Dynamics365CustomerEngagementPlanEnterpriseEdition'}
        'ENTERPRISEPREMIUM_NOPSTNCONF'	        {'EnterpriseE5(withoutAudioConferencing)'}
        'ENTERPRISEPREMIUM'				        {'EnterpriseE5(withAudioConferencing)'}
        'MCOSTANDARD'					        {'SkypeforBusinessOnlineStandalonePlan2'}
        'PROJECT_MADEIRA_PREVIEW_IW_SKU'	    {'Dynamics365forFinancialsforIWs'}
        'STANDARDWOFFPACK_IW_STUDENT'	        {'Office365EducationforStudents'}
        'STANDARDWOFFPACK_IW_FACULTY'	        {'Office365EducationforFaculty'}
        'EOP_ENTERPRISE_FACULTY'			    {'ExchangeOnlineProtectionforFaculty'}
        'EXCHANGESTANDARD_STUDENT'		        {'ExchangeOnline(Plan1)forStudents'}
        'OFFICESUBSCRIPTION_STUDENT'		    {'OfficeProPlusStudentBenefit'}
        'STANDARDWOFFPACK_FACULTY'		        {'Office365EducationE1forFaculty'}
        'STANDARDWOFFPACK_STUDENT'		        {'MicrosoftOffice365(PlanA2)forStudents'}
        'DYN365_FINANCIALS_BUSINESS_SKU'	    {'Dynamics365forFinancialsBusinessEdition'}
        'DYN365_FINANCIALS_TEAM_MEMBERS_SKU'    {'Dynamics365forTeamMembersBusinessEdition'}
        'FLOW_FREE'						        {'MicrosoftFlowFree'}
        'POWER_BI_PRO'					        {'PowerBIPro'}
        'O365_BUSINESS'					        {'Office365Business'}
        'DYN365_ENTERPRISE_SALES'		        {'DynamicsOffice365EnterpriseSales'}
        'RIGHTSMANAGEMENT'				        {'RightsManagement'}
        'PROJECTPROFESSIONAL'			        {'ProjectProfessional'}
        'VISIOONLINE_PLAN1'				        {'VisioOnlinePlan1'}
        'EXCHANGEENTERPRISE'				    {'ExchangeOnlinePlan2'}
        'DYN365_ENTERPRISE_P1_IW'		        {'Dynamics365P1TrialforInformationWorkers'}
        'DYN365_ENTERPRISE_TEAM_MEMBERS'	    {'Dynamics365ForTeamMembersEnterpriseEdition'}
        'CRMSTANDARD'					        {'MicrosoftDynamicsCRMOnlineProfessional'}
        'EXCHANGEARCHIVE_ADDON'			        {'ExchangeOnlineArchivingForExchangeOnline'}
        'EXCHANGEDESKLESS'				        {'ExchangeOnlineKiosk'}
        'SPZA_IW'						        {'AppConnect'}
        'WINDOWS_STORE'					        {'WindowsStoreforBusiness'}
        'MCOEV'							        {'MicrosoftPhoneSystem'}
        'VIDEO_INTEROP'					        {'PolycomSkypeMeetingVideoInteropforSkypeforBusiness'}
        'SPE_E5'							    {'Microsoft365E5'}
        'SPE_E3'							    {'Microsoft365E3'}
        'ATA'							        {'AdvancedThreatAnalytics'}
        'MCOPSTN2'						        {'DomesticandInternationalCallingPlan'}
        'FLOW_P1'						        {'MicrosoftFlowPlan1'}
        'FLOW_P2'						        {'MicrosoftFlowPlan2'}
        'CRMSTORAGE'						    {'MicrosoftDynamicsCRMOnlineAdditionalStorage'}
        'SMB_APPS'						        {'MicrosoftBusinessApps'}
        'MICROSOFT_BUSINESS_CENTER'		        {'MicrosoftBusinessCenter'}
        'DYN365_TEAM_MEMBERS'			        {'Dynamics365TeamMembers'}
        'STREAM'							    {'MicrosoftStreamTrial'}
        'EMSPREMIUM'                            {'ENTERPRISEMOBILITY+SECURITYE5'}
        default                                 {'Unknown'}
    }
}
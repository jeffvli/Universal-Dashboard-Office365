function Convert-LicenseString {
    # Remove first set of characters matching "$STRING:"
    $String = $args[0] -Replace '^[^:]+:\s*', ''
    # Rename licenses to human-readable format
    switch ($String) {
        'DESKLESSPACK'                      {'K[1]'}
        'DESKLESSWOFFPACK'                  {'K[2]'}
        'FLOW_FREE'                         {'Flow Free'}
        'LITEPACK'                          {'{P[1]'}
        'EXCHANGESTANDARD'                  {'Exchange Online'}
        'O365_BUSINESS_ESSENTIALS'          {'Business Essentials'}
        'STANDARDPACK'                      {'E[1]'}
        'STANDARDWOFFPACK'                  {'E[2]'}
        'ENTERPRISEPACK'                    {'E[3]'}
        'ENTERPRISEPACKLRG'                 {'E[3]'}
        'ENTERPRISEWITHSCAL'                {'E[4]'}
        'STANDARDPACK_STUDENT'              {'A[1] for Students'}
        'STANDARDWOFFPACKPACK_STUDENT'      {'A[2] for Students'}
        'ENTERPRISEPACK_STUDENT'            {'A[3] for Students'}
        'ENTERPRISEWITHSCAL_STUDENT'        {'A[4] for Students'}
        'STANDARDPACK_FACULTY'              {'A[1] for Faculty'}
        'STANDARDWOFFPACKPACK_FACULTY'      {'A[2] for Faculty'}
        'ENTERPRISEPACK_FACULTY'            {'A[3] for Faculty'}
        'ENTERPRISEWITHSCAL_FACULTY'        {'A[4] for Faculty'}
        'ENTERPRISEPACK_B_PILOT'            {'Enterprise Preview'}
        'STANDARD_B_PILOT'                  {'Small Business Preview'}
        'VISIOCLIENT'                       {'Visio'}
        'POWER_BI_ADDON'                    {'Power BI Addon'}
        'POWER_BI_INDIVIDUAL_USE'           {'Power BI Individual'}
        'POWER_BI_STANDALONE'               {'Power BI Stand-Alone'}
        'POWER_BI_STANDARD'                 {'Power BI Standard'}
        'PROJECTESSENTIALS'                 {'Project Lite'}
        'PROJECTCLIENT'                     {'Project Professional'}
        'PROJECTPROFESSIONAL'               {'Project Professional'}
        'PROJECTONLINE_PLAN_1'              {'Project P[1]'}
        'PROJECTONLINE_PLAN_2'              {'Project P[2]'}
        'ECAL_SERVICES'                     {'ECAL'}
        'EMS'                               {'Enterprise Mobility Suite'}
        'RIGHTSMANAGEMENT_ADHOC'            {'Windows Azure Rights Management'}
        'MCOMEETADV'                        {'Skype'}
        'SHAREPOINTSTORAGE'                 {'SharePoint'}
        'PLANNERSTANDALONE'                 {'Planner'}
        'CRMIUR'                            {'CMRIUR'}
        'BI_AZURE_P1'                       {'Power BI Reporting and Analytics'}
        'INTUNE_A'                          {'Windows Intune Plan A'}
        'MICROSOFT_BUSINESS_CENTER'         {'Microsoft Business Center'}
        default                             {'Unknown'}
    }
}
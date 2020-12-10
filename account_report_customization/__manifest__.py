# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

{
    'name': 'Account Report Customization',
    'version': '13.0.1.0',
    'summary': 'Account Report Customization',
    'category': '',
    'description': """
        Account Report Customization
    """,
    'author': 'Caret IT Solutions Pvt. Ltd.',
    'website': 'http://www.caretit.com',
    'depends': ['analytic','account_accountant','account_reports'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/account_custom_common_wizard.xml',
        'report/general_ledger_report.xml',
        'views/account_custom_view.xml',
    ],
    'qweb': [],
    'installable': True,
    'application': False,
}

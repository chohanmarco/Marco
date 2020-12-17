# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

{
    'name': 'Account Trial Balance',
    'version': '13.0.1.0',
    'summary': 'Account Trial Balance',
    'category': '',
    'description': """
        Adds Trial Balance report in PDF and Xls Format
    """,
    'author': 'Caret IT Solutions Pvt. Ltd.',
    'website': 'http://www.caretit.com',
    'depends': ['account_report_customization'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/account_trial_balance_wizard.xml',
        'report/trial_balance_report.xml',
        'views/account_tb_custom_view.xml',
    ],
    'qweb': [],
    'installable': True,
    'application': False,
}

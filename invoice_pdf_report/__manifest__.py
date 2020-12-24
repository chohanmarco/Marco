# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

{
    'name': 'Invoice PDF Report',
    'version': '13.0.1.0',
    'summary': 'Invoice PDF Custom Report',
    'category': '',
    'description': """
        This module adds PDF custom report for invoice
    """,
    'author': 'Caret IT Solutions Pvt. Ltd.',
    'website': 'http://www.caretit.com',
    'depends': ['account'],
    'data': [
        'report/report_invoice_marco.xml',
        'views/account_move_custom_view.xml',
        'views/res_company_view.xml'
    ],
    'qweb': [],
    'installable': True,
    'application': False,
}

# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

{
    'name': 'Invoice Account ID Customization',
    'version': '13.0.1.0',
    'summary': 'Invoice Account Customization',
    'category': '',
    'description': """
        This module selects Account ID Based on Set in Income Account field in Partner
    """,
    'author': 'Caret IT Solutions Pvt. Ltd.',
    'website': 'http://www.caretit.com',
    'depends': ['account'],
    'data': [
        'views/partner_view.xml',
    ],
    'qweb': [],
    'installable': True,
    'application': False,
}

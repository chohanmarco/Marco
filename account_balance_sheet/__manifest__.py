# -*- coding: utf-8 -*-
{
    'name': "account_balance_sheet",

    'summary': """
        Balance Sheet""",

    'description': """
        Adds Balance Sheet report in PDF and Xls Format
    """,

    'author': "My Company",
    'website': "http://www.yourcompany.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/13.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '0.1',

    # any module necessary for this one to work correctly
    'depends': ['base','account','analytic','account_accountant',],

    # always loaded
    'data': [
        # 'security/ir.model.access.csv',
        'views/views.xml',
        'wizard/balance_sheet_wizard.xml',
        'report/balance_sheet_report.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
        'demo/demo.xml',
    ],
}

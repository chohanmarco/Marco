# -*- coding: utf-8 -*-
# from odoo import http


# class AccountBalanceSheet(http.Controller):
#     @http.route('/account_balance_sheet/account_balance_sheet/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/account_balance_sheet/account_balance_sheet/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('account_balance_sheet.listing', {
#             'root': '/account_balance_sheet/account_balance_sheet',
#             'objects': http.request.env['account_balance_sheet.account_balance_sheet'].search([]),
#         })

#     @http.route('/account_balance_sheet/account_balance_sheet/objects/<model("account_balance_sheet.account_balance_sheet"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('account_balance_sheet.object', {
#             'object': obj
#         })

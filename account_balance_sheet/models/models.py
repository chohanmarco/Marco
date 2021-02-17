# -*- coding: utf-8 -*-

from odoo import models, fields, api


class AccountAccountsInherit(models.Model):
    _inherit = "account.account"

    temp_accounts = fields.Boolean(string= 'Select', default=False)
   

class AccountAnalyticAccounts(models.Model):
    _inherit = "account.analytic.account"

    temp_analytics = fields.Boolean(string= 'Select', default=False)


class AccountMoveLine(models.Model):
    _inherit = "account.move.line"

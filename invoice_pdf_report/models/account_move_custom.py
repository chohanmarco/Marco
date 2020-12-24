# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################
from odoo import models, fields, api, _


class AccountMove(models.Model):
    _inherit = "account.move"

    purchase_order_number = fields.Char(string= 'Purchase Order No')
    projectname = fields.Text(string= 'Project Name')
    bank_account_id = fields.Many2one('res.partner.bank',string="Bank Account Number")
    notes = fields.Text(string="Notes")


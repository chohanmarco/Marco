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
    notes = fields.Text(string="Description")
    is_contract =   fields.Boolean(string="For Contract?")
    contract_order_number = fields.Char(string= 'Contract No')
    internal_custom_invoice_no = fields.Char(string="Invoice No")

    def amount_to_text(self, amount, currency):
        convert_amount_in_words = self.company_id.currency_id.amount_to_text(amount)
        convert_amount_in_words = convert_amount_in_words.replace('And', '')
        convert_amount_in_words = convert_amount_in_words.replace('Riyal', 'Saudi Riyals')
        return convert_amount_in_words
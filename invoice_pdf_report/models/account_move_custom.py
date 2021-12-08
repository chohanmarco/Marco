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
        convert_amount_in_words = self.currency_id.amount_to_text(amount)
        convert_amount_in_words = convert_amount_in_words.replace('And', '')
        convert_amount_in_words = convert_amount_in_words.replace('Riyal', 'Saudi Riyals')
        convert_amount_in_words = convert_amount_in_words.replace('Dollars', 'US Dollars')
        return convert_amount_in_words

    def invoice_amount_in_words(self, lang, amount):
        convert_amount_in_words_arabic = self.currency_id.with_context(lang='ar_001').amount_to_text(amount)
        convert_amount_in_words_arabic = convert_amount_in_words_arabic.replace('Riyal', 'Saudi Riyals')
        convert_amount_in_words_arabic = convert_amount_in_words_arabic.replace('Dollars', 'US Dollars')
        convert_amount_in_words_arabic = convert_amount_in_words_arabic.replace('Saudi Riyals', 'ريال سعودي')
        convert_amount_in_words_arabic = convert_amount_in_words_arabic.replace('Halala', 'هللة')
        convert_amount_in_words_arabic = convert_amount_in_words_arabic.replace('US Dollars', 'دولار أمريكي')
        convert_amount_in_words_arabic = convert_amount_in_words_arabic.replace('Cents', 'هللة')
        return convert_amount_in_words_arabic
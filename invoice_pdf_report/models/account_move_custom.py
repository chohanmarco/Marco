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
        # convert_amount_in_words = convert_amount_in_words.replace('Dollars', 'US Dollars')
        return convert_amount_in_words


class AccountMoveLine(models.Model):
    _inherit = "account.move.line"

    @api.onchange('product_id')
    def _onchange_product_id(self):
        for line in self:
            if not line.product_id or line.display_type in ('line_section', 'line_note'):
                continue

            line.name = line._get_computed_name()
            if line.move_id.partner_id.income_account:
                line.account_id = line.move_id.partner_id.income_account
            else:
                line.account_id = line._get_computed_account()
            taxes = line._get_computed_taxes()
            if taxes and line.move_id.fiscal_position_id:
                taxes = line.move_id.fiscal_position_id.map_tax(taxes, partner=line.partner_id)
            line.tax_ids = taxes
            line.product_uom_id = line._get_computed_uom()
            line.price_unit = line._get_computed_price_unit()

        if len(self) == 1:
            return {'domain': {'product_uom_id': [('category_id', '=', self.product_uom_id.category_id.id)]}}

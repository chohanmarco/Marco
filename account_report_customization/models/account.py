# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################
from odoo import models, fields, api, _


class AccountAccount(models.Model):
    _inherit = "account.account"

    temp_for_report = fields.Boolean(string= 'Select', default=False)

class AccountMoveLine(models.Model):
    _inherit = "account.move.line"


    wraplabel = fields.Text(string= 'Lable.', compute='_get_wraplabel')

    @api.depends('name')
    def _get_wraplabel(self):
        for line in self:
            NewLabel = ''
            a_string = line.name
            if a_string:
                split_strings = []
                n  = 40
                for index in range(0, len(a_string), n):
                    split_strings.append(a_string[index : index + n])
                NewLabel = '\n'.join(split_strings)
                line.wraplabel = NewLabel
            else:
                line.wraplabel = NewLabel
                
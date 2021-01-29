# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

from odoo import api, models, _
from odoo.exceptions import UserError
from collections import defaultdict
from datetime import datetime


class TrialBalanceReport(models.AbstractModel):
    _name = 'report.account_trial_balance.report_trial_balance'
    _description = 'Trial Balance Report'

    @api.model
    def _get_report_values(self, docids, data=None):
       
        if not data.get('form'):
            raise UserError(_("Form content is missing, this report cannot be printed."))
        
        return {
            'stockData': data.get('get_trial_balance'),
            'data': data,
        }

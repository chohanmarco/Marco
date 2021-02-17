from odoo import api, models, _
from odoo.exceptions import UserError
from collections import defaultdict
from datetime import datetime


class BalanceSheetReport(models.AbstractModel):
    _name = 'report.account_balance_sheet.account_balance_sheet'
    _description = 'Balance Sheet Report'

    # @api.model
    # def _get_report_values(self, docids, data=None):
       
    #     if not data.get('form'):
    #         raise UserError(_("Form content is missing, this report cannot be printed."))
        
    #     return {
    #         'stockData': data.get('get_trial_balance'),
    #         'data': data,
    #     }
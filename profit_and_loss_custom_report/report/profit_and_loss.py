from odoo import api, models, _
from odoo.exceptions import UserError
from collections import defaultdict
from datetime import datetime


class ProfitLossReport(models.AbstractModel):
    _name = 'report.profit_and_loss_custom_report.report_profit_loss'
    _description = 'Profit And Loss Report'



    # @api.model
    # def _get_report_values(self, docids, data=None):
    #     if not data.get('form'):
    #         raise UserError(_("Form content is missing, this report cannot be printed."))

    #     return {
    #         'stockdata': data['get_profit_loss'],
    #         'data': data.get('form'),
    #     }

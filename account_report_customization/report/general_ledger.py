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


class GeneralLedgerReport(models.AbstractModel):
    _name = 'report.account_report_customization.report_general_ledger'
    _description = 'General Ledger Report'

    @api.model
    def _get_report_values(self, docids, data=None):
        if not data.get('form'):
            raise UserError(_("Form content is missing, this report cannot be printed."))

        # dateFrom = data['form'].get('date_from')
        # dateTo = data['form'].get('date_to')

        # AccountIds = data['form'].get('account_ids')
        # Status = ['posted']
        # MoveLines = False
        # if AccountIds:
        #     self.env.cr.execute("""
        #         SELECT aml.id
        #         FROM account_move_line aml
        #         LEFT JOIN account_move am ON (am.id=aml.move_id)
        #         WHERE (aml.date >= %s) AND
        #             (aml.date <= %s) AND
        #             (aml.account_id in %s) AND
        #             (am.state in %s) ORDER BY aml.date""",
        #         (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountIds), tuple(Status),))
        #     MoveLines = [x[0] for x in self.env.cr.fetchall()]
        # mainDict = defaultdict(list)
        # Balance = 0.0
        # for ml in self.env['account.move.line'].sudo().browse(MoveLines):
        #     Balance = Balance + (ml.debit - ml.credit)
        #     mainDict[ml.account_id.display_name or '-'].append({'date': ml.date or '',
        #                                                 'move': ml.move_id and ml.move_id.name or '',
        #                                                 'ref' : ml.ref or '',
        #                                                 'name' : ml.name or '',
        #                                                 'debit': ml.debit or 0.0,
        #                                                 'credit': ml.credit or 0.0,
        #                                                 'balance': Balance or 0.0,
        #                                                 'project': ml.project_id and ml.project_id.name or ''
        #                                             })
        return {
            # 'stockData': mainDict,
            'stockData': data.get('get_general_ledger'),
            'data': data,
        }

# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

from odoo import api, fields, models, _
from odoo.exceptions import UserError
from collections import defaultdict
from datetime import datetime
import xlwt
import io
import numpy as np
import pandas as pd
from PIL import Image as PILImage
from odoo.tools.float_utils import float_round

class AccountTrialBalanceReport(models.TransientModel):
    _name = 'account.trial.balance.report'
    _description = "Account Trial Balance Report"

    date_from = fields.Date(string="From Date")
    date_to = fields.Date(string="To Date")
    account_ids = fields.Many2many('account.account', string='Accounts')
    account_without_transaction = fields.Boolean(string= 'Show Accounts without transactions', default=False)
    account_zero_closing_balance = fields.Boolean(string= 'Show Accounts with zero closing balance', default=False)
    dimension_wise_project = fields.Selection([('none','None'),('month','Month Wise'),('dimension', 'Dimension')],
                                              string='Dimension',
                                              default='none')
    projectwise = fields.Selection([('project', 'Project')],string='Project',default='project')
    detail_report = fields.Boolean(string= 'Show Detail Report(Accounting)', default=False)
    show_dr_cr_separately = fields.Boolean(string= 'Show Opening Dr/Cr Separately', default=False)
    analytic_account_ids = fields.Many2many('account.analytic.account', string='Analytic Accounts')

    def action_redirect_to_aml_view(self):
        action = self.env.ref('account_report_customization.action_account_moves_all_filter_with_report_tree').read()[0]
        dateFrom = self.date_from
        dateTo = self.date_to
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        Status = ['posted']
        MoveLines = []
        if AccountIds:
            self.env.cr.execute("""
                SELECT aml.id
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountIds), tuple(Status),))
            MoveLines = [x[0] for x in self.env.cr.fetchall()]
        action['context'] = {'create': False}
        action['domain'] = [('id', 'in', MoveLines)]
        return action   

    def print_report_trial_balance(self):
        if self.date_from >= self.date_to:
            raise UserError(_("Start Date is greater than or equal to End Date."))
        datas = {'form': self.read()[0],
                 'get_trial_balance': self.get_trial_balance_detail()
            }
        return self.env.ref('account_trial_balance.action_report_trial_balance').report_action([], data=datas)

    def get_trial_balance_detail(self):
        dateFrom = self.date_from
        dateTo = self.date_to
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        Status = ['posted']
        MoveLineIds = []
        mainDict = defaultdict(list)
        DynamicList = [analytic_account.name for analytic_account in self.env['account.analytic.account'].search([])]
        for Account in self.env['account.account'].browse(AccountIds):
            Balance = 0.0
            self.env.cr.execute("""
                SELECT aml.date as date,
                       aml.debit as debit,
                       aml.credit as credit,
                       aml.id as movelineid
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(Status),))
            MoveLineIds = self.env.cr.fetchall()

            self.env.cr.execute("""
                SELECT sum(aml.debit) as debit,
                       sum(aml.credit) as credit
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date < %s) AND
                    (aml.account_id = %s) AND
                    (am.state in %s)""",
                (str(dateFrom) + ' 00:00:00', Account.id, tuple(Status),))
            OpeningMove = self.env.cr.fetchone()
            
            OpeningDebit = 0.0
            if OpeningMove[0] is None:
                OpeningDebit = 0.0
            else:
                OpeningDebit = OpeningMove[0]

            OpeningCredit = 0.0
            if OpeningMove[1] is None:
                OpeningCredit = 0.0
            else:
                OpeningCredit = OpeningMove[1]

            FinalOpeningDebit = 0.0
            FinalOpeningCredit = 0.0
            OpeningBalance = OpeningDebit - OpeningCredit
            if OpeningBalance > 0.0:
                FinalOpeningDebit = OpeningBalance
            elif OpeningBalance < 0.0:
                FinalOpeningCredit = abs(OpeningBalance)
            ClosingDebit = 0.0
            ClosingCredit = 0.0
            NetBalance = 0.0
            if MoveLineIds:
                total_op_debit = 0.0
                total_op_credit = 0.0
                total_tr_debit = 0.0
                total_tr_credit = 0.0
                total_closing_debit = 0.0
                total_closing_credit = 0.0
                total_netbalance = 0.0
                for ml in MoveLineIds:
                    AnalyticVals = []
                    #Balance = Balance + (ml[3] - ml[4])
                    total_tr_debit += ml[1]
                    total_tr_credit += ml[2]

                ClosingBalance = OpeningBalance + (total_tr_debit - total_tr_credit)
                if ClosingBalance > 0.0:
                    ClosingDebit = ClosingBalance or 0.0
                elif ClosingBalance < 0.0:
                    ClosingCredit = abs(ClosingBalance)
                NetBalance = ClosingDebit - ClosingCredit

                if self.account_zero_closing_balance and Balance == 0.0:
                    Vals = {'acccode': Account.code or '',
                            'accname': Account.name or '',
                            'openingdr': FinalOpeningDebit,
                            'openingcr': FinalOpeningCredit,
                            'trdebit': total_tr_debit or 0.0,
                            'trcredit': total_tr_credit or 0.0,
                            'closingdr':ClosingDebit or 0.0,
                            'closingcr':ClosingCredit or 0.0,
                            'netbalance': NetBalance or 0.0,
                            }
                    mainDict[Account.name or '-'].append(Vals)
    #                         if self.dimension_wise_project:
    #                             for analytic in DynamicList:
    #                                 if ml.analytic_account_id.name == analytic:
    #                                     AnalyticVals.append({analytic:ml.debit-ml.credit or 0.0})
    #                                 else:
    #                                     AnalyticVals.append({analytic:0.0})
    #                         Vals.update({'analytic_vals':AnalyticVals})
                else:
                    Vals = {'acccode':Account.code or '',
                            'accname': Account.name or '',
                            'openingdr': FinalOpeningDebit,
                            'openingcr': FinalOpeningCredit,
                            'trdebit': total_tr_debit or 0.0,
                            'trcredit': total_tr_credit or 0.0,
                            'closingdr':ClosingDebit or 0.0,
                            'closingcr':ClosingCredit or 0.0,
                            'netbalance': NetBalance or 0.0,
                            }
                    mainDict[Account.name or '-'].append(Vals)
    #                         if self.dimension_wise_project:
    #                             for analytic in DynamicList:
    #                                 if ml.analytic_account_id.name == analytic:
    #                                     bala = ml.debit-ml.credit
    #                                     AnalyticVals.append({analytic:ml.debit-ml.credit or 0.0})
    #                                 else:
    #                                     AnalyticVals.append({analytic:0.0})
    #                         Vals.update({'analytic_vals':AnalyticVals})
            if self.account_without_transaction and not MoveLineIds:
                AnalyticVals = []
                Vals = {'acccode':Account.code,
                        'accname':Account.name,
                        'openingdr':0.0,
                        'openingcr':0.0,
                        'trdebit': 0.0,
                        'trcredit':0.0,
                        'closingdr':0.0,
                        'closingcr':0.0,
                        'netbalance':0.0,
                        }
                mainDict[Account.name or '-'].append(Vals)
#                 if self.dimension_wise_project:
#                     for analytic in DynamicList:
#                         AnalyticVals.append({analytic:0.0})
#                 Vals.update({'analytic_vals':AnalyticVals})
        return mainDict

    @api.model
    def default_get(self, fields):
        vals = super(AccountTrialBalanceReport, self).default_get(fields)
        ac_ids = self.env['account.account'].search([])
        analytic_ids = self.env['account.analytic.account'].search([])
        self.env.cr.execute('update account_account set temp_for_report=False')
        self.env.cr.execute('update account_analytic_account set temp_analytic_report=False')
        if 'account_ids' in fields and not vals.get('account_ids') and ac_ids:
            iids = []
            for ac_id in ac_ids:
                iids.append(ac_id.id)
            vals['account_ids'] = [(6, 0, iids)]
        if 'analytic_account_ids' in fields and not vals.get('analytic_account_ids') and analytic_ids:
            aniids = []
            for ana_ac in analytic_ids:
                aniids.append(ana_ac.id)
            vals['analytic_account_ids'] = [(6, 0, aniids)]
        return vals

    def trial_balance_export_excel(self):
        """
        This methods make list of dict to Export in Dailybook Excel
        """
        CompanyImage = self.env.company.logo
        dateFrom = self.date_from
        dateTo = self.date_to
        mainDict = defaultdict(list)
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        Status = ['posted']
        Projectwise = self.dimension_wise_project
        MoveLines = []

        AllAnalyticAccounts = self.analytic_account_ids
        FilteredAnalyticAccountIds = AllAnalyticAccounts.filtered(lambda a: a.temp_analytic_report)
        AnalyticAccountIds = FilteredAnalyticAccountIds
        if not AnalyticAccountIds:
            AnalyticAccountIds = AllAnalyticAccounts

        mainDict = defaultdict(list)
        DynamicList = [analytic_account.name for analytic_account in self.env['account.analytic.account'].search([])]
        for Account in self.env['account.account'].browse(AccountIds):
            Balance = 0.0
            self.env.cr.execute("""
                SELECT aml.date as date,
                       aml.debit as debit,
                       aml.credit as credit,
                       aa.name as analytic,
                       aml.id as movelineid
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                LEFT JOIN account_analytic_account aa ON (aa.id=aml.analytic_account_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(Status),))
            MoveLineIds = self.env.cr.fetchall()

            self.env.cr.execute("""
                SELECT sum(aml.debit) as debit,
                       sum(aml.credit) as credit
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date < %s) AND
                    (aml.account_id = %s) AND
                    (am.state in %s)""",
                (str(dateFrom) + ' 00:00:00', Account.id, tuple(Status),))
            OpeningMove = self.env.cr.fetchone()
            
            OpeningDebit = 0.0
            if OpeningMove[0] is None:
                OpeningDebit = 0.0
            else:
                OpeningDebit = OpeningMove[0]

            OpeningCredit = 0.0
            if OpeningMove[1] is None:
                OpeningCredit = 0.0
            else:
                OpeningCredit = OpeningMove[1]

            FinalOpeningDebit = 0.0
            FinalOpeningCredit = 0.0
            OpeningBalance = OpeningDebit - OpeningCredit
            if OpeningBalance > 0.0:
                FinalOpeningDebit = OpeningBalance
            elif OpeningBalance < 0.0:
                FinalOpeningCredit = abs(OpeningBalance)
            ClosingDebit = 0.0
            ClosingCredit = 0.0
            NetBalance = 0.0
            if MoveLineIds:
                total_op_debit = 0.0
                total_op_credit = 0.0
                total_tr_debit = 0.0
                total_tr_credit = 0.0
                total_closing_debit = 0.0
                total_closing_credit = 0.0
                total_netbalance = 0.0
                AnalyticVals = []
                Vals = {}
                for ml in MoveLineIds:
                    AnalyticVals = []
                    #Balance = Balance + (ml[3] - ml[4])
                    total_tr_debit += ml[1]
                    total_tr_credit += ml[2]
                    if Projectwise == 'dimension':
                        for analytic in DynamicList:
                            if ml[3] == analytic and ml[3] is not None:
                                AnalyticVals.append({analytic:ml[1]-ml[2] or 0.0})
                            else:
                                AnalyticVals.append({analytic:0.0})
                    Vals.update({'analytic_vals':AnalyticVals})


                ClosingBalance = OpeningBalance + (total_tr_debit - total_tr_credit)
                if ClosingBalance > 0.0:
                    ClosingDebit = ClosingBalance or 0.0
                elif ClosingBalance < 0.0:
                    ClosingCredit = abs(ClosingBalance)
                NetBalance = ClosingDebit - ClosingCredit

                if self.account_zero_closing_balance and Balance == 0.0:
                    Vals.update({'acccode': Account.code or '',
                            'accname': Account.name or '',
                            'openingdr': FinalOpeningDebit,
                            'openingcr': FinalOpeningCredit,
                            'trdebit': total_tr_debit or 0.0,
                            'trcredit': total_tr_credit or 0.0,
                            'closingdr':ClosingDebit or 0.0,
                            'closingcr':ClosingCredit or 0.0,
                            'netbalance': NetBalance or 0.0,
                            })
                    mainDict[Account.name or '-'].append(Vals)
                else:
                    Vals.update({'acccode':Account.code or '',
                            'accname': Account.name or '',
                            'openingdr': FinalOpeningDebit,
                            'openingcr': FinalOpeningCredit,
                            'trdebit': total_tr_debit or 0.0,
                            'trcredit': total_tr_credit or 0.0,
                            'closingdr':ClosingDebit or 0.0,
                            'closingcr':ClosingCredit or 0.0,
                            'netbalance': NetBalance or 0.0,
                            })
                    mainDict[Account.name or '-'].append(Vals)
#                     if Projectwise == 'dimension':
#                         for analytic in DynamicList:
#                             if ml.analytic_account_id.name == analytic:
#                                 AnalyticVals.append({analytic:ml.debit-ml.credit or 0.0})
#                             else:
#                                 AnalyticVals.append({analytic:0.0})
#                     Vals.update({'analytic_vals':AnalyticVals})
            if self.account_without_transaction and not MoveLineIds:
                AnalyticVals = []
                Vals = {'acccode':Account.code,
                        'accname':Account.name,
                        'openingdr':0.0,
                        'openingcr':0.0,
                        'trdebit': 0.0,
                        'trcredit':0.0,
                        'closingdr':0.0,
                        'closingcr':0.0,
                        'netbalance':0.0,
                        }
                mainDict[Account.name or '-'].append(Vals)
                if Projectwise == 'dimension':
                    for analytic in DynamicList:
                        AnalyticVals.append({analytic:0.0})
                Vals.update({'analytic_vals':AnalyticVals})

        import base64
        filename = 'Trial Balance.xls'
        start = dateFrom.strftime('%d %b, %Y')
        end = dateTo.strftime('%d %b, %Y')
        form_name = 'Trial Balance Between ' + str(start) + ' to ' + str(end)
        workbook = xlwt.Workbook()
        style = xlwt.XFStyle()
        tall_style = xlwt.easyxf('font:height 720;') # 36pt
        # Create a font to use with the style
        font = xlwt.Font()
        font.name = 'Times New Roman'
        font.bold = True
        font.height = 250
        style.font = font
        xlwt.add_palette_colour("custom_colour", 0x21)
        workbook.set_colour_RGB(0x21, 105, 105, 105)

        xlwt.add_palette_colour("dark_blue", 0x3A)
        workbook.set_colour_RGB(0x3A, 0,0,139)  

        xlwt.add_palette_colour("gainsboro", 0x15)
        workbook.set_colour_RGB(0x15,205,205,205)  

        worksheet = workbook.add_sheet("Trial Balance", cell_overwrite_ok=True)
        worksheet.show_grid = False

        styleheader = xlwt.easyxf('font: bold 1, colour black, height 300;')
        stylecolumnheader = xlwt.easyxf('font: bold 1, colour white, height 200;pattern: pattern solid, fore_colour custom_colour')
        linedata = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        stylecolaccount = xlwt.easyxf('font: bold 1, colour white, height 200; \
                                      pattern: pattern solid, fore_colour dark_blue; \
                                      align: vert centre, horiz centre; \
                                      borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        analytic_st_col = xlwt.easyxf('font: bold 1, colour black, height 200; \
                                    pattern: pattern solid, fore_colour gainsboro; \
                                    align: vert centre, horiz centre; \
                                    borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        general = xlwt.easyxf('font: bold 1, colour black, height 210;')
        dateheader = xlwt.easyxf('font: bold 1, colour black, height 200;')
        maintotal = xlwt.easyxf('font: bold 1, colour black, height 200; \
                borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;')
        finaltotalheader = xlwt.easyxf('pattern: fore_color white; font: bold 1, colour black; align: horiz right; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;')
        rightfont = xlwt.easyxf('pattern: fore_color white; font: color black; align: horiz right; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;')
        floatstyle = xlwt.easyxf("borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;", "#,###.00")
        finaltotalheaderbold = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour black; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;", "#,###.00")

        zero_col = worksheet.col(0)
        zero_col.width = 236 * 20
        first_col = worksheet.col(1)
        first_col.width = 236 * 40
        second_col = worksheet.col(2)
        second_col.width = 236 * 20
        third_col = worksheet.col(3)
        third_col.width = 236 * 20
        fourth_col = worksheet.col(4)
        fourth_col.width = 236 * 20
        fifth_col = worksheet.col(5)
        fifth_col.width = 236 * 20
        sixth_col = worksheet.col(6)
        sixth_col.width = 236 * 20
        seventh_col = worksheet.col(7)
        seventh_col.width = 236 * 20
        eigth_col = worksheet.col(8)
        eigth_col.width = 236 * 20
        #HEADER
        worksheet.row(4).height_mismatch = True
        worksheet.row(4).height = 360
        
        worksheet.write_merge(0, 1, 2, 5, self.env.company.name,styleheader)
        worksheet.write_merge(2, 2, 2, 5, 'Trial Balance',general)
        headerstring = 'From :' + str(self.date_from.strftime('%d %b %Y') or '') + ' To :' + str(self.date_to.strftime('%d %b %Y') or '')
        worksheet.write_merge(3, 3, 2, 5, headerstring,dateheader)
        
        #SUB-HEADER
        row = 4
        ColIndexes = {}
        if self.show_dr_cr_separately:
            ColIndexes = {'Account Code':0,
                          'Account Name':1,
                          'Op.Debit':2,
                          'Op.Credit':3,
                          'Tr.Debit':4,
                          'Tr.Credit':5,
                          'Closing Debit':6,
                          'Closing Credit':7,
                          'Net Balance':8}
        else:
            ColIndexes = {'Account Code':0,
                          'Account Name':1,
                          'Tr.Debit':2,
                          'Tr.Credit':3,
                          'Closing Debit':4,
                          'Closing Credit':5,
                          'Net Balance':6}            
        worksheet.write(row, 0, 'Account Code', stylecolaccount)
        worksheet.write(row, 1, 'Account Name', stylecolaccount)
        opcol=2
        if self.show_dr_cr_separately:
            worksheet.write(row, opcol, 'Op.Debit', stylecolaccount)
            opcol+=1
            worksheet.write(row, opcol, 'Op.Credit', stylecolaccount)
            opcol+=1
        worksheet.write(row, opcol, 'Tr.Debit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Tr.Credit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Closing Debit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Closing Credit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Net Balance', stylecolaccount)
        opcol+=1

        if self.show_dr_cr_separately:
            calc = 8
        else:
            calc = 6
        if Projectwise == 'dimension':
            if self.show_dr_cr_separately:
                col = 9
            else:
                col = 7
            for analytic in AnalyticAccountIds:
                dictval = {analytic.name:col}
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, analytic.name, analytic_st_col)
                col+=1
                calc+=1
        row = 5
        totalopeningdebit = 0.0
        totalopeningcredit = 0.0
        totaltrdebit = 0.0
        totaltrcredit = 0.0
        totalclosingdebit = 0.0
        totalclosingcredit = 0.0
        totalnetbalance = 0.0
        FinalDict = {}
        for ac in mainDict:
            newDict = {}
            worksheet.row(row).height_mismatch = True
            worksheet.row(row).height = 310
            for line in mainDict.get(ac):
                if self.show_dr_cr_separately:
                    totalopeningdebit+= line.get('openingdr')
                    totalopeningcredit+= line.get('openingcr')
                totaltrdebit+= line.get('trdebit')
                totaltrcredit+= line.get('trcredit')
                totalclosingdebit += line.get('closingdr')
                totalclosingcredit += line.get('closingdr')
                totalnetbalance += line.get('netbalance')
                worksheet.write(row, 0, line.get('acccode',False),linedata)
                worksheet.write(row, 1, line.get('accname',''),linedata)
                linecol = 2
                if self.show_dr_cr_separately:
                    if line.get('openingdr',0.0) == 0.0:
                        worksheet.write(row, linecol, 0.0,rightfont)
                    else:
                        worksheet.write(row, linecol, line.get('openingdr',0.0),floatstyle)
                    linecol+=1
                    if line.get('openingcr',0.0) == 0.0:
                        worksheet.write(row, linecol, 0.0,rightfont)
                    else:
                        worksheet.write(row, linecol, line.get('openingcr',0.0),floatstyle)
                    linecol+=1
                
                if line.get('trdebit',0.0) == 0.0:
                    worksheet.write(row, linecol, 0.0 ,rightfont)
                else:
                    worksheet.write(row, linecol, line.get('trdebit',0.0),floatstyle)
                linecol+=1
                
                if line.get('trcredit',0.0) == 0.0:
                    worksheet.write(row, linecol, 0.0 ,rightfont)
                else:
                    worksheet.write(row, linecol, line.get('trcredit',0.0),floatstyle)
                linecol+=1
                
                if line.get('closingdr',0.0) == 0.0:
                    worksheet.write(row, linecol, 0.0 ,rightfont)
                else:
                    worksheet.write(row, linecol, line.get('closingdr',0.0),floatstyle)
                linecol+=1
                
                if line.get('closingcr',0.0) == 0.0:
                    worksheet.write(row, linecol, 0.0 ,rightfont)
                else:
                    worksheet.write(row, linecol, line.get('closingcr',0.0),floatstyle)
                linecol+=1

                if line.get('netbalance',0.0) == 0.0:
                    worksheet.write(row, linecol, 0.0 ,rightfont)
                else:
                    worksheet.write(row, linecol, line.get('netbalance',0.0),floatstyle)
                linecol+=1
                if Projectwise == 'dimension':
                    for nl in line.get('analytic_vals'):
                        for k,v in nl.items():
                            if ColIndexes.get(k) is not None:
                                if v == 0.0:
                                    worksheet.write(row, ColIndexes.get(k), 0.0,rightfont)
                                else:
                                    worksheet.write(row, ColIndexes.get(k), v or 0.0,floatstyle)
                                keylocation = ColIndexes.get(k)
                                if keylocation in newDict:
                                    newDict[keylocation] += v
                                else:
                                    newDict.update({
                                        keylocation: v or 0.0,
                                                    })
            
            for newkey,newval in newDict.items():
                if newkey in FinalDict:
                    FinalDict[newkey] += newval
                else:
                    FinalDict.update({newkey:newval})
            row+=1
#            worksheet.write_merge(row, row, 0, 1, 'REPORT TOTAL', style = dateheader)
            worksheet.write(row, 0, 'REPORT TOTAL', style = maintotal)
            worksheet.write(row, 1, '-', style = maintotal)
            finalcol = 2
            if self.show_dr_cr_separately:
                if totalopeningdebit == 0.0:
                    worksheet.write(row, finalcol, 0.0,finaltotalheader)
                else:
                    worksheet.write(row, finalcol, totalopeningdebit,finaltotalheaderbold)
                finalcol+=1
                
                if totalopeningcredit == 0.0:
                    worksheet.write(row, finalcol, 0.0,finaltotalheader)
                else:
                    worksheet.write(row, finalcol, totalopeningcredit,finaltotalheaderbold)
                finalcol+=1
                
            if totaltrdebit == 0.0:
                worksheet.write(row, finalcol, 0.0,finaltotalheader)
            else:
                worksheet.write(row, finalcol, totaltrdebit,finaltotalheaderbold)
            finalcol+=1
            if totaltrcredit == 0.0:
                worksheet.write(row, finalcol, 0.0,finaltotalheader)
            else:
                worksheet.write(row, finalcol, totaltrcredit,finaltotalheaderbold)
            finalcol+=1
            if totalclosingcredit == 0.0:
                worksheet.write(row, finalcol, 0.0,finaltotalheader)
            else:
                worksheet.write(row, finalcol, totalclosingcredit,finaltotalheaderbold)
            finalcol+=1
            if totalclosingdebit == 0.0:
                worksheet.write(row, finalcol, 0.0,finaltotalheader)
            else:
                worksheet.write(row, finalcol, totalclosingdebit,finaltotalheaderbold)
            finalcol+=1
            if totalnetbalance == 0.0:
                worksheet.write(row, finalcol, 0.0,finaltotalheader)
            else:
                worksheet.write(row, finalcol, totalnetbalance,finaltotalheaderbold)
            finalcol+=1
            for finalkey,finalval in FinalDict.items():
                if finalval == 0.0:
                    worksheet.write(row, finalkey, 0.0,style=finaltotalheader)
                else:
                    worksheet.write(row, finalkey, finalval,style=finaltotalheaderbold)
                        
        row+=2
        buffer = io.BytesIO()
        workbook.save(buffer)
        export_id = self.env['trial.balance.excel'].create(
                        {'excel_file': base64.encodestring(buffer.getvalue()), 'file_name': filename})
        buffer.close()
    
        return {
            'name': form_name,
            'view_mode': 'form',
            'res_id': export_id.id,
            'res_model': 'trial.balance.excel',
            'view_mode': 'form',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

class trial_balance_export_excel(models.TransientModel):
    _name= "trial.balance.excel"
    _description = "Trial Balance Excel Report"

    excel_file = fields.Binary('Report for Trial Balance')
    file_name = fields.Char('File', size=64)

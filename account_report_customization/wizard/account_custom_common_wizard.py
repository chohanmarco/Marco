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

class AccountCustomReport(models.TransientModel):
    _name = 'account.custom.report'
    _description = "Account Custom Report"

    date_from = fields.Date(string="From Date")
    date_to = fields.Date(string="To Date")
    account_ids = fields.Many2many('account.account', string='Accounts')
    account_without_transaction = fields.Boolean(string= 'Show Accounts without transactions', default=False)
    account_zero_closing_balance = fields.Boolean(string= 'Show Accounts with zero closing balance', default=False)
    dimension_wise_project = fields.Boolean(string= 'Dimension Wise Project', default=False)
    dimensions = fields.Selection([('project', 'Project')],string='Dimension',default='project')
    detail_report = fields.Boolean(string= 'Show Detail Report(Accounting)', default=False)
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

        AllAnalyticAccounts = self.analytic_account_ids
        FilteredAnalyticAccountIds = AllAnalyticAccounts.filtered(lambda a: a.temp_analytic_report)
        AnalyticAccountIds = FilteredAnalyticAccountIds.ids
        if not AnalyticAccountIds:
            AnalyticAccountIds = AllAnalyticAccounts.ids
            
        Status = ['posted']
        MoveLines = []
        if self.dimension_wise_project:
            self.env.cr.execute("""
                SELECT aml.id
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (aml.analytic_account_id in %s) AND
                    (am.state in %s) """,
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountIds), tuple(AnalyticAccountIds), tuple(Status),))
            MoveLines = [x[0] for x in self.env.cr.fetchall()]
        else:
            self.env.cr.execute("""
                SELECT aml.id
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) """,
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountIds), tuple(Status),))
            MoveLines = [x[0] for x in self.env.cr.fetchall()]            
        action['context'] = {'create': False}
        action['domain'] = [('id', 'in', MoveLines)]
        return action   

    def print_report(self):
        if self.date_from >= self.date_to:
            raise UserError(_("Start Date is greater than or equal to End Date."))
        datas = {'form': self.read()[0],
                 'get_general_ledger': self.get_general_ledger_detail()
            }
        return self.env.ref('account_report_customization.action_report_general_ledger').report_action([], data=datas)

    def get_general_ledger_detail(self):
        dateFrom = self.date_from
        dateTo = self.date_to
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        Status = ['posted']
        MoveLines = []
        mainDict = defaultdict(list)
        StaticList = ['date','move','name','debit','credit','balance']
        DynamicList = [analytic_account.name for analytic_account in self.env['account.analytic.account'].search([])]
        for Account in self.env['account.account'].browse(AccountIds):
            Balance = 0.0
            self.env.cr.execute("""
                SELECT aml.id
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(Status),))
            MoveLines = [x[0] for x in self.env.cr.fetchall()]
            MoveLineIds = self.env['account.move.line'].sudo().browse(MoveLines)
            if MoveLineIds:
                for ml in MoveLineIds:
                    AnalyticVals = []
                    Balance = Balance + (ml.debit - ml.credit)
                    if self.account_zero_closing_balance and Balance == 0.0:
                        Vals = {'date': str(ml.date.strftime('%d/%b/%Y')) or '',
                                'move': ml.move_id and ml.move_id.name or '',
                                'name' : ml.name or '',
                                'debit': ml.debit or 0.0,
                                'credit': ml.credit or 0.0,
                                'balance': Balance or 0.0,
                                }
                        mainDict[Account.name or '-'].append(Vals)
                    else:
                        Vals = {'date': str(ml.date.strftime('%d/%b/%Y')) or '',
                                'move': ml.move_id and ml.move_id.name or '',
                                'name' : ml.name or '',
                                'debit': ml.debit or 0.0,
                                'credit': ml.credit or 0.0,
                                'balance': Balance or 0.0,
                                }
                        mainDict[Account.name or '-'].append(Vals)
            if self.account_without_transaction and not MoveLines:
                AnalyticVals = []
                Vals = {'date':False,
                        'move':'',
                        'name':'',
                        'debit':0.0,
                        'credit':0.0,
                        'balance':0.0,
                      }
                mainDict[Account.name or '-'].append(Vals)
        return mainDict

    @api.model
    def default_get(self, fields):
        vals = super(AccountCustomReport, self).default_get(fields)
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

    def get_initials(self,account):
        AccountAccountObj = self.env['account.account']
        AccountTypeObj = self.env['account.account.type']
        Account = AccountAccountObj.search([('name','=',account)])
        ExcludeTypes = AccountTypeObj.search([('name','in',['Expenses','Depreciation','Cost of Revenue'])])
        Status = ['posted']
        if Account.user_type_id.id in ExcludeTypes.ids:
            return 0,0,0
        if Account:
            self.env.cr.execute("""
                SELECT sum(aml.debit)
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date < %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s)""",
                (str(self.date_from) + ' 00:00:00', tuple([Account.id]), tuple(Status),))
            InitialDebit = self.env.cr.fetchone()
            self.env.cr.execute("""
                SELECT sum(aml.credit)
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date < %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s)""",
                (str(self.date_from) + ' 00:00:00', tuple([Account.id]), tuple(Status),))
            InitialCredit = self.env.cr.fetchone()

            if InitialDebit[0] is None:
                InitialDebit = 0.0 
            else:
                InitialDebit = InitialDebit[0]
                
            if InitialCredit[0] is None:
                InitialCredit = 0.0
            else:
                InitialCredit = InitialCredit[0]
            InitialBalance = InitialDebit - InitialCredit
            return InitialDebit,InitialCredit,InitialBalance

    def general_ledger_export_excel(self):
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

        DynamicList = [analytic_account.name for analytic_account in AnalyticAccountIds]
        MoveLines = []
        for Account in self.env['account.account'].browse(AccountIds):
            Balance = 0.0
            if self.dimension_wise_project:
                self.env.cr.execute("""
                    SELECT aml.id
                    FROM account_move_line aml
                    LEFT JOIN account_move am ON (am.id=aml.move_id)
                    WHERE (aml.date >= %s) AND
                        (aml.date <= %s) AND
                        (aml.account_id in %s) AND
                        (aml.analytic_account_id in %s) AND
                        (am.state in %s) ORDER BY aml.date""",
                    (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(AnalyticAccountIds.ids), tuple(Status),))
                MoveLines = [x[0] for x in self.env.cr.fetchall()]
            else:
                self.env.cr.execute("""
                    SELECT aml.id
                    FROM account_move_line aml
                    LEFT JOIN account_move am ON (am.id=aml.move_id)
                    WHERE (aml.date >= %s) AND
                        (aml.date <= %s) AND
                        (aml.account_id in %s) AND
                        (am.state in %s) ORDER BY aml.date""",
                    (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(Status),))
                MoveLines = [x[0] for x in self.env.cr.fetchall()]
            MoveLineIds = self.env['account.move.line'].sudo().browse(MoveLines)
            if MoveLineIds:
                for ml in MoveLineIds:
                    AnalyticVals = []
                    Balance = Balance + (ml.debit - ml.credit)
                    if self.account_zero_closing_balance and Balance == 0.0:
                        Vals = {'date': str(ml.date.strftime('%d/%b/%Y')) or '',
                                'move': ml.move_id and ml.move_id.name or '',
                                'name' : ml.name or '',
                                'debit': ml.debit or 0.0,
                                'credit': ml.credit or 0.0,
                                'balance': Balance or 0.0,
                                }
                        mainDict[Account.name or '-'].append(Vals)
                        if self.dimension_wise_project:
                            for analytic in DynamicList:
                                if ml.analytic_account_id.name == analytic:
                                    AnalyticVals.append({analytic:ml.debit-ml.credit or 0.0})
                                else:
                                    AnalyticVals.append({analytic:0.0})
                        Vals.update({'analytic_vals':AnalyticVals})
                    else:
                        Vals = {'date': str(ml.date.strftime('%d/%b/%Y')) or '',
                                'move': ml.move_id and ml.move_id.name or '',
                                'name' : ml.name or '',
                                'debit': ml.debit or 0.0,
                                'credit': ml.credit or 0.0,
                                'balance': Balance or 0.0,
                                }
                        mainDict[Account.name or '-'].append(Vals)
                        if self.dimension_wise_project:
                            for analytic in DynamicList:
                                if ml.analytic_account_id.name == analytic:
                                    bala = ml.debit-ml.credit
                                    AnalyticVals.append({analytic:ml.debit-ml.credit or 0.0})
                                else:
                                    AnalyticVals.append({analytic:0.0})
                        Vals.update({'analytic_vals':AnalyticVals})
            if self.account_without_transaction and not MoveLines:
                AnalyticVals = []
                Vals = {'date':'',
                        'move':'',
                        'name':'',
                        'debit':0.0,
                        'credit':0.0,
                        'balance':0.0,
                      }
                mainDict[Account.name or '-'].append(Vals)
                if self.dimension_wise_project:
                    for analytic in DynamicList:
                        AnalyticVals.append({analytic:0.0})
                Vals.update({'analytic_vals':AnalyticVals})

        import base64
        filename = 'General Ledger.xls'
        start = dateFrom.strftime('%d %b, %Y')
        end = dateTo.strftime('%d %b, %Y')
        form_name = 'General Ledger Between ' + str(start) + ' to ' + str(end)
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

        worksheet = workbook.add_sheet("General Ledger", cell_overwrite_ok=True)
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
        left thin, right thin, top thin, bottom thin;", "#,##0.00")
        finaltotalheaderbold = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour black; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;", "#,##0.00")
        zero_col = worksheet.col(0)
        zero_col.width = 236 * 22
        first_col = worksheet.col(1)
        first_col.width = 236 * 40
        second_col = worksheet.col(2)
        second_col.width = 236 * 40
        third_col = worksheet.col(3)
        third_col.width = 236 * 25
        fourth_col = worksheet.col(4)
        fourth_col.width = 236 * 20
        fifth_col = worksheet.col(5)
        fifth_col.width = 236 * 20
        sixth_col = worksheet.col(6)
        sixth_col.width = 236 * 20
        seventh_col = worksheet.col(7)
        seventh_col.width = 236 * 20
        #HEADER
        worksheet.row(4).height_mismatch = True
        worksheet.row(4).height = 360
        
        worksheet.write_merge(0, 1, 2, 5, self.env.company.name,styleheader)
        worksheet.write_merge(2, 2, 2, 5, 'General Ledger',general)
        headerstring = 'From :' + str(self.date_from.strftime('%d %b %Y') or '') + ' To :' + str(self.date_to.strftime('%d %b %Y') or '')
        worksheet.write_merge(3, 3, 2, 5, headerstring,dateheader)
        
        #SUB-HEADER
        row = 4
        ColIndexes = {'Date':0,
                      'Journal Entry':1,
                      'Reference':2,
                      'Label':3,
                      'Project':4,
                      'Debit':5,
                      'Credit':6,
                      'Balance':7}
        worksheet.write(row, 0, 'Voucher Date', stylecolaccount)
        worksheet.write(row, 1, 'Voucher Number', stylecolaccount)
        worksheet.write(row, 2, 'Remarks', stylecolaccount)
        worksheet.write(row, 3, 'Debit', stylecolaccount)
        worksheet.write(row, 4, 'Credit', stylecolaccount)
        worksheet.write(row, 5, 'Balance', stylecolaccount)
        analytic_account_ids = self.env['account.analytic.account'].search([])
        calc = 5
        if self.dimension_wise_project:
            col = 6
            for analytic in AnalyticAccountIds:
                dictval = {analytic.name:col}
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, analytic.name, analytic_st_col)
                col+=1
                calc+=1
        row = 5
        totaldebit = 0.0
        totalcredit = 0.0
        totalanalytic = 0.0
        FinalDict = {}
        for ac in mainDict:
            newDict = {}
            worksheet.row(row).height_mismatch = True
            worksheet.row(row).height = 310
            initial_debit = 0.0
            initial_credit = 0.0
            initial_balance = 0.0
            worksheet.write_merge(row, row, 0, calc, str(ac),stylecolumnheader)
            if not self.dimension_wise_project:
                row+=1
                worksheet.write(row, 0, 'Initial Balance', linedata)
                worksheet.write(row, 1,'', rightfont)
                worksheet.write(row, 2,'', rightfont)
                initial_debit,initial_credit,initial_balance = self.get_initials(ac)
                
                if initial_debit == 0.0:
                    worksheet.write(row, 3, 0.0,rightfont)
                else:
                    worksheet.write(row, 3, initial_debit,floatstyle)
                if initial_credit == 0.0:
                    worksheet.write(row, 4, 0.0,rightfont)
                else:
                    worksheet.write(row, 4, initial_credit,floatstyle)
                if initial_balance == 0.0:
                    worksheet.write(row, 5, 0.0,rightfont)
                else:
                    worksheet.write(row, 5, initial_balance,floatstyle)
            subtotal_debit = 0.0
            subtotal_credit = 0.0
            if not self.dimension_wise_project:
                subtotal_debit += initial_debit
                subtotal_credit += initial_credit
            Subvals = {}
            for line in mainDict.get(ac):
                row+=1
                totaldebit+= line.get('debit')
                totalcredit+= line.get('credit')
                worksheet.write(row, 0, line.get('date',False),linedata)
                worksheet.write(row, 1, line.get('move',''),linedata)
                worksheet.write(row, 2, line.get('name',''),linedata)
                if line.get('debit',0.0) == 0.0:
                    worksheet.write(row, 3, 0.0,rightfont)
                else:
                    worksheet.write(row, 3, line.get('debit',0.0),floatstyle)
                    
                if line.get('credit',0.0) == 0.0:
                    worksheet.write(row, 4, 0.0,rightfont)
                else:
                    worksheet.write(row, 4, line.get('credit',0.0),floatstyle)
                    
                if line.get('balance',0.0) == 0.0:
                    worksheet.write(row, 5, 0.0 ,rightfont)
                else:
                    worksheet.write(row, 5, line.get('balance',0.0),floatstyle)
                    
                subtotal_debit += line.get('debit')
                subtotal_credit += line.get('credit')
                
                if Projectwise:
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
            row+=1
            worksheet.write_merge(row, row, 0, 2, 'Subtotal '+ '(' + str(ac) +')',maintotal)
            if subtotal_debit == 0.0:
                worksheet.write(row, 3, 0.0,finaltotalheader)
            else:
                worksheet.write(row, 3, subtotal_debit,finaltotalheaderbold)

            if subtotal_credit == 0.0:
                worksheet.write(row, 4, 0.0,finaltotalheader)
            else:
                worksheet.write(row, 4, subtotal_credit,finaltotalheaderbold)
            
            balancesub = subtotal_debit - subtotal_credit
            if balancesub == 0.0:
                worksheet.write(row, 5, 0.0,finaltotalheader)
            else:
                worksheet.write(row, 5, balancesub,finaltotalheaderbold)
            
            for newkey,newval in newDict.items():
                if newval == 0.0:
                    worksheet.write(row, newkey, 0.0,finaltotalheader)                        
                else:
                    worksheet.write(row, newkey, newval,finaltotalheaderbold)

                if newkey in FinalDict:
                    FinalDict[newkey] += newval
                else:
                    FinalDict.update({newkey:newval})

            row += 1
            worksheet.write_merge(row, row, 0, 2, 'REPORT TOTAL', style = maintotal)
            if totaldebit == 0.0:
                worksheet.write(row, 3, 0.0,finaltotalheader)
            else:
                worksheet.write(row, 3, totaldebit,finaltotalheaderbold)

            if totalcredit == 0.0:
                worksheet.write(row, 4, 0.0,finaltotalheader)
            else:
                worksheet.write(row, 4, totalcredit,finaltotalheaderbold)
            
            totalbala = totaldebit - totalcredit
            if totalbala == 0.0:
                worksheet.write(row, 5, 0.0,finaltotalheader)
            else:
                worksheet.write(row, 5, totalbala,finaltotalheaderbold)

            for finalkey,finalval in FinalDict.items():
                if finalval == 0.0:
                    worksheet.write(row, finalkey, 0.0,style=finaltotalheader)
                else:
                    worksheet.write(row, finalkey, finalval,style=finaltotalheaderbold)
                    
        row+=2
        buffer = io.BytesIO()
        workbook.save(buffer)
        export_id = self.env['general.ledger.excel'].create(
                        {'excel_file': base64.encodestring(buffer.getvalue()), 'file_name': filename})
        buffer.close()
    
        return {
            'name': form_name,
            'view_mode': 'form',
            'res_id': export_id.id,
            'res_model': 'general.ledger.excel',
            'view_mode': 'form',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

class general_ledger_export_excel(models.TransientModel):
    _name= "general.ledger.excel"
    _description = "General Ledger Excel Report"

    excel_file = fields.Binary('Report for General Ledger')
    file_name = fields.Char('File', size=64)
